from __future__ import annotations
import os
import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import (
    LEFT, RIGHT, X, Y, BOTH, END, W, CENTER,
    PRIMARY, WARNING, SECONDARY, SUCCESS, INFO, VERTICAL, ROUND, STRIPED,
)
from ttkbootstrap.toast import ToastNotification

from pub_settlement import BookstoreSettlementProcessor
from utils import timestamp, detect_bookstore

# Optional Drag & Drop (tkinterdnd2)
_DND_AVAILABLE = True
try:
    from tkinterdnd2 import DND_FILES  # type: ignore
except Exception:
    _DND_AVAILABLE = False


class SettlementGUI_ttk:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("출판사/서점 정산 통합 - Pub Settlement")
        # Honor theme from the created root (main.py). Avoid re-setting here.
        self.root.geometry("880x640")

        self.processor = BookstoreSettlementProcessor()
        self.files: list[str] = []
        self.q: queue.Queue = queue.Queue()

        # Keep references for enable/disable during processing
        self.btn_add: tb.Button | None = None
        self.btn_remove: tb.Button | None = None
        self.btn_clear: tb.Button | None = None
        self.btn_process: tb.Button | None = None
        self.menu: tk.Menu | None = None
        self.drop_frame: tk.Widget | None = None

        self._build_ui()

    # ---------------- UI ---------------- #
    def _build_ui(self):
        # Top bar
        top = tb.Frame(self.root, padding=10)
        top.pack(fill=X)

        tb.Label(
            top,
            text="📚 Pub Settlement (YES24/교보/알라딘)",
            font=("Segoe UI", 14, "bold"),
        ).pack(side=LEFT)

        btns = tb.Frame(top)
        btns.pack(side=RIGHT)

        self.btn_add = tb.Button(btns, text="파일 추가", bootstyle=PRIMARY, command=self._add_files)
        self.btn_add.pack(side=LEFT, padx=4)
        self.btn_remove = tb.Button(btns, text="선택 제거", bootstyle=WARNING, command=self._remove_selected)
        self.btn_remove.pack(side=LEFT, padx=4)
        self.btn_clear = tb.Button(btns, text="전체 초기화", bootstyle=SECONDARY, command=self._clear_all)
        self.btn_clear.pack(side=LEFT, padx=4)
        self.btn_process = tb.Button(btns, text="정산 통합 처리", bootstyle=SUCCESS, command=self._process_async)
        self.btn_process.pack(side=LEFT, padx=4)

        # Drop zone / 안내
        dz = tb.Labelframe(self.root, text="드래그앤드롭", padding=10, bootstyle=INFO)
        dz.pack(fill=X, padx=10, pady=6)
        self.drop_frame = dz

        dz_label_text = (
            "여기에 파일을 드래그앤드롭 하세요" if _DND_AVAILABLE else "ⓘ tkinterdnd2 미설치 — [파일 추가] 버튼을 사용하세요"
        )
        self.drop_label = tb.Label(dz, text=dz_label_text, anchor=CENTER)
        self.drop_label.pack(fill=X)

        # Bind DnD on the drop zone, not the root window
        if _DND_AVAILABLE:
            try:
                if hasattr(dz, "drop_target_register"):
                    dz.drop_target_register(DND_FILES)
                    dz.dnd_bind("<<Drop>>", self._on_drop)
                elif hasattr(self.drop_label, "drop_target_register"):
                    self.drop_label.drop_target_register(DND_FILES)
                    self.drop_label.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        # Treeview (파일 목록)
        treebox = tb.Frame(self.root, padding=(10, 0, 10, 0))
        treebox.pack(fill=BOTH, expand=True)

        cols = ("name", "path", "store")
        self.tree = tb.Treeview(treebox, columns=cols, show="headings", bootstyle=INFO)
        self.tree.heading("name", text="파일명")
        self.tree.heading("path", text="경로")
        self.tree.heading("store", text="서점 감지")
        self.tree.column("name", width=220, anchor=W)
        self.tree.column("path", width=520, anchor=W)
        self.tree.column("store", width=100, anchor=CENTER)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)

        yscroll = tb.Scrollbar(treebox, orient=VERTICAL, command=self.tree.yview, bootstyle=ROUND)
        self.tree.configure(yscroll=yscroll.set)
        yscroll.pack(side=RIGHT, fill=Y)

        # 컨텍스트 메뉴
        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="선택 제거", command=self._remove_selected)
        self.menu.add_command(label="경로 복사", command=self._copy_path)
        self.tree.bind("<Button-3>", self._popup_menu)

        # Status & progress
        statusbar = tb.Frame(self.root, padding=10)
        statusbar.pack(fill=X)

        self.status_var = tk.StringVar(value="파일을 추가해주세요")
        tb.Label(statusbar, textvariable=self.status_var).pack(side=LEFT)

        self.prog = tb.Progressbar(statusbar, mode="determinate", bootstyle=STRIPED)
        self.prog.pack(side=RIGHT, fill=X, expand=True, padx=10)

    # ---------------- Handlers ---------------- #
    def _popup_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            assert self.menu is not None
            self.menu.tk_popup(event.x_root, event.y_root)

    def _copy_path(self):
        sel = self.tree.selection()
        if not sel:
            return
        path = self.tree.set(sel[0], "path")
        self.root.clipboard_clear()
        self.root.clipboard_append(path)
        try:
            ToastNotification(title="복사됨", message="경로가 클립보드에 복사되었습니다", duration=2000).show_toast()
        except Exception:
            pass

    def _on_drop(self, event):
        try:
            files = self.root.tk.splitlist(event.data)
        except Exception:
            files = [event.data]
        self._add_paths(files)

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="정산 파일 선택",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if paths:
            self._add_paths(paths)

    def _add_paths(self, paths):
        added = 0
        for p in paths:
            p = str(p).strip("{}")
            if os.path.isfile(p) and p not in self.files:
                self.files.append(p)
                store = detect_bookstore(p) or "-"
                self.tree.insert("", END, values=(os.path.basename(p), p, store))
                added += 1
        if added:
            self.status_var.set(f"{len(self.files)}개 파일 준비됨")

    def _remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        for iid in sel:
            path = self.tree.set(iid, "path")
            if path in self.files:
                self.files.remove(path)
            self.tree.delete(iid)
        self.status_var.set(f"{len(self.files)}개 파일 준비됨")

    def _clear_all(self):
        self.files.clear()
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.status_var.set("목록이 초기화되었습니다.")

    def _set_busy(self, busy: bool):
        state = tk.DISABLED if busy else tk.NORMAL
        for btn in (self.btn_add, self.btn_remove, self.btn_clear, self.btn_process):
            if btn is not None:
                btn.configure(state=state)
        if self.menu is not None:
            try:
                # Disable/enable menu entries by index
                end_index = self.menu.index('end') or -1
                for i in range(end_index + 1):
                    self.menu.entryconfigure(i, state=state)
            except Exception:
                pass
        self.root.configure(cursor="watch" if busy else "")

    # ---------------- Processing (threaded) ---------------- #
    def _process_async(self):
        if not self.files:
            messagebox.showwarning("경고", "처리할 파일이 없습니다.")
            return
        self._set_busy(True)
        self.status_var.set("처리 중…")
        self.prog.configure(maximum=len(self.files), value=0)
        t = threading.Thread(target=self._process_worker, daemon=True)
        t.start()
        self.root.after(100, self._poll_queue)

    def _process_worker(self):
        # 각 처리마다 프로세서를 초기화(중복 처리 방지)
        self.processor = BookstoreSettlementProcessor()
        results: list[str] = []
        errors: list[str] = []
        done = 0
        for fp in list(self.files):
            cnt, err = self.processor.process_file(fp)
            name = os.path.basename(fp)
            if err:
                errors.append(f"· {name}: {err}")
            else:
                results.append(f"· {name}: {cnt}건 처리")
            done += 1
            self.q.put(("progress", done))

        # 메인 스레드에서 후처리
        self.q.put(("finished", (results, errors)))

    def _poll_queue(self):
        drained = False
        while not self.q.empty():
            kind, payload = self.q.get_nowait()
            if kind == "progress":
                self.prog.configure(value=payload)
            elif kind == "finished":
                drained = True
                results, errors = payload
                self._on_finished(results, errors)
        if not drained:
            self.root.after(100, self._poll_queue)

    def _on_finished(self, results: list[str], errors: list[str]):
        try:
            df = self.processor.get_unified_dataframe()
            total = len(df)
            if total == 0:
                msg = "\n".join(errors) if errors else "처리 가능한 데이터가 없습니다."
                messagebox.showerror("오류", msg)
                self.status_var.set("처리 실패")
                return

            default_name = f"Settlement_{timestamp()}.xlsx"
            save_path = filedialog.asksaveasfilename(
                title="저장 위치 선택",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=default_name,
            )
            if not save_path:
                self.status_var.set("저장 취소됨")
                return

            saved = False
            if save_path.lower().endswith(".csv"):
                saved = self.processor.save_to_csv(save_path)
            else:
                saved = self.processor.save_to_excel(save_path)

            # 요약 만들기 (가능하면 서점/수량/정산금액 기준)
            summary = "\n\n정산 통합 결과:\n"
            try:
                store_col, qty_col, amt_col = "서점명", "입고수량", "정산액"
                grp = df.groupby(store_col, dropna=False).agg({qty_col: "sum", amt_col: "sum"}).reset_index()
                lines = [f"- {r[store_col]}: 수량 {int(r[qty_col])} / 금액 {int(r[amt_col]):,}원"
                        for _, r in grp.iterrows()]
                summary += f"총 {len(df)}건 처리\n" + "\n".join(lines)
            except Exception as e:
                summary += f"총 {len(df)}건 처리 (요약 계산 실패: {e})"

            msg = ""
            if results:
                msg += "\n".join(results)
            if errors:
                msg += ("\n\n" if msg else "") + "\n".join(errors)
            msg += summary

            if saved:
                msg += f"\n\n저장 완료:\n{save_path}"
                messagebox.showinfo("처리 완료", msg)
                self.status_var.set(f"처리 완료: {total}건")
            else:
                messagebox.showerror("오류", "파일 저장에 실패했습니다.")
                self.status_var.set("저장 실패")
        except PermissionError:
            messagebox.showerror("오류", "파일이 열려 있어 저장할 수 없습니다. 닫은 뒤 다시 시도하세요.")
            self.status_var.set("저장 실패")
        finally:
            self._set_busy(False)

