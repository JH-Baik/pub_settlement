# main.py
from tkinterdnd2 import TkinterDnD
import ttkbootstrap as tb
from gui_ttk import SettlementGUI_ttk

def main():
    root = TkinterDnD.Tk()
    tb.Style(theme="darkly")  # 테마 적용
    SettlementGUI_ttk(root)
    root.mainloop()

if __name__ == "__main__":
    main()
