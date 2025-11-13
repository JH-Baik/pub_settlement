# main.py
from __future__ import annotations
import ttkbootstrap as tb
from gui_ttk import SettlementGUI_ttk

def main():
    root = tb.Window(themename="darkly")
    app = SettlementGUI_ttk(root)
    root.mainloop()

if __name__ == "__main__":
    main()
