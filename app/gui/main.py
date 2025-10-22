# app/gui/main.py
from app.gui.app import WBSApp  # ← IMPORT ABSOLUTO

def main():
    app = WBSApp()
    app.mainloop()

if __name__ == "__main__":
    main()
