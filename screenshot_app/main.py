from app.main_window import ListAppWindow
from PySide6 import QtWidgets

def main():
    app = QtWidgets.QApplication([])
    w = ListAppWindow()
    w.show()
    app.exec()

if __name__ == "__main__":
    main()
