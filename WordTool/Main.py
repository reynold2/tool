from WordTool.ui.MainWindow import *
import sys


class CommonRes:
    def __init__(self):
        pass

    @staticmethod
    def readQss(style):
        with open(style, 'r') as f:
            return f.read()



if __name__ == '__main__':
    app = QApplication(sys.argv)

    win = MainUi()
    styleFile = './resources/style.qss'
    qssStyle = CommonRes.readQss(styleFile)
    win.setStyleSheet(qssStyle)
    win.show()
    sys.exit(app.exec_())

    print("11")
