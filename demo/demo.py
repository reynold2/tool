import sys
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

import xpinyin
class TableView(QTableView):
    def __init__(self):
        super(TableView, self).__init__()

    def paintEvent(self, even):
        QTableView.paintEvent(self, even)


        painter = QPainter(self)
        painter.save()

        pen = painter.pen()
        pen.setWidth(1)
        pen.setColor(QColor(166, 66, 250))
        painter.setPen(pen)

        painter.drawRect(200,200, 300, 558)
        painter.restore()

class PreviewWidget(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.resize(1047, 611)
        self.setWindowTitle('Table')

        h_layout = QHBoxLayout()
        self.table_view = TableView()

        h_layout.addWidget(self.table_view)
        self.setLayout(h_layout)





    def paintEvent(self, even):
        pass
        # painter = QPainter(self)
        # painter.setPen(QColor(166, 66, 250))
        # painter.begin(self)
        # painter.drawLine(100, 100, 200, 200)
        # painter.end()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PreviewWidget()
    ex.show()
    sys.exit(app.exec_())
