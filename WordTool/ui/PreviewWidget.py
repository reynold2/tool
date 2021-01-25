import sys
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *


class StandardModel(QStandardItemModel):
    def __init__(self, indexs=[]):
        super(QStandardItemModel, self).__init__()
        self.index_s = indexs

    def data(self, index, role=None):
        if role == Qt.TextAlignmentRole:
            return Qt.AlignCenter

        if role == Qt.BackgroundRole:
            index_m = [index.row(), index.column()]
            if index_m in self.index_s:
                return QColor(255, 0, 0)

        return QStandardItemModel.data(self, index, role)


class PreviewWidget(QWidget):

    def __init__(self):
        super(PreviewWidget, self).__init__()
        self.initUI()
        self.model = StandardModel()

    def initUI(self):
        self.resize(1047, 611)
        self.setWindowTitle('Table')
        self.h_layout = QHBoxLayout()
        self.table_view = QTableView()
        self.table_view.horizontalHeader().setStretchLastSection(True)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.h_layout.addWidget(self.table_view)
        self.setLayout(self.h_layout)

    def set_activate_table_model(self, activate_table, index=None, labels=None):
        if activate_table is None:
            return
        if labels is not None:
            self.model.setHorizontalHeaderLabels(labels)

        for x, row in enumerate(activate_table.rows):  # 遍历表格的所有行
            for y, cell in enumerate(row.cells):
                self.model.setItem(x, y, QStandardItem(cell.text.replace(" ", "").replace("\n", "")))
        if index is not None:
            # self.table_view.setItemDelegate(TLabelDelegate(index))
            self.model.index_s = index
        else:
            self.span_cells(activate_table)
        self.table_view.setModel(self.model)

    def set_table_data_model(self, labels, excle_data):

        self.model.setHorizontalHeaderLabels(labels)
        for cell in excle_data:  # 遍历表格的所有行
            self.model.setItem(cell[0], cell[1], QStandardItem(cell[2].replace(" ", "").replace("\n", "")))
        self.table_view.setModel(self.model)

    def get_current_model(self):

        return self.model

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Message', 'You sure to quit?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def span_cells(self, activate_table):
        for x, row in enumerate(activate_table.rows):  # 遍历表格的所有行
            rowcelltext = ""
            rowcellcount = 1
            start_pos = []
            for y, cell in enumerate(row.cells):
                if rowcelltext == cell.text:
                    if (len(start_pos) == 0):
                        start_pos.append(x)
                        start_pos.append(y - 1)
                    rowcellcount = rowcellcount + 1
                else:
                    start_pos.clear()
                    rowcellcount = 1
                rowcelltext = cell.text
                if (len(start_pos) == 2):
                    self.table_view.setSpan(start_pos[0], start_pos[1], 1, rowcellcount)
        for x, colunm in enumerate(activate_table.columns):  # 遍历表格的所有列
            rowcelltext = ""
            rowcellcount = 1
            start_pos = []
            for y, cell in enumerate(colunm.cells):
                if rowcelltext == cell.text:
                    if (len(start_pos) == 0):
                        start_pos.append(x)
                        start_pos.append(y - 1)
                    rowcellcount = rowcellcount + 1
                else:
                    start_pos.clear()
                    rowcellcount = 1
                rowcelltext = cell.text
                if (len(start_pos) == 2):
                    self.table_view.setSpan(start_pos[1], start_pos[0], rowcellcount, 1)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PreviewWidget()
    sys.exit(app.exec_())
