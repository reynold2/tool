# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ui_mianwindow.ui'
#
# Created by: PyQt5 UI code generator 5.13.0
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1047, 611)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setTitle("")
        self.groupBox.setObjectName("groupBox")
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout(self.groupBox)
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_5.addItem(spacerItem)
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.sourcepath = QtWidgets.QLineEdit(self.groupBox)
        self.sourcepath.setReadOnly(True)
        self.sourcepath.setObjectName("sourcepath")
        self.horizontalLayout_3.addWidget(self.sourcepath)
        self.Butpath = QtWidgets.QPushButton(self.groupBox)
        self.Butpath.setObjectName("Butpath")
        self.horizontalLayout_3.addWidget(self.Butpath)
        self.verticalLayout_3.addLayout(self.horizontalLayout_3)
        self.horizontalLayout_5.addLayout(self.verticalLayout_3)
        self.issavesourcepath = QtWidgets.QCheckBox(self.groupBox)
        self.issavesourcepath.setChecked(True)
        self.issavesourcepath.setObjectName("issavesourcepath")
        self.horizontalLayout_5.addWidget(self.issavesourcepath)
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem1)
        self.butrun = QtWidgets.QPushButton(self.groupBox)
        self.butrun.setObjectName("butrun")
        self.horizontalLayout_2.addWidget(self.butrun)
        self.verticalLayout_4.addLayout(self.horizontalLayout_2)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem2)
        self.butstop = QtWidgets.QPushButton(self.groupBox)
        self.butstop.setObjectName("butstop")
        self.horizontalLayout_4.addWidget(self.butstop)
        self.verticalLayout_4.addLayout(self.horizontalLayout_4)
        self.horizontalLayout_5.addLayout(self.verticalLayout_4)
        self.verticalLayout_5.addWidget(self.groupBox)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName("tabWidget")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.verticalLayout_12 = QtWidgets.QVBoxLayout(self.tab_3)
        self.verticalLayout_12.setObjectName("verticalLayout_12")
        self.verticalLayout_11 = QtWidgets.QVBoxLayout()
        self.verticalLayout_11.setObjectName("verticalLayout_11")
        self.horizontalLayout_21 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_21.setObjectName("horizontalLayout_21")
        self.groupBox_3 = QtWidgets.QGroupBox(self.tab_3)
        self.groupBox_3.setObjectName("groupBox_3")
        self.verticalLayout_6 = QtWidgets.QVBoxLayout(self.groupBox_3)
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.verticalLayout_8 = QtWidgets.QVBoxLayout()
        self.verticalLayout_8.setObjectName("verticalLayout_8")
        self.horizontalLayout_11 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_11.setObjectName("horizontalLayout_11")
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.label = QtWidgets.QLabel(self.groupBox_3)
        self.label.setObjectName("label")
        self.horizontalLayout_6.addWidget(self.label)
        self.tabstart = QtWidgets.QSpinBox(self.groupBox_3)
        self.tabstart.setMinimum(1)
        self.tabstart.setProperty("value", 1)
        self.tabstart.setObjectName("tabstart")
        self.horizontalLayout_6.addWidget(self.tabstart)
        self.horizontalLayout_11.addLayout(self.horizontalLayout_6)
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.label_2 = QtWidgets.QLabel(self.groupBox_3)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_7.addWidget(self.label_2)
        self.tabend = QtWidgets.QSpinBox(self.groupBox_3)
        self.tabend.setMinimum(1)
        self.tabend.setMaximum(800)
        self.tabend.setProperty("value", 500)
        self.tabend.setObjectName("tabend")
        self.horizontalLayout_7.addWidget(self.tabend)
        self.horizontalLayout_11.addLayout(self.horizontalLayout_7)
        self.verticalLayout_8.addLayout(self.horizontalLayout_11)
        self.horizontalLayout_12 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_12.setObjectName("horizontalLayout_12")
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.label_4 = QtWidgets.QLabel(self.groupBox_3)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_8.addWidget(self.label_4)
        self.tabrow = QtWidgets.QSpinBox(self.groupBox_3)
        self.tabrow.setMinimum(1)
        self.tabrow.setObjectName("tabrow")
        self.horizontalLayout_8.addWidget(self.tabrow)
        self.horizontalLayout_12.addLayout(self.horizontalLayout_8)
        spacerItem3 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_12.addItem(spacerItem3)
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        self.label_5 = QtWidgets.QLabel(self.groupBox_3)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_9.addWidget(self.label_5)
        self.tabcolumn = QtWidgets.QSpinBox(self.groupBox_3)
        self.tabcolumn.setMinimum(1)
        self.tabcolumn.setObjectName("tabcolumn")
        self.horizontalLayout_9.addWidget(self.tabcolumn)
        self.horizontalLayout_12.addLayout(self.horizontalLayout_9)
        self.verticalLayout_8.addLayout(self.horizontalLayout_12)
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.label_3 = QtWidgets.QLabel(self.groupBox_3)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_10.addWidget(self.label_3)
        self.tabval = QtWidgets.QLineEdit(self.groupBox_3)
        self.tabval.setObjectName("tabval")
        self.horizontalLayout_10.addWidget(self.tabval)
        self.verticalLayout_8.addLayout(self.horizontalLayout_10)
        self.horizontalLayout_15 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_15.setObjectName("horizontalLayout_15")
        self.bsave_conf = QtWidgets.QPushButton(self.groupBox_3)
        self.bsave_conf.setObjectName("bsave_conf")
        self.horizontalLayout_15.addWidget(self.bsave_conf)
        self.advanced_conf = QtWidgets.QPushButton(self.groupBox_3)
        self.advanced_conf.setObjectName("advanced_conf")
        self.horizontalLayout_15.addWidget(self.advanced_conf)
        self.verticalLayout_8.addLayout(self.horizontalLayout_15)
        self.verticalLayout_6.addLayout(self.verticalLayout_8)
        self.line = QtWidgets.QFrame(self.groupBox_3)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.verticalLayout_6.addWidget(self.line)
        spacerItem4 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_6.addItem(spacerItem4)
        self.horizontalLayout_38 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_38.setObjectName("horizontalLayout_38")
        self.label_27 = QtWidgets.QLabel(self.groupBox_3)
        self.label_27.setObjectName("label_27")
        self.horizontalLayout_38.addWidget(self.label_27)
        self.comboBox_macro = QtWidgets.QComboBox(self.groupBox_3)
        self.comboBox_macro.setObjectName("comboBox_macro")
        self.comboBox_macro.addItem("")
        self.comboBox_macro.addItem("")
        self.comboBox_macro.addItem("")
        self.horizontalLayout_38.addWidget(self.comboBox_macro)
        self.pushButton_macro = QtWidgets.QPushButton(self.groupBox_3)
        self.pushButton_macro.setObjectName("pushButton_macro")
        self.horizontalLayout_38.addWidget(self.pushButton_macro)
        self.verticalLayout_6.addLayout(self.horizontalLayout_38)
        spacerItem5 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_6.addItem(spacerItem5)
        self.horizontalLayout_21.addWidget(self.groupBox_3)
        self.tabWidget_2 = QtWidgets.QTabWidget(self.tab_3)
        self.tabWidget_2.setObjectName("tabWidget_2")
        self.tab_5 = QtWidgets.QWidget()
        self.tab_5.setObjectName("tab_5")
        self.verticalLayout_10 = QtWidgets.QVBoxLayout(self.tab_5)
        self.verticalLayout_10.setObjectName("verticalLayout_10")
        self.verticalLayout_9 = QtWidgets.QVBoxLayout()
        self.verticalLayout_9.setObjectName("verticalLayout_9")
        self.groupBox_2 = QtWidgets.QGroupBox(self.tab_5)
        self.groupBox_2.setObjectName("groupBox_2")
        self.layoutWidget = QtWidgets.QWidget(self.groupBox_2)
        self.layoutWidget.setGeometry(QtCore.QRect(20, 30, 281, 82))
        self.layoutWidget.setObjectName("layoutWidget")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.layoutWidget)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.horizontalLayout_37 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_37.setObjectName("horizontalLayout_37")
        self.horizontalLayout_13 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_13.setObjectName("horizontalLayout_13")
        self.label_6 = QtWidgets.QLabel(self.layoutWidget)
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_13.addWidget(self.label_6)
        self.lineEdit_replace_s = QtWidgets.QLineEdit(self.layoutWidget)
        self.lineEdit_replace_s.setObjectName("lineEdit_replace_s")
        self.horizontalLayout_13.addWidget(self.lineEdit_replace_s)
        self.horizontalLayout_37.addLayout(self.horizontalLayout_13)
        self.horizontalLayout_14 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_14.setObjectName("horizontalLayout_14")
        self.label_7 = QtWidgets.QLabel(self.layoutWidget)
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_14.addWidget(self.label_7)
        self.lineEdit_replace_t = QtWidgets.QLineEdit(self.layoutWidget)
        self.lineEdit_replace_t.setObjectName("lineEdit_replace_t")
        self.horizontalLayout_14.addWidget(self.lineEdit_replace_t)
        self.horizontalLayout_37.addLayout(self.horizontalLayout_14)
        self.verticalLayout_2.addLayout(self.horizontalLayout_37)
        self.horizontalLayout_17 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_17.setObjectName("horizontalLayout_17")
        spacerItem6 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_17.addItem(spacerItem6)
        self.checkBox_blank = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBox_blank.setChecked(False)
        self.checkBox_blank.setObjectName("checkBox_blank")
        self.horizontalLayout_17.addWidget(self.checkBox_blank)
        self.verticalLayout_2.addLayout(self.horizontalLayout_17)
        self.horizontalLayout_36 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_36.setObjectName("horizontalLayout_36")
        self.label_25 = QtWidgets.QLabel(self.layoutWidget)
        self.label_25.setObjectName("label_25")
        self.horizontalLayout_36.addWidget(self.label_25)
        self.spinBox_del1 = QtWidgets.QSpinBox(self.layoutWidget)
        self.spinBox_del1.setObjectName("spinBox_del1")
        self.horizontalLayout_36.addWidget(self.spinBox_del1)
        self.label_26 = QtWidgets.QLabel(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_26.sizePolicy().hasHeightForWidth())
        self.label_26.setSizePolicy(sizePolicy)
        self.label_26.setMinimumSize(QtCore.QSize(30, 0))
        self.label_26.setAlignment(QtCore.Qt.AlignCenter)
        self.label_26.setObjectName("label_26")
        self.horizontalLayout_36.addWidget(self.label_26)
        self.spinBox_del2 = QtWidgets.QSpinBox(self.layoutWidget)
        self.spinBox_del2.setObjectName("spinBox_del2")
        self.horizontalLayout_36.addWidget(self.spinBox_del2)
        self.verticalLayout_2.addLayout(self.horizontalLayout_36)
        self.verticalLayout_9.addWidget(self.groupBox_2)
        self.groupBox_4 = QtWidgets.QGroupBox(self.tab_5)
        self.groupBox_4.setObjectName("groupBox_4")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.groupBox_4)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout_23 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_23.setObjectName("horizontalLayout_23")
        self.label_13 = QtWidgets.QLabel(self.groupBox_4)
        self.label_13.setObjectName("label_13")
        self.horizontalLayout_23.addWidget(self.label_13)
        spacerItem7 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_23.addItem(spacerItem7)
        self.comboBox_shop = QtWidgets.QComboBox(self.groupBox_4)
        self.comboBox_shop.setObjectName("comboBox_shop")
        self.comboBox_shop.addItem("")
        self.comboBox_shop.addItem("")
        self.comboBox_shop.addItem("")
        self.comboBox_shop.addItem("")
        self.comboBox_shop.addItem("")
        self.horizontalLayout_23.addWidget(self.comboBox_shop)
        self.verticalLayout.addLayout(self.horizontalLayout_23)
        self.horizontalLayout_22 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_22.setObjectName("horizontalLayout_22")
        self.label_12 = QtWidgets.QLabel(self.groupBox_4)
        self.label_12.setObjectName("label_12")
        self.horizontalLayout_22.addWidget(self.label_12)
        spacerItem8 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_22.addItem(spacerItem8)
        self.comboBox_font = QtWidgets.QComboBox(self.groupBox_4)
        self.comboBox_font.setObjectName("comboBox_font")
        self.comboBox_font.addItem("")
        self.comboBox_font.addItem("")
        self.comboBox_font.addItem("")
        self.horizontalLayout_22.addWidget(self.comboBox_font)
        self.verticalLayout.addLayout(self.horizontalLayout_22)
        self.horizontalLayout_16 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_16.setObjectName("horizontalLayout_16")
        self.label_8 = QtWidgets.QLabel(self.groupBox_4)
        self.label_8.setObjectName("label_8")
        self.horizontalLayout_16.addWidget(self.label_8)
        spacerItem9 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_16.addItem(spacerItem9)
        self.comboBox_font_alignment = QtWidgets.QComboBox(self.groupBox_4)
        self.comboBox_font_alignment.setObjectName("comboBox_font_alignment")
        self.comboBox_font_alignment.addItem("")
        self.comboBox_font_alignment.addItem("")
        self.comboBox_font_alignment.addItem("")
        self.comboBox_font_alignment.addItem("")
        self.comboBox_font_alignment.addItem("")
        self.horizontalLayout_16.addWidget(self.comboBox_font_alignment)
        self.verticalLayout.addLayout(self.horizontalLayout_16)
        self.horizontalLayout_35 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_35.setObjectName("horizontalLayout_35")
        self.label_9 = QtWidgets.QLabel(self.groupBox_4)
        self.label_9.setObjectName("label_9")
        self.horizontalLayout_35.addWidget(self.label_9)
        spacerItem10 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_35.addItem(spacerItem10)
        self.comboBox_font_bold = QtWidgets.QComboBox(self.groupBox_4)
        self.comboBox_font_bold.setObjectName("comboBox_font_bold")
        self.comboBox_font_bold.addItem("")
        self.comboBox_font_bold.addItem("")
        self.comboBox_font_bold.addItem("")
        self.horizontalLayout_35.addWidget(self.comboBox_font_bold)
        self.verticalLayout.addLayout(self.horizontalLayout_35)
        self.horizontalLayout_18 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_18.setObjectName("horizontalLayout_18")
        self.label_10 = QtWidgets.QLabel(self.groupBox_4)
        self.label_10.setObjectName("label_10")
        self.horizontalLayout_18.addWidget(self.label_10)
        spacerItem11 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_18.addItem(spacerItem11)
        self.checkBox_bold = QtWidgets.QCheckBox(self.groupBox_4)
        self.checkBox_bold.setChecked(False)
        self.checkBox_bold.setObjectName("checkBox_bold")
        self.horizontalLayout_18.addWidget(self.checkBox_bold)
        self.verticalLayout.addLayout(self.horizontalLayout_18)
        self.verticalLayout_9.addWidget(self.groupBox_4)
        self.verticalLayout_10.addLayout(self.verticalLayout_9)
        self.tabWidget_2.addTab(self.tab_5, "")
        self.tab_6 = QtWidgets.QWidget()
        self.tab_6.setObjectName("tab_6")
        self.verticalLayout_15 = QtWidgets.QVBoxLayout(self.tab_6)
        self.verticalLayout_15.setObjectName("verticalLayout_15")
        self.verticalLayout_14 = QtWidgets.QVBoxLayout()
        self.verticalLayout_14.setObjectName("verticalLayout_14")
        self.horizontalLayout_24 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_24.setObjectName("horizontalLayout_24")
        self.label_15 = QtWidgets.QLabel(self.tab_6)
        self.label_15.setObjectName("label_15")
        self.horizontalLayout_24.addWidget(self.label_15)
        self.ltitle1 = QtWidgets.QLineEdit(self.tab_6)
        self.ltitle1.setObjectName("ltitle1")
        self.horizontalLayout_24.addWidget(self.ltitle1)
        self.horizontalLayout_20 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_20.setObjectName("horizontalLayout_20")
        self.label_14 = QtWidgets.QLabel(self.tab_6)
        self.label_14.setObjectName("label_14")
        self.horizontalLayout_20.addWidget(self.label_14)
        self.title1row = QtWidgets.QSpinBox(self.tab_6)
        self.title1row.setMinimum(1)
        self.title1row.setMaximum(999)
        self.title1row.setProperty("value", 1)
        self.title1row.setObjectName("title1row")
        self.horizontalLayout_20.addWidget(self.title1row)
        self.horizontalLayout_24.addLayout(self.horizontalLayout_20)
        self.horizontalLayout_19 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_19.setObjectName("horizontalLayout_19")
        self.label_11 = QtWidgets.QLabel(self.tab_6)
        self.label_11.setObjectName("label_11")
        self.horizontalLayout_19.addWidget(self.label_11)
        self.title1column = QtWidgets.QSpinBox(self.tab_6)
        self.title1column.setMinimum(1)
        self.title1column.setMaximum(999)
        self.title1column.setProperty("value", 2)
        self.title1column.setObjectName("title1column")
        self.horizontalLayout_19.addWidget(self.title1column)
        self.horizontalLayout_24.addLayout(self.horizontalLayout_19)
        self.verticalLayout_14.addLayout(self.horizontalLayout_24)
        self.horizontalLayout_25 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_25.setObjectName("horizontalLayout_25")
        self.label_16 = QtWidgets.QLabel(self.tab_6)
        self.label_16.setObjectName("label_16")
        self.horizontalLayout_25.addWidget(self.label_16)
        self.ltitle2 = QtWidgets.QLineEdit(self.tab_6)
        self.ltitle2.setObjectName("ltitle2")
        self.horizontalLayout_25.addWidget(self.ltitle2)
        self.horizontalLayout_27 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_27.setObjectName("horizontalLayout_27")
        self.label_18 = QtWidgets.QLabel(self.tab_6)
        self.label_18.setObjectName("label_18")
        self.horizontalLayout_27.addWidget(self.label_18)
        self.title2row = QtWidgets.QSpinBox(self.tab_6)
        self.title2row.setMinimum(1)
        self.title2row.setMaximum(999)
        self.title2row.setProperty("value", 1)
        self.title2row.setObjectName("title2row")
        self.horizontalLayout_27.addWidget(self.title2row)
        self.horizontalLayout_25.addLayout(self.horizontalLayout_27)
        self.horizontalLayout_26 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_26.setObjectName("horizontalLayout_26")
        self.label_17 = QtWidgets.QLabel(self.tab_6)
        self.label_17.setObjectName("label_17")
        self.horizontalLayout_26.addWidget(self.label_17)
        self.title2column = QtWidgets.QSpinBox(self.tab_6)
        self.title2column.setMinimum(1)
        self.title2column.setMaximum(999)
        self.title2column.setProperty("value", 5)
        self.title2column.setObjectName("title2column")
        self.horizontalLayout_26.addWidget(self.title2column)
        self.horizontalLayout_25.addLayout(self.horizontalLayout_26)
        self.verticalLayout_14.addLayout(self.horizontalLayout_25)
        self.horizontalLayout_28 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_28.setObjectName("horizontalLayout_28")
        self.label_19 = QtWidgets.QLabel(self.tab_6)
        self.label_19.setObjectName("label_19")
        self.horizontalLayout_28.addWidget(self.label_19)
        self.ltitle3 = QtWidgets.QLineEdit(self.tab_6)
        self.ltitle3.setObjectName("ltitle3")
        self.horizontalLayout_28.addWidget(self.ltitle3)
        self.horizontalLayout_30 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_30.setObjectName("horizontalLayout_30")
        self.label_21 = QtWidgets.QLabel(self.tab_6)
        self.label_21.setObjectName("label_21")
        self.horizontalLayout_30.addWidget(self.label_21)
        self.title3row = QtWidgets.QSpinBox(self.tab_6)
        self.title3row.setMinimum(1)
        self.title3row.setMaximum(999)
        self.title3row.setProperty("value", 4)
        self.title3row.setObjectName("title3row")
        self.horizontalLayout_30.addWidget(self.title3row)
        self.horizontalLayout_28.addLayout(self.horizontalLayout_30)
        self.horizontalLayout_29 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_29.setObjectName("horizontalLayout_29")
        self.label_20 = QtWidgets.QLabel(self.tab_6)
        self.label_20.setObjectName("label_20")
        self.horizontalLayout_29.addWidget(self.label_20)
        self.title3column = QtWidgets.QSpinBox(self.tab_6)
        self.title3column.setMinimum(1)
        self.title3column.setMaximum(999)
        self.title3column.setProperty("value", 4)
        self.title3column.setObjectName("title3column")
        self.horizontalLayout_29.addWidget(self.title3column)
        self.horizontalLayout_28.addLayout(self.horizontalLayout_29)
        self.verticalLayout_14.addLayout(self.horizontalLayout_28)
        self.horizontalLayout_31 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_31.setObjectName("horizontalLayout_31")
        self.label_22 = QtWidgets.QLabel(self.tab_6)
        self.label_22.setObjectName("label_22")
        self.horizontalLayout_31.addWidget(self.label_22)
        self.ltitle4 = QtWidgets.QLineEdit(self.tab_6)
        self.ltitle4.setObjectName("ltitle4")
        self.horizontalLayout_31.addWidget(self.ltitle4)
        self.horizontalLayout_33 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_33.setObjectName("horizontalLayout_33")
        self.label_24 = QtWidgets.QLabel(self.tab_6)
        self.label_24.setObjectName("label_24")
        self.horizontalLayout_33.addWidget(self.label_24)
        self.title4row = QtWidgets.QSpinBox(self.tab_6)
        self.title4row.setMinimum(1)
        self.title4row.setMaximum(999)
        self.title4row.setProperty("value", 5)
        self.title4row.setObjectName("title4row")
        self.horizontalLayout_33.addWidget(self.title4row)
        self.horizontalLayout_31.addLayout(self.horizontalLayout_33)
        self.horizontalLayout_32 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_32.setObjectName("horizontalLayout_32")
        self.label_23 = QtWidgets.QLabel(self.tab_6)
        self.label_23.setObjectName("label_23")
        self.horizontalLayout_32.addWidget(self.label_23)
        self.title4column = QtWidgets.QSpinBox(self.tab_6)
        self.title4column.setMinimum(1)
        self.title4column.setMaximum(999)
        self.title4column.setProperty("value", 4)
        self.title4column.setObjectName("title4column")
        self.horizontalLayout_32.addWidget(self.title4column)
        self.horizontalLayout_31.addLayout(self.horizontalLayout_32)
        self.verticalLayout_14.addLayout(self.horizontalLayout_31)
        self.verticalLayout_15.addLayout(self.verticalLayout_14)
        self.horizontalLayout_34 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_34.setObjectName("horizontalLayout_34")
        spacerItem12 = QtWidgets.QSpacerItem(13, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_34.addItem(spacerItem12)
        spacerItem13 = QtWidgets.QSpacerItem(13, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_34.addItem(spacerItem13)
        self.postionview = QtWidgets.QPushButton(self.tab_6)
        self.postionview.setIconSize(QtCore.QSize(16, 16))
        self.postionview.setAutoRepeatDelay(300)
        self.postionview.setObjectName("postionview")
        self.horizontalLayout_34.addWidget(self.postionview)
        spacerItem14 = QtWidgets.QSpacerItem(13, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_34.addItem(spacerItem14)
        self.resultview = QtWidgets.QPushButton(self.tab_6)
        self.resultview.setObjectName("resultview")
        self.horizontalLayout_34.addWidget(self.resultview)
        self.exceladvanced = QtWidgets.QPushButton(self.tab_6)
        self.exceladvanced.setObjectName("exceladvanced")
        self.horizontalLayout_34.addWidget(self.exceladvanced)
        self.verticalLayout_15.addLayout(self.horizontalLayout_34)
        spacerItem15 = QtWidgets.QSpacerItem(17, 170, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_15.addItem(spacerItem15)
        self.tabWidget_2.addTab(self.tab_6, "")
        self.horizontalLayout_21.addWidget(self.tabWidget_2)
        self.tablistWidget = QtWidgets.QListWidget(self.tab_3)
        self.tablistWidget.setObjectName("tablistWidget")
        self.horizontalLayout_21.addWidget(self.tablistWidget)
        self.verticalLayout_11.addLayout(self.horizontalLayout_21)
        self.verticalLayout_12.addLayout(self.verticalLayout_11)
        self.progressBar = QtWidgets.QProgressBar(self.tab_3)
        self.progressBar.setProperty("value", 100)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout_12.addWidget(self.progressBar)
        self.label_log = QtWidgets.QLabel(self.tab_3)
        self.label_log.setText("")
        self.label_log.setObjectName("label_log")
        self.verticalLayout_12.addWidget(self.label_log)
        self.tabWidget.addTab(self.tab_3, "")
        self.tab_4 = QtWidgets.QWidget()
        self.tab_4.setObjectName("tab_4")
        self.tabWidget.addTab(self.tab_4, "")
        self.horizontalLayout.addWidget(self.tabWidget)
        self.verticalLayout_5.addLayout(self.horizontalLayout)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1047, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        self.tabWidget_2.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.sourcepath.setText(_translate("MainWindow", "D:\\ykq\\code\\ykq\\WordTool\\data\\新建 Microsoft Word 文档.docx"))
        self.Butpath.setText(_translate("MainWindow", "选择文件"))
        self.issavesourcepath.setText(_translate("MainWindow", "是否保留源文件"))
        self.butrun.setText(_translate("MainWindow", "启动"))
        self.butstop.setText(_translate("MainWindow", "停止"))
        self.groupBox_3.setTitle(_translate("MainWindow", "基础配置"))
        self.label.setText(_translate("MainWindow", " 表开始"))
        self.label_2.setText(_translate("MainWindow", "表结束"))
        self.label_4.setText(_translate("MainWindow", "行"))
        self.label_5.setText(_translate("MainWindow", "列"))
        self.label_3.setText(_translate("MainWindow", "单元格值"))
        self.tabval.setText(_translate("MainWindow", "用例编号"))
        self.bsave_conf.setText(_translate("MainWindow", "保存配置"))
        self.advanced_conf.setText(_translate("MainWindow", "高级配置"))
        self.label_27.setText(_translate("MainWindow", "  高级操作"))
        self.comboBox_macro.setItemText(0, _translate("MainWindow", "无"))
        self.comboBox_macro.setItemText(1, _translate("MainWindow", "自动生成表头"))
        self.comboBox_macro.setItemText(2, _translate("MainWindow", "选中全部的表"))
        self.pushButton_macro.setText(_translate("MainWindow", "宏执行"))
        self.groupBox_2.setTitle(_translate("MainWindow", "表内容处理"))
        self.label_6.setText(_translate("MainWindow", "数据查找"))
        self.label_7.setText(_translate("MainWindow", "数据替换"))
        self.checkBox_blank.setText(_translate("MainWindow", "去除中间空格"))
        self.label_25.setText(_translate("MainWindow", "删除行选择:"))
        self.label_26.setText(_translate("MainWindow", " 至"))
        self.groupBox_4.setTitle(_translate("MainWindow", "表格式处理"))
        self.label_13.setText(_translate("MainWindow", "字号"))
        self.comboBox_shop.setItemText(0, _translate("MainWindow", "默认"))
        self.comboBox_shop.setItemText(1, _translate("MainWindow", "小四"))
        self.comboBox_shop.setItemText(2, _translate("MainWindow", "五号"))
        self.comboBox_shop.setItemText(3, _translate("MainWindow", "小五"))
        self.comboBox_shop.setItemText(4, _translate("MainWindow", "四号"))
        self.label_12.setText(_translate("MainWindow", "字体"))
        self.comboBox_font.setItemText(0, _translate("MainWindow", "默认"))
        self.comboBox_font.setItemText(1, _translate("MainWindow", "宋体"))
        self.comboBox_font.setItemText(2, _translate("MainWindow", "黑体"))
        self.label_8.setText(_translate("MainWindow", "对齐方式"))
        self.comboBox_font_alignment.setItemText(0, _translate("MainWindow", "默认"))
        self.comboBox_font_alignment.setItemText(1, _translate("MainWindow", "居中"))
        self.comboBox_font_alignment.setItemText(2, _translate("MainWindow", "左边对齐"))
        self.comboBox_font_alignment.setItemText(3, _translate("MainWindow", "右对对齐"))
        self.comboBox_font_alignment.setItemText(4, _translate("MainWindow", "两端对齐"))
        self.label_9.setText(_translate("MainWindow", " 字体加粗"))
        self.comboBox_font_bold.setItemText(0, _translate("MainWindow", "默认"))
        self.comboBox_font_bold.setItemText(1, _translate("MainWindow", "加粗"))
        self.comboBox_font_bold.setItemText(2, _translate("MainWindow", "去加粗"))
        self.label_10.setText(_translate("MainWindow", "表边框"))
        self.checkBox_bold.setText(_translate("MainWindow", "是否加粗"))
        self.tabWidget_2.setTabText(self.tabWidget_2.indexOf(self.tab_5), _translate("MainWindow", "word"))
        self.label_15.setText(_translate("MainWindow", "标题1"))
        self.ltitle1.setText(_translate("MainWindow", "标题1"))
        self.label_14.setText(_translate("MainWindow", "行"))
        self.label_11.setText(_translate("MainWindow", "列"))
        self.label_16.setText(_translate("MainWindow", "标题2"))
        self.ltitle2.setText(_translate("MainWindow", "标题2"))
        self.label_18.setText(_translate("MainWindow", "行"))
        self.label_17.setText(_translate("MainWindow", "列"))
        self.label_19.setText(_translate("MainWindow", "标题3"))
        self.ltitle3.setText(_translate("MainWindow", "标题3"))
        self.label_21.setText(_translate("MainWindow", "行"))
        self.label_20.setText(_translate("MainWindow", "列"))
        self.label_22.setText(_translate("MainWindow", "标题4"))
        self.ltitle4.setText(_translate("MainWindow", "标题4"))
        self.label_24.setText(_translate("MainWindow", "行"))
        self.label_23.setText(_translate("MainWindow", "列"))
        self.postionview.setText(_translate("MainWindow", "提取位置预览"))
        self.resultview.setText(_translate("MainWindow", "提取结果预览"))
        self.exceladvanced.setText(_translate("MainWindow", " 高级设置"))
        self.tabWidget_2.setTabText(self.tabWidget_2.indexOf(self.tab_6), _translate("MainWindow", "Excel"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), _translate("MainWindow", "table"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_4), _translate("MainWindow", "段落"))