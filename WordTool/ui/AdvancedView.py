from WordTool.ui.ui_fiel.ui_advancedsettings import *
from PyQt5.QtWidgets import *
from WordTool.DataManger.DataManger import *
from PyQt5.QtGui import *


class ContextWidget(QListWidget):
    combination_changed_signal = pyqtSignal()

    def __init__(self, list_widget):
        super(ContextWidget, self).__init__()
        self.setFlow(QListView.TopToBottom)
        self.setAcceptDrops(True)
        self.setMovement(QListView.Free)
        list_widget.itemPressed.connect(self.current_update)
        self.external = False
        self.item_widget_kw = {}
        self.current_drag_item = None

    def contextMenuEvent(self, event):
        if self.count() < 1:
            return
        menu = QMenu()
        action_up = QAction(u'上移', self)
        action_up.triggered.connect(self.item_up)
        action_down = QAction(u'下移', self)
        action_down.triggered.connect(self.item_down)
        action_del = QAction(u'删除', self)
        action_del.triggered.connect(self.del_item)
        menu.addAction(action_up)
        menu.addAction(action_down)
        menu.addAction(action_del)
        menu.exec(QCursor.pos())

    def item_up(self):
        current_index = self.currentRow()
        if current_index - 1 < 0:
            return
        item_widget = self.itemWidget(self.currentItem())
        new_item = self.takeItem(current_index)
        self.insertItem(current_index - 1, new_item)
        self.setCurrentRow(current_index - 1)
        self.setItemWidget(new_item, self.create_widget_to_obj(item_widget))

        self.combination_changed_signal.emit()

    def item_down(self):
        current_index = self.currentRow()
        if self.count() <= current_index + 1:
            return
        item_widget = self.itemWidget(self.currentItem())
        new_item = self.takeItem(current_index)
        self.insertItem(current_index + 1, new_item)
        self.setCurrentRow(current_index + 1)
        self.setItemWidget(new_item, self.create_widget_to_obj(item_widget))

        self.combination_changed_signal.emit()

    def del_item(self):
        self.takeItem(self.currentRow())
        # self.setCurrentRow(self.currentRow() - 1)
        self.combination_changed_signal.emit()

    def dropEvent(self, event):
        if self.external:
            item = self.current_drag_item
            self.external = False
            self.append_widget(item.text())
        else:
            item = QListWidgetItem()
            item.setText(self.currentItem().text())
            self.add_item_widget(item, self.itemWidget(self.currentItem()))
            self.del_item()

        self.combination_changed_signal.emit()

    def append_widget(self, name):
        item = QListWidgetItem()
        item.setText(name)
        self.add_item_widget(item, self.create_widget_to_enum_type(self.type2name(name)))

    def add_item_widget(self, item, widget):
        self.addItem(item)
        self.setItemWidget(item, widget)

    def append_data_to_items(self, combination_date):

        for widget_type, val in combination_date:
            item = QListWidgetItem()
            item.setText(self.name2type(widget_type))
            self.addItem(item)
            self.setItemWidget(item, self.create_widget_to_enum_type(widget_type, val))

    def get_combination_date(self):
        combination_data = []
        for item_row in range(self.count()):
            combination_data.append(self.get_val_to_obj(self.itemWidget(self.item(item_row))))

        return combination_data

    def current_update(self, item):
        self.external = True
        self.current_drag_item = item

    @staticmethod
    def create_widget_to_enum_type(enum_widget_type, value=""):
        if enum_widget_type == EnumWidget.Label:
            label = QLabel()
            label.setMaximumWidth(60)
            # label.setMaximumHeight(60)
            label.setText(str(value))
            label.setToolTip(str(value))
            return label
        elif enum_widget_type == EnumWidget.LineEdit:
            line_edit = QLineEdit()
            line_edit.setMaximumWidth(60)
            # lineEdit.setMaximumHeight(60)
            line_edit.setText(str(value))
            line_edit.setToolTip(str(value))
            return line_edit
        elif enum_widget_type == EnumWidget.SpinBox:
            spinbox = QSpinBox()
            spinbox.setMaximumWidth(60)
            # spinBox.setMaximumHeight(60)

            spinbox.setValue(value if type(value) is int else 0)
            spinbox.setToolTip(str(value))
            return spinbox

    @staticmethod
    def create_widget_to_obj(obj_type):
        if type(obj_type) == QLabel:
            label = QLabel()
            label.setMaximumWidth(60)
            # label.setMaximumHeight(60)
            label.setText(obj_type.text())
            label.setToolTip(obj_type.text())
            return label
        elif type(obj_type) == QLineEdit:
            line_edit = QLineEdit()
            line_edit.setMaximumWidth(60)
            # lineEdit.setMaximumHeight(60)
            line_edit.setText(obj_type.text())
            line_edit.setToolTip(obj_type.text())
            return line_edit
        elif type(obj_type) == QSpinBox:
            spinbox = QSpinBox()
            spinbox.setMaximumWidth(60)
            # spinBox.setMaximumHeight(60)
            spinbox.setValue(obj_type.value())
            spinbox.setToolTip(str(obj_type.value()))
            return spinbox

    @staticmethod
    def get_val_to_obj(obj):
        if type(obj) == QLabel:
            return [EnumWidget.Label, obj.text()]
        elif type(obj) == QLineEdit:
            return [EnumWidget.LineEdit, obj.text()]
        elif type(obj) == QSpinBox:
            return [EnumWidget.SpinBox, obj.value()]

    @staticmethod
    def type2name(name):

        if name == u"原始内容":
            widget_type = EnumWidget.Label

        elif name == u"自增序号":
            widget_type = EnumWidget.SpinBox

        elif name == u"补充值":
            widget_type = EnumWidget.LineEdit
        return widget_type

    @staticmethod
    def name2type(widget_type):

        if widget_type == EnumWidget.Label:
            name = u"原始内容"
        elif widget_type == EnumWidget.SpinBox:
            name = u"自增序号"
        elif widget_type == EnumWidget.LineEdit:
            name = u"补充值"
        return name


class AdvancedView(QWidget, Ui_AdvancedView):

    def __init__(self):
        super(AdvancedView, self).__init__()
        self.setupUi(self)
        self.init_ui()
        self.connect_init()
        self.m_context = "我叫示例"

    def connect_init(self):
        """界面信号手动连接"""

        self.One_Column.valueChanged.connect(self.location_value_changed)
        self.Multiple_Column.stateChanged.connect(self.location_value_changed)
        self.Combination_Column.textChanged.connect(self.location_value_changed)

        self.repeat1.clicked.connect(self.repeat_value_changed)
        self.repeat2.clicked.connect(self.repeat_value_changed)
        self.repeat3.clicked.connect(self.repeat_value_changed)

        self.slice1.valueChanged.connect(self.slice_value_changed)
        self.slice2.valueChanged.connect(self.slice_value_changed)

        self.pushtest.clicked.connect(self.update_preview)

        self.textEdit_deal.textChanged.connect(self.eval_value_changed)

        self.checkBox_eval.stateChanged.connect(self.eval_value_changed)

        self.com_conversion.currentIndexChanged.connect(self.conversion_value_changed)

    def init_ui(self):

        self.contextlistWidget.setDragEnabled(True)

        self.ContextWidget = ContextWidget(self.contextlistWidget)
        self.ContextWidget.combination_changed_signal.connect(self.combination_changed)
        self.containerlistWidget.addWidget(self.ContextWidget)

        self.kw = RULES_HANDLER_INSTANCE.get(RULES_ID.ADVANCED_SETTINGS)
        if self.kw is None:
            self.kw = {}
            self.__location_default()
            self.__repeat_default()
            self.__slice_default()
            self.__combination_default()
            self.__eval_default()
            self.__conversion_default()
            return

        self.__data_to_view()

    def __data_to_view(self):

        self.location(self.kw["location"])
        self.repeat(self.kw["repeat"])
        self.slice(self.kw["slice"])
        self.combination(self.kw["combination"])
        self.eval(self.kw.get("eval"))
        self.conversion(self.kw.get("conversion"))

        self.update_preview()

    def closeEvent(self, close_event):
        self.set_parameter()
        super().closeEvent(close_event)

    def set_parameter(self):
        self.kw["combination"] = self.ContextWidget.get_combination_date()
        RULES_HANDLER_INSTANCE.set_advanced_settings(self.kw)

    def __location_default(self):
        if self.kw.get("location") is None:
            self.kw["location"] = {EnumLocation.One_Column: -1, EnumLocation.Multiple_Column: True,
                                   EnumLocation.Combination_Column: ""}

        self.location(self.kw["location"])

    def location_value_changed(self):
        if self.kw["location"] == {EnumLocation.One_Column: self.One_Column.value(),
                                   EnumLocation.Multiple_Column: self.Multiple_Column.isChecked(),
                                   EnumLocation.Combination_Column: self.Combination_Column.text()}:
            return

        if type(self.sender()) == QSpinBox:
            self.kw["location"] = {EnumLocation.One_Column: self.One_Column.value(),
                                   EnumLocation.Multiple_Column: False,
                                   EnumLocation.Combination_Column: ""}
        elif type(self.sender()) == QCheckBox:
            if self.Multiple_Column.isChecked():
                self.kw["location"] = {EnumLocation.One_Column: 0,
                                       EnumLocation.Multiple_Column: self.Multiple_Column.isChecked(),
                                       EnumLocation.Combination_Column: ""}
        elif type(self.sender()) == QLineEdit:
            self.kw["location"] = {EnumLocation.One_Column: 0,
                                   EnumLocation.Multiple_Column: False,
                                   EnumLocation.Combination_Column: self.Combination_Column.text()}

        self.location(self.kw["location"])

    def location(self, location):
        self.One_Column.setValue(location[EnumLocation.One_Column])
        self.Multiple_Column.setChecked(location[EnumLocation.Multiple_Column])
        self.Combination_Column.setText(location[EnumLocation.Combination_Column])

    def __repeat_default(self):
        if self.kw.get("repeat") is None:
            self.kw["repeat"] = EnumRepeat.No_Processing
        self.repeat(self.kw["repeat"])

    def repeat_value_changed(self):
        if self.sender().text() == "无处理":
            self.kw["repeat"] = EnumRepeat.No_Processing
        elif self.sender().text() == "单例去重":
            self.kw["repeat"] = EnumRepeat.One_Processing
        elif self.sender().text() == "整条去重":
            self.kw["repeat"] = EnumRepeat.All_Processing

        self.repeat(self.kw["repeat"])

    def repeat(self, m_repeat):
        if m_repeat == EnumRepeat.No_Processing:
            self.repeat1.setChecked(True)
        elif m_repeat == EnumRepeat.One_Processing:
            self.repeat2.setChecked(True)
        elif m_repeat == EnumRepeat.All_Processing:
            self.repeat3.setChecked(True)

    def __slice_default(self):
        if self.kw.get("slice") is None:
            self.kw["slice"] = [0, 99]
        self.slice(self.kw["slice"])

    def slice_value_changed(self):
        self.kw["slice"] = [self.slice1.value(), self.slice2.value()]

        self.slice(self.kw["slice"])
        self.update_preview()

    def slice(self, m_slice):
        start = m_slice[0]
        end = m_slice[1]
        self.slice1.setValue(start)
        self.slice2.setValue(end)

    def __combination_default(self):
        if self.kw.get("combination"):
            self.combination(self.kw.get("combination"))
        else:
            self.kw["combination"] = []

    def combination(self, m_combination):

        if m_combination:
            self.ContextWidget.append_data_to_items(m_combination)
        # self.update_preview()

    def combination_changed(self):
        self.kw["combination"] = self.ContextWidget.get_combination_date()

        self.update_preview()

    def __eval_default(self):
        if self.kw.get("eval") is None:
            self.kw["eval"] = [False, ""]
        self.eval(self.kw["eval"])

    def eval(self, eval_val):
        if eval_val:
            self.checkBox_eval.setChecked(eval_val[0])
            self.textEdit_deal.setText(eval_val[1])

    def eval_value_changed(self):

        self.kw["eval"] = [self.checkBox_eval.isChecked(), self.textEdit_deal.toPlainText()]

    def __conversion_default(self):
        if self.kw.get("conversion") is None:
            self.kw["conversion"] = 0
        self.conversion(self.kw["conversion"])

    def conversion(self, conversion_val):
        if conversion_val is None:
            self.kw["conversion"] = 0
            conversion_val = 0
        self.com_conversion.setCurrentIndex(conversion_val)

    def conversion_value_changed(self, conversion_val):

        self.kw["conversion"] = conversion_val
        self.update_preview()

    def get_parameter(self):
        return self.kw

    def update_preview(self):
        data = "我是示例"
        """判断响应来源，当是点击按钮触发时，不进行状态验证"""
        is_push_test = False
        if self.sender() is not None:
            if type(self.sender()) == QPushButton:
                is_push_test = True
        """依据配置类型进行演示数据的实时显示处理"""
        if self.checkBox_eval.isChecked() or is_push_test:
            eval_text = self.kw.get("eval")
            val = AdvancedManger.eval_val_cal(data, eval_text[1])
        else:
            combination = self.kw.get("combination")
            slice_val = self.kw.get("slice")
            conversion = self.kw.get("conversion")

            val = AdvancedManger.configuration_val_cal(data, combination, slice_val, True, conversion)
        self.contextlabel.setText(val)
