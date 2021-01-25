from PyQt5.QtWidgets import *
from WordTool.ui.ui_fiel.ui_mianwindow import *
from WordTool.ui.AdvancedView import *
import os
from WordTool.DataManger.DataManger import *
from WordTool.ui.PreviewWidget import PreviewWidget
from PyQt5.QtCore import *


# 加载 ui 文件
# Ui_MainWindow, QtBaseClass = uic.loadUiType('ui/ui_mianwindow.ui')


class MainUi(QMainWindow, Ui_MainWindow):

    def __init__(self):
        QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.__init_self()

    def __init_self(self):

        self.M_DataManger = DataManger()
        self.__init__connect()
        self.__init__attribute()
        self.__init_rules_view()
        self.m_Advanced_View = AdvancedView()

    def __init_rules_view(self):
        if RULES_HANDLER_INSTANCE.load_rules_handler() is False:
            return
        """设置默认路径"""
        s_path = RULES_HANDLER_INSTANCE.get_s_path()
        if s_path:
            self.sourcepath.setText(s_path)
        "设置默认需要处理的数据"
        table_source_cell = RULES_HANDLER_INSTANCE.get_table_source_cell()
        if table_source_cell:
            self.tabrow.setValue(table_source_cell[0] + 1)
            self.tabcolumn.setValue(table_source_cell[1] + 1)
            self.tabval.setText(table_source_cell[2])
        "设置默认需要保存的数据"
        required_cells = RULES_HANDLER_INSTANCE.get_table_required_cells()
        if required_cells:
            self.set_init_required_cells(required_cells)
        "设置标签页"
        label = RULES_HANDLER_INSTANCE.get_table_label()
        if label:
            self.set_init_table_label(label)
        "设置有效表范围"
        range_table_index = RULES_HANDLER_INSTANCE.get_table_range()
        if range_table_index:
            self.tabstart.setValue(range_table_index[0])
            self.tabend.setValue(range_table_index[1])
        "设置表格式"
        table_style = RULES_HANDLER_INSTANCE.get_table_style()
        if table_style:
            self.comboBox_shop.setCurrentIndex(table_style[0])
            self.comboBox_font.setCurrentIndex(table_style[1])
            self.comboBox_font_alignment.setCurrentIndex(table_style[2])
            self.comboBox_font_bold.setCurrentIndex(table_style[3])
            self.checkBox_bold.setChecked(table_style[4])
        "设置表内容"
        table_content = RULES_HANDLER_INSTANCE.get_table_content()
        if table_content and len(table_content) == 5:
            self.lineEdit_replace_s.setText(table_content[0])
            self.lineEdit_replace_t.setText(table_content[1])
            self.checkBox_blank.setChecked(table_content[2])
            self.spinBox_del1.setValue(table_content[3])
            self.spinBox_del2.setValue(table_content[4])

    def __set_rules(self):
        """设置默认路径"""
        RULES_HANDLER_INSTANCE.set_s_path(self.sourcepath.text())
        "设置默认需要处理的数据"
        table_source_cell = [int(self.tabrow.text()) - 1, int(self.tabcolumn.text()) - 1, self.tabval.text()]
        RULES_HANDLER_INSTANCE.set_table_source_cell(table_source_cell)
        "设置默认需要保存的数据"
        RULES_HANDLER_INSTANCE.set_table_required_cells(self.get__init_required_cells())
        "设置标签页"
        RULES_HANDLER_INSTANCE.set_table_label(list(self.tag_label_index.keys()))
        "设置保存路径"
        path = os.path.splitext(self.sourcepath.text())[
                   0] + self.M_DataManger.get_Method_Type_Suffix(
            self.get_method()) if self.issavesourcepath.isChecked() else ""

        RULES_HANDLER_INSTANCE.set_t_path(path)
        "设置有效表范围"
        RULES_HANDLER_INSTANCE.set_table_range([int(self.tabstart.text()), int(self.tabend.text())])
        "设置表格式"
        RULES_HANDLER_INSTANCE.set_table_style(
            [self.comboBox_shop.currentIndex(), self.comboBox_font.currentIndex(),
             self.comboBox_font_alignment.currentIndex(),
             self.comboBox_font_bold.currentIndex(), self.checkBox_bold.isChecked()])
        "设置表内容"
        RULES_HANDLER_INSTANCE.set_table_content(
            [self.lineEdit_replace_s.text(), self.lineEdit_replace_t.text(), self.checkBox_blank.isChecked(),
             self.spinBox_del1.value(), self.spinBox_del2.value()])

    def update_required_cells(self):
        """重新从界面上的输入获取需要的位置"""
        RULES_HANDLER_INSTANCE.set_table_required_cells(self.get__init_required_cells())

    def __init__connect(self):
        self.butrun.clicked.connect(self.run)
        self.Butpath.clicked.connect(self.get_filepath)
        self.bsave_conf.clicked.connect(self.activate_table_preview)
        self.tablistWidget.currentTextChanged.connect(self.set_current_text)
        self.tablistWidget.doubleClicked.connect(self.preview_current)
        self.resultview.clicked.connect(self.preview_generate_data)
        self.postionview.clicked.connect(self.preview_tag)
        self.M_DataManger.S_THREAD_BEGIN.connect(self.run_thread_begin)
        self.M_DataManger.S_THREAD_END.connect(self.run_thread_end)
        self.exceladvanced.clicked.connect(self.advanced_view)
        self.pushButton_macro.clicked.connect(self.macro_run)

    def __init__attribute(self):
        self.activate_map_table = {}
        self.current_text = ""
        self.tag_label_index = {}

    def advanced_view(self):

        self.m_Advanced_View.show()

    def set_current_text(self, m_current_text):

        self.current_text = m_current_text

    def preview_current(self):
        self.preview(self.activate_map_table[self.get_current_text_key(self.current_text)])

    def get_current_text_key(self, m_current_text=None):
        key = 0
        if len(self.activate_map_table) == 0:
            return None
        if m_current_text is not None:
            try:
                key = int(m_current_text.split(":")[0])
            except:
                pass
        return key

    def get_current_text(self):
        return self.current_text

    def update_rules(self):
        self.__set_rules()

    def get__init_required_cells(self):
        """提取需要保存的位置数据，依据界面设置，进行初始化"""
        self.tag_label_index.clear()
        if self.ltitle1.text() != "":
            self.tag_label_index[self.ltitle1.text()] = [int(self.title1row.text()) - 1,
                                                         int(self.title1column.text()) - 1]
        if self.ltitle2.text() != "":
            self.tag_label_index[self.ltitle2.text()] = [int(self.title2row.text()) - 1,
                                                         int(self.title2column.text()) - 1]
        if self.ltitle3.text() != "":
            self.tag_label_index[self.ltitle3.text()] = [int(self.title3row.text()) - 1,
                                                         int(self.title3column.text()) - 1]
        if self.ltitle4.text() != "":
            self.tag_label_index[self.ltitle4.text()] = [int(self.title4row.text()) - 1,
                                                         int(self.title4column.text()) - 1]

        return list(self.tag_label_index.values())

    def set_init_required_cells(self, values=None):
        """将需求定位位置显示到视图中"""
        if values is None:
            return
        for index, value in enumerate(values):
            eval(" self.title%srow.setValue(value[0]+1)" % str(index + 1))
            eval(" self.title%scolumn.setValue(value[1]+1)" % str(index + 1))

    def set_init_table_label(self, labels=None):
        """将label显示到视图中"""
        if labels is None:
            return
        for index, value in enumerate(labels):
            eval(" self.ltitle%s.setText(value)" % str(index + 1))

    def run(self):

        self.__set_rules()

        if RULES_HANDLER_INSTANCE.get_s_path() != "":
            self.M_DataManger.thread_run(self.get_method())

    def run_thread_begin(self):
        self.timer = QTimer(self)  # 初始化一个定时器
        self.timer.timeout.connect(self.label_log_display)  # 每次计时到时间时发出信号
        self.timer.start(1000)  # 设置计时间隔并启动；单位毫秒
        self.num = 0
        self.butrun.setEnabled(False)
        self.progressBar.setValue(0)

    def label_log_display(self):
        self.num = self.num + 1
        self.label_log.setText("正在生成:" + str(self.num) + "秒！！！")
        self.progressBar.setValue(self.num * 10 if self.num < 10 else 100)

    def run_thread_end(self):
        self.timer.stop()
        self.label_log.setText("生成完成！耗时：" + str(self.num) + "秒！！！")
        self.num = 0
        self.butrun.setEnabled(True)
        self.progressBar.setValue(100)

    def get_filepath(self):
        file_name, file_type = QFileDialog.getOpenFileName(self, "选取文件", os.getcwd(),
                                                           "All Files(*);;Text Files(*.txt)")
        self.set_source_filepath(file_name)
        self.sourcepath.setText(file_name)

    def get_method(self):
        current_father_title = self.tabWidget.currentIndex()
        current_sub_title = self.tabWidget_2.currentIndex()
        m_method = 0
        if current_father_title == 0 and current_sub_title == 1:
            m_method = Method_Type.M_WordTableToExcle

        elif current_father_title == 0 and current_sub_title == 0:
            m_method = Method_Type.M_WordTableToWordTable
        return m_method

    def set_source_filepath(self, file_name):
        self.__filepath = file_name

    def get_source_filepath(self):
        return self.__filepath

    def activate_table_preview(self):
        self.__set_rules()
        self.tablistWidget.clear()
        self.activate_map_table.clear()
        activate_map_table = self.M_DataManger.get_activate_map_table()
        count = len(activate_map_table)

        for index in activate_map_table.keys():
            """获取的单元格信息作为标签名"""
            index = count + 1 - index
            label = activate_map_table[index].cell(1, 2).text

            if len(label) > 8:
                table_label = str(index) + ":" + label[:8]

            else:
                table_label = str(index) + ":" + label

            self.tablistWidget.insertItem(0, table_label)
            self.tablistWidget.item(0).setToolTip(label)
            """默认加载时第一个为选择"""
            if index == 1:
                self.tablistWidget.item(0).setSelected(True)

        self.activate_map_table = activate_map_table

    def get_current_activate_table(self):
        """ 默认为当前列表中的第一个"""
        table_index = 1
        if len(self.current_text.strip()) != 0:
            table_index = self.get_current_text_key(self.current_text)
        return self.activate_map_table.get(table_index)

    def preview_tag(self):
        """预览生成数据在原始位置界面"""
        if not bool(self.activate_map_table):
            return

        self.update_required_cells()
        if not RULES_HANDLER_INSTANCE.get_table_required_cells():
            return

        self.preview_tag_widget = PreviewWidget()
        self.update_required_cells()

        self.preview_tag_widget.set_activate_table_model(self.get_current_activate_table(),
                                                         RULES_HANDLER_INSTANCE.get_table_required_cells())

        self.preview_tag_widget.show()

    def preview_generate_data(self):
        """预览Excel生成数据"""
        self.update_required_cells()
        label = list(self.tag_label_index.keys())

        data = self.M_DataManger.get_generate_data()
        if not data:
            return
        self.preview_generate_data_widget = PreviewWidget()

        self.preview_generate_data_widget.save = QPushButton("保存")
        self.preview_generate_data_widget.h_layout.addWidget(self.preview_generate_data_widget.save)

        # 高级数据处理

        self.preview_generate_data_widget.set_table_data_model(label, data)

        # 得到转换后的excel_data
        excel_data = self.model_to_exl(self.preview_generate_data_widget.get_current_model())
        # 保存预览表中全部数据
        self.preview_generate_data_widget.save.clicked.connect(lambda: self.save_all(excel_data))

        self.preview_generate_data_widget.show()

    def preview(self, activate_table):
        """预览生成数据在原始位置"""
        self.preview_widget = PreviewWidget()
        self.update_required_cells()
        self.preview_widget.set_activate_table_model(activate_table)

        self.preview_widget.show()

    def save_all(self, excel_data, labels=[]):
        """保存excel_data到默认位置中"""
        self.M_DataManger.data_to_excel(excel_data, labels)

    def model_to_exl(self, model=None):
        """将视图中数据转换为Excel数据"""
        excel_data = []

        columnCount = model.columnCount()
        rowCount = model.rowCount()

        for index in range(columnCount):
            excel_data.append([0, index, model.horizontalHeaderItem(index).text()])
        for row in range(rowCount):
            for column in range(columnCount):
                excel_data.append(
                    [row + 1, column, model.item(row, column).text()])

        return excel_data

    def closeEvent(self, *args, **kwargs):
        """关闭主界面同时关闭所有界面"""
        self.close()
        if hasattr(self, "preview_widget"):
            self.preview_widget.close()
        elif hasattr(self, "preview_generate_data_widget"):
            self.preview_generate_data_widget.close()
        elif hasattr(self, "preview_tag_widget"):
            self.preview_tag_widget.close()
        elif hasattr(self, "m_Advanced_View"):
            self.m_Advanced_View.close()

        """序列化规则对象"""
        self.__set_rules()
        RULES_HANDLER_INSTANCE.dump_rules_handler()

    def macro_run(self):

        self.M_DataManger.macro(self.comboBox_macro.currentIndex())


if __name__ == '__main__':
    pass
