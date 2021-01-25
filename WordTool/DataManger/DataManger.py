from WordTool.DataManger.RulesHandler import *
from WordTool.TypeHandle.PdfData import pdftest
from WordTool.TypeHandle.ExcleOperate import ExcleOperate
from WordTool.TypeHandle.WordOperate import WordOperate
from WordTool.Thread.ThreadHandler import ThreadHandler
from WordTool.TypeHandle.WordMacro import *
from PyQt5.QtCore import *

from WordTool.DataManger.AdvancedManger import *


class Method_Type:
    M_WordTableToExcle = 0
    M_PdfTableToExcle = 1
    M_ExcletoExcle = 2
    M_WordTableToWordTable = 3


class DataManger(QObject):
    S_THREAD_BEGIN = pyqtSignal()
    S_THREAD_END = pyqtSignal()

    def __init__(self):
        super(QObject, self).__init__()

        self.M_WordOperate = WordOperate()
        self.M_ExcleOperate = ExcleOperate()
        self.m_method_suffix = 0
        self.activate_table_list = []

    def thread_run(self, m_type):
        if m_type == Method_Type.M_WordTableToExcle:
            try:
                return self.word_table_to_excle()
            except IOError:
                return -1
        elif m_type == Method_Type.M_PdfTableToExcle:
            pass

        elif m_type == Method_Type.M_ExcletoExcle:
            pass

        elif m_type == Method_Type.M_WordTableToWordTable:
            try:
                return self.word_table_to_word()
            except IOError:
                return -1

    def word_table_to_excle(self):
        if self.activate_table_list is False:
            self.get_activate_tables()
        self.ThreadHandler = ThreadHandler(self.M_WordOperate)
        self.ThreadHandler.SIGNAL_THREAD_BEGIN.connect(self.run_thread_begin)
        self.ThreadHandler.SIGNAL_THREAD_END.connect(self.run_thread_end)
        self.ThreadHandler.start()

    def word_table_to_word(self):
        if self.activate_table_list is False:
            self.get_activate_tables()
        self.ThreadHandler = ThreadHandler(self.M_WordOperate, 0)
        self.ThreadHandler.SIGNAL_THREAD_BEGIN.connect(self.run_thread_begin)
        self.ThreadHandler.SIGNAL_THREAD_END.connect(self.run_thread_end)
        self.ThreadHandler.start()

    def run_thread_begin(self):
        self.S_THREAD_BEGIN.emit()

    def run_thread_end(self):
        self.S_THREAD_END.emit()

    def get_Method_Type_Suffix(self, m_method_type):
        suffix = ".xls"
        if (m_method_type == Method_Type.M_WordTableToExcle):
            pass
        elif (m_method_type == Method_Type.M_PdfTableToExcle):
            pass
        elif (m_method_type == Method_Type.M_PdfTableToExcle):
            suffix = ".docx"
        return suffix

    def get_activate_tables(self):
        word_path = RULES_HANDLER_INSTANCE.get_s_path()
        if word_path == None:
            return None

        tables = RULES_HANDLER_INSTANCE.get_table_source_cell()

        self.M_WordOperate.set_word_path(word_path)
        self.M_WordOperate.get_tables(*tables)
        self.M_WordOperate.set_activate_table_range(*RULES_HANDLER_INSTANCE.get_table_range())
        self.activate_table_list = self.M_WordOperate.get_activate_table_list()

    def get_generate_data(self):
        self.get_activate_tables()
        self.M_WordOperate.table_relational_mapping(RULES_HANDLER_INSTANCE.get_table_required_cells())
        data = self.M_WordOperate.excle_data
        if RULES_HANDLER_INSTANCE.get_advanced_settings() is not None:
            return AdvancedManger.advanced_data_deal(data)
        return data

    def get_activate_tables_list(self):
        self.get_activate_tables()
        return self.activate_table_list

    def get_activate_map_table(self):

        activate_map_table = {}

        for index, table in enumerate(self.get_activate_tables_list()):
            activate_map_table[index + 1] = table

        return activate_map_table

    def data_to_excel(self, ecl_data, label=[]):

        self.M_ExcleOperate.save_excle(ecl_data, label)

    def source_object_handling(self):
        for table in self.get_activate_tables_list():
            self.M_WordOperate.table_all_paragraphs_replace(table, "测试", "天下")
            self.M_WordOperate.del_table_row(table, 1)
            self.M_WordOperate.table_all_paragraphs_style(table, ContentStyle)
            self.M_WordOperate.align_table(table, "CENTER")
        self.M_WordOperate.save("D:\\ykq\\code\\ykq\\WordTool\\data\\test.docx")

    @staticmethod
    def macro(index):

        WordMacro.macro(index, RULES_HANDLER_INSTANCE.get_s_path())


if __name__ == '__main__':
    pass
