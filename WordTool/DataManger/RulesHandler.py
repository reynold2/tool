from typing import Any

from WordTool.Common.GlobalVar import *
import pickle


class RULES_ID:
    " 源文件路径 "

    DOC_SPATH = 0
    " 目标保存路径"
    DOC_TPATH = 1

    "  word中源位置单元格（定位相同位置单元格）同一位位置出现值等于多少的方式去匹配"
    TAB_SOURCE = 2
    " 需要提取出来的单元格数据"
    TAB_REQUIREDCELLS = 3

    "需要处理的表的范围，在已经提取的出来的中表中"
    TAB_ACTIVATE_LEN = 4

    "需要提取出来的单元格数据_标签名"
    TAB_REQUIREDCELLS_LABEL = 5

    "结果高级设置参数"
    ADVANCED_SETTINGS = 6

    "表内容处理设置参数"

    TABLE_CONTENT_PROCESSING = 7

    "表格式处理设置参数"

    TABLE_FORMAT_PROCESSING = 8


class RulesHandler(object):
    __INSTANCE = None
    "类属性"

    def __new__(cls, *args, **kwargs):
        if cls.__INSTANCE == None:
            cls.__INSTANCE = object.__new__(cls)
            cls.__kw_Rules = {}
        return cls.__INSTANCE

    " 源文件路径"

    def set_s_path(self, path):
        self.__kw_Rules[RULES_ID.DOC_SPATH] = path

    "目标文件路径"

    def set_t_path(self, path):
        self.__kw_Rules[RULES_ID.DOC_TPATH] = path

    "源文件路径"

    def get_s_path(self):
        return self.__kw_Rules.get(RULES_ID.DOC_SPATH)

    " 目标文件路径"

    def get_t_path(self):
        return self.__kw_Rules[RULES_ID.DOC_TPATH]

    "获取全部规则"

    def get_rules(self):
        return self.__kw_Rules

    "目标单元格将word中所有单元格中包含该位置值等于时查找出来"

    def set_table_source_cell(self, source):
        self.__kw_Rules[RULES_ID.TAB_SOURCE] = source

    " 目标位置和需要获取值的字段位置"

    def get_table_source_cell(self):

        return self.__kw_Rules.get(RULES_ID.TAB_SOURCE, None)

    '设置需要的位置用于生成Excel的位置'

    def set_table_required_cells(self, requiredcells):
        self.__kw_Rules[RULES_ID.TAB_REQUIREDCELLS] = requiredcells

    " 目标位置和需要获取值的字段位置"

    def get_table_required_cells(self):
        return self.__kw_Rules.get(RULES_ID.TAB_REQUIREDCELLS, None)

    " 用于控制表范围"

    def set_table_range(self, index):
        self.__kw_Rules[RULES_ID.TAB_ACTIVATE_LEN] = index

    def get_table_range(self):
        return self.__kw_Rules.get(RULES_ID.TAB_ACTIVATE_LEN, None)

    "目标标签页"

    def set_table_label(self, label):
        self.__kw_Rules[RULES_ID.TAB_REQUIREDCELLS_LABEL] = label

    def get_table_label(self):

        return self.__kw_Rules.get(RULES_ID.TAB_REQUIREDCELLS_LABEL, None)

    "高级设置"

    def set_advanced_settings(self, advanced_settings):
        self.__kw_Rules[RULES_ID.ADVANCED_SETTINGS] = advanced_settings

    def get_advanced_settings(self):
        return self.__kw_Rules.get(RULES_ID.ADVANCED_SETTINGS, None)

    "表内容"

    def set_table_content(self, content):
        self.__kw_Rules[RULES_ID.TABLE_CONTENT_PROCESSING] = content

    def get_table_content(self):

        return self.__kw_Rules.get(RULES_ID.TABLE_CONTENT_PROCESSING, None)

    "表格式"

    def set_table_style(self, style):
        self.__kw_Rules[RULES_ID.TABLE_FORMAT_PROCESSING] = style

    def get_table_style(self):

        return self.__kw_Rules.get(RULES_ID.TABLE_FORMAT_PROCESSING, None)

    " 清空设置"

    def clear_kw_rules(self):
        self.__kw_Rules.clear()

    "字典获取"

    def get(self, k):
        return self.__kw_Rules.get(k)

    "序列化对象规则"

    def dump_rules_handler(self):
        try:
            with open("Rules.data", "wb") as f:
                pickle.dump(self.__kw_Rules, f)
        except IOError:
            return None

    "反序列化对象规则"

    def load_rules_handler(self):
        self.clear_kw_rules()
        try:
            with open(r"Rules.data", "rb") as f:
                self.__kw_Rules = pickle.load(f)
                return True
        except IOError:
            return False


'创建规则实例对象'
RULES_HANDLER_INSTANCE = RulesHandler()
if __name__ == '__main__':
    RULES_HANDLER_INSTANCE.get_s_path()
