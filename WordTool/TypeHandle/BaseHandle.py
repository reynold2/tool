import abc
from PyQt5.QtCore import *
from WordTool.DataManger.RulesHandler import *


class BaseHandle(metaclass=abc.ABCMeta):

    def __init__(self):
        self.SIGNAL_WORK_END = pyqtSignal()
        self.SIGNAL_WORK_END.connect(self.work_stop)

    @abc.abstractmethod
    def run(self, run_param=None):
        self.set_word_path(RULES_HANDLER_INSTANCE.get_s_path())
        self.get_tables(*RULES_HANDLER_INSTANCE.get_table_source_cell())
        self.table_relational_mapping(RULES_HANDLER_INSTANCE.get_table_required_cells())
        self.save_excle(RULES_HANDLER_INSTANCE.get_table_label(), RULES_HANDLER_INSTANCE.get_t_path())

    @abc.abstractmethod
    def set_rules(self, rule):
        pass

    @abc.abstractmethod
    def get_rules(self):
        pass

    @abc.abstractmethod
    def work_stop(self):
        pass


if __name__ == '__main__':
    print(1)
