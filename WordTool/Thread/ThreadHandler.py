from PyQt5.QtCore import *

from WordTool.TypeHandle.BaseHandle import *


class ThreadHandler(QThread):
    SIGNAL_THREAD_BEGIN = pyqtSignal()
    SIGNAL_THREAD_END = pyqtSignal()

    def __init__(self, obj, run_param=None):
        super(QThread, self).__init__()
        self.obj = obj
        self.run_param = run_param

    def run(self):
        self.SIGNAL_THREAD_BEGIN.emit()
        self.obj.run(self.run_param)
        self.SIGNAL_THREAD_END.emit()

    def quit(self):
        # 线程退出
        self.quit()
        # 工作停止

        self.obj.SIGNAL_WORK_END.emit()
