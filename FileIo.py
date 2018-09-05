'''
Created on 2018年7月4日

@author: Administrator
'''
import xlrd
import configparser
import xlwt
import os
from xlutils.copy import copy
import logging



class Excel(object):
    instance=None
    init_falg=False
    def __new__(cls, *args, **kwargs):
        if cls.instance is None:
            cls.instance=super().__new__(cls)
            return cls.instance
    def __init__(self,path,*args,**kwargs):
        if Excel.init_falg is False:

            self._excelpath=path
            self.list_exceldata = []
            Excel.init_falg=True
            return
    @property
    def excelpath(self):
        return self._excelpath

    @excelpath.setter
    def excelpath(self, value):
        self._excelpath = value
        return self._excelpath



    def ReadEXcleData(self):
        try:
            self.exceldata = xlrd.open_workbook(self._excelpath)
            table = self.exceldata.sheet_by_index(0)
            for h in range(table.nrows):
                for l in range(table.ncols):
                    self.list_exceldata.append(h)
                    self.list_exceldata.append(l)
                    self.list_exceldata.append(str(table.cell(h, l).value))
            return self.list_exceldata
        except Exception as res:
            logging.exception("Read ExcelFile opening exception:%s"%(res))
    def WriteEXcleData(self, path=None,listdata=[]):
        if (listdata and path) is not None:
            try:
                wb = xlwt.Workbook(encoding='utf-8')
                sh = wb.add_sheet("Report")
                for i in range(0, len(listdata), 3):
                    b = listdata[i:i + 3]
                    sh.write(b[0], b[1],
                                 b[2])
            except IndexError:
                logging.exception(
                        "File data is lost, incoming data cannot be triples, illegal")
            finally:
                    wb.save(self._excelpath)
        elif path is not None and listdata is None:
            wb = copy(self.exceldata)
            wb.save(path)
        elif path is None and listdata is not None:
            try:
                wb = xlwt.Workbook(encoding='utf-8')
                sh = wb.add_sheet("Report")
                for i in range(0, len(listdata), 3):
                    b = listdata[i:i + 3]
                    sh.write(b[0], b[1],
                                 b[2])
            except IndexError:
                logging.exception(
                        "File data is lost, incoming data cannot be triples, illegal")
            finally:
                    wb.save(self._excelpath)
        else:
            pass


class config_io(object):
    def __init__(self, defaultconfigpath="config.ini", **kw):
        self._configpath = defaultconfigpath
        self.confdata = {}

    @property
    def configpath(self):
        return self._configpath

    @configpath.setter
    def configpath(self, value):
        self._configpath = value
# 读取ini返回一个字典

    def ReadConfigData(self):
        try:
            conf = configparser.ConfigParser()
            conf.read(self._configpath, encoding='utf-8-sig')
            hander = conf.sections()
            if "Config" in hander:
                k_v = conf.items('Config')
                self.confdata = dict(k_v)
                return self.confdata
            else:
                k_v = conf.items(hander[0])
                self.confdata = dict(k_v)
                return self.confdata
        except:
            self.confdata = {}
            return self.confdata

    def WriteConfigData(self, outconfdata={}, **kw):
        section = "Config"
        if outconfdata == self.confdata:
            logging.info("Configuration files are not changed to write")
        else:
            self.confdata = outconfdata
            conf = configparser.ConfigParser()
            try:
                conf.add_section(section)
                for key, value in self.confdata.items():
                    conf.set(section, str(key), str(value))
                    try:
                        with open(self._configpath, 'w') as fw:
                            conf.write(fw)
                    except IOError:
                        logging.info(
                            "The file path does not exist and cannot be saved")

            except:
                logging(
                    "File content exception incorrect writing erro")


if __name__ == "__main__":
    c = config_io()
    e = Excel(path="report.xls")
    print(c.ReadConfigData())
    print(e.ReadEXcleData())


    # print(e.ReadEXcleData())
    c.configpath = "config1.ini"
    c.WriteConfigData(outconfdata={"f": 2221})
    e.WriteEXcleData(listdata=[1, 3, 12])
    print(c.ReadConfigData())
