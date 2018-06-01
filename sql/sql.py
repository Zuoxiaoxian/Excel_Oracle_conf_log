# -*- coding: utf-8 -*-
# @Time     : 2018/5/25 10:13
# @Author   : xiao xian.zuo
# @Email    : 1980179070@163.com
# @File     : sql.py
# @Software : PyCharm


import cx_Oracle
import time
import os
import config

from excel_01.excel_01 import ParseSheetZxx
from log.error_log import error_log
from sql_sql import select
sql = select()
print(sql)

cwd_path = os.getcwd()
parert_path = os.path.abspath(os.path.dirname(cwd_path) + os.path.sep + '.')
excel_path = os.path.join(parert_path, 'excel_workerbook')
conf_path = os.path.join(parert_path, 'config.conf')
host, port, user, pass_, db, proc = config.start(conf_path)


class Orcale(object):
    def __init__(self, host, port, user, passwd, db):
        self.host = host
        self.port = port
        self.user = user
        self.passwd = passwd
        self.db = db
        self.dsn = cx_Oracle.makedsn(
            self.host,
            self.port,
            self.db,
        )

    def ping_server(self):
        try:
            self.conn = cx_Oracle.connect(
                self.user,
                self.passwd,
                self.dsn,
            )
            print('已连接')
        except Exception as e:
            print("Error______________", e)
            time.sleep(2)
            error_log(e)
            return self.ping_server()

    def select(self, sql):
        try:
            self.ping_server()
            cursor = self.conn.cursor()
            cursor.execute(sql)
            result = cursor.fetchall()
            return result
            # print result
        except Exception as e:
            error_log(e)
            print(e)

    def execut(self, sql, params=''):
        try:
            self.ping_server()
            cursor = self.conn.cursor()
            if params:
                cursor.execute(sql, params)
                result = cursor.fetchall()
            else:
                cursor.execute(sql)
                result = cursor.fetchall()
            self.conn.commit()
            # print result
            return True
        except Exception as e:
            print(e)
            error_log(e)
        finally:
            try:
                cursor.close()
                self.conn.close()
            except Exception:
                pass
    # 需要调用存储过程！

    def call_proc(self, proc, in_param, var_num=0, num=0):
        '''
        :param proc: 存储过程名，str
        :param in_param:
        :param var_num:
        :param num:
        :return:
        '''
        try:
            self.ping_server()
            cursor = self.conn.cursor()
            # 声明变量,主要是 out 参数的声明
            # rtnId = [cursor.var(cx_Oracle.NUMBER)
            #          for v in xrange(num)]
            # rtnMsg = [cursor.var(cx_Oracle.STRING)
            #           for v in range(var_num)]
            I_ExcelID = cursor.var(cx_Oracle.NUMBER)
            I_ExcelID.setvalue(0, 2)
            print(I_ExcelID, "*********************")
            param = (I_ExcelID, ) + in_param
            print(param, "*********************")
            cursor.callproc(proc, param)
            print("I_ExcelID.getvalue(),----------", I_ExcelID.getvalue())
            return I_ExcelID.getvalue()
        except cx_Oracle.Error as exc:
            # error, =  exc.args
            error_log(exc)
            print("错误!", exc)
            return None
        finally:
            try:
                cursor.close()
                self.conn.close()
            except Exception:
                pass


def start(host, port, user, pass_, db, proc):
    orcale = Orcale(host, port, user, pass_, db)
    '''
    orcale = Orcale()
    # orcale.ping_server()
    sql = sql_sql.select_sql
    # print sql
    orcale.execut(sql)

    '''

    parse_sheet = ParseSheetZxx(excel_path)
    wb_sheetnames = parse_sheet.get_sheets()
    zong_lists = parse_sheet.get_merger_and_all(wb_sheetnames)
    proc = proc
    for zong_list in zong_lists:
        for zong in zong_list:
            zong = zong[1:]
            print("zong[1: ]", zong)
            in_param = tuple(zong)
            print("in_param", in_param, type(in_param))
            orcale.call_proc(proc, in_param, 0, 0)
        # time.sleep(2)

if __name__ == '__main__':
    # start(host, port, user, pass_, db, proc) #存储过程!
    o = Orcale(host, port, user, pass_, db)
    result = o.select(sql)
    print(result)
    print(len(result))
