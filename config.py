# -*- coding: utf-8 -*-
# @Time     : 2018/5/30 15:38
# @Author   : xiao xian.zuo
# @Email    : 1980179070@163.com
# @File     : start.py
# @Software : PyCharm

import configparser
import os


def start(path_name):
    cp = configparser.ConfigParser()
    # cp.read('config.conf')
    cp.read(path_name)


    host = cp.get('db', 'host')
    port = cp.get('db', 'port')
    user = cp.get('db', 'user')
    pass_ = cp.get('db', 'pass_')
    db = cp.get('db', 'db')
    proc = cp.get('db', 'proc')
    return host, port, user, pass_, db, proc


if __name__ == '__main__':
    cwd_path = os.getcwd()
    path_name = os.path.join(cwd_path, 'config.conf')
    print(path_name)
    # C:\Users\xiaoxian.zuo\Desktop\py3_excel\config.conf
    start(path_name)
