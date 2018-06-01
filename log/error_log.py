# -*- coding: utf-8 -*-
# @Time     : 2018/6/1 10:56
# @Author   : xiao xian.zuo
# @Email    : 1980179070@163.com
# @File     : error_log.py
# @Software : PyCharm


import logging
import os

def error_log(error):
    cwd_path = os.getcwd()
    parert_path = os.path.abspath(os.path.dirname(cwd_path) + os.path.sep + '.')
    filename = os.path.join(parert_path, 'log')
    filename = os.path.join(filename, 'error.log')

    # 返回一个指定名称的日志记录器！
    logger = logging.getLogger("AppName")

    # 指定 logger的输出格式
    formatter = logging.Formatter('%(asctime)s %(levelname)-8s:  %(message)s')

    # 文件日志
    file_handler = logging.FileHandler(filename)

    # 通过setFormatter 指定输出格式
    file_handler.setFormatter(formatter)

    # 为logger添加日志处理器
    logger.addHandler(file_handler)

    # reeor
    logger.error(error)
