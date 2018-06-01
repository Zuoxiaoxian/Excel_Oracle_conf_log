# -*- coding: utf-8 -*-
# 作者      ：xiaoxianzuo.zuo
# QQ        ：1980179070
# 文件名    ： 1.py
# 新建时间   ：2018/5/31/031 21:36


from excel_01 import excel_01
print("121212")
zong_lists = excel_01.excel_01()
for zong_list in zong_lists:
    for zong in zong_list:
        print("zong: ", zong)
