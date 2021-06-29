#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2021/6/18 16:28
# @Author  : Zhang Shanxiu
import os
import re
import sys


def main():
    path = r'E:\Projects\PyCharm\08Automation\jpgs'
    old_names = os.listdir(path)
    print(old_names)
    for old_name in old_names:
        jpgs = os.listdir(path + '/' + old_name)
        print(jpgs)
        pre = old_name.split('_')[0]
        for jpg in jpgs:
            old_one = path + '/' + old_name + '/' + jpg
            new_one = path + '/' + pre + '.JPG'
            os.rename(old_one, new_one)
        # print(jpgs)
    # print(old_names)


if __name__ == '__main__':
    main()