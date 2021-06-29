#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2021/6/18 9:11
# @Author  : Zhang Shanxiu
import os
from win32com.client import Dispatch, constants, gencache, DispatchEx
from main import PDFConverter
from excel import Xlsxs


def get_number(path):
    base = os.path.basename(path)
    return int(base.split('_')[0])


def main():
    input_folder = 'ppts'
    output_folder = "pdfs"
    pathname = os.path.join(os.path.abspath('.'), input_folder)
    pdfConverter = PDFConverter(pathname, output_folder)
    print(pdfConverter._filename_list)
    ppt_res = list(map(get_number, pdfConverter._filename_list))
    print(ppt_res)

    # 打印日志
    log = open('./logs/log.txt', 'w')

    xlsx = Xlsxs('./xlsxs/RCAR 2021.xlsx')
    xlsx.get_numbers()
    # print(xlsx.repeat)
    log.write("重复序号：" + xlsx.repeat + '\n')

    res, errors = list(), list()
    for number in ppt_res:
        if number not in xlsx.set:
            res.append(str(number))
    for number in xlsx.set:
        if number not in ppt_res:
            errors.append(str(number))
    ans = ", ".join(res)
    print(len(ppt_res) == len(set(ppt_res)))
    print(len(ppt_res))
    print(len(res))
    print(len(xlsx.set))
    print(ans)
    print('error', errors)
    log.write("遗漏序号：" + ans + '\n')
    log.write("错误序号：" + ', '.join(errors) + '\n')

    # 关闭日志
    log.close()


if __name__ == "__main__":
    main()