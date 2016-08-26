#!/usr/bin/env python
# -*- coding:utf8 -*-

import xlrd
def readExcel(file):
    cases = []
    try:
        excel = xlrd.open_workbook(file)
        sheet =  excel.sheet_by_index(0)
    except Exception,e:
        print '测试用例不存在',e
    else:
        rows =  sheet.nrows
        for row in range(rows):
            if row !=0:
                cases.append(sheet.row_values(row))
            return cases

def interfectTest(Cases):
    for case in cases:
        project = case[0]
        case_id = case[1]
        interfaceName = case[2]
        desc = case[3]
        method = case[4]
        url = case[5]
        request = case[6]
        res_chect = case[7]

        print cases
        print project,case_id,interfaceName

readExcel('test_case.xlsx')