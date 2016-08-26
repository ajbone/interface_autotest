# -*- coding:utf-8 -*-
import requests, xlrd, time, sys
#import MySQLdb
#导入需要用到的模块
from xlutils import copy
#从xlutils模块中导入copy这个函数
def readExcel(file_path):
    try:
        book = xlrd.open_workbook(file_path)#打开excel
    except Exception,e:
        #如果路径不在或者excel不正确，返回报错信息
        print '路径不在或者excel不正确',e
        return e
    else:
        sheet = book.sheet_by_index(0)#取第一个sheet页
        rows= sheet.nrows#取这个sheet页的所有行数
        case_list = []#保存每一条case
        for i in range(rows):
            if i !=0:
                #把每一条测试用例添加到case_list中
                case_list.append(sheet.row_values(i))
        #调用接口测试的函数，把存所有case的list和excel的路径传进去，因为后面还需要把返回报文和测试结果写到excel中，
        #所以需要传入excel测试用例的路径，interfaceTest函数在下面有定义
        interfaceTest(case_list,file_path)

def interfaceTest(case_list,file_path):
    res_flags = []
    #存测试结果的list
    request_urls = []
    #存请求报文的list
    responses = []
    #存返回报文的list
    for case in case_list:
        '''
        先遍历excel中每一条case的值，然后根据对应的索引取到case中每个字段的值
        '''
        try:
            '''
            这里捕捉一下异常，如果excel格式不正确的话，就返回异常
            '''
            #项目，提bug的时候可以根据项目来提
            product = case[0]
            #用例id，提bug的时候用
            case_id = case[1]
            #接口名称，也是提bug的时候用
            interface_name = case[2]
            #用例描述
            case_detail = case[3]
            #请求方式
            method = case[4]
            #请求url
            url = case[5]
            #入参
            param = case[6]
            #预期结果
            res_check = case[7]
            #测试人员
            tester = case[10]

            #print product,interface_name,case_detail,method,url
            print "###############",res_check
        except Exception,e:
            return '测试用例格式不正确！%s'%e
        if param== '':
            '''
            如果请求参数是空的话，请求报文就是url，然后把请求报文存到请求报文list中
            '''
            new_url = url#请求报文
            request_urls.append(new_url)
        else:
            '''
            如果请求参数不为空的话，请求报文就是url+?+参数，格式和下面一样
            http://127.0.0.1:8080/rest/login?oper_no=marry&id=100，然后把请求报文存到请求报文list中
            '''
            new_url = url+'?'+urlParam(param)#请求报文
            '''
            excel里面的如果有多个入参的话，参数是用;隔开，a=1;b=2这样的，请求的时候多个参数要用&连接，
            要把;替换成&，所以调用了urlParam这个函数，把参数中的;替换成&，函数在下面定义的
            '''
            request_urls.append(new_url)
        if method.upper() == 'GET':
            '''
            如果是get请求就调用requests模块的get方法，.text是获取返回报文，保存返回报文，
            把返回报文存到返回报文的list中
            '''
            print new_url
            results = requests.get(new_url).text
            print "############33333333",results
            responses.append(results)

            #results = json.loads(json.dumps(results))
            #res_check = json.loads(json.dumps(res_check))
            #print "results is %s,res_check is %s" % results,res_check
            res = readRes(results,res_check)

            print "############44444444",res_check
        else:

            results = requests.post(new_url).text
            responses.append(results)

            res = readRes(results,res_check)
        if 'pass' in res:
            '''
            判断测试结果，然后把通过或者失败插入到测试结果的list中
            '''
            res_flags.append('pass')
        else:
            res_flags.append('fail')

            #writeBug(case_id,interface_name,new_url,results,res_check)

    copy_excel(file_path,res_flags,request_urls,responses)

import json

def readRes(res,res_check):

    print "########88888888888",res  
    print "########99999999999",res_check
    res = json.loads(json.dumps(res))
    res_check = json.loads(json.dumps(res_check))

    res = res.replace('":"',"=").replace('":',"=")
    res_check = res_check.split(';')
    for s in res_check:
        if s in res:
            pass
        else:
            return  '错误，返回参数和预期结果不一致'+str(s)
    return 'pass'


def urlParam(param):
   
    return param.replace(';','&')

def copy_excel(file_path,res_flags,request_urls,responses):
   
    #打开原来的excel，获取到这个book对象
    book = xlrd.open_workbook(file_path)
    #复制一个new_book
    new_book = copy.copy(book)
    #然后获取到这个复制的excel的第一个sheet页
    sheet = new_book.get_sheet(0)
    i = 1
    for request_url,response,flag in zip(request_urls,responses,res_flags):

        sheet.write(i,8,u'%s'%request_url)
        sheet.write(i,9,u'%s'%response)
        sheet.write(i,11,u'%s'%flag)
        i+=1
    #写完之后在当前目录下(可以自己指定一个目录)保存一个以当前时间命名的测试结果，time.strftime()是格式化日期
    new_book.save('%s_测试结果.xls'%time.strftime('%Y%m%d%H%M%S'))
'''
def writeBug(bug_id,interface_name,request,response,res_check):

    bug_id = bug_id.encode('utf-8')
    interface_name = interface_name.encode('utf-8')
    res_check = res_check.encode('utf-8')
    response = response.encode('utf-8')
    request = request.encode('utf-8')

    #取当前时间，作为提bug的时间
    now = time.strftime("%Y-%m-%d %H:%M:%S")
    #bug标题用bug编号加上接口名称然后加上_结果和预期不符，可以自己随便定义要什么样的bug标题
    bug_title = bug_id + '_' + interface_name + '_结果和预期不符'
    #复现步骤就是请求报文+预期结果+返回报文
    step = '[请求报文]<br />'+request+'<br/>'+'[预期结果]<br/>'+res_check+'<br/>'+'<br/>'+'[响应报文]<br />'+'<br/>'+response
    #拼sql，这里面的项目id，创建人，严重程度，指派给谁，都在sql里面写死，使用的时候可以根据项目和接口
    # 来判断提bug的严重程度和提交给谁
    sql = "INSERT INTO `bf_bug_info` (`created_at`, `created_by`, `updated_at`, `updated_by`, `bug_status`, `assign_to`, `title`, `mail_to`, `repeat_step`, `lock_version`, `resolved_at`, `resolved_by`, `closed_at`, `closed_by`, `related_bug`, `related_case`, `related_result`, " \
          "`productmodule_id`, `modified_by`, `solution`, `duplicate_id`, `product_id`, " \
          "`reopen_count`, `priority`, `severity`) VALUES ('%s', '1', '%s', '1', 'Active', '1', '%s', '系统管理员', '%s', '1', NULL , NULL, NULL, NULL, '', '', '', NULL, " \
          "'1', NULL, NULL, '1', '0', '1', '1');"%(now,now,bug_title,step)
    #建立连接，使用MMySQLdb模块的connect方法连接mysql，传入账号、密码、数据库、端口、ip和字符集
    coon = MySQLdb.connect(user='root',passwd='123456',db='bugfree',port=3306,host='127.0.0.1',charset='utf8')
    #建立游标
    cursor = coon.cursor()
    #执行sql
    cursor.execute(sql)
    #提交
    coon.commit()
    #关闭游标
    cursor.close()
    #关闭连接
    coon.close()

'''
if __name__ == '__main__':
    try:
        filename = sys.argv[1]
    except IndexError,e:
        print 'Please enter a correct testcase! \n e.x: python gkk.py test_case.xls'
    else:
        readExcel(filename)
    print 'Done!'