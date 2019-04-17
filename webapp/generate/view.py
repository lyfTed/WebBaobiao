# _*_ coding: utf-8 _*_
from . import _generate

from flask import render_template, request, send_from_directory, abort, flash, redirect, send_file
from flask_login import login_required, current_user
import os
import xlrd, xlwt
from xlutils.copy import copy
from openpyxl import Workbook, load_workbook
from .form import GenerateForm, excels
from .. import conn
from pypinyin import lazy_pinyin

pardir = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
basedir = os.path.abspath(os.path.dirname(__file__))

FILE_TO_DOWNLOAD = {'1': '资金期限表', '2': 'G25', '3': 'Q02'}


@_generate.route('/')
@login_required
def generate():
    form = GenerateForm()
    generatelist = request.values.getlist('excels')
    generatedate = request.values.get('generatedate')
    if generatelist == []:
        return render_template('generate.html', form=form)
    else:
        filedir = os.path.join(pardir, 'static', 'upload')
        # print(filedir)
        # print(basedir)
        generatedate = generatedate.split('-')[0] + '_' + generatedate.split('-')[1]
        for generatefile in generatelist:
            filetogenerate_chinese = FILE_TO_DOWNLOAD[generatefile]
            generateFile(filetogenerate_chinese, generatedate)
        return render_template('generate.html', form=form)


def generateFile(filetogenerate_chinese, generatedate):
    conn.ping(reconnect=True)
    cursor = conn.cursor()
    filetogenerate = ''.join(lazy_pinyin(filetogenerate_chinese))
    # 创建新表
    sql = 'create table if not exists ' + filetogenerate + '_' + generatedate + \
          ' select * from ' + filetogenerate + ';'
    cursor.execute(sql)
    # 新增一列解释
    sql = 'alter table ' + filetogenerate + '_' + generatedate + \
          ' add contentexplain VARCHAR(500);'
    print(sql)
    cursor.execute(sql)
    sql = 'update ' + filetogenerate + '_' + generatedate + \
          ' set contentexplain=content;'
    cursor.execute(sql)
    # 从模板拿需要填写的格子
    sql = 'select distinct position, content from ' + filetogenerate + ' where editable=True;'
    cursor.execute(sql)
    conn.commit()
    sqlresult = cursor.fetchall()
    for i in range(len(sqlresult)):
        # 获取哪个格子
        position = sqlresult[i][0]
        print(position)
        userlist = []
        userset = {}
        alertlist = []
        # 获取用户和内容
        content_list = sqlresult[i][1].lstrip('|').split('|')
        for content in content_list:
            userandvalue = content.split('：')
            if len(userandvalue) == 1:
                userandvalue = content.split(':')
            user = ''.join(lazy_pinyin(userandvalue[0]))
            if len(userandvalue) > 1:
                value = userandvalue[1]
            else:
                value = None
            if user not in userlist:
                userlist.append(user)
                userset[user] = []
            userset[user].append((position, value))
        positionvaluelist = []
        for user in userlist:
            for i in range(len(userset[user])):
                position = userset[user][i][0]
                # value = userset[auth][i][0]
                try:
                    sql = 'select value from ' + user + \
                        ' where baobiao="' + filetogenerate_chinese + '" and position="' + position + '";'
                    # print(sql)
                    cursor.execute(sql)
                    result = cursor.fetchall()
                    value = result[0][0]
                    positionvaluelist.append(value)
                    if value is None:
                        alertlist.append(user)
                except:
                    alertlist.append(user)
                finally:
                    pass
        positionvalue = sum([x if x is not None else 0 for x in positionvaluelist])
        print(alertlist)
        sql = 'update ' + filetogenerate + '_' + generatedate + ' set content="' + str(positionvalue) + \
              '" where position="' + str(position) + '";'
        print(sql)
        cursor.execute(sql)
    conn.commit()
    ######################
    # 生成excel
    # 计算行数列数
    wb = load_workbook(pardir + '/static/upload/' + filetogenerate_chinese + '/' + filetogenerate_chinese + '.xlsx')
    sh = wb.active
    sql = 'select distinct position, content from ' + filetogenerate + '_' + generatedate + ' where editable=TRUE;'
    cursor.execute(sql)
    conn.commit()
    sqlresult = cursor.fetchall()
    for x in sqlresult:
        sh[x[0]] = float(x[1])
    # 把带公式计算的格子填入公式，自动计算
    sql = 'select distinct position, content from ' + filetogenerate + ' where content like "=%";'
    cursor.execute(sql)
    conn.commit()
    sqlresult = cursor.fetchall()
    for x in sqlresult:
        print(x[1])
        sh[x[0]] = str(x[1])
    filedir = os.path.join(pardir, 'static', 'Generate', filetogenerate_chinese)
    if not os.path.exists(filedir):
        os.mkdir(filedir)
    wb.save(filedir + '/' + filetogenerate_chinese + '_' + generatedate + '.xlsx')
