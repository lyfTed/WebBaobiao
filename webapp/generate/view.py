# _*_ coding: utf-8 _*_
from . import _generate

from flask import render_template, request, send_from_directory, abort, flash, redirect, send_file
from flask_login import login_required, current_user
import os
import xlrd
import xlwt
from .form import GenerateForm, excels
from .. import conn
from pypinyin import lazy_pinyin

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
        filedir = os.path.join(basedir, 'upload')
        # print(filedir)
        # print(basedir)
        for generatefile in generatelist:
            print(generatefile)
            filetogenerate_chinese = FILE_TO_DOWNLOAD[generatefile]
            # call the function to generate filetogenerate
            print(filetogenerate_chinese)
            generateFile(filetogenerate_chinese)
        return render_template('generate.html', form=form)


def generateFile(filetogenerate_chinese):
    conn.ping(reconnect=True)
    cursor = conn.cursor()
    filetogenerate = ''.join(lazy_pinyin(filetogenerate_chinese))
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
                # value = userset[user][i][0]
                try:
                    sql = 'select value from ' + user + \
                        ' where baobiao="' + filetogenerate_chinese + '" and position="' + position + '";'
                    print(sql)
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
        print(positionvaluelist)
        print(positionvalue)
        print(alertlist)
