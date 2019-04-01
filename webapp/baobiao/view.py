# _*_ coding: utf-8 _*_
from . import _baobiao

from flask import render_template, request, send_from_directory, abort, flash, redirect, send_file
from flask_login import login_required, current_user
import os
import xlrd
import xlwt
from .form import BaobiaoForm, TianxieForm, excels
from pypinyin import lazy_pinyin
import pymysql
import re
from .. import conn

FILE_TO_SET = {'1': '资金期限表', '2': 'G25', '3': 'Q02'}

basedir = os.path.abspath(os.path.dirname(__file__))


@_baobiao.route('/')
@login_required
def setbaobiao():
    form = BaobiaoForm()
    conn.ping(reconnect=True)
    cursor = conn.cursor()
    filetosetlist = request.values.getlist("excel")
    # print(filetosetlist)
    if filetosetlist == []:
        return render_template("baobiao.html", form=form)
    else:
        for file in filetosetlist:
            baobiao_split(cursor, file)
            conn.commit()
            conn.close()
            flash('Baobiao(s) successfully Splitted')
            return render_template('baobiao.html', form=form)


def baobiao_split(cursor, file):
    filetoset_chinese = FILE_TO_SET[file]
    filetoset = ''.join(lazy_pinyin(FILE_TO_SET[file]))
    sql = 'select distinct position, content from ' + filetoset + ' where editable=True;'
    cursor.execute(sql)
    conn.commit()
    sqlresult = cursor.fetchall()
    userlist = []
    userset = {}
    for i in range(len(sqlresult)):
        # 获取哪个格子
        position = sqlresult[i][0]
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
    for user in userlist:
        sql = 'create table if not exists ' + user + \
              '(baobiao VARCHAR(100), position VARCHAR(100), content VARCHAR(500), ' \
              'value DOUBLE, primary key (baobiao, position));'
        cursor.execute(sql)
        for i in range(len(userset[user])):
            print(userset[user][i])
            position = userset[user][i][0]
            value = userset[user][i][1]
            try:
                sql = 'insert into ' + user + ' (baobiao, position, content) values ("' + \
                      filetoset_chinese + '", "' + str(position) + '", "' + str(value) + '");'
                cursor.execute(sql)
            except:
                sql = 'update ' + user + ' set content="' + str(value) + '" where baobiao="' + \
                      filetoset_chinese + '" and position="' + str(position) + '";'
                cursor.execute(sql)
            else:
                pass


@_baobiao.route('/tianxie/', methods=['GET', 'POST'])
@login_required
def baobiao_tianxie():
    form = TianxieForm()
    username = current_user.username.lower()
    conn.ping(reconnect=True)
    cursor = conn.cursor()
    sql = 'select distinct * from ' + username + ';'
    cursor.execute(sql)
    sqlresult = cursor.fetchall()
    if request.method == 'GET':
        print('GET')
        return render_template("baobiao_tianxie.html", form=form, sqlresult=sqlresult)
    elif request.method == 'POST':
        tianxie = request.form.getlist("values")
        try:
            for i in range(len(tianxie)):
                baobiao = sqlresult[i][0]
                gezi = sqlresult[i][1]
                value = str(tianxie[i])
                sql = 'update ' + username + ' set value=' + value + ' where baobiao="' + \
                        baobiao + '" and position="' + gezi + '";'
                print(sql)
                if value != '':
                    cursor.execute(sql)
            conn.commit()
            conn.close()
            flash("数据提交成功")
            return redirect('/baobiao/tianxie/')
        except:
            flash("有个格子不是数字格式，只能填写数字格式")
            return redirect('/baobiao/tianxie/')
        finally:
            return redirect('/baobiao/tianxie/')
