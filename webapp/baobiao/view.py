# _*_ coding: utf-8 _*_
from . import _baobiao

from flask import render_template, request, send_from_directory, abort, flash, redirect, send_file
from flask_login import login_required, current_user
import os
from .form import BaobiaoForm, TianxieForm, QueryForm, excels
from pypinyin import lazy_pinyin
import pyexcel
from openpyxl import Workbook, load_workbook
from  win32com.client import Dispatch
from .form import GenerateForm, excels
from .. import conn
from ..models import BaobiaoToSet
import pythoncom

basedir = os.path.abspath(os.path.dirname(__file__))
pardir = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))


# 获取数据库中报表名
def get_baobiao_name():
    result = BaobiaoToSet.query.order_by(BaobiaoToSet.id).all()
    FILE_TO_SET = {}
    for i in range(len(result)):
        FILE_TO_SET[str(i+1)] = str(result[i])
    # print(FILE_TO_SET)
    return FILE_TO_SET


@_baobiao.route('/split/')
@login_required
def split():
    form = BaobiaoForm()
    form.excels.choices = [(a.id, a.file) for a in BaobiaoToSet.query.all()]
    conn.ping(reconnect=True)
    cursor = conn.cursor()
    filetosetlist = request.values.getlist("excels")
    # print(filetosetlist)
    if filetosetlist == []:
        return render_template("baobiao.html", form=form)
    else:
        for file in filetosetlist:
            baobiao_split(cursor, file)
            conn.commit()
        flash('Baobiao(s) Successfully Splitted')
        conn.close()
        return render_template('baobiao.html', form=form)


def baobiao_split(cursor, file):
    FILE_TO_SET = get_baobiao_name()
    filetoset_chinese = FILE_TO_SET[file]
    filetoset = ''.join(lazy_pinyin(FILE_TO_SET[file])).lower()
    # print(filetoset)
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
            # print(userset[user][i])
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


@_baobiao.route('/fill/', methods=['GET', 'POST'])
@login_required
def fill():
    form = TianxieForm()
    username = current_user.username.lower()
    conn.ping(reconnect=True)
    cursor = conn.cursor()
    try:
        sql = 'select distinct * from ' + username + ';'
        cursor.execute(sql)
        sqlresult = cursor.fetchall()
    except:
        sqlresult = None
    if request.method == 'GET' and sqlresult is not None:
        return render_template("baobiao_tianxie.html", form=form, sqlresult=sqlresult)
    elif request.method == 'GET' and sqlresult is None:
        return render_template("baobiao_tianxie.html", form=form)
    elif request.method == 'POST':
        tianxie = request.form.getlist("values")
        try:
            for i in range(len(tianxie)):
                baobiao = sqlresult[i][0]
                gezi = sqlresult[i][1]
                value = str(tianxie[i])
                sql = 'update ' + username + ' set value=' + value + ' where baobiao="' + \
                        baobiao + '" and position="' + gezi + '";'
                # print(sql)
                if value != '':
                    cursor.execute(sql)
            conn.commit()
            conn.close()
            flash("数据提交成功")
            return redirect('/baobiao/fill/')
        except:
            flash("有个格子不是数字格式，只能填写数字格式")
            return redirect('/baobiao/fill/')
        finally:
            return redirect('/baobiao/fill/')


@_baobiao.route('/generate')
@login_required
def generate():
    form = GenerateForm()
    form.excels.choices = [(a.id, a.file) for a in BaobiaoToSet.query.all()]
    FILE_TO_SET = get_baobiao_name()
    generatelist = request.values.getlist('excels')
    generatedate = request.values.get('generatedate')
    filedir = os.path.join(pardir, 'Files', 'generate')
    if not os.path.exists(filedir):
        os.mkdir(filedir)
    if generatelist == []:
        return render_template('generate.html', form=form)
    else:
        # print(generatedate)
        generatedate = generatedate.split('-')[0] + '_' + generatedate.split('-')[1]
        allcomplete = True
        for generatefile in generatelist:
            filetogenerate_chinese = FILE_TO_SET[generatefile]
            alert = generateFile(filetogenerate_chinese, generatedate)
            if len(alert) != 0:
                allcomplete = False
                alertmsg = filetogenerate_chinese + ': ' + ','.join(alert)
                flash('以下用户还未完成对应报表：' + alertmsg)
        if allcomplete:
            flash('报表生成成功')
        else:
            flash('报表生成成功但数据还未完整')
        return render_template('generate.html', form=form)


def generateFile(filetogenerate_chinese, generatedate):
    pythoncom.CoInitialize()
    conn.ping(reconnect=True)
    cursor = conn.cursor()
    filetogenerate = ''.join(lazy_pinyin(filetogenerate_chinese))
    tablenamenew = filetogenerate + '_' + generatedate
    # 创建新表
    sql = 'create table if not exists ' + tablenamenew + \
          '(tablename VARCHAR(100), position VARCHAR(100), content VARCHAR(500),' \
          ' editable Boolean, contentexplain VARCHAR(500), primary key (position));'
    cursor.execute(sql)
    conn.commit()
    try:
        sql = 'insert into ' + tablenamenew + ' (tablename, position, content, editable, contentexplain) ' \
              'select tablename, position, content, editable, content from ' + filetogenerate + ';'
        cursor.execute(sql)
        conn.commit()
    except:
        print('已经初始化过本表')
    finally:
        pass

    # 从模板拿需要填写的格子
    sql = 'select distinct position, content from ' + filetogenerate + ' where editable=True;'
    cursor.execute(sql)
    conn.commit()
    sqlresult = cursor.fetchall()
    # 用来提示哪些用户还未填写此张报表
    alertset = set()
    for i in range(len(sqlresult)):
        # 获取哪个格子
        position = sqlresult[i][0]
        # print(position)
        userlist = []
        userset = {}
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
            # print(user)
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
                        alertset.add(user)
                except:
                    alertset.add(user)
                finally:
                    pass
        positionvalue = sum([x if x is not None else 0 for x in positionvaluelist])
        sql = 'update ' + filetogenerate + '_' + generatedate + ' set content="' + str(positionvalue) + \
              '" where position="' + str(position) + '";'
        # print(sql)
        cursor.execute(sql)
    conn.commit()
    ######################
    # 生成excel
    # 计算行数列数
    wb = load_workbook(pardir + '/Files/upload/' + filetogenerate_chinese + '/' + filetogenerate_chinese + '.xlsx')
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
        sh[x[0]] = str(x[1])
    filedir = os.path.join(pardir, 'Files', 'Generate', filetogenerate_chinese)
    if not os.path.exists(filedir):
        os.mkdir(filedir)
    # 保存带公式的xlsx
    wb.save(filedir + '/' + filetogenerate_chinese + '_' + generatedate + '.xlsx')
    #### https://www.cnblogs.com/vhills/p/8327918.html
    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = False
    xlBook = xlApp.Workbooks.Open(filedir + '/' + filetogenerate_chinese + '_' + generatedate + '.xlsx')
    xlBook.Save()
    xlBook.Close()
    wb = load_workbook(filedir + '/' + filetogenerate_chinese + '_' + generatedate + '.xlsx', data_only=True)
    wb.save(filedir + '/' + filetogenerate_chinese + '_' + generatedate + '.xlsx')
    return alertset


@_baobiao.route('/query/', methods=['GET', 'POST'])
@login_required
def query():
    form = QueryForm()
    form.excel.choices = [(a.id, a.file) for a in BaobiaoToSet.query.all()]
    if request.method == 'GET':
        return render_template("baobiao_query.html", form=form)
    else:
        FILE_TO_SET = get_baobiao_name()
        baobiao = FILE_TO_SET[str(form.excel.data)]
        generatedate = request.values.get('generatedate')
        generatedate = generatedate.split('-')[0] + '_' + generatedate.split('-')[1]
        lastdate = request.values.get('lastdate')
        lastdate = lastdate.split('-')[0] + '_' + lastdate.split('-')[1]
        baobiao_generate = "".join(baobiao + '_' + generatedate)
        baobiao_last = "".join(baobiao + '_' + lastdate)
        filedir = os.path.join(pardir, 'Files', 'Generate', '资金期限表')
        destdir = os.path.join(pardir, 'templates')

        pyexcel.save_as(file_name=filedir + '/' + baobiao_generate + '.xlsx', dest_file_name=destdir+'/query.handsontable.html')
        pyexcel.save_as(file_name=filedir + '/' + baobiao_last + '.xlsx', dest_file_name=destdir+'/last.handsontable.html')
        result = baobiao_compare(baobiao, generatedate, lastdate)
        return render_template("baobiao_query_result.html", form=form, result=result, baobiao=baobiao)


def baobiao_compare(baobiao, generatedate, lastdate):
    baobiao_generate = "".join(lazy_pinyin(baobiao + '_' + generatedate.replace('-', '_')))
    baobiao_last = "".join(lazy_pinyin(baobiao + '_' + lastdate.replace('-', '_')))
    conn.ping(reconnect=True)
    cursor = conn.cursor()
    # this term baobiao
    sql = 'select distinct position, content, contentexplain from ' + baobiao_generate + ' where editable=True;'
    cursor.execute(sql)
    conn.commit()
    sqlresult_generate = cursor.fetchall()
    result_generate = dict((x, [z, y]) for x, y, z in sqlresult_generate)
    # print(result_generate)
    # last term baobiao
    sql = 'select distinct position, content from ' + baobiao_last + ' where editable=True;'
    cursor.execute(sql)
    conn.commit()
    sqlresult_last = cursor.fetchall()
    result_last= dict((x, y) for x, y in sqlresult_last)
    changedict = {}
    for key in result_generate:
        try:
            value_explain = result_generate.get(key)[0].split('|')
            value_explain = "，".join(value_explain).lstrip('，')
            value_generate = result_generate.get(key)[1]
            value_last = result_last.get(key)
            # print(value_generate)
            print(value_last)
            if float(value_last) != 0:
                pctchange = float(value_generate) / float(value_last) - 1
                show_result = str(round(pctchange * 100, 2)) + '%'
            else:
                pctchange = None
                show_result = 'Cannot Compare Percentage Change'
            changedict[key] = [value_explain, value_last, value_generate, show_result, pctchange]
            print(pctchange)
        except KeyError:
            print('No key in last table')
            value_generate = result_generate.get(key)
            value_last = None
            pctchange = None
            changedict[key] = [value_explain, value_last, value_generate, None, pctchange]
        finally:
            pass
    # print(changedict)
    return changedict


@_baobiao.route('/query/generate', methods=['GET', 'POST'])
@login_required
def show_generate_baobiao():
    return render_template("query.handsontable.html")


@_baobiao.route('/query/last', methods=['GET', 'POST'])
@login_required
def show_last_baobiao():
    return render_template("last.handsontable.html")
