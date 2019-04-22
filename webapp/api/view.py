# _*_ coding: utf-8 _*_
from . import _api

from werkzeug.utils import secure_filename
from flask import render_template, request, send_from_directory, abort, flash, redirect, send_file
from flask_login import login_required, current_user
import os
import re
import zipfile
import xlrd
import pymysql
from pypinyin import lazy_pinyin
from .form import UploadForm, excels, DownloadForm
from .. import conn
import pandas as pd
from openpyxl import load_workbook
from ..models import BaobiaoToSet

pardir = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
basedir = os.path.abspath(os.path.dirname(__file__))
ALLOWED_EXTENSIONS = set(['xlsx'])


# 用于判断文件后缀
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# 获取数据库中报表名
def get_baobiao_name():
    result = BaobiaoToSet.query.order_by(BaobiaoToSet.id).all()
    FILE_TO_SET = {}
    for i in range(len(result)):
        FILE_TO_SET[str(i+1)] = str(result[i])
    # print(FILE_TO_SET)
    return FILE_TO_SET


@_api.route('/upload/')
@login_required
def upload():
    return render_template('upload.html')


@_api.route('/upload_file/', methods=['POST'])
@login_required
def upload_file():
    if request.method == 'POST':
        # get current auth
        username = current_user.username
        # print(username)
        # check if the post request has the file part
        filedir = os.path.join(pardir, 'Files', 'upload')
        if not os.path.exists(filedir):
            os.mkdir(filedir)
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No file selected for uploading')
            return redirect(request.url)
        files = request.files.getlist("file")
        for file in files:
            if file and allowed_file(file.filename):
                rawfilename = re.split('[_.]', file.filename)
                filename = rawfilename[0] + '.' + rawfilename[-1]
                # 文件名加上用户名（不启用）
                # filename = rawfilename[0] + '_' + username + '.' + rawfilename[-1]
                filedir = os.path.join(pardir, 'Files', 'upload', rawfilename[0])
                if not os.path.exists(filedir):
                    os.mkdir(filedir)
                file.save(os.path.join(filedir, filename))
                importintodb(os.path.join(filedir, filename), re.split('[_.]', filename)[0])
        flash('File(s) Successfully Uploaded')
        return redirect('/api/upload')


def importintodb(file_to_generate, filename):
    conn.ping(reconnect=True)
    wb = xlrd.open_workbook(file_to_generate)
    sh = wb.sheet_by_index(0)
    nrows = sh.nrows  # 行数
    ncols = sh.ncols  # 列数
    title = sh.cell_value(0, 0)
    cols = [chr(i + ord('A')) for i in range(ncols)]
    rows = [str(i + 1) for i in range(nrows)]

    wb2 = load_workbook(file_to_generate)
    sheet_names = wb2.get_sheet_names()
    name = sheet_names[0]
    sheet_ranges = wb2[name]
    df = pd.DataFrame(sheet_ranges.values)
    df = df.fillna("")
    cursor = conn.cursor()
    # 创建table
    # 用第一行第一列做表明
    tablename_chinese = filename
    tablename = ''.join(lazy_pinyin(filename))
    sql = 'create table if not exists ' + tablename + \
          '(tablename VARCHAR(100), position VARCHAR(100), content VARCHAR(500), editable Boolean, ' \
          'primary key (position));'
    cursor.execute(sql)
    try:
        for i in range(nrows):
            for j in range(ncols):
                position = cols[j] + rows[i]
                content = str(df.iloc[i, j])
                editable = False
                if len(content) > 0 and content[0] == "|":
                    editable = True
                sql = 'insert into ' + tablename + ' (tablename, position, content, editable) values ("' + \
                      tablename_chinese + '","' + position + '", "' + str(content) + '", ' + str(editable) + ");"
                cursor.execute(sql)
        conn.commit()
    except:
        print('Table already exists')
    finally:
        pass
    conn.close()


@_api.route('/download/', methods=['GET', 'POST'])
@login_required
def download():
    FILE_TO_SET = get_baobiao_name()
    form = DownloadForm()
    form.excels.choices = [(a.id, a.file) for a in BaobiaoToSet.query.all()]
    downloadlist = request.values.getlist('excels')
    if downloadlist == []:
        return render_template('download.html', form=form)
    else:
        generatedate = request.values.get('generatedate')
        generatedate = generatedate.split('-')[0] + '_' + generatedate.split('-')[1]
        filedir = os.path.join(pardir, 'Files', 'Generate')
        if os.path.exists(filedir+'/Baobiao.zip'):
            os.remove(filedir+'/Baobiao.zip')
        zipf = zipfile.ZipFile(filedir+'/Baobiao.zip', 'w', zipfile.ZIP_DEFLATED)
        for filetodownload in downloadlist:
            filefolder = FILE_TO_SET[filetodownload]
            filename = filefolder + '_' + generatedate + '.xlsx'
            if os.path.isfile(os.path.join(filedir, filefolder, filename)):
                zipf.write(filedir + '/' + filefolder + '/' + filename, filename)
        zipf.close()
        return send_file(filedir+'/'+'Baobiao.zip', mimetype='zip', attachment_filename='Baobiao.zip', as_attachment=True)


