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
from .form import UploadForm, excels, DownloadForm, SplitForm
from .. import conn
import pandas as pd
from openpyxl import load_workbook

pardir = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
basedir = os.path.abspath(os.path.dirname(__file__))
ALLOWED_EXTENSIONS = set(['xlsx'])
FILE_TO_DOWNLOAD = {'1': '资金期限表', '2': 'G25', '3': 'Q02'}


# 用于判断文件后缀
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# # 上传文件
### 运用wtf quick form，处理不了多文件上传。
# @_api.route('/upload/', methods=['Get', 'POST'], strict_slashes=False)
# @login_required
# def upload():
#     form = UploadForm()
#     if form.validate_on_submit():
#         # filename = excels.save(request.files.get('excels'))
#         # print(filename)
#         filename = secure_filename(form.excels.data.filename)
#         excels.save(form.excels.data, name=form.excels.data.filename)
#     else:
#         filename = None
#     return render_template('upload.html', form=form, filename=filename)


@_api.route('/upload/')
@login_required
def upload():
    return render_template('upload.html')


@_api.route('/upload_file/', methods=['POST'])
@login_required
def upload_file():
    if request.method == 'POST':
        # get current user
        username = current_user.username
        # print(username)
        # check if the post request has the file part
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
                filedir = os.path.join(basedir, 'upload', rawfilename[0])
                if not os.path.exists(filedir):
                    os.mkdir(filedir)
                file.save(os.path.join(filedir, filename))
                importintodb(os.path.join(filedir, filename), re.split('[_.]', filename)[0])
        flash('File(s) successfully uploaded')
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
          '(tablename VARCHAR(100), position VARCHAR(100), content VARCHAR(500), editable Boolean);'
    cursor.execute(sql)
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
    conn.close()


@_api.route('/download/', methods=['GET', 'POST'])
@login_required
def download():
    form = DownloadForm()
    downloadlist = request.values.getlist('excels')
    if downloadlist == []:
        return render_template('download.html', form=form)
    else:
        generatedate = request.values.get('generatedate')
        generatedate = generatedate.split('-')[0] + '_' + generatedate.split('-')[1]
        filedir = os.path.join(pardir, 'generate')
        if os.path.exists(filedir+'/Baobiao.zip'):
            os.remove(filedir+'/Baobiao.zip')
        zipf = zipfile.ZipFile(filedir+'/Baobiao.zip', 'w', zipfile.ZIP_DEFLATED)
        for filetodownload in downloadlist:
            filefolder = FILE_TO_DOWNLOAD[filetodownload]
            filename = filefolder + '_' + generatedate + '.xlsx'
            print(filefolder)
            print(filename)
            if os.path.isfile(os.path.join(filedir, filefolder, filename)):
                print(12345)
                zipf.write(filedir + '/' + filefolder + '/' + filename, filename)
        zipf.close()
        return send_file(filedir+'\\'+'Baobiao.zip', mimetype='zip', attachment_filename='Baobiao.zip', as_attachment=True)


