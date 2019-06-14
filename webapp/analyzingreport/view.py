# _*_ coding: utf-8 _*_
from . import _analyzingreport

from werkzeug.utils import secure_filename
from flask import render_template, request, send_from_directory, abort, flash, redirect, send_file
from flask_login import login_required, current_user
import os
import re
import zipfile
import xlrd
import pymysql
from pypinyin import lazy_pinyin
from .form import UploadForm, excels, QueryForm
from .. import conn
import pandas as pd
from openpyxl import load_workbook
from ..models import BaobiaoToSet

pardir = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
basedir = os.path.abspath(os.path.dirname(__file__))
ALLOWED_EXTENSIONS = set(['xlsx', 'xls'])


# 用于判断文件后缀
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# 获取数据库中报表名
def get_baobiao_name():
    result = BaobiaoToSet.query.order_by(BaobiaoToSet.id).all()
    FILE_TO_SET = {}
    for i in range(len(result)):
        rs = str(result[i]).split(',')
        file = str(rs[0].strip('"').strip("'"))
        freq = str(rs[1].strip('"').strip("'"))
        FILE_TO_SET[str(i+1)] = file
    return FILE_TO_SET


def get_baobiao_freq():
    result = BaobiaoToSet.query.order_by(BaobiaoToSet.id).all()
    FREQ_OF_FILE = {}
    for i in range(len(result)):
        rs = str(result[i]).split(',')
        file = str(rs[0].strip('"').strip("'"))
        freq = str(rs[1].strip('"').strip("'"))
        FREQ_OF_FILE[file] = freq
    return FREQ_OF_FILE


@_analyzingreport.route('/', methods=['GET', 'POST'])
@login_required
def main():
    form = UploadForm()
    kemu = form.kemu
    institution = form.institution
    currency = form.currency
    date = form.date
    if request.method == 'GET':
        return render_template('analyzing_report.html', form=form)
    if request.method == 'POST':
        filedir = os.path.join(pardir, 'Files', 'upload', 'analyzingreport')
        if not os.path.exists(filedir):
            os.mkdir(filedir)
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No file selected for uploading')
            return redirect(request.url)
        file = request.files.getlist("file")[0]
        print(file)
        if file and allowed_file(file.filename):
            filename = re.split('[_.]', file.filename)[0] + '.xlsx'
            print(filename)
            file.save(os.path.join(filedir, filename))
        flash('科目表上传成功')
        return redirect('/analyzingreport')







