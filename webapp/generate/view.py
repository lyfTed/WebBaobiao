# _*_ coding: utf-8 _*_
from . import _generate

from flask import render_template, request, send_from_directory, abort, flash, redirect, send_file
from flask_login import login_required, current_user
import os
import xlrd
import xlwt
from .form import GenerateForm, excels

basedir = os.path.abspath(os.path.dirname(__file__))
FILE_TO_DOWNLOAD = {'1': '资金期限表', '2': 'G25', '3': 'Q02'}


@_generate.route('/')
@login_required
def generate():
    form = GenerateForm()
    generatelist = request.values.getlist('excels')
    if generatelist == []:
        return render_template('generate.html', form=form)
    else:
        filedir = os.path.join(basedir, 'upload')
        # print(filedir)
        # print(basedir)
        for filetogenerate in generatelist:
            # call the function to generate filetogenerate
            generateFile(filetogenerate + '.xlsx')
        return render_template('generate.html', form=form)


def generateFile(filetogenerate):
    pass
