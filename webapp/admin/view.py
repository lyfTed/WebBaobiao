from flask import render_template
from . import _admin
from flask_admin import Admin, AdminIndexView
from flask_admin.contrib.sqla import ModelView


@_admin.route('/')
def index():
    return render_template('admin.html')
