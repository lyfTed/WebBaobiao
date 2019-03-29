# _*_ coding: utf-8 _*_
# filename: __init__.py
from flask import Blueprint

_admin = Blueprint('admin', __name__)
# 蓝本名为“_admin”, 这里为了避免名字与系统变量名重叠，在变量名前添加下划线区分， Blueprint函数里的admin参数表示路由名，
# 例如（url_for('admin.index'), 不能调用为url_for('_admin.index'）)

from . import view  # 这一步很重要， 作用是把所有的视图都写入到蓝本中去