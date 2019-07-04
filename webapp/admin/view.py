from flask import render_template
from . import _admin
from flask_admin import Admin, AdminIndexView, BaseView, expose
from flask_login import current_user
from flask_admin.contrib.sqla import ModelView
from flask import url_for
from .form import BaobiaoTeSetForm

# 定制一个页面，用自己的模板（仅仅是页面）然后再程序中加入代码 admin.add_view(MyNews(name=u'发表新闻'))
class MyAdminView(BaseView):
    def is_accessible(self):
        if current_user.is_authenticated and current_user.username.lower() == 'admin':
            return True
        return False

    @expose('/', methods=['GET', 'POST'])
    def index(self):
        return self.render('myadmin.html')


class MyBaseView(BaseView):
    def is_accessible(self):
        if current_user.is_authenticated and current_user.username.lower() == 'admin':
            return True
        return False

    @expose('/', methods=['GET', 'POST'])
    def index(self):
        return self.render('index.html')


# 管理数据库表，设置表显示哪些字段
class MyUserView(ModelView):
    def is_accessible(self):
        if current_user.is_authenticated and current_user.username.lower() == 'admin':
            return True
        return False

    # Disable model creation
    can_create = False
    can_edit = True
    can_delete = True
    # Override displayed fields
    column_list = ('id', 'username', 'email', 'dept')

    def __init__(self, table, session, **kwargs):
        # You can pass name and other parameters if you want to
        super(MyUserView, self).__init__(table, session, **kwargs)


class MyBaobiaoView(ModelView):
    def is_accessible(self):
        if current_user.is_authenticated and current_user.username.lower() == 'admin':
            return True
        return False

    # Disable model creation
    can_create = True
    can_edit = True
    can_delete = True
    # Override displayed fields
    column_list = ('id', 'file', 'freq', 'auditor')

    form = BaobiaoTeSetForm

    def __init__(self, table, session, **kwargs):
        # You can pass name and other parameters if you want to
        super(MyBaobiaoView, self).__init__(table, session, **kwargs)


