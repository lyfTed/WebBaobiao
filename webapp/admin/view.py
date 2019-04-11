from flask import render_template
from . import _admin
from flask_admin import Admin, AdminIndexView, BaseView, expose
from flask_login import current_user
from flask_admin.contrib.sqla import ModelView

# 定制一个页面，用自己的模板（仅仅是页面）然后再程序中加入代码 admin.add_view(MyNews(name=u'发表新闻'))
class MyAdminView(BaseView):
    def is_accessible(self):
        if current_user.is_authenticated and current_user.username.lower() == 'admin':
            return True
        return False

    @expose('/', methods=['GET', 'POST'])
    def index(self):
        return self.render('myadmin.html')


# 管理数据库表，设置表显示哪些字段
class MyView(ModelView):
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
        super(MyView, self).__init__(table, session, **kwargs)



