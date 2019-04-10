from flask import render_template
from . import _admin
from flask_admin import Admin, AdminIndexView, BaseView, expose
from flask_admin.contrib.sqla import ModelView

# 定制一个页面，用自己的模板（仅仅是页面）然后再程序中加入代码 admin.add_view(MyNews(name=u'发表新闻'))
class MyAdminView(BaseView):
    @expose('/', methods=['GET', 'POST'])
    def index(self):
        return self.render('myadmin.html')


# 管理数据库表，设置表显示哪些字段
class MyView(ModelView):
    # Disable model creation
    can_create = False

    # Override displayed fields
    column_list = ('login', 'email')

    # def __init__(self, session, **kwargs):
    #     # You can pass name and other parameters if you want to
    #     super(MyView, self).__init__(User, session, **kwargs)


flask_admin = Admin(name="管理员页面")
