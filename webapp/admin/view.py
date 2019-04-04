from flask import render_template
from . import _admin
from flask_admin import Admin, AdminIndexView

from flask_admin.contrib.sqla import ModelView
from flask_admin import BaseView, expose

# 定制一个页面，用自己的模板（仅仅是页面）然后再程序中加入代码 admin.add_view(MyNews(name=u'发表新闻'))
class MyAdminView(BaseView):
    @expose('/', methods=['GET', 'POST'])
    def index(self):
        return self.render('admin/myadmin.html')
    # 调用的一个子url
    @expose('/test/', methods=['GET', 'POST'])
    def test(self):
        return ".test"


# 管理数据库表，设置表显示哪些字段
class MyV1(ModelView):
    # can_create = False

    column_labels = {
        'id': u'序号',
        'name': u'名称',
    }
    column_list = ('id', 'name')

    def __init__(self, models_name, session, **kwargs):
        super(MyV1, self).__init__(models_name, session, **kwargs)
