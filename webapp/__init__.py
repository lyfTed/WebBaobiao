# _*_ coding: utf-8 _*_
# filename: __init.py
import os
from flask import Flask
from flask_admin import Admin, AdminIndexView
from flask_bootstrap import Bootstrap
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from flask_mail import Mail
from flask_moment import Moment
from flask_uploads import UploadSet, configure_uploads, DOCUMENTS
from webapp.admin.view import MyAdminView, MyUserView, MyBaobiaoView, MyBaseView
from flask_admin.contrib.fileadmin import FileAdmin
from flask import Blueprint


from config import config
import pymysql

staticfilepath = os.path.join(os.path.dirname(__file__), 'Files')
bootstrap = Bootstrap()
db = SQLAlchemy()
flask_admin = Admin(index_view=AdminIndexView(name="后台管理", template='myadmin.html', url='/admin'))
mail = Mail()
moment = Moment()
login_manager = LoginManager()
# 会话保护等级
login_manager.session_protection = 'basic'
# 设置登录页面端点
login_manager.login_view = 'auth.login'
excels = UploadSet('Excels', DOCUMENTS)
conn = pymysql.connect(host='localhost', user='root', passwd='lyfTeddy3.14', db='baobiaodb', use_unicode=True,
                           charset='utf8')
FILE_TO_SET = {'1': '资金期限表', '2': 'G25', '3': 'Q02'}
from webapp.models import db, User, BaobiaoToSet

def create_app(config_name):
    # __name__ 决定应用根目录
    # app = Flask(__name__, static_url_path='', static_folder='')
    app = Flask(__name__)
    # 实例化flask_admin
    # 初始化app配置
    app.config.from_object(config[config_name])
    config[config_name].init_app(app)
    # 扩展应用初始化
    bootstrap.init_app(app)
    mail.init_app(app)
    moment.init_app(app)
    db.init_app(app)
    login_manager.init_app(app)
    configure_uploads(app, excels)
    flask_admin.init_app(app)
    flask_admin.add_view(MyBaseView(name='报表主页', endpoint='index'))
    flask_admin.add_view(MyUserView(User, db.session, name='用户管理'))
    flask_admin.add_view(MyBaobiaoView(BaobiaoToSet, db.session, name='报表名管理'))
    flask_admin.add_view(FileAdmin(staticfilepath, name='报表模板与生成文件管理'))


    # 注册蓝本
    from .main import _main as main_blueprint
    app.register_blueprint(main_blueprint, url_prefix='/main')
    from .auth import _auth as auth_blueprint
    app.register_blueprint(auth_blueprint, url_prefix='/auth')
    from .api import _api as api_blueprint
    app.register_blueprint(api_blueprint, url_prefix='/api')
    from .baobiao import _baobiao as baobiao_blueprint
    app.register_blueprint(baobiao_blueprint, url_prefix='/baobiao')
    from .baobiao import _baobiao as baobiao_blueprint
    return app
