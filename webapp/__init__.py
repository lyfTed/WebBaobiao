# _*_ coding: utf-8 _*_
# filename: __init.py
import os
from flask import Flask
from flask_bootstrap import Bootstrap
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from flask_mail import Mail
from flask_moment import Moment
from flask_uploads import UploadSet, configure_uploads, DOCUMENTS
from config import config
import pymysql

bootstrap = Bootstrap()
db = SQLAlchemy()
mail = Mail()
moment = Moment()
login_manager = LoginManager()
# 会话保护等级
login_manager.session_protection = 'basic'
# 设置登录页面端点
login_manager.login_view = 'user.login'
excels = UploadSet('Excels', DOCUMENTS)
conn = pymysql.connect(host='localhost', user='root', passwd='lyfTeddy3.14', db='baobiaodb', use_unicode=True,
                           charset='utf8')

def create_app(config_name):
    # __name__ 决定应用根目录
    app = Flask(__name__)
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

    # 注册蓝本
    from .main import _main as main_blueprint
    app.register_blueprint(main_blueprint, url_prefix='/main')
    from .admin import _admin as admin_blueprint
    app.register_blueprint(admin_blueprint, url_prefix='/admin')
    from .user import _user as user_blueprint
    app.register_blueprint(user_blueprint, url_prefix='/user')
    from .api import _api as api_blueprint
    app.register_blueprint(api_blueprint, url_prefix='/api')
    from .generate import _generate as generate_blueprint
    app.register_blueprint(generate_blueprint, url_prefix='/generate')
    from .baobiao import _baobiao as baobiao_blueprint
    app.register_blueprint(baobiao_blueprint, url_prefix='/baobiao')
    return app
