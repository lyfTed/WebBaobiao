# _*_ coding: utf-8 _*_
# filename: manage.py
import os
from webapp import db, create_app
from flask_script import Manager, Shell
from flask_migrate import Migrate, MigrateCommand
from webapp.models import *

app = create_app(os.environ.get('FLASKY_CONFIG') or 'default')  # 实例化app
manager = Manager(app)  # 实例化manager
migrate = Migrate(app, db)  # 实例化migrate（数据库调试命令）


def make_shell_context():
    return dict(app=app, db=db)


manager.add_command('shell', Shell(make_context=make_shell_context))  # 添加shell参数命令（python manage.py shell）
manager.add_command('db', MigrateCommand)  # 添加db参数命令（python manage.py db init）


if __name__ == '__main__':
    manager.run()
