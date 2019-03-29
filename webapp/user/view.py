from flask import render_template, redirect, request, url_for, flash
from flask_login import login_user, logout_user, login_required, current_user
###从Flask_login导入login_user, logout_user, login_required 函数
from . import _user
###从本级目录中导入_user蓝本
from ..models import User
###从上级目录中的models.py导入User模型
from .form import LoginForm, RegistrationForm
#从本级目录中的forms.py中导入LoginForm类
from .. import db


@_user.route('/login/', methods=['GET', 'POST'])
###当请求为GET时，直接渲染模板，当请求是POST提交时，验证表格数据，然后尝试登入用户。
def login():
    form = LoginForm()
    if form.validate_on_submit():###表格中填入了数据，执行下面操作
        user = User.query.filter_by(id=form.id.data).first()
        ###视图函数使用表单中填写的email加载用户
        if user is not None and user.verify_password(form.password.data):
        ###如果user不是空的，而且验证表格中的密码正确，执行下面的语句，调用Flask_Login中的login_user（）函数，在用户会话中把用户标记为登录。
        ###否则直接执行flash消息和跳转到新表格中。
            login_user(user, form.remember_me.data)
            ###login_user函数的参数是要登录的用户，以及可选的‘记住我’布尔值。
            return redirect(request.args.get('next') or url_for('main.index'))
            ###用户访问未授权的ＵＲＬ时会显示登录表单，Flask-Login会把原地址保存在查询字符串的next参数中，这个参数可从request.args字典中读取。如果查询字符串中没有next参数，则重定向到首页。
        flash('Invalid username or password.')
    return render_template('login.html', form=form)


@_user.route('/logout/')
###退出路由
@login_required
###用户要求已经登录
def logout():
    logout_user()
    ###登出用户，这个视图函数调用logout_user()函数，删除并重设用户会话。
    flash('You have been logged out.')
    ###显示flash消息
    return redirect(url_for('main.index'))
    ###重定向到首页

@_user.route('/register/', methods=['GET', 'POST'])
def register():
    form = RegistrationForm()
    if form.validate_on_submit():
        user = User(id=form.id.data,
                    email=form.email.data,
                    username=form.username.data,
                    password=form.password.data)
        db.session.add(user)
        flash('You can now login.')
        return redirect(url_for('user.login'))
    return render_template('register.html', form=form)


