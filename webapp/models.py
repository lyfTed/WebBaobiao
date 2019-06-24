# -*- coding: utf-8 -*-
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import UserMixin

from . import db, login_manager


class User(UserMixin,  db.Model):
    __tablename__ = 'users'
    id = db.Column(db.String(6), primary_key=True)
    email = db.Column(db.String(64), unique=True, index=True)
    username = db.Column(db.String(64), unique=True, index=True)
    dept = db.Column(db.String(64))
    password_hash = db.Column(db.String(128))

    @property
    def password(self):
        raise AttributeError('password is not a readable attribute')
    ###试图读取password属性，返回错误。

    @password.setter
    ###计算密码散列值的函数通过名为password的只写属性实现。
    def password(self, password):
        self.password_hash = generate_password_hash(password)
    ###设置password的属性时，赋值方法会调用Werkzeug提供的generate_password_hash()函数，并把结果赋值给password_hash字段。

    def verify_password(self, password):
        return check_password_hash(self.password_hash, password)
    ###verify_password方法接受一个参数（密码），将其传给Werkzeug提供的check_password_hash()函数，和储存在User模型中的密码散列值进行对比。密码正确返回True，密码错误返回False。

    def reset_password(self):
        self.password_hash = generate_password_hash("123456")

    def change_password(self, password):
        self.password_hash = generate_password_hash(password)

    def __repr__(self):
        return '<User %r>' % self.username


# 加载用户的回调函数
@login_manager.user_loader
def load_user(user_id):
    # print('call load_user')
    return User.query.get(user_id)


class Report(db.Model):
    __tablename__ = 'reports'
    id = db.Column(db.String(6), primary_key=True)
    name = db.Column(db.String(64), unique=True, index=True)
    content = db.Column(db.String(5000))

    def __repr__(self):
        return '<Report %r>' % self.name


class BaobiaoToSet(db.Model):
    __tablename__ = 'baobiaotoset'
    id = db.Column(db.Integer, primary_key=True)
    file = db.Column(db.String(100), unique=True, index=True)
    freq = db.Column(db.String(10))

    def __repr__(self):
        return '%r,%r' % (self.file, self.freq)
