from flask_wtf import Form
###从Flask-WTF扩展导入Form基类
from wtforms import StringField, PasswordField, BooleanField, SubmitField
###从WTForms包中导入字段类
from wtforms.validators import DataRequired, Length, Email, Regexp, EqualTo
###从WTForms导入验证函数
from wtforms import ValidationError
from ..models import User


class LoginForm(Form):
    id = StringField('ID', validators=[DataRequired(), Length(1, 5)])
    # email = StringField('Email', validators=[DataRequired(), Length(1, 64), Email()])
    ###StringField构造函数中的可选参数validators指定一个有验证函数组成的列表，在接受用户提交的数据之前验证数据。
    ###电子邮件字段用到了WTForms提供的Length（）和Email（）验证函数。
    password = PasswordField('Password', validators=[DataRequired()])
    ###PasswordField类表示属性为type="password"的<input>元素。
    remember_me = BooleanField('Keep me logged in')
    ###BooleanField类表示复选框。
    submit = SubmitField('Log In')


class RegistrationForm(Form):
    id = StringField('ID', validators=[DataRequired(), Length(1, 5)])
    email = StringField('Email', validators=[DataRequired(), Length(1, 64),
                                           Email()])
    username = StringField('Username', validators=[
        DataRequired(), Length(1, 64), Regexp('^[A-Za-z][A-Za-z0-9_.]*$', 0,
                                          'Usernames must have only letters, '
                                          'numbers, dots or underscores')])
    ###WTForms提供的Regexp验证函数，确保username字段只包含字母，数字，下划线和点号。这个验证函数中的正则表达式后面的两个参数分别是正则表达式的旗标和验证失败时显示的错误消息。
    password = PasswordField('Password', validators=[
        DataRequired(), EqualTo('password2', message='Passwords must match.')])
        ###EqualTo验证函数可以验证两个密码字段中的值是否一致，他附属在两个密码字段上，另一个字段作为参数传入。
    password2 = PasswordField('Confirm password', validators=[DataRequired()])
    dept = StringField('Department',  validators=[DataRequired(), Length(1, 64)])
    submit = SubmitField('Register')

    def validate_email(self, field):
        if User.query.filter_by(email=field.data).first():
            raise ValidationError('Email already registered.')

    def validate_username(self, field):
        if User.query.filter_by(username=field.data).first():
            raise ValidationError('Username already in use.')
    ###表单类中定义了以validate_开头且后面跟着字段名的方法，这种方法就和常规的验证函数一起调用。

