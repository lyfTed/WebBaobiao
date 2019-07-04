from flask_wtf import FlaskForm
###从Flask-WTF扩展导入Form基类
from wtforms import SubmitField, SelectField,  SelectMultipleField, StringField, FloatField
from wtforms.fields.html5 import DateField
###从WTForms包中导入字段类
from wtforms.validators import DataRequired, Length
from flask_wtf.file import FileField, FileAllowed, FileRequired
from datetime import datetime, date, timedelta
from .. import excels


class BaobiaoForm(FlaskForm):
    excels = SelectMultipleField('拆分报表（多选）', validators=[DataRequired()], coerce=int)
    submit = SubmitField(u'拆分')


class TianxieForm(FlaskForm):
    value = FloatField(u'值', validators=[DataRequired()])
    submit = SubmitField(u'提交')


class PreviewForm(FlaskForm):
    excel = SelectField('报表模板', validators=[DataRequired()], coerce=int)
    preview = SubmitField(u'预览报表模板')


class GenerateForm(FlaskForm):
    excels = SelectMultipleField('生成报表（可多选）', validators=[DataRequired()], coerce=int)
    submit = SubmitField('生成')


class QueryForm(FlaskForm):
    excel = SelectField('报表名', validators=[DataRequired()], coerce=int)
    query1 = SubmitField((date(date.today().year, date.today().month, 1) - timedelta(days=1)).strftime("%Y/%m"))
    query2 = SubmitField((date(date.today().year, date.today().month-1, 1) - timedelta(days=1)).strftime("%Y/%m"))
    query3 = SubmitField((date(date.today().year, date.today().month-2, 1) - timedelta(days=1)).strftime("%Y/%m"))
    query4 = SubmitField((date(date.today().year, date.today().month-3, 1) - timedelta(days=1)).strftime("%Y/%m"))
    query5 = SubmitField((date(date.today().year, date.today().month-4, 1) - timedelta(days=1)).strftime("%Y/%m"))
    generatedate = DateField(u'自选日期', format='%Y-%m-%d')
    submit = SubmitField('查询自选日期')



