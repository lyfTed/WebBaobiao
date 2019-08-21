from flask_wtf import FlaskForm
###从Flask-WTF扩展导入Form基类
from wtforms import SubmitField, SelectField,  SelectMultipleField, StringField, FloatField, HiddenField
from wtforms.fields.html5 import DateField
###从WTForms包中导入字段类
from wtforms.validators import DataRequired, Length
from wtforms import widgets
from flask_wtf.file import FileField, FileAllowed, FileRequired
from datetime import datetime, date, timedelta
from .. import excels


class BaobiaoForm(FlaskForm):
    excels = SelectMultipleField(u'拆分报表（多选）', validators=[DataRequired()], coerce=int)
    submit = SubmitField(u'拆分')


class TianxieForm(FlaskForm):
    value = FloatField(u'值', validators=[DataRequired()])
    submit = SubmitField(u'提交')


class PreviewForm(FlaskForm):
    excel = SelectField(u'报表模板', validators=[DataRequired()], coerce=int)
    preview = SubmitField(u'预览第一个Sheet')
    download = SubmitField(u'下载报表模板')


class GenerateForm(FlaskForm):
    excels = SelectMultipleField(u'生成报表（可多选）', validators=[DataRequired()], coerce=int)
    submit = SubmitField(u'生成')


class QueryForm(FlaskForm):
    form_name = HiddenField('Form Name')
    excel = SelectField(u'报表名', validators=[DataRequired()], coerce=int, id='select_excel')
    querydate = SelectField(u'查询日期', coerce=str, id='select_query_date')
    customizeddate = DateField(u'或自选日期(IE用户请输入YYYY-MM-DD)', format='%Y-%m-%d')
    submit = SubmitField(u'查询(显示第一个Sheet)')



