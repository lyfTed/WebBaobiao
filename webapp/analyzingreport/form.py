from flask_wtf import FlaskForm
###从Flask-WTF扩展导入Form基类
from wtforms import SubmitField, SelectField,  SelectMultipleField
from wtforms.fields.html5 import DateField
###从WTForms包中导入字段类
from wtforms.validators import DataRequired
from flask_wtf.file import FileField, FileAllowed, FileRequired
from .. import excels


class UploadForm(FlaskForm):
    excel = FileField('上传报表', validators=[FileAllowed(excels, u'文件格式不对'), FileRequired()])
    curr = SelectField('币种', validators=[DataRequired()], choices=[(1, 'EUR'), (2, 'GBP'), (3, 'AUD'), (4, 'USD'),
                                                        (5, 'CAD'), (6, 'SGD'), (7, 'HKD'), (8, 'JPY')], coerce=int)
    date = DateField(u'报表日期（YYYY-MM）', validators=[DataRequired()], format='%Y-%m-%d')
    submit = SubmitField('上传')


class QueryForm(FlaskForm):
    excels = SelectMultipleField('下载报表（可多选）', validators=[DataRequired()], coerce=int)
    generatedate = DateField(u'报表日期（YYYY-MM）', validators=[DataRequired()], format='%Y-%m-%d')
    submit = SubmitField('下载')


