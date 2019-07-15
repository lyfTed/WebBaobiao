from flask_wtf import FlaskForm
###从Flask-WTF扩展导入Form基类
from wtforms import SubmitField, SelectField,  SelectMultipleField
from wtforms.fields.html5 import DateField
###从WTForms包中导入字段类
from wtforms.validators import DataRequired
from flask_wtf.file import FileField, FileAllowed, FileRequired
from .. import excels


class UploadForm(FlaskForm):
    excels = FileField('上传报表', validators=[FileAllowed(excels, u'文件格式不对'), FileRequired()])
    submit = SubmitField('上传')


class DownloadForm(FlaskForm):
    excels = SelectMultipleField('下载报表（可多选）', validators=[DataRequired()], coerce=int)
    generatedate = DateField(u'报表日期(YYYY-MM-DD)', validators=[DataRequired()], format='%Y-%m-%d')
    submit = SubmitField('下载')


