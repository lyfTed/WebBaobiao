from flask_wtf import FlaskForm
###从Flask-WTF扩展导入Form基类
from wtforms import SubmitField, SelectField,  SelectMultipleField, DateField
###从WTForms包中导入字段类
from wtforms.validators import DataRequired
from flask_wtf.file import FileField, FileAllowed, FileRequired
from .. import excels


class UploadForm(FlaskForm):
    excels = FileField('上传报表', validators=[FileAllowed(excels, u'文件格式不对'), FileRequired()])
    submit = SubmitField('上传')


class DownloadForm(FlaskForm):
    excels = SelectMultipleField('下载报表（可多选）', choices=[('1', '资金期限表'), ('2', 'G25'), ('3', 'Q02')],
                                 validators=[DataRequired()], coerce=int)
    generatedate = DateField(u'报表日期（YYYY-MM）', validators=[DataRequired()], format='%Y-%m-%d')
    submit = SubmitField('下载')

#
# class SplitForm(FlaskForm):
#     excels = SelectMultipleField('拆分报表', choices=[('1', '资金期限表'), ('2', 'G25'), ('3', 'Q02')],
#                                  validators=[DataRequired()], coerce=int)
#     submit = SubmitField('Split')
