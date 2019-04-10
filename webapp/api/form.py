from flask_wtf import FlaskForm
###从Flask-WTF扩展导入Form基类
from wtforms import SubmitField, SelectField,  SelectMultipleField, DateField
###从WTForms包中导入字段类
from wtforms.validators import DataRequired
from flask_wtf.file import FileField, FileAllowed, FileRequired
from .. import excels


class UploadForm(FlaskForm):
    excels = FileField('File to Upload', validators=[FileAllowed(excels, u'文件格式不对'), FileRequired()])
    submit = SubmitField('Upload')


class DownloadForm(FlaskForm):
    excels = SelectMultipleField('Files to Download', choices=[('1', '资金期限表'), ('2', 'G25'), ('3', 'Q02')],
                                 validators=[DataRequired()], coerce=int)
    generatedate = DateField(u'生成日期', validators=[DataRequired()], format='%Y-%m-%d')
    submit = SubmitField('Download')


class SplitForm(FlaskForm):
    excels = SelectMultipleField('Files to Split', choices=[('1', '资金期限表'), ('2', 'G25'), ('3', 'Q02')],
                                 validators=[DataRequired()], coerce=int)
    submit = SubmitField('Split')
