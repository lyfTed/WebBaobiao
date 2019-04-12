from flask_wtf import FlaskForm
###从Flask-WTF扩展导入Form基类
from wtforms import SubmitField, SelectField,  SelectMultipleField, DateField
###从WTForms包中导入字段类
from wtforms.validators import DataRequired
from flask_wtf.file import FileField, FileAllowed, FileRequired
from .. import excels


class GenerateForm(FlaskForm):
    excels = SelectMultipleField('生成报表（可多选）', choices=[('1', '资金期限表'), ('2', 'G25'), ('3', 'Q02')],
                                 validators=[DataRequired()], coerce=int)
    generatedate = DateField(u'报表日期（YYYY-MM）', validators=[DataRequired()], format='%Y-%m-%d')
    submit = SubmitField('生成')
