from flask_wtf import FlaskForm
###从Flask-WTF扩展导入Form基类
from wtforms import SubmitField, SelectField,  SelectMultipleField
###从WTForms包中导入字段类
from wtforms.validators import DataRequired
from flask_wtf.file import FileField, FileAllowed, FileRequired
from .. import excels


class GenerateForm(FlaskForm):
    excels = SelectMultipleField('File to Download', choices=[('1', '资金期限表'), ('2', 'G25'), ('3', 'Q02')],
                                 validators=[DataRequired()], coerce=int)
    submit = SubmitField('Generate')
