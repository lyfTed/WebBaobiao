from flask_wtf import FlaskForm
###从Flask-WTF扩展导入Form基类
from wtforms import SubmitField, SelectField,  SelectMultipleField, StringField, FloatField,DateField
###从WTForms包中导入字段类
from wtforms.validators import DataRequired, Length
from flask_wtf.file import FileField, FileAllowed, FileRequired
from .. import excels


class BaobiaoForm(FlaskForm):
    excel = SelectMultipleField('拆分报表（多选）', choices=[('1', '资金期限表'), ('2', 'G25'), ('3', 'Q02')],
                                 validators=[DataRequired()], coerce=int)
    submit = SubmitField(u'拆分')


class TianxieForm(FlaskForm):
    value = FloatField(u'值', validators=[DataRequired()])
    submit = SubmitField(u'提交')


class QueryForm(FlaskForm):
    excel = SelectField('报表名', choices=[('1', '资金期限表'), ('2', 'G25'), ('3', 'Q02')],
                                 validators=[DataRequired()], coerce=int)
    generatedate = DateField(u'报表日期（YYYY-MM）', validators=[DataRequired()], format='%Y-%m-%d')
    lastdate = DateField(u'上期日期（YYYY-MM）', validators=[DataRequired()], format='%Y-%m-%d')
    submit = SubmitField('查询结果')



