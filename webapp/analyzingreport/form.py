from flask_wtf import FlaskForm
###从Flask-WTF扩展导入Form基类
from wtforms import SubmitField, SelectField,  SelectMultipleField
from wtforms.fields.html5 import DateField
###从WTForms包中导入字段类
from wtforms.validators import DataRequired
from flask_wtf.file import FileField, FileAllowed, FileRequired
from .. import excels


class UploadForm(FlaskForm):
    # excel = FileField('上传报表', validators=[FileAllowed(excels, u'文件格式不对'), FileRequired()])
    kemu = SelectField('科目', validators=[DataRequired()], choices=[(1, '一级科目'), (2, '二级科目'), (3, '三级科目')],
                       coerce=int)
    institution = SelectField('机构', validators=[DataRequired()], choices=[(1, '全业务'), (2, '自贸')], coerce=int)
    currency = SelectField('币种', validators=[DataRequired()], choices=[(1, 'EUR'), (2, 'GBP'), (3, 'AUD'), (4, 'USD'),
                        (5, 'CAD'), (6, 'SGD'), (7, 'HKD'), (8, 'JPY'), (9, '外币折人民币'), (10, '外币折美元'),
                        (11, '本外币折人民币')], default=9, coerce=int)
    date = DateField(u'报表日期', validators=[DataRequired()], format='%Y-%m-%d')
    submit = SubmitField('上传')


class QueryForm(FlaskForm):
    kemu = SelectField('科目', validators=[DataRequired()], choices=[(1, '一级科目'), (2, '二级科目'), (3, '三级科目')],
                       coerce=int)
    institution = SelectField('机构', validators=[DataRequired()], choices=[(1, '全业务'), (2, '自贸')], coerce=int)
    currency = SelectField('币种', validators=[DataRequired()], choices=[(1, 'EUR'), (2, 'GBP'), (3, 'AUD'), (4, 'USD'),
                                                                       (5, 'CAD'), (6, 'SGD'), (7, 'HKD'), (8, 'JPY'),
                                                                       (9, '外币折人民币'), (10, '外币折美元'),
                                                                       (11, '本外币折人民币')], default=9, coerce=int)
    date = DateField(u'报表日期', validators=[DataRequired()], format='%Y-%m-%d')
    submit = SubmitField('上传')



