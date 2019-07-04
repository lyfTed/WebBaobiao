from wtforms import StringField
from flask_admin.form import BaseForm


class BaobiaoTeSetForm(BaseForm):
    id = StringField()
    file = StringField()
    freq = StringField()
    auditor = StringField()
