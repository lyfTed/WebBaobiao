from flask import render_template
from . import _main


@_main.route('/')
def index():
    return render_template('index.html')
