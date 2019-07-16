# _*_ coding: utf-8 _*_
import os
from webapp import create_app
from flask import render_template

# app = create_app(os.environ.get('FLASKY_CONFIG'))  # 实例化app
app = create_app(os.environ.get('FLASKY_CONFIG') or 'default')  # 实例化app


@app.route('/')
def index():
    return render_template('index.html')


if __name__ == '__main__':
    app.run(host='0.0.0.0', threaded=True, processes=4)




