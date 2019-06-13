# _*_ coding: utf-8 _*_
# filename: __init__.py
from flask import Blueprint

_analyzingreport = Blueprint('analyzingreport', __name__)

from . import view