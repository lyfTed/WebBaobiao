# _*_ coding: utf-8 _*_
# filename: __init__.py
from flask import Blueprint

_main = Blueprint('main', __name__)

from . import view