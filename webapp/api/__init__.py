# _*_ coding: utf-8 _*_
# filename: __init__.py
from flask import Blueprint

_api = Blueprint('api', __name__)

from . import view