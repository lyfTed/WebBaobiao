# _*_ coding: utf-8 _*_
# filename: __init__.py
from flask import Blueprint

_auth = Blueprint('auth', __name__)

from . import view
