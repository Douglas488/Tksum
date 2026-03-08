# -*- coding: utf-8 -*-
# PythonAnywhere: 在 Web 里把 "WSGI configuration file" 指到这个文件，例如 /home/Douglas488/TkPy/wsgi.py
# 并确保 "Project" 的目录是 TkPy 所在目录，例如 /home/Douglas488/TkPy

import sys
import os

# 当前文件所在目录加入路径（即 TkPy 目录）
this_dir = os.path.dirname(os.path.abspath(__file__))
if this_dir not in sys.path:
    sys.path.insert(0, this_dir)

# 可选：若项目在子目录，改用下面这一行并改成你的实际路径
# sys.path.insert(0, '/home/Douglas488/TkPy')

os.chdir(this_dir)

from app import app as application
