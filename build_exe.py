# -*- coding: utf-8 -*-
"""PyInstaller 打包脚本"""
import PyInstaller.__main__
import os

here = os.path.dirname(os.path.abspath(__file__))
icon = os.path.join(here, 'calcbook.ico')  # 山水图标（软件图标 + .gcls 工程文件图标）

args = [
    os.path.join(here, 'main.py'),
    '--name=工程量计算书',
    '--onefile',
    '--windowed',
    '--clean',
    '--noconfirm',
    f'--distpath={os.path.join(here, "dist")}',
    f'--workpath={os.path.join(here, "build")}',
    f'--specpath={os.path.join(here, "build")}',
    f'--add-data={os.path.join(here, "core")}{os.pathsep}core',
    f'--add-data={icon}{os.pathsep}.',
    '--hidden-import=openpyxl',
    '--hidden-import=PyQt5.QtCore',
    '--hidden-import=PyQt5.QtGui',
    '--hidden-import=PyQt5.QtWidgets',
]
if os.path.exists(icon):
    args.append(f'--icon={icon}')
PyInstaller.__main__.run(args)
