#!/usr/bin/env,python
# ,','-*-,coding:,utf-8','-*-

import PyInstaller.__main__

PyInstaller.__main__.run([
    'main.py',  # ,替换为您的脚本文件名
    '--onefile',
    '--console',
    "--name={}".format("FileMerge"),
    "--distpath={}".format("dist"),
    # 指定模块依赖
    # '--hidden-import=clr', '--hidden-import=cssselect'
])
