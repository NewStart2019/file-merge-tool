# 安装环境编译可执行文件
## 前置准备
### 1. 安装 Python 环境（建议用虚拟环境）
cmd:: 创建虚拟环境（隔离依赖，避免打包体积过大）
```pwsh
python -m venv venv
venv\Scripts\activate
```

### 2. 安装依赖和 PyInstaller
```pwsh
pip install pyinstaller
pip install python-docx openpyxl pdfplumber pdf2docx 
pip install python-office
```
## 生成 .spec 文件
### 3. 先用命令行生成初始 spec（不直接打包）
```powershell
pyi-makespec --onefile --console --name word_tables_to_excel word_tables_to_excel.py

# 强制收集所有需要的依赖包
pyi-makespec `
  --onefile `
  --console `
  --collect-all python-office `
  --collect-all office `
  --collect-all pdf2docx `
  --collect-all pdfplumber `
  --name word_tables_to_excel `
  word_tables_to_excel.py
```
这会在当前目录生成 word_tables_to_excel.spec，不会立即编译。

## 编辑 .spec 文件
### 4. 打开并修改 spec 文件
用记事本或 VSCode 打开 word_tables_to_excel.spec，替换为以下完整内容：
```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['word_tables_to_excel.py'],       # 主脚本
    pathex=['.'],                       # 脚本所在目录
    binaries=[],
    datas=[
        # 如果你有额外资源文件，在这里添加，格式：('源路径', '目标目录')
        # ('config.json', '.'),
    ],
    hiddenimports=[
        # pdf2docx 和 pdfplumber 用了动态导入，需要手动声明
        'pdf2docx',
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        'pdfminer.converter',
        'pdfminer.pdfpage',
        'pdfminer.pdfinterp',
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'lxml',
        'lxml.etree',
        'PIL',                          # pdf2docx 依赖 Pillow
        'cv2',                          # pdf2docx 依赖 opencv
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',       # 不需要 GUI，排除减小体积
        'matplotlib',
        'numpy',         # 如果你的脚本不直接用 numpy 可排除
        'pandas',
        'scipy',
        'IPython',
        'jupyter',
        'notebook',
        'test',
        'unittest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='word_tables_to_excel',    # 生成的 exe 文件名
    debug=False,                     # 改为 True 可看详细错误
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,                        # 启用 UPX 压缩（需安装 UPX）
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,                    # True=命令行窗口，False=无窗口（GUI）
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,                       # 可指定图标: icon='app.ico'
)
```

## 编译为 exe
### 5. 执行编译
```shell
pyinstaller .\word_tables_to_excel.spec
```

### 6. 找到生成的 exe
```plaintext
your_project\
├── word_tables_to_excel.spec
├── build\                        ← 临时编译文件（可删除）
│   └── word_tables_to_excel\
└── dist\                         ← ✅ 最终输出在这里
    └── word_tables_to_excel.exe  ← 就是这个
```

## 验证运行
### 7. 测试 exe
```shell
dist\word_tables_to_excel.exe --help

:: 提取全部表格
dist\word_tables_to_excel.exe report.docx output.xlsx

:: 关键字过滤
dist\word_tables_to_excel.exe report.docx output.xlsx -k 合同 金额

:: PDF 输入
dist\word_tables_to_excel.exe report.pdf output.xlsx
```

## 常见问题排查
|问题|原因|解决办法|
|-|-|-|
|ModuleNotFoundError|动态导入未被发现| 在 spec 的 hiddenimports 里添加缺失模块名|
|exe 闪退无报错 |运行时错误被吞掉 | spec 里设 debug=True，或在 cmd 运行看输出 |
| lxml / cv2 报错 | 二进制依赖未打包 | 改用 --collect-all pdf2docx 重新生成 spec |
| 体积过大（>100MB）| pdf2docx/cv2 较重|在 excludes 里排除不用的库 |
| UPX 报错 | 未安装 UPX | 下载 upx.github.io 放到 PATH，或 spec 里设 upx=False|

## 如果有隐藏导入报错，用这个命令找缺失模块：
cmd:: debug 模式运行 exe，看具体哪个模块找不到
```shell
dist\word_tables_to_excel.exe report.docx out.xlsx 2>&1 | more
```
或重新打包时加 --collect-all：
cmd:: 把 pdf2docx 所有子模块都打包进去（最保险）
```shell
pyinstaller --onefile --collect-all pdf2docx --collect-all pdfplumber word_tables_to_excel.py
```

完整目录结构参考
# 常用的python工具库：

## 本文件是将指定目录下面的所有java文件合并到一个word文档。然后把根据指定目录下面的目录结构设计word目录等级。
## 打包成window可执行文件方法。
 * 使用管理员权限执行，安装pip install pyinstaller 
 * 找到可执行文件pyinstaller.exe
 * 打包：D:\AllServer\Anaconda\Scripts\pyinstaller.exe  --onefile main.py --name JavaFileMerge.exe

## 方法二：执行build.py配置文件 python  build.py
详细介绍请参考个人博客（语雀）

打包遇见：
The 'pathlib' package is an obsolete backport of a standard library package and is incompatible with PyInstaller. Please remove this package (located in D:\AllServer\Anaconda\lib\site-packages) using
    conda remove
then try again.
解决方法：conda remove pathlib
执行失败，配置国内镜像源
conda config --set show_channel_urls yes
conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/free/
conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/main/

手动配置镜像源： 用户目录下\.condarc 文件，添加下面的内容
channels:
  - https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/main/
  - https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/free/
