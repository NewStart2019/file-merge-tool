# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files, collect_all, copy_metadata

block_cipher = None

# ── 自动递归收集关键包 ──────────────────────────────────────────────
hidden = []
datas_extra = []
binaries_extra = []

module = [
  'office', # python-office 及其依赖
  'imageio', # imageio（解决 No package metadata 问题）
  'pdf2docx',
  'pdfplumber',
  'pdfminer',
  'matplotlib',
]
datas_extra += copy_metadata('imageio')          # imageio 的 .dist-info
datas_extra += copy_metadata('python-office')    # python-office 的 .dist-info
datas_extra += copy_metadata('openpyxl')
datas_extra += copy_metadata('pdf2docx')
datas_extra += copy_metadata('pdfplumber')
datas_extra += copy_metadata('pdfminer.six')
datas_extra += copy_metadata('Pillow')
datas_extra += copy_metadata('numpy')
datas_extra += copy_metadata('requests')
datas_extra += copy_metadata('lxml')
datas_extra += copy_metadata('pymupdf')
datas_extra += collect_data_files('akshare')
datas_extra += copy_metadata('akshare')

for module_name in module:
  tmp_ret = collect_all(module_name)
  datas_extra += tmp_ret[0]; binaries_extra += tmp_ret[1]; hidden += tmp_ret[2]

# pkg_resources（解决元数据检查问题）
hidden += collect_submodules('pkg_resources')
datas_extra += collect_data_files('pkg_resources')

# setuptools（pkg_resources 依赖）
hidden += collect_submodules('setuptools')

# pymupdf（pdf2docx 核心依赖）
tmp_ret = collect_all('pymupdf')
datas_extra += tmp_ret[0]; binaries_extra += tmp_ret[1]; hidden += tmp_ret[2]

a = Analysis(
    ['word_tables_to_excel.py'],
    pathex=['.'],
    binaries=binaries_extra,
    datas=datas_extra + [
        # 如有自定义资源文件在此追加，例如：
        # ('config.json', '.'),
    ],
    hiddenimports=hidden + [
        # docx / openpyxl
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',

        # lxml
        'lxml',
        'lxml.etree',
        'lxml.html',
        'lxml._elementpath',

        # Pillow
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',

        # multiprocessing（pdf2docx 多进程转换）
        'multiprocessing',
        'multiprocessing.pool',
        'multiprocessing.managers',
        'multiprocessing.synchronize',
        'multiprocessing.sharedctypes',

        # Windows 特有
        'win32timezone',
        'win32api',
        'win32con',

        # pypdf 加密（可选，避免启动报错）
        'pypdf._crypt_providers._cryptography',

        # 其他运行时依赖
        'cv2',
        'numpy',
        'requests',
        'urllib3',
        'charset_normalizer',
        'certifi',
        'jieba',
        'imageio',
        'imageio.plugins',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 明确排除不需要的大型包，减小体积
        'tkinter',
        'IPython',
        'jupyter',
        'notebook',
        'mypy',
        'pytest',
        '_pytest',
        'PySide6',
        'PyQt4',
        'PyQt5',
        'tornado',
        'sqlalchemy',
        'pyarrow',
        'numba',
        'paddle',
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
    name='word_tables_to_excel',
    debug=False,          # 改 True 可在 cmd 看详细启动日志
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)