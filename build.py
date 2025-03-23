import PyInstaller.__main__
import os

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 定义需要打包的文件
main_script = os.path.join(current_dir, 'main.py')

# 定义需要包含的数据文件
datas = [
    ('输出模版.xlsx', '.'),
    ('excel_processor.py', '.'),
    ('file_preview.py', '.')
]

# 构建datas参数
datas_args = []
for src, dst in datas:
    datas_args.extend(['--add-data', f'{src}{os.pathsep}{dst}'])

# PyInstaller参数
options = [
    main_script,
    '--name=Excel合并工具',
    '--windowed',
    '--noconsole',
    '--clean',
    '--onefile',
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--hidden-import=PySide6',
] + datas_args

# 运行PyInstaller
PyInstaller.__main__.run(options) 