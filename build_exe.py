import PyInstaller.__main__
import os

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# PyInstaller参数
args = [
    'document_converter.py',
    '--name=文档批量转图片工具',
    '--windowed',
    '--onefile',
    '--clean',
    '--noconfirm',
    '--add-data', 'requirements.txt;requirements.txt',
    '--icon', 'NONE',  # 如果需要图标，可以添加.ico文件路径
    '--hidden-import', 'PyQt5.sip',
    '--hidden-import', 'PyQt5.QtCore',
    '--hidden-import', 'PyQt5.QtGui',
    '--hidden-import', 'PyQt5.QtWidgets',
    '--collect-all', 'PyQt5',
    '--collect-all', 'PIL',
    '--collect-all', 'fitz',
    '--collect-all', 'docx',
    '--exclude-module', 'matplotlib',
    '--exclude-module', 'numpy',
    '--exclude-module', 'scipy',
    '--exclude-module', 'pandas',
    '--workpath', os.path.join(current_dir, 'build'),
    '--distpath', os.path.join(current_dir, 'dist'),
    '--specpath', current_dir
]

if __name__ == '__main__':
    PyInstaller.__main__.run(args)
    print("打包完成！")
    print("生成的exe文件在 dist 目录中")