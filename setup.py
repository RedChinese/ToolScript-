# -*- coding: utf-8 -*-
# @Time : 2023/3/10 10:32
# @Author : zhenyu
# @Software: PyCharm
import setuptools

setuptools.setup(
    name="bttool",
    version="1.0",
    # 维护者信息
    dseription="This is the Blue Film report generator tool",
    install_requires=['python-docx', 'xlsxwriter'],
    # package_dir=
    packages=['BTTool', 'BTTool/lib'],
    # py_modules=["BTTool", 'BTTool'],
    python_requrires=">=3.0",
    platforms="mac",
    license="MIT",
    entry_points={
        'console_scripts': [
            'bttool = BTTool.btreportmain:main'
        ]
    },

    # scripts=""
)
