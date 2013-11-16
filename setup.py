#!/usr/bin/python

from distutils.core import setup
from pyexcelerate import __version__

setup(
    name="PyExcelerate",
    version=__version__,
    author="Kevin Wang and Kevin Zhang",
    author_email="zhangk@uchicago.edu",
    maintainer="Kevin Zhang",
    maintainer_email="zhangk@uchicago.edu",
    url="https://github.com/whitehat2k9/PyExcelerate",
    description="Accelerated Excel XLSX Writing Library for Python",
    license="LICENSE",
    install_requires=[
        'Jinja2',
		'six'
    ],
    packages=[	
        'pyexcelerate'
    ]
)
