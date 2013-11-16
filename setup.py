#!/usr/bin/python

from setuptools import setup
from pyexcelerate import __version__

setup(
    name="PyExcelerate",
    version=__version__,
    author="Kevin Wang and Kevin Zhang",
    author_email="kevin@kevinzhang.me",
    maintainer="Kevin Zhang",
    maintainer_email="kevin@kevinzhang.me",
    url="https://github.com/kz26/PyExcelerate",
    description="Accelerated Excel XLSX Writing Library for Python 2/3",
    install_requires=[
        'Jinja2',
		'six'
    ],
    packages=[	
        'pyexcelerate'
    ]
)
