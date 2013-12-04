#!/usr/bin/python

from setuptools import setup

setup(
	name="PyExcelerate",
	version='0.4.1',
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
	],
	package_data={
		'pyexcelerate': [
			'templates/*.xml',
			'templates/_rels/.rels',
			'templates/docProps/*.xml',
			'templates/xl/*.xml',
			'templates/xl/_rels/*',
			'templates/xl/worksheets/*.xml',
		]
	}
)
