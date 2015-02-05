#!/usr/bin/python

from setuptools import setup

setup(
	name="PyExcelerate",
	version='0.6.7',
	author="Kevin Wang, Kevin Zhang",
	author_email="kevin+pyexcelerate@kevinzhang.me",
	maintainer="Kevin Zhang",
	maintainer_email="kevin+pyexcelerate@kevinzhang.me",
	url="https://github.com/kz26/PyExcelerate",
	description="Accelerated Excel XLSX Writing Library for Python 2/3",
	long_description=open('README.rst').read(),
	install_requires=[
		'Jinja2',
		'six>=1.4.0'
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
	},
	classifiers=[
		'Development Status :: 4 - Beta',
		'Intended Audience :: Developers',
		'License :: OSI Approved :: BSD License',
		'Programming Language :: Python :: 2.6',
		'Programming Language :: Python :: 2.7',
		'Programming Language :: Python :: 3.3',
		'Programming Language :: Python :: 3.4',
	]

)
