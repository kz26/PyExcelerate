from distutils.core import setup

setup(
    name="PyExcelerate",
    version="0.2.0",
    author="Kevin Wang and Kevin Zhang",
    author_email="zhangk@uchicago.edu",
    maintainer="Kevin Zhang",
    maintainer_email="zhangk@uchicago.edu",
    url="https://github.com/whitehat2k9/PyExcelerate",
    description="Accelerated Excel XLSX Writing Library for Python",
    license="LICENSE",
    install_requires=[
        "Jinja2"
    ],
    packages=[
        'pyexcelerate'
    ]
)
