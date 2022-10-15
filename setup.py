from setuptools import setup
from setuptools import find_packages

setup(
    name='excelsigncheck',
    version='1.0.0',
    author='Henrique da Silva Santos',
    author_email='rique_dev@hotmail.com',
    packages=find_packages(exclude=('excelsigncheck.tests*',)),
    url='https://github.com/riquedev/excelsigncheck',
    description='Need to verify the digital signature of an Excel? Try this package.',
    install_requires=[
        "pywin32 == 304"
    ],
    test_suite="excelsigncheck.tests"

)
