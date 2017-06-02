# -*- coding: utf-8 -*-

from setuptools import setup, find_packages
from codecs import open
from os import path

pwd = path.abspath(path.dirname(__file__))

with open(path.join(pwd, 'README.md')) as f:
    readme = f.read()

with open(path.join(pwd, 'LICENSE')) as f:
    license = f.read()

setup(
    name='PyVisio',
    version='0.1.beta1',
    description='Visio document manipulating library',
    long_description=readme,
    author='Ivo Velcovsky',
    author_email='velcovsky@email.cz',
    url='https://github.com/i-wan/pyvisio',
    license=license,
    classifiers=[
        'Environment :: Win32 (MS Windows)',
        'Operating System :: Microsoft',
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7'
    ],
    packages=find_packages(exclude=('tests', 'docs')),
    install_requires=['pypiwin32']
)
