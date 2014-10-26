#!/usr/bin/env python
# encoding: utf-8
#
# Copyright (c) 2012, Peter Hillerström <peter.hillerstrom@gmail.com>
# All rights reserved. This software is licensed under 3-clause BSD license.
#
# For the full copyright and license information, please view the LICENSE
# file that was distributed with this source code.

from __future__ import with_statement


from setuptools import setup


PACKAGE_NAME = 'excelsior'
PACKAGE_VERSION = '0.0.4'
PACKAGES = ['excelsior']

def striplines(lines):
    return [line.strip() for line in lines]

with open('requirements.pip', 'rb') as reqs:
    install_requires = list(filter(None, striplines(reqs.readlines())))

with open('README.rst', 'rb') as readme:
    long_description = "\n" + readme.read()


setup(
    name=PACKAGE_NAME,
    version=PACKAGE_VERSION,
    packages=PACKAGES,
    requires=[],
    install_requires=install_requires,
    scripts=['bin/excelsior'],

    description="Convert Excel spreadsheets into TSV, CSV, Json or Yaml format.",
    long_description=long_description,
    author='Peter Hillerström',
    author_email='peter.hillerstrom@gmail.com',
    license="BSD License",
    url='https://github.com/peterhil/excelsior',

    classifiers=[
        'Development Status :: 4 - Beta',
        'Environment :: Console',
        'Intended Audience :: Developers',
        'Intended Audience :: End Users/Desktop',
        'Intended Audience :: Information Technology',
        'Intended Audience :: Science/Research',
        'License :: OSI Approved :: BSD License',
        'Operating System :: MacOS :: MacOS X',
        'Operating System :: POSIX',
        'Operating System :: Unix',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2.6',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Topic :: Text Processing :: Filters',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
        'Topic :: Utilities',
    ],
)
