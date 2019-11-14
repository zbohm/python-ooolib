#!/usr/bin/env python
# -*- coding: utf-8 -*-
from setuptools import find_packages, setup
import ooolib

with open("README.rst", "r") as fh:
    long_description = fh.read()


setup(
    name="ooolib-python",
    version=ooolib.version_number(),
    description="Package for creating OpenDocument Format(ODF) spreadsheets.",
    author="Joseph Colton",
    author_email="joseph@colton.byuh.edu",
    maintainer=u"Zdeněk Böhm",
    maintainer_email="zdenek.bohm@nic.cz",
    license='MIT',
    url="https://github.com/zbohm/python-ooolib",
    long_description=long_description,
    packages=find_packages(),
    package_data={
        "ooolib": ["tests/fixtures/test-cells.ods"],
    },
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Programming Language :: Python :: 2.7",
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ]
)
