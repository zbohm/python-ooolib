#!/usr/bin/env python
# -*- coding: utf-8 -*-
import setuptools
import ooolib

with open("README.md", "r") as fh:
    long_description = fh.read()


setuptools.setup(
      name="ooolib-python",
      version=ooolib.version_number(),
      description="Package for creating OpenDocument Format(ODF) spreadsheets.",
      author="Joseph Colton",
      author_email="joseph@colton.byuh.edu",
      maintainer=u"Zdeněk Böhm",
      maintainer_email="zdenek.bohm@nic.cz",
      license='GNU LGPL',
      url="https://github.com/zbohm/python-ooolib",
      long_description=long_description,
      packages=['ooolib', 'ooolib.tests'],
      classifiers=[
         "Programming Language :: Python :: 2.7",
         "Programming Language :: Python :: 3",
         "License :: OSI Approved :: GNU Lesser General Public License v3 (LGPLv3)",
         "Operating System :: OS Independent",
      ]
)
