#!/usr/bin/env python
# -*- coding: utf-8 -*-

from distutils.core import setup
import distutils.core
import ooolib

setup(name="ooolib-python",
      version=ooolib.version_number(),
      description="Package for creating OpenDocument Format(ODF) spreadsheets.",
      author="Joseph Colton",
      author_email="joseph@colton.byuh.edu",
      maintainer=u"Zdeněk Böhm",
      maintainer_email="zdenek.bohm@nic.cz",
      license = 'GNU LGPL',
      url = "https://github.com/zbohm/python-ooolib",
      packages=['ooolib']
      )
