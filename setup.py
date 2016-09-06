#!/usr/bin/env python

from distutils.core import setup
import distutils.core
import ooolib

setup(name="ooolib-python",
      version=ooolib.version_number(),
      description="Package for creating OpenDocument Format(ODF) spreadsheets.",
      author="Joseph Colton",
      author_email="joseph@colton.byuh.edu",
      maintainer="Joseph Colton",
      maintainer_email="joseph@colton.byuh.edu",
      license = 'GNU LGPL',
      url = "http://sourceforge.net/projects/ooolib/",
      packages=['ooolib']
      )
