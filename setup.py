# -*- coding: UTF8 -*-
from distutils.core import setup

import __version__

classifiers = [
  'Development Status :: 5 - Production/Stable',
  'Environment :: Win32 (MS Windows)',
  'Intended Audience :: Developers',
  'Intended Audience :: System Administrators',
  'Programming Language :: Python :: 2',
  'License :: PSF',
  'Natural Language :: English',
  'Operating System :: Microsoft :: Windows :: Windows 95/98/2000',
  'Topic :: System :: Systems Administration'
]

setup (
  name = "active_directory",
  version = __version__.__VERSION__,
  description = "Active Directory",
  long_description = open ("readme.txt").read (),
  author = "Tim Golden",
  author_email = "mail@timgolden.me.uk",
  url = "http://timgolden.me.uk/python/active_directory.html",
  license = "http://www.opensource.org/licenses/mit-license.php",
  py_modules = ["active_directory"]
)

