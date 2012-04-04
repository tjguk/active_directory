# -*- coding: UTF8 -*-
from distutils.core import setup
import __active_directory_version__ as __version__

classifiers = [
  'Development Status :: 5 - Production/Stable',
  'Environment :: Win32 (MS Windows)',
  'Intended Audience :: Developers',
  'Intended Audience :: System Administrators',
  "Programming Language :: Python :: 2",
  "Programming Language :: Python :: 3",
  'License :: PSF',
  'Natural Language :: English',
  'Operating System :: Microsoft :: Windows :: Windows 95/98/2000',
  'Topic :: System :: Systems Administration'
]

#
# setup wants a long description which we'd like to read
# from README.rst; setup also wants a file called README
# github, however, wants a file called readme.rst. This
# is the compromise:
#
try:
  long_description = open ("README.rst").read ()
  open ("README", "w").write (long_description)
except (OSError, IOError):
   long_description = ""

setup (
  name = "active_directory",
  version = __version__.__VERSION__ + __version__.__RELEASE__,
  description = "Active Directory",
  author = "Tim Golden",
  author_email = "mail@timgolden.me.uk",
  url = "https://github.com/tjguk/active_directory",
  license = "http://www.opensource.org/licenses/mit-license.php",
  py_modules = ["active_directory", "__active_directory_version__"],
  long_description = long_description,
)

