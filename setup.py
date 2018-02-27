#
# Initially copied from:
# https://raw.githubusercontent.com/pypa/sampleproject/master/setup.py
#

from setuptools import setup, find_packages
import os
import codecs
import __active_directory_version__ as _version

here = os.path.abspath(os.path.dirname(__file__))

with codecs.open(os.path.join(here, 'README.rst'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='active_directory',

    version=_version.__VERSION__,

    description='Active Directory',
    long_description=long_description,

    url='https://github.com/tjguk/active_directory',

    author='Tim Golden',
    author_email='mail@timgolden.me.uk',

    license='MIT',

    classifiers=[
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
    ],

    py_modules=["active_directory", "__active_directory_version__"],
)
