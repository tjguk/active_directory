******************************
Python Active Directory Module
******************************

What is it?
===========

Active Directory (AD) is Microsoft's answer to LDAP, the industry-standard
 directory service holding information about users, computers and
 other resources in a tree structure, arranged by departments or
 geographical location, and optimized for searching.

The Python Active Directory module is a lightweight wrapper on top of the
pywin32 extensions, and hides some of the plumbing needed to get Python to
talk to the AD API. It's pure Python and should work with any version of
Python from 2.2 onwards (generators) and any recent version of pywin32.


Where do I get it?
==================

http://timgolden.me.uk/python/active_directory.html


How do I install it?
====================

When all's said and done, it's just a module. But for those
who like setup programs:

python setup.py install


Prerequisites
=============

If you're running a recent Python (2.2+) on a recent Windows (2k, 2k3, XP)
and you have Mark Hammond's pywin32 extensions installed, you're probably
up-and-running already. Otherwise...

  Windows
  -------
  If you're running Win9x / NT4 you'll need to get AD support
  from Microsoft. Microsoft URLs change quite often, so I suggest you
  do this: 
  http://www.google.com/search?q=site%3Amicrosoft.com+active+directory+downloads

  Python
  ------
  http://www.python.org/ (just in case you didn't know)

  pywin32 (was win32all)
  ----------------------
  http://pywin32.sf.net


How do I use it?
================

There are examples at: http://timgolden.me.uk/python/ad_cookbook.html
but as a quick taster, try this, to list all users' display names:

import active_directory

for person in active_directory.search ("objectCategory='Person'"):
  print person.displayName

What License is it released under?
==================================
Licensed under the Python Software Foundation license:
 http://www.python.org/psf/license.html
