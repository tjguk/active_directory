active_directory2 - Active Directory Management
===============================================

What is it?
-----------

Active Directory (AD) is Microsoft's answer to LDAP, the industry-standard
directory service holding information about users, computers and
other resources in a tree structure, arranged by departments or
geographical location, and optimized for searching.

The Python Active Directory module is a lightweight wrapper on top of the
pywin32 extensions, and hides some of the plumbing needed to get Python to
talk to the AD API. It's pure Python and should work with any recent combination
of python and pywin32.


Where do I get it?
------------------

* http://svn.timgolden.me.uk/active_directory/branches/rework
* http://pypi.python.org/...


How do I install it?
--------------------

* Using pip or easy_install: pip install active_directory2
* Using the Windows installers on PyPI
* From the source zipfile: unzip and then python setup.py install


How do I use it?
----------------

Have a look at the :doc:`tutorial` or the :doc:`cookbook`. For a quick
taster, this is how you would find all users in Domain Admins who haven't
logged on for 6 months::

  import datetime
  from active_directory2 import ad, schema

  root = ad.AD ()
  for user in root.search (
    schema.objectCategory="person",
    schema.objectClass="user",
    schema.memberOf="Domain Admins",
    schema.lastLogon < (datetime.datetime.now () - datetime.timedelta (days=30*6))
  ):
    print user


What's Changed?
---------------

See the :doc:`changes` document

Copyright & License?
--------------------

* Copyright Tim Golden <mail@timgolden.me.uk> 20011

* Licensed under the (GPL-compatible) MIT License:
  http://www.opensource.org/licenses/mit-license.php
