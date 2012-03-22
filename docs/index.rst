The active_directory Module
***************************

..  note::
    While the module itself is fairly mature, these docs are definitely
    a work in progress. I'll try to get more examples into the cookbook
    which should help people get started.


What is it?
-----------

The active_directory module is a light wrapper around the
AD functionality. It allows easy searching of common objects
(users, groups, ou) and browsing of their contents. AD objects
are wrapped in Python objects which ease their use in Python
code while allowing the underlying object to be accessed easily.

* :doc:`searching`

There's also list of cookbook examples:

* :doc:`cookbook`


Where do I get it?
------------------

* pip install active_directory
* github: http://github.com/tjguk/active_directory
* PyPI: http://pypi.python.org/pypi/active_directory


Copyright & License?
--------------------

* Copyright Tim Golden <mail@timgolden.me.uk> 2012

* Licensed under the (GPL-compatible) MIT License:
  http://www.opensource.org/licenses/mit-license.php


Prerequisites & Compatibility
-----------------------------

The module has been tested on versions of Python from 2.4 to 2.7 plus Python 3.2
running on WinXP, Win7 & Win2k3. It may also work on older (or newer) versions.
It's tested with the most recent pywin32 extensions, but the functionality
it uses from those libraries has been in place for many versions.
