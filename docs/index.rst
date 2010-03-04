active_directory - Active Directory Management
==============================================

What is it?
-----------

Blha blah


Where do I get it?
------------------

* **Subversion**: http://svn.timgolden.me.uk/wmi/trunk/
* **Windows installer**: http://timgolden.me.uk/python/downloads/WMI-1.4.6.win32.exe
* **Zipped-up source**: http://timgolden.me.uk/python/downloads/WMI-1.4.6.zip

* **Older Versions**: http://timgolden.me.uk/python/downloads

How do I install it?
--------------------

When all's said and done, it's just a module. But for those who like setup programs::

  python setup.py install

Or download the Windows installer and double-click.


How do I use it?
----------------

Have a look at the :doc:`tutorial` or the :doc:`cookbook`. As a quick
taster, try this, to find all Automatic services which are not running
and offer the option to restart each one::

  import active_directory

  c = wmi.WMI ()
  for s in c.Win32_Service (StartMode="Auto", State="Stopped"):
    if raw_input ("Restart %s? " % s.Caption).upper () == "Y":
      s.StartService ()

What's Changed?
---------------

See the :doc:`changes` document

Copyright & License?
--------------------

* Copyright Tim Golden <mail@timgolden.me.uk> 2003 - 2010

* Licensed under the (GPL-compatible) MIT License:
  http://www.opensource.org/licenses/mit-license.php

Prerequisites
-------------

If you're running a recent Python (2.5+) on a recent Windows (2k, 2k3, XP)
and you have Mark Hammond's win32 extensions installed, you're probably
up-and-running already. Otherwise...

Python
~~~~~~
http://www.python.org/ (just in case you didn't know)

pywin32 (was win32all)
~~~~~~~~~~~~~~~~~~~~~~
http://sourceforge.net/projects/pywin32/files/
