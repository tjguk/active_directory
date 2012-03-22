Changelog
=========

..  module:: active_directory

0.8
---

16th March 2012

* Added Python 3.x compatibility
* Bumped Python 2.x support up to Python 2.3+ only (sets & datetime)
* Added walk method to any container

0.7.1
-----

3rd April 2008

* Removed any features which post-date Python 2.2

0.7
---

12th Jan 2008

* Added general-purpose find_XXX function to _AD_object instances.
  The idea is that .find_abc_def ('xxx') translates to a call to
  .search (objectClass="abcDef", name="xxx") and returns the first
  item found.
* Added ability to clear cache
* Some tidying-up and commenting
* Added hashability to allow for inclusion in sets
* Added general-purpose .search_xxx function which operates in
  the same way as find_xxx

0.6.6
-----

27th Apr 2007

* Escaped slash character in LDAP moniker
  (Thanks for Jason Erickson for bug report and patch)

0.6.5
-----

16th Mar 2007

* Really corrected bug in search clause handling

0.6.4
-----

16th Mar 2007

* Corrected bug in search clause handling

0.6.3
-----

12 Mar 2007

* Fixed bug in find_user / search

0.6.2
-----

12th Mar 2007


* Slight refactoring
* Added find_ou method to AD_objects and at module level
* Added find_public_folder method to AD_objects and at module level

0.6.1
-----

11th Mar 2007

* Bundle-bugfix release

0.6
---

11th Mar 2007

* Reasonably substantial overhaul
* Added useful converters to many properties.
* Separated out common types of AD objects
* Added find_group method to AD_objects and at module level
* Moved find_user / find_computer to AD Object; module-level now proxies
* Added os.walk-style .walk method to AD_group
* Made AD_object a factory function, doing useful things with
  path or object.

0.4
---

12th May 2005

* Added ADS_GROUP constants to support cookbook examples.
* Added .dump method to AD_object to allow easy viewing of all fields.
* Allowed find_user / find_computer to have default values,
  meaning the logged-on user and current machine.
* Added license: PSF

0.3
---

20th Oct 2004

* Added "Page Size" param to query to allow resultsets of > 1000.
* Refactored search mechanisms to module-level and switched to SQL queries.

0.2
---

19th Oct 2004

* Added support for attribute assignment (see AD_object.__setattr__)
* Added module-level functions:
  + root - returns a default AD instance
  + search - calls root's search
  + find_user - returns first match for a user/fullname
  + find_computer - returns first match for a computer
* Now runs under 2.2 (removed reference to basestring)

0.1
---

15th Oct 2004

* Initial release
