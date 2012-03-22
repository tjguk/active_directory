active_directory
================

Active Directory (AD) is Microsoft's answer to LDAP, the industry-standard
directory service holding information about users, computers and
other resources in a tree structure, arranged by departments or
geographical location, and optimized for searching.

The active_directory module is a light wrapper around the
AD functionality. It allows easy searching of common objects
(users, groups, ou) and browsing of their contents. AD objects
are wrapped in Python objects which ease their use in Python
code while allowing the underlying object to be accessed easily.

active_directory is tested on all versions of Python from 2.4 to 3.2.
It makes heavy use of the adsi modules in the pywin32 extensions.

Docs are hosted at: http://active_directory.readthedocs.org/
