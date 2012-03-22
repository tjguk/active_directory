active_directory
================

active_directory - a lightweight wrapper around COM support
for Microsoft's Active Directory

Active Directory is Microsoft's answer to LDAP, the industry-standard
directory service holding information about users, computers and
other resources in a tree structure, arranged by departments or
geographical location, and optimized for searching.

There are several ways of attaching to Active Directory. This
module uses the Dispatchable LDAP:// objects and wraps them
lightly in helpful Python classes which do a bit of the
otherwise tedious plumbing. The module is quite naive, and
has only really been developed to aid searching, but since
you can always access the original COM object, there's nothing
to stop you using it for any AD operations.

+ The active directory class (_AD_object or a subclass) will determine
  its properties and allow you to access them as instance properties.

   eg
     import active_directory
     goldent = active_directory.find_user ("goldent")
     print ad.displayName

+ Any object returned by the AD object's operations is themselves
  wrapped as AD objects so you get the same benefits.

  eg
    import active_directory
    users = active_directory.root ().child ("cn=users")
    for user in users.search ("displayName='Tim*'"):
      print user.displayName

+ To search the AD, there are two module-level general
  search functions, and module-level convenience functions
  to find a user, computer etc. Usage is illustrated below:

   import active_directory as ad

   for user in ad.search (
     "objectClass='User'",
     "displayName='Tim Golden' OR sAMAccountName='goldent'"
   ):
     #
     # This search returns an AD_object
     #
     print user

   query = \"""
     SELECT Name, displayName
     FROM 'LDAP://cn=users,DC=gb,DC=vo,DC=local'
     WHERE displayName = 'John*'
   \"""
   for user in ad.search_ex (query):
     #
     # This search returns an ADO_object, which
     #  is faster but doesn't give the convenience
     #  of the AD methods etc.
     #
     print user

   print ad.find_user ("goldent")

   print ad.find_computer ("vogbp200")

   users = ad.AD ().child ("cn=users")
   for u in users.search ("displayName='Tim*'"):
     print u

+ Typical usage will be:

import active_directory

for computer in active_directory.search ("objectClass='computer'"):
  print computer.displayName

(c) Tim Golden <active-directory@timgolden.me.uk> October 2004-2012
Licensed under the (GPL-compatible) MIT License:
http://www.opensource.org/licenses/mit-license.php

Many thanks, obviously to Mark Hammond for creating
the pywin32 extensions without which this wouldn't
have been possible.