..  module:: active_directory

Tutorial
========

Basics
------

Find by name
~~~~~~~~~~~~

To find the first item wtih a certain name::

  import active_directory

  item1 = active_directory.find("item1")

To narrow the search::

  import active_directory

  user1 = active_directory.find("user1", objectClass="user", objectCategory="person")

Find
~~~~

To find a specific user, group or ou::

  import active_directory

  u1 = active_directory.find_user("u1")
  g1 = active_directory.find_group("g1")
  o1 = active_directory.find_ou("o1")

If any doesn't exist, None is returned. If more than one
item matches, the first is returned.

Search
~~~~~~

To search for all items matching a set of criteria::

  import active_directory

  for tim in active_directory.search(
      objectCategory="person",
      displayName="Tim*"
  ):
      print (tim)

Display one attribute
~~~~~~~~~~~~~~~~~~~~~

To see one of the attributes of an AD object::

  import active_directory

  john_smith = active_directory.find_user("John Smith")
  print(john_smith.sAMAccountName)
  print(john_smith.displayName)
  print(john_smith.distinguishedName)

Display all attributes
~~~~~~~~~~~~~~~~~~~~~~

To see a quick display of all of an AD object's attributes::

  import active_directory

  john_smith = active_directory.find_user("John Smith")
  john_smith.dump()



Slightly More Advanced
----------------------

Find the root of a domain
~~~~~~~~~~~~~~~~~~~~~~~~~

To determine the root of the default domain::

  import active_directory

  domain = active_directory.AD()

To determine the root of a domain from one of its DCs::

  import active_directory

  domain = active_directory.AD("SVR-DC1")

Search or Find from a particular point
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

To start searching (or finding) from a particular point in
the AD tree::

  import active_directory

  users = active_directory.AD().find_ou("Users")
  for tim in users.search(displayName="Tim*"):
      print(tim)

Search with more complex criteria
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

To search when the criteria are more complex than equal-to conditions,
and-ed together::

  import active_directory

  for person in active_directory.search (
      "(displayName='Tim*' AND logonCount > 0) OR displayName='Fred'",
      objectCategory="person"
  ):
      print (person)

..  note::
    The query mechanism which underlies all the searches is using
    the SQL form of querying, so any positional parameters such as
    the above must fit that style. To send an LDAP query string, use
    the :func:`query` function directly, optionally wrapping the
    resulting records via the :func:`AD_object` function.

Raw Search
~~~~~~~~~~

To perform a search with a predetermined query string, and without
converting the results to AD objects::

  import active_directory

  base = "<LDAP://%s>" % active_directory.AD()
  for item in active_directory.query (
        base + ";(objectClass=group);distinguishedName,displayName,sAMAccountName"
  ):
      print (item.distinguishedName)
