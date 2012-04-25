..  module:: active_directory

Tutorial
========

The `active_directory` module has, essentially, three areas of functionality:
searching and browsing; reading and writing attributes; and managing objects.
These are covered  below in that order. To simplify things the examples assume that you
are logged on to a domain account with suitable authority to perform the actions
on the default Active Directory instance.

Searching & Browsing
--------------------

Before you can do anything within Active Directory, you must have hold of a
suitable object. You have two easy options to start with: either find the
root of the nearest available Active Directory instance or search for
an object by name and any other characteristics.

Active Directory Root
~~~~~~~~~~~~~~~~~~~~~

This can be accessed by a call to the module-level :func:`AD` function.
This finds the root of the Domain you're logged onto and returns an
object which you can then use for further browsing and searching::

    import active_directory

    domain = active_directory.AD()

Finding one object
------------------

The `find_*` family of functions will search according to the criteria supplied
and will return the first object found or `None` if no object matches the criteria.
The functions below exist as module-level functions and as methods of any container-like
AD object. The module functions construct a cached object representing the root of the
default domain and search from there.

All the find_* family of functions use Ambiguous Name Resolution
when searching, so the name will be compared against display name, NT account
name and canonical name to find a match.

Find an Object
~~~~~~~~~~~~~~

To find a particular object by name, use the :func:`find`
function which takes the name of an object and, optionally, other search
parameters which might include the object's class or any other characteristics::

    import active_directory

    domain_users = active_directory.find("Domain Users", objectClass="group")


Find a User, Group, Computer or OU
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Since finding users, groups, computers & organisational units is such
a common requirements there are convenience functions to find objects
of these types::

    import active_directory

    domain_admins = active_directory.find_group("Domain Admins")
    tim = active_directory.find_user("Tim Golden")
    users = active_directory.find_ou("Users")
    dc = active_directory.find_computer("SVR-DC1")

Note that these more focused functions take no further parameters; you
can only specify the name to be found.

Find an object from an existing container
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If you already have an object containing an Active Directory
container, for example an organisational unit, you can call its
:meth:`AD_object.find` method in the same way::

    import active_directory

    users = active_directory.find_ou("Users")
    tim = users.find_user("tim")


Searching
---------

If you need to identify all the objects matching a set of criteria,
use the module- or object-level :func:`search` function. This function
takes a combination of keyword and positional parameters. The keyword
parameters are converted to equal-to filters while the positional
parameters are passed through unchanged to the underlying query
functionality. Note that the active_directory module currently uses
the SQL-esque query language rather than the more conventional LDAP
syntax. If you prefer to use that syntax, see the :func:`query` function
below.

The `search` function iterates over its results, yielding each one
as an :class:`AD_object` instance.

Search with simple criteria
~~~~~~~~~~~~~~~~~~~~~~~~~~~

To search for all items matching a set of simple criteria::

  import active_directory

  for tim in active_directory.search(
      objectCategory="person",
      displayName="Tim*"
  ):
      print (tim.sAMAccountName)

Search with non-simple criteria
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

To search for items where the criteria are more complex::

    import active_directory

    #
    # FIXME
    #
    for inactive in active_directory.search(
        "userAccountControl:1.2.840.113556.1.4.803:=2",
        objectCategory="person",
        objectClass="user"
    ):
        print(inactive.displayName)

Run a raw query
~~~~~~~~~~~~~~~

If you need or prefer to use LDAP-style queries, or if you have some other
query which is difficult to carry out with the :func:`search` function,
you can call the lower-level :func:`query` function which the search & find
functions call under the covers.

Note that this function expects you to pass a correctly-formatted ADO
query string and returns an :class:`ADO_record` object. You can convert
this into a wrapped AD object by calling the :func:`AD_object` function
with

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
