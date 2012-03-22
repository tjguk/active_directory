..  module:: active_directory

Searching
=========

Parameters
----------

The search & find functions below all accept the following parameter styles
(in addition to any which are specific to the function in question):

* keyword args - these translate into an equals-to search in the resulting
  LDAP query, so search (objectClass='user') will return an iterable of all
  objects whose class is 'user'.
* positional args - these strings can be more complex, even compound, AD query components and
  will be AND-ed together. So search ("logonCount > 0", "displayName = Tim*")
  will return an iterable of all objects whose displayName starts with "Tim"
  and who have logged on at least once.
* both - it's perfectly possible to combine both parameter styles. According
  to normal Python argument rules, the keyword arguments must come after the
  positional arguments. So search ("logonCount > 0", objectClass="user")
  will return an iterable of all users who have logged on at least once.

Finding
-------

If you only want to find one thing (or the first thing which matches) then
the use the find_xxx family of methods. Each of these takes a name and
the more general one accepts other parameters. They then return the first
item which matches, or None if no match is found. There is no error if
more than one object matches; the first found will be returned.

The find functions are methods of the :class:`_AD_object` class but there
are module-level convenience counterparts which construct a default root
object to query against, cacheing it for future searches.

..  py:function:: find (name, *args, **kwargs)

    Search from the root of Active Directory for an object whose name (according
    to Ambiguous Name Resolution) matches the `name` parameter and which matches
    the othe search criteria, if any.

    :param name: the name of the object being searched for.
    :param args: further positional search criteria
    :param kwargs: further keyword search criteria
    :returns: a :class:`_AD_object` object or `None`

..  py:function:: find_user (name)

    Search from the root of Active Directory for an user whose name (according
    to Ambiguous Name Resolution) matches the `name` parameter.

    :param name: the name of the user being searched for.
    :returns: a :class:`_AD_user` object or `None`

..  py:function:: find_group (name)

    Search from the root of Active Directory for an group whose name (according
    to Ambiguous Name Resolution) matches the `name` parameter.

    :param name: the name of the group being searched for.
    :returns: a :class:`_AD_group` object or `None`

..  py:function:: find_ou (name)

    Search from the root of Active Directory for an OU whose name (according
    to Ambiguous Name Resolution) matches the `name` parameter.

    :param name: the name of the OU being searched for.
    :returns: a :class:`_AD_organisational_unit` object or `None`

..  py:function:: find_computer (name)

    Search from the root of Active Directory for a computer whose name (according
    to Ambiguous Name Resolution) matches the `name` parameter.

    :param name: the name of the computer being searched for.
    :returns: a :class:`_AD_computer` object or `None`

Searching
---------

If you want to find the set of objects matching some criteria the use
the search method. It accepts arbitrary parameters from which it constructs
a valid search string. It returns a (possibly empty) iterator over the matches
returned from Active Directory.

The search function is a method of the :class:`_AD_object` class but there
is a module-level counterpart which constructs a default root
object to query against, cacheing it for future searches.

..  py:function:: search (*args, **kwargs)

    Search from the root of Active Directory for all objects which
    match the criteria given.

    :param args: further positional search criteria
    :param kwargs: further keyword search criteria
    :returns: an iterator of :class:`_AD_object` objects

Raw Searching
-------------

The quickest searching, but requiring the most work up front, is to use
the :func:`search_ex` function whose only parameter is a well-formed Active
Directory search string and which returns an iterator of :class:`ADO_object`.
The search string can be conventional LDAP format or a sort of stunted SQL
accepted by Active Directory.

This is the easiest way to run an existing query (eg from a mailing list
or a webpage) against Active Directory::

  import active_directory

  root = "LDAP://dc=local,dc=westpark"
  query_string = """SELECT
    distinguishedName
  FROM
    %s
  WHERE
    objectClass = 'user'
  """ % root
  for result in active_directory.search_ex (query_string):
    print result.distinguishedName

