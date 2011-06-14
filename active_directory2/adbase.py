# -*- coding: iso-8859-1 -*-
import os, sys
import socket

import win32api
from win32com import adsi
import win32com.client

from . import core
from . import constants
from . import exc
from . import support
from . import utils

class NotAContainerError (exc.ActiveDirectoryError):
  pass

class NoFilterError (exc.ActiveDirectoryError):
  pass

class _ADContainer (object):
  ur"""A support object which takes an existing AD COM object
  which implements the IADsContainer interface and provides
  a corresponding iterator.

  It is not expected to be called by user code (although it
  can be). It is the basis of the :meth:`ADBase.__iter__` method
  of :class:`ADBase` and its subclasses.
  """

  def __init__ (self, ad_com_object):
    try:
      self.container = exc.wrapped (ad_com_object.QueryInterface, adsi.IID_IADsContainer)
    except exc.ActiveDirectoryError, (error_code, _, _):
      if error_code == exc.E_NOINTERFACE:
        raise NotAContainerError

  def __iter__ (self):
    enumerator = exc.wrapped (adsi.ADsBuildEnumerator, self.container)
    while True:
      items = exc.wrapped (adsi.ADsEnumerateNext, enumerator, 10)
      if items:
        for item in items:
          yield item
      else:
        break

class ADBase (object):
  """A slender wrapper around an AD COM object.

  Attributes can be read & set. If the underlying object is a
  container it can be iterated over and subobjects can be retrieved, added
  and removed. It can also be walked (in the style of Python's os.walk) and
  flattened: :meth:`walk`, :meth:`flat`.

  Pretty-printing to stdout is via the :meth:`dump` method.

  There is simple searching with filters, returning the results as objects:
  :meth:`search`, :meth:`find`.

  The underlying object can itself be deleted via :meth:`delete`

  No validation is done of the attribute names and no conversions of the
  values. It can be used alone (most easily via the :func:`adbase` function
  which takes an AD path and returns an ADBase object). It also provides the
  basis for the :class:`ADObject` class.

  Attributes can be set by normal Python attribute access. Since for identifier
  names AD uses hyphens which aren't valid in Python identifiers, underscores
  will be converted to hyphens. To avoid Python keyword issues, trailing underscores
  are also stripped::

    from active_directory2 import ad

    me = ad.find_user ()
    me.displayName
    me.title = "Senior Programmer"
    me.department = None

  Since AD caches attribute lookup, a second retrieval of the same attribute
  for the same object won't hit the AD server and won't see any remote changes.
  To force the cache to be refreshed, use the :meth:`get` method which will
  refresh the cache and return the new attribute value.

  To set several attribute values at once, use the :meth:`set` method, which
  takes keyword args and commits the changes once all the attributes are set.

  Objects can be created, retrieved and removed under this object by item access.
  A relative distinguished name (rdn) must be given for the item. For object retrieval,
  this rdn can be several levels deep. For creation and deletion, only one level is
  allowed.

  Assigning to an item of an AD container creates a new object with the rdn given.
  Additional data is supplied as a dictionary-like object which must contain an
  entry for the object's Class and may contain other information. Exactly what must
  be passed will vary from one object type to another.

  ::

    from active_directory2 import ad

    my_ou = ad.find_ou ("MyOU")
    my_ou['cn=Tim'] = dict (Class="user", sAMAccountName="tim", displayName="Tim Golden", sn="TEST")
    my_ou['cn=Minimal'] = dict (Class="user", sn="TEST")
    for obj in my_ou.search (sn="TEST"):
      del my_ou[obj.Name]

  You can move an object from one container to another or (a slightly
  specialised version of the same thing) rename it within the same container.
  Note that an object to be moved must be a leaf node or an empty container::

    from active_directory2 import ad

    root = ad.AD ()
    my_ou = root['ou=MyOU']
    archive = root['ou=archive']
    my_ou.move ("cn=User1", archive)
    my_ou.rename ("cn=User2", "cn=User3")
  """

  #
  # For speed, hardcode the known properties of the IADs class
  #
  _properties = ["ADsPath", "Class", "GUID", "Name", "Parent", "Schema"]
  _schemas = {}

  def __init__ (self, obj):
    utils._set (self, "properties", set ())
    self.com_object = com_object = adsi._get_good_ret (obj)
    self.path = path = com_object.ADsPath
    scheme, server, dn = utils.parse_moniker (path)
    self.server = server.rstrip ("/")
    self.cls = cls = com_object.Class
    if cls not in self._schemas:
      schema_path = com_object.Schema
      try:
        #
        # Relying on the fact that by this time a valid
        # bind must have been created on this connection
        # which ADSI will reuse without prompting.
        #
        self._schemas[cls] = core.open_object (schema_path)
      except exc.BadPathnameError:
        self._schemas[cls] = None
    self.schema = self._schemas[cls]
    if self.schema:
      self.properties.update (self.schema.MandatoryProperties + self.schema.OptionalProperties)
    self.dn = self.distinguishedName

  def _put (self, name, value):
    #
    # The only way to clear an AD value is by using the
    # extended PutEx mechanism.
    #
    if value is None:
      exc.wrapped (self.com_object.PutEx, constants.ADS_PROPERTY.CLEAR, name, None)
    else:
      exc.wrapped (self.com_object.Put, name, value)

  @staticmethod
  def _munged_attribute (name):
    #
    # AD names are either camelCase or hyphen-separated, never underscored
    # Since Python identifiers can't include hypens but can include underscores,
    # translate underscores to hyphens.
    #
    return u"-".join (name.rstrip ("_").split (u"_"))

  def __getattr__ (self, name):
    #
    # Attempt to get the attribute by attribute access and the Get
    # method, with and without attribute name munging.
    #
    def _getattr (name):
      try:
        return exc.wrapped (getattr, self.com_object, name)
      except AttributeError:
        return exc.wrapped (self.com_object.Get, name)
    try:
      return _getattr (self._munged_attribute (name))
    except AttributeError:
      return _getattr (self._munged_attribute (name))

  def __setattr__ (self, name, value):
    munged_name = self._munged_attribute (name)
    if munged_name in self.properties:
      self._put (munged_name, value)
      exc.wrapped (self.com_object.SetInfo)
    else:
      super (ADBase, self).__setattr__ (name, value)

  def _get_object (self, rdn):
    container = exc.wrapped (self.com_object.QueryInterface, adsi.IID_IADsContainer)
    obj = exc.wrapped (container.GetObject, None, rdn)
    return adsi._get_good_ret (obj)

  def __getitem__ (self, rdn):
    return self.__class__ (self._get_object (rdn))

  def __setitem__ (self, rdn, info):
    #
    # If CopyHere were actually implemented in ADSI, this method could
    # be overloaded to call it, but only MoveHere is implemented and
    # it seemed counterintuitive to implement a move via the __setitem__
    # protocol.
    #
    try:
      cls = info.pop ('Class')
    except KeyError:
      raise exc.ActiveDirectoryError ("Must specify at least Class for new AD object")
    obj = exc.wrapped (self.com_object.Create, cls, rdn)
    exc.wrapped (obj.SetInfo)
    for k, v in info.items ():
      setattr (obj, k, v)
    exc.wrapped (obj.SetInfo)
    return self.__class__ (obj)

  def __delitem__ (self, rdn):
    #
    # Although the docs say you can pass NULL as the first param
    # to Delete, it doesn't appear to be supported. To keep the
    # interface in line, we'll do a GetObject (which does support
    # a NULL class) and then use the Class attribute to fill in
    # the Delete method.
    #
    exc.wrapped (self.Delete, self._get_object (rdn).Class, rdn)

  def __repr__ (self):
    return u"<%s: %s>" % (self.__class__.__name__, self.as_string ())

  def __str__ (self):
    return self.as_string ()

  def __iter__(self):
    try:
      for item in _ADContainer (self.com_object):
        yield self.__class__ (item)
    except NotAContainerError:
      raise TypeError ("%r is not iterable" % self)

  def __eq__ (self, other):
    return self.com_object.GUID == other.com_object.GUID

  def __hash__ (self):
    return hash (self.com_object.GUID)

  @classmethod
  def from_path (cls, path):
    ur"""Create an object of this class from an AD path and, optionally, credentials

    :param obj_or_path: a valid LDAP moniker
    :param cred: anything accepted by :func:`credentials.credentials`
    :returns: a :class:`ADBase` object
    """
    return cls (core.open_object (path))

  @classmethod
  def factory (cls, obj_or_path=None):
    ur"""Return an :class:`ADBase` object corresponding to `obj_or_path`.

    * If `obj_or_path` is an existing :class:`ADBase` object, return it
    * If `obj_or_path` is a Python COM object, return an :class:`ADBase` object which wraps it
    * Otherwise, assume that `obj_or_path` is an LDAP path and return the
      corresponding :class:`ADBase` object

    :param obj_or_path: an existing :class:`ADBase` object, a Python COM object or an LDAP moniker
    :param cred: anything accepted by :func:`credentials.credentials`
    :returns: a :class:`ADBase` object
    """
    if isinstance (obj_or_path, cls):
      return obj_or_path
    elif isinstance (obj_or_path, win32com.client.CDispatch):
      return cls (obj_or_path)
    else:
      return cls.from_path (obj_or_path)

  def as_string (self):
    return self.path

  def dump (self, ofile=sys.stdout):
    ur"""Pretty-print the contents of this object, starting with the
    AD class definition, and followed by the attributes of this particular
    instance.

    :param ofile: the open file to write output to [`sys.stdout`]
    """
    def munged (value):
      if isinstance (value, unicode):
        value = value.encode ("ascii", "backslashreplace")
      return value
    ofile.write (self.as_string () + u"\n")
    ofile.write ("[\n")
    for property in self.__class__._properties:
      value = exc.wrapped (getattr, self, property, None)
      if value:
        ofile.write ("  %s => %r\n" % (unicode (property).encode ("ascii", "backslashreplace"), munged (value)))
    ofile.write ("]\n")
    ofile.write ("{\n")
    for property in sorted (self.properties):
      value = exc.wrapped (getattr, self, property, None)
      if value:
        ofile.write ("  %s => %r\n" % (unicode (property).encode ("ascii", "backslashreplace"), munged (value)))
    ofile.write ("}\n")

  def get (self, attr):
    ur"""Force an attribute value to be read from AD, not from
    the local AD cache. (NB This is a system cache, not a Python
    one. No cacheing of attributes is done at Python level).

    :param attr: a valid attribute name for this object
    """
    exc.wrapped (self.com_object.GetInfo)
    return getattr (self, attr)

  def set (self, **kwargs):
    ur"""Set several properties at once. This should be slightly faster than setting
    the properties individually as SetInfo is called only once, at the end::

      from active_directory2 import ad
      user01 = ad.find_user ("user01")
      user01.set (displayName="User One", sAMAccountName="user-01")
    """
    for name, value in kwargs.items ():
      self._put (name, value)
    exc.wrapped (self.com_object.SetInfo)

  def move0 (self, rdn, elsewhere, new_rdn=None):
    ur"""Move a child object to another container, optionally
    renaming it on the way.

    :param rdn: the rdn of an object within this container
    :param elsewhere: another container
    :type elsewhere: anything accepted by :meth:`factory`
    :param new_rdn: the new rdn of the object if it is to be renamed as well
                    as moved.
    """
    elsewhere_obj = self.__class__.factory (elsewhere)
    exc.wrapped (
      elsewhere_obj.com_object.MoveHere,
      self._get_object (rdn).ADsPath,
      new_rdn or rdn
    )

  def move (self, elsewhere, new_rdn=None):
    ur"""Move an object to another container, optionally renaming it
    on the way.

    :param elsewhere: another container
    :type elsewhere: anything accepted by :meth:`factory`
    :param new_rdn: the rdn of the object in its new container if it is to
                    be renamed as well as moved
    """
    elsewhere_obj = self.__class__.factory (elsewhere)
    exc.wrapped (
      elsewhere_obj.com_object.MoveHere,
      self.path,
      new_rdn or self.Name
    )

  def rename0 (self, rdn, new_rdn):
    ur"""Rename a child within this container (the underlying action is
    a move to the same container).

    :param rdn: the rdn of an object within this container
    :param new_rdn: the new rdn of the object
    """
    self.move (rdn, self, new_rdn)

  def rename (self, new_rdn):
    ur"""Rename this object within its current container

    :param new_rdn: the new rdn of this object
    """
    self.move (self.parent, new_rdn)

  def delete (self):
    ur"""Delete this object and all its descendants. The :class:`ADBase`
    object will persist but any attempt to read its properties will fail.
    """
    exc.wrapped (self.com_object.QueryInterface, adsi.IID_IADsDeleteOps).DeleteObject (0)

  query = core.query

  def search (self, *args, **kwargs):
    """Return an iterator of :class:`ADBase` objects corresponding to
    the LDAP filter formed from the positional and keyword params.
    The :func:`active_directory2.core.and_` and :func:`active_directory2.core.or_`
    functions can be convenient ways of building up the filter.

    The filter is constructed as follows:

    * All params are and-ed together, producing a &(...) filter
    * Positional params are taken as-is (and so can be fully-fledged filters)
    * Keyword params become equi-filters of the form k=v

    So a call like this::

      obj.search (core.or_ (cn="tim", sn="golden"), "logonCount >= 0", objectCategory="person")

    would generate this filter::

      &(|((cn=tim)(sn=golden))(logonCount >= 0)(objectCategory=person))
    """
    if not (args or kwargs):
      raise NoFilterError
    filter = support.and_ (*args, **kwargs)
    for result in core.query (self.com_object, filter, ['distinguishedName']):
      rdn = support.rdn (self.distinguishedName, result['distinguishedName'][0])
      if not rdn:
        yield self
      else:
        yield self[rdn]

  def find (self, *args, **kwargs):
    ur"""Hand off arguments to :meth:`search` and return the first result
    """
    for result in self.search (*args, **kwargs):
      return result

  #
  # Common convenience functions
  #
  def find_user (self, name=None):
    ur"""Return the first user object matching `name`. Ambiguous name resolution
    is used, so `name` can match display name or account name. If no name is
    passed, the logged-on user is found.
    """
    name = name or exc.wrapped (win32api.GetUserName)
    return self.find (anr=name, objectClass="user", objectCategory="person")

  def find_computer (self, name=None):
    ur"""Return the first computer object matching `name`. Ambiguous name resolution
    is used, so `name` can match display name or account name. If no name is
    passed, this computer is found.
    """
    name = name or exc.wrapped (socket.gethostname)
    return self.find (anr=name, objectCategory="Computer")

  def find_group (self, name):
    ur"""Return the first group object matching `name`. Ambiguous name resolution
    is used, so `name` can match display name or account name.
    """
    return self.find (anr=name, objectCategory="group")

  def find_ou (self, name):
    ur"""Return the first organizational unit object matching `name`. Ambiguous name resolution
    is used, so `name` can match display name or account name.
    """
    return self.find (anr=name, objectCategory="organizationalUnit")

  def walk (self, level=0):
    ur"""Mimic the behaviour of Python `os.walk` iterator.

    :return: iterator of level, this container, containers, items
    """
    subordinates = [(s, s.schema.Container) for s in self]
    yield (
      level,
      self,
      (s for s, is_container in subordinates if is_container),
      (s for s, is_container in subordinates if not is_container)
    )
    for s, is_container in subordinates:
      if is_container:
        for walked in s.walk (level+1):
          yield walked

  def flat (self):
    ur"""Return a flat iteration over all items under this container.
    """
    for level, container, containers, items in self.walk ():
      for item in items:
        yield item

adbase = ADBase.factory
