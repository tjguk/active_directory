# -*- coding: iso-8859-1 -*-
import os, sys

from win32com.adsi import adsi

from . import core
from . import constants
from . import exc
from . import utils

class NotAContainerError (exc.ActiveDirectoryError):
  pass

class ADContainer (object):
  """A support object which takes an existing AD COM object
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
          yield exc.wrapped (item.QueryInterface, adsi.IID_IADs)
      else:
        break

class ADBase (object):
  """A slender wrapper around an AD COM object.

  Attributes can be read, set and cleared. If the underlying object is a
  container it can be iterated over and subobjects can be retrieved, added
  and removed. It can also be walked (in the style of Python's os.walk) and
  flattened: :meth:`walk`, :meth:`flat`.

  Pretty-printing to stdout is via the :meth:`dump` method.

  There is simple searching with filters, returning the
  results as objects if its own type: :meth:`search`, :meth:`find`.

  The underlying object can itself be deleted via :meth:`delete`

  No validation is done of the attribute names and no conversions of the
  values. It can be used alone (most easily via the :func:`adbase` function
  which takes an AD path and returns an ADBase object). It also provides the
  basis for the :class:`ADObject` class.
  """

  #
  # For speed, hardcode the known properties of the IADs class
  #
  _properties = ["ADsPath", "Class", "GUID", "Name", "Parent", "Schema"]
  _schemas = {}

  def __init__ (self, obj, cred=None):
    utils._set (self, "properties", set ())
    self.com_object = com_object = obj.QueryInterface (adsi.IID_IADs)
    self.cred = cred
    self.path = com_object.ADsPath
    scheme, server, dn = utils.parse_moniker (com_object.ADsPath)
    self.server = server.rstrip ("/")
    self.cls = cls = exc.wrapped (getattr, com_object, "Class")
    if cls not in self._schemas:
      schema_path = exc.wrapped (getattr, com_object, "Schema")
      try:
        self._schemas[cls] = core.open_object (schema_path, cred=cred)
      except exc.BadPathnameError:
        self._schemas[cls] = None
    self.schema = self._schemas[cls]
    if self.schema:
      self.properties.update (self.schema.MandatoryProperties + self.schema.OptionalProperties)

  def _put (self, name, value):
    operation = constants.ADS_PROPERTY.CLEAR if value is None else constants.ADS_PROPERTY.UPDATE
    exc.wrapped (self.com_object.PutEx, operation, name, value)

  def __getattr__ (self, name):
    #
    # AD names are either camelCase or hyphen-separated, never underscored
    # Since Python identifiers can't include hypens but can
    # include underscores, translate underscores to hyphens.
    #
    name = "_".join (name.split ("-"))
    try:
      return exc.wrapped (getattr, self.com_object, name)
    except (AttributeError, NotImplementedError):
      try:
        return exc.wrapped (self.com_object.Get, name)
      except NotImplementedError:
        raise AttributeError

  def __setattr__ (self, name, value):
    if name in self.properties:
      self._put (name, value)
      exc.wrapped (self.com_object.SetInfo)
    else:
      super (ADBase, self).__setattr__ (name, value)

  def __delattr__ (self, name):
    self._put (name, None)
    exc.wrapped (self.com_object.SetInfo)

  def _item_identifier (self, ad_class, item_identifier):
    item_namer = core.class_schema (ad_class, self.server, self.cred).NamingProperties
    if item_identifier.startswith ("%s=" % item_namer):
      return item_identifier
    else:
      return "%s=%s" % (item_namer, item_identifier)

  def __getitem__ (self, item):
    item_type, item_identifier = item
    item_identifier = self._item_identifier (item_type, item_identifier)
    container = exc.wrapped (self.com_object.QueryInterface, adsi.IID_IADsContainer)
    obj = exc.wrapped (container.GetObject, item_type, item_identifier)
    return self.__class__ (obj, self.cred)

  def __setitem__ (self, item, info):
    item_type, item_identifier = item
    item_identifier = self._item_identifier (item_type, item_identifier)
    obj = exc.wrapped (self.com_object.Create, item_type, item_identifier)
    exc.wrapped (obj.SetInfo)
    for k, v in info.items ():
      setattr (obj, k, v)
    exc.wrapped (obj.SetInfo)
    return self.__class__ (obj, self.cred)

  def __delitem__ (self, item):
    item_type, item_identifier = item
    item_identifier = self._item_identifier (item_type, item_identifier)
    exc.wrapped (self.com_object.Delete, item_type, item_identifier)

  def as_string (self):
    return self.path

  def __repr__ (self):
    return u"<%s: %s>" % (self.__class__.__name__, self.as_string ())

  def __str__ (self):
    return self.as_string ()

  def __iter__(self):
    try:
      for item in ADContainer (self.com_object):
        yield self.__class__ (item, self.cred)
    except NotAContainerError:
      raise TypeError ("%r is not iterable" % self)

  def __eq__ (self, other):
    return self.com_object.GUID == other.com_object.GUID

  def __hash__ (self):
    return hash (self.com_object.GUID)

  @classmethod
  def from_path (cls, path, cred=None):
    return cls (core.open_object (path, cred))

  def delete (self):
    """Delete this object and all its descendants. The :class:`ADBase`
    object will persist but any attempt to read its properties will fail.
    """
    exc.wrapped (self.com_object.QueryInterface, adsi.IID_IADsDeleteOps).DeleteObject (0)

  query = core.dquery

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
    filter = core.and_ (*args, **kwargs)
    for result in self.query (filter, ['ADsPath']):
      yield self.__class__ (core.open_object (result['ADsPath'][0], cred=self.cred))

  def find (self, *args, **kwargs):
    for result in self.search (*args, **kwargs):
      return result

  def walk (self, level=0):
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
    for level, container, containers, items in self.walk ():
      for item in items:
        yield item

  def as_string (self):
    return self.path

  def dump (self, ofile=sys.stdout):
    def munged (value):
      if isinstance (value, unicode):
        value = value.encode ("ascii", "backslashreplace")
      return value
    ofile.write (self.as_string () + u"\n")
    ofile.write ("[\n")
    for property in self._properties:
      value = exc.wrapped (getattr, self, property, None)
      if value:
        ofile.write ("  %s => %s\n" % (unicode (property).encode ("ascii", "backslashreplace"), munged (value)))
    ofile.write ("]\n")
    ofile.write ("{\n")
    for property in sorted (self.properties):
      value = exc.wrapped (getattr, self, property, None)
      if value:
        ofile.write ("  %s => %s\n" % (unicode (property).encode ("ascii", "backslashreplace"), munged (value)))
    ofile.write ("}\n")

def adbase (obj_or_path=None, cred=None):
  if obj_or_path is None:
    return ADBase (core.root_obj (), cred=cred)
  elif isinstance (obj_or_path, ADBase):
    return obj_or_path
  else:
    return ADBase.from_path (obj_or_path, cred)
