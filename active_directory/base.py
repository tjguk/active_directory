# -*- coding: iso-8859-1 -*-
import os, sys

import win32api
from win32com import adsi

from . import core
from . import exc
from . import utils

class ADSimple (object):

  _properties = []

  def __init__ (self, obj):
    utils._set (self, u"com_object", obj)
    utils._set (self, u"properties", self._properties)
    self.path = obj.ADsPath

  def __getattr__ (self, name):
    try:
      return exc.wrapped (getattr, self.com_object, name)
    except AttributeError:
      try:
        return exc.wrapped (self.com_object.GetEx, name)
      except NotImplementedError:
        raise AttributeError

  def as_string (self):
    return self.path

  def dump (self, ofile=sys.stdout):
    def encode (text):
      if isinstance (text, unicode):
        return unicode (text).encode (sys.stdout.encoding, "backslashreplace")
      else:
        return text

    ofile.write (self.as_string () + u"\n")
    ofile.write ("{\n")
    for name in self.properties:
      try:
        value = getattr (self, name)
      except:
        raise
        value = "Unable to get value"
      if value:
        if isinstance (name, unicode):
          name = encode (name)
        if isinstance (value, (tuple, list)):
          value = "[(%d items)]" % len (value)
        if isinstance (value, unicode):
          value = encode (value)
          if len (value) > 60:
            value = value[:25] + "..." + value[-25:]
        ofile.write ("  %s => %s\n" % (encode (name), encode (value)))
    ofile.write ("}\n")

class ADBase (ADSimple):
  u"""Wrap an active-directory object for easier access
   to its properties and children. May be instantiated
   either directly from a COM object or from an ADs Path.

   Every IADs-derived object has at least the following attributes:

   Name, Class, GUID, ADsPath, Parent, Schema

   eg,

     import active_directory as ad
     users = ad.ad ("LDAP://cn=Users,DC=gb,DC=vo,DC=local")
  """

  _default_properties = [u"Name", u"Class", u"GUID", u"ADsPath", u"Parent", u"Schema"]
  _schema_cache = {}

  def __init__ (self, obj, username=None, password=None, parse_schema=True):
    super (ADBase, self).__init__ (obj)
    schema = None
    if parse_schema:
      try:
        schema = exc.wrapped (adsi.ADsGetObject, exc.wrapped (getattr, obj, u"Schema", None))
      except exc.ActiveDirectoryError:
        schema = None
    properties, is_container = self._schema (schema)
    utils._set (self, u"properties", properties)
    self.is_container = is_container

    #
    # At this point, __getattr__ & __setattr__ have enough
    # to decide whether an attribute belongs to the delegated
    # object or not.
    #
    self.username = username
    self.password = password
    self.connection = core.connect (username=username, password=password)
    self.dn = exc.wrapped (getattr, self.com_object, u"distinguishedName", None) or self.com_object.name
    self._delegate_map = dict ()

  def __getitem__ (self, key):
    return getattr (self, key)

  def __getattr__ (self, name):
    #
    # Special-case find_... methods to search for
    # corresponding object types.
    #
    if name.startswith (u"find_"):
      names = name[len (u"find_"):].lower ().split ("_")
      first, rest = names[0], names[1:]
      object_class = "".join ([first] + [n.title () for n in rest])
      return self._find (object_class)

    if name.startswith (u"search_"):
      names = name[len (u"search_"):].lower ().split ("_")
      first, rest = names[0], names[1:]
      object_class = u"".join ([first] + [n.title () for n in rest])
      return self._search (object_class)

    if name.startswith (u"get_"):
      names = name[len (u"get_"):].lower ().split (u"_")
      first, rest = names[0], names[1:]
      object_class = u"".join ([first] + [n.title () for n in rest])
      return self._get (object_class)

    #
    # Allow access to object's properties as though normal
    # Python instance properties. Some properties are accessed
    # directly through the object, others by calling its Get
    # method. Not clear why.
    #
    if name not in self._delegate_map:
      value = super (ADBase, self).__getattr__ (name)
      converter = get_converter (name)
      self._delegate_map[name] = converter (value)
    return self._delegate_map[name]

  def __setitem__ (self, key, value):
    from_ad, to_ad = types.types.get (name, (None, None))
    if to_ad:
      setattr (self, key, converter (value))
    else:
      setattr (self, key, value)

  def __setattr__ (self, name, value):
    #
    # Allow attribute access to the underlying object's
    #  fields.
    #
    if name in self.properties:
      exc.wrapped (self.com_object.Put, name, value)
      exc.wrapped (self.com_object.SetInfo)
      #
      # Invalidate to ensure map is refreshed on next get
      #
      if name in self._delegate_map:
        del self._delegate_map[name]
    else:
      super (ADBase, self).__setattr__ (name, value)

  def as_string (self):
    return self.path

  def __str__ (self):
    return self.as_string ()

  def __repr__ (self):
    return u"<%s: %s>" % (exc.wrapped (getattr, self.com_object, u"Class") or u"AD", self.dn)

  def __eq__ (self, other):
    return self.com_object.Guid == other.com_object.Guid

  def __hash__ (self):
    return hash (self.com_object.Guid)

  class AD_iterator:
    u""" Inner class for wrapping iterated objects
    (This class and the __iter__ method supplied by
    Stian Søiland <stian@soiland.no>)
    """
    def __init__ (self, com_object):
      self._iter = iter (com_object)
    def __iter__ (self):
      return self
    def next (self):
      return ad (self._iter.next ())

  def __iter__(self):
    return self.AD_iterator (self.com_object)

  def _get_parent (self):
    return self.__class__ (self.com_object.Parent)
  parent = property (_get_parent)

  @classmethod
  def _schema (cls, cschema):
    if cschema is None:
      return cls._default_properties, False

    if cschema.ADsPath not in cls._schema_cache:
      properties = \
        exc.wrapped (getattr, cschema, u"mandatoryProperties", []) + \
        exc.wrapped (getattr, cschema, u"optionalProperties", [])
      cls._schema_cache[cschema.ADsPath] = properties, exc.wrapped (getattr, cschema, u"Container", False)
    return cls._schema_cache[cschema.ADsPath]

  def refresh (self):
    exc.wrapped (self.com_object.GetInfo)

  def walk (self):
    u"""Analogous to os.walk, traverse this AD subtree,
    depth-first, and yield for each container:

    container, containers, items
    """
    children = list (self)
    this_containers = [c for c in children if c.is_container]
    this_items = [c for c in children if not c.is_container]
    yield self, this_containers, this_items
    for c in this_containers:
      for container, containers, items in c.walk ():
        yield container, containers, items

  def flat (self):
    for container, containers, items in self.walk ():
      for item in items:
        yield item

  def set (self, **kwds):
    u"""Set a number of values at one time. Should be
     a little more efficient than assigning properties
     one after another.

    eg,

      import active_directory
      user = active_directory.find_user ("goldent")
      user.set (displayName = "Tim Golden", description="SQL Developer")
    """
    for k, v in kwds.items ():
      exc.wrapped (self.com_object.Put, k, v)
    exc.wrapped (self.com_object.SetInfo)

  def _find (self, object_class):
    u"""Helper function to allow general-purpose searching for
    objects of a class by calling a .find_xxx_yyy method.
    """
    def _find (name):
      for item in self.search (objectClass=object_class, name=name):
        return item
    return _find

  def _search (self, object_class):
    u"""Helper function to allow general-purpose searching for
    objects of a class by calling a .search_xxx_yyy method.
    """
    def _search (*args, **kwargs):
      return self.search (objectClass=object_class, *args, **kwargs)
    return _search

  def _get (self, object_class):
    u"""Helper function to allow general-purpose retrieval of a
    child object by class.
    """
    def _get (rdn):
      return self.get (object_class, rdn)
    return _get

  def find (self, name):
    for item in self.search (name=name):
      return item

  def find_user (self, name=None):
    u"""Make a special case of (the common need of) finding a user
    either by username or by display name
    """
    name = name or exc.wrapped (win32api.GetUserName)
    filter = core.and_ (
      core.or_ (sAMAccountName=name, displayName=name, cn=name),
      sAMAccountType=constants.SAM_ACCOUNT_TYPES.USER_OBJECT
    )
    for user in self.search (filter):
      return user

  def find_ou (self, name):
    u"""Convenient alias for find_organizational_unit"""
    return self.find_organizational_unit (name)

  def search (self, *args, **kwargs):
    filter = core.and_ (*args, **kwargs)
    #~ query_string = core.qs (base=self.ADsPath, filter=filter, attributes=["objectGuid"])
    query_string = u"<%s>;(%s);objectGuid;Subtree" % (self.ADsPath, filter)
    for result in core.query (query_string, connection=self.connection):
      guid = u"".join (u"%02X" % ord (i) for i in result['objectGuid'])
      yield ad (u"LDAP://<GUID=%s>" % guid, username=self.username, password=self.password)

  def get (self, object_class, relative_path):
    return ad (exc.wrapped (self.com_object.GetObject, object_class, relative_path))

  def new_ou (self, name, description=None, **kwargs):
    obj = exc.wrapped (self.com_object.Create, u"organizationalUnit", u"ou=%s" % name)
    exc.wrapped (obj.Put, u"description", description or name)
    exc.wrapped (obj.SetInfo)
    for name, value in kwargs.items ():
      exc.wrapped (obj.Put, name, value)
    exc.wrapped (obj.SetInfo)
    return ad (obj)

  def new_group (self, name, type=constants.GROUP_TYPES.DOMAIN_LOCAL | constants.GROUP_TYPES.SECURITY_ENABLED, **kwargs):
    obj = exc.wrapped (self.com_object.Create, u"group", u"cn=%s" % name)
    exc.wrapped (obj.Put, u"sAMAccountName", name)
    exc.wrapped (obj.Put, u"groupType", type)
    exc.wrapped (obj.SetInfo)
    for name, value in kwargs.items ():
      exc.wrapped (obj.Put, name, value)
    exc.wrapped (obj.SetInfo)
    return ad (obj)

  def new (self, object_class, sam_account_name, **kwargs):
    obj = exc.wrapped (self.com_object.Create, object_class, u"cn=%s" % sam_account_name)
    exc.wrapped (obj.Put, u"sAMAccountName", sam_account_name)
    exc.wrapped (obj.SetInfo)
    for name, value in kwargs.items ():
      exc.wrapped (obj.Put, name, value)
    exc.wrapped (obj.SetInfo)
    return ad (obj)

class WinNT (ADBase):

  def __eq__ (self, other):
    return self.com_object.ADsPath.lower () == other.com_object.ADsPath.lower ()

  def __hash__ (self):
    return hash (self.com_object.ADsPath.lower ())

class Group (ADBase):

  def _get_members (self):
    return _Members (self)
  def _set_members (self, members):
    original = self.members
    new_members = set (ad (m) for m in members)
    print u"original", original
    print u"new members", new_members
    print u"new_members - original", new_members - original
    for member in (new_members - original):
      print u"Adding", member
      exc.wrapped (self.com_object.Add, member.AdsPath)
    print u"original - new_members", original - new_members
    for member in (original - new_members):
      print u"Removing", member
      exc.wrapped (self.com_object.Remove, member.AdsPath)
  members = property (_get_members, _set_members)

  def walk (self):
    """Override the usual .walk method by returning instead:

    group, groups, users
    """
    members = self.members
    groups = [m for m in members if m.Class == u'group']
    users = [m for m in members if m.Class == u'user']
    yield (self, groups, users)
    for group in groups:
      for result in group.walk ():
        yield result

  def flat (self):
    for group, groups, members in self.walk ():
      for member in members:
        yield member

class WinNTGroup (WinNT, Group):
  pass

def namespaces ():
  return ADBase (adsi.ADsGetObject (u"ADs:"), parse_schema=False)

_CLASS_MAP = {
  u"group" : Group,
}
_WINNT_CLASS_MAP = {
  u"group" : WinNTGroup
}
_namespace_names = None
def ad (obj_or_path, username=None, password=None):
  u"""Factory function for suitably-classed Active Directory
  objects from an incoming path or object. NB The interface
  is now  intended to be:

    ad (obj_or_path)

  @param obj_or_path Either an COM AD object or the path to one. If
  the path doesn't start with "LDAP://" this will be prepended.

  @return An _AD_object or a subclass proxying for the AD object
  """
  if isinstance (obj_or_path, ADBase):
    return obj_or_path

  global _namespace_names
  if _namespace_names is None:
    _namespace_names = [u"GC:"] + [ns.Name for ns in adsi.ADsGetObject (u"ADs:")]
  matcher = re.compile ("(" + "|".join (_namespace_names)+ ")?(//)?([A-za-z0-9-_]+/)?(.*)")
  if isinstance (obj_or_path, basestring):
    #
    # Special-case the "ADs:" moniker which isn't a child of IADs
    #
    if obj_or_path == u"ADs:":
      return namespaces ()

    scheme, slashes, server, dn = matcher.match (obj_or_path).groups ()
    if scheme is None:
        scheme, slashes = u"LDAP:", u"//"
    if scheme == u"WinNT:":
      moniker = dn
    else:
      moniker = escaped_moniker (dn)
    obj_path = scheme + (slashes or u"") + (server or u"") + (moniker or u"")
    obj = exc.wrapped (adsi.ADsOpenObject, obj_path, username, password, DEFAULT_BIND_FLAGS)
  else:
    obj = obj_or_path
    scheme, slashes, server, dn = matcher.match (obj_or_path.AdsPath).groups ()

  if dn == u"rootDSE":
    return ADBase (obj, username, password, parse_schema=False)

  if scheme == u"WinNT:":
    class_map = _WINNT_CLASS_MAP.get (obj.Class.lower (), WinNT)
  else:
    class_map = _CLASS_MAP.get (obj.Class.lower (), ADBase)
  return class_map (obj)
