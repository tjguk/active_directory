# -*- coding: iso-8859-1 -*-
u"""active_directory - a lightweight wrapper around COM support
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

Key functions are:

* :func:`connection`, :func:`query` and :func:`query_string` - these offer the
  most raw functionality: slightly assisting an ADO query and returning a
  Python dictionary of results::

    import datetime
    import active_directory as ad
    #
    # Find all objects created this month in creation order
    #
    this_month = datetime.date.today ().replace (day=1)
    query_string = ad.query_string (
      filter=ad.schema.whenCreated >= this_month,
      attributes=["distinguishedName", "whenCreated"]
    )
    for new_object in ad.query (query_string, sort_on="whenCreated"):
      print "%(distinguishedName)s => %(whenCreated)s" % new_object

* :func:`ad` - this is the wrap-all function which transforms an LDAP: moniker
  into a Python object which offers the existing properties and members in
  Pythonic wrappers. It will also convert an existing LDAP COM Object::

    import active_directory as ad

    me =

* :func:`find_user`, :func:`find_group`, :func:`find_ou` - these are module-level
  convenience functions which each return a Python object corresponding to the
  user, group or ou of the name passed in::

    import active_directory as ad

    camden_users = (obj for obj in ad.find_ou ("Camden") if obj.Class == "User")

* The active directory class (ADBase or a subclass) will determine
  its properties and allow you to access them as instance properties::

     import active_directory as ad
     goldent = ad.find_user ("goldent")
     print goldent.displayName

* Any object returned by the AD object's operations is itself
  wrapped as an AD object so you get the same benefits::

    import active_directory as  ad
    users = ad.root ().child ("cn=users")
    for user in users.search (displayName='Tim*'):
      print user.displayName

* To search the AD, there are two module-level general
  search functions, and module-level convenience functions
  to find a user, computer etc. Usage is illustrated below::

   import active_directory as ad

   for user in ad.search (
     objectClass='User',
     ad.core.or_ (displayName='Tim Golden', sAMAccountName='goldent')
   ):
     #
     # This search returns an ADUser object
     #
     print user

* Typical usage will be::

    import active_directory as ad

    for computer in ad.search (objectClass='computer'):
      print computer.displayName

(c) Tim Golden <mail@timgolden.me.uk> October 2004-2010
Licensed under the (GPL-compatible) MIT License:
http://www.opensource.org/licenses/mit-license.php

Many thanks, obviously, to Mark Hammond for creating
the pywin32 extensions without which this wouldn't
have been possible. (Or would at least have been much
more work...)
"""
__VERSION__ = u"1.0rc1"

import os, sys
import datetime
import logging
import re
import struct

import pythoncom
import pywintypes
import win32api
import win32com.client
import win32security
from win32com import adsi
from win32com.adsi import adsicon

from . import attributes
from . import converters
from . import constants
from . import core
from . import exc
from . import utils

logger = logging.getLogger ("active_directory")
def enable_debugging ():
  logger.addHandler (logging.StreamHandler (sys.stdout))
  logger.setLevel (logging.DEBUG)

try:
  import collections
  SetBase = collections.MutableSet
except (ImportError, AttributeError):
  logger.warn ("Unable to use collections.MutableSet; using object instead")
  SetBase = object

DEFAULT_BIND_FLAGS = adsicon.ADS_SECURE_AUTHENTICATION

def get_converter (name):
  if name not in converters.converters:
    obj = None ## attribute (name)
    if obj and obj.attributeSyntax in TYPE_CONVERTERS:
      converters.register_converter (name, from_ad=TYPE_CONVERTERS[obj.attributeSyntax])
    elif name.endswith ("GUID"):
      converters.register_converter (name, from_ad=converters.convert_to_guid)
  from_ad, _ = converters.converters.get (name, (None, None))
  return from_ad or (lambda x : x)

def attribute (attribute_name, root=None):
  root_dse = root or win32com.client.GetObject ("LDAP://rootDSE")
  schemaNamingContext = root_dse.Get ("schemaNamingContext")
  qs = core.query_string (
    base="LDAP://%s" % schemaNamingContext,
    filter="ldapDisplayName=%s" % attribute_name,
    attributes="*"
  )
  for item in core.query (qs):
    return item
  else:
    return {}

schema = attributes._Attributes ()

def search_ex (query_string=u"", username=None, password=None):
  u"""FIXME: Historical version of :func:`query`"""
  return core.query (query_string, connection=connect (username, password))

class _Members (set):

  def __init__ (self, group):
    super (_Members, self).__init__ (ad (i) for i in iter (exc.wrapped (group.com_object.members)))
    self._group = group

  def _effect (self, original):
    group = self._group.com_object
    for member in (self - original):
      print u"Adding", member
      #~ group.Add (member.AdsPath)
      exc.wrapped (group.Add, member.AdsPath)
    for member in (original - self):
      print u"Removing", member
      #~ group.Remove (member.AdsPath)
      exc.wrapped (group.Remove, member.AdsPath)

  def update (self, *others):
    original = set (self)
    for other in others:
      super (_Members, self).update (ad (o) for o in other)
    self._effect (original)

  def __ior__ (self, other):
    return self.update (other)

  def intersection_update (self, *others):
    original = set (self)
    for other in others:
      super (_Members, self).intersection_update (ad (o) for o in other)
    self._effect (original)

  def __iand__ (self, other):
    return self.intersection_update (self, other)

  def difference_update (self, *others):
    original = set (self)
    for other in others:
      self.difference_update (ad (o) for o in other)
    self._effect (original)

  def symmetric_difference_update (self, *others):
    original = set (self)
    for other in others:
      self.symmetric_difference_update (ad (o) for o in others)
    self._effect (original)

  def add (self, elem):
    original = set (self)
    result = super (_Members, self).add (ad (elem))
    self._effect (original)
    return result

  def remove (self, elem):
    original = set (self)
    result = super (_Members, self).remove (ad (elem))
    self._effect (original)
    return result

  def discard (self, elem):
    original = set (self)
    result = super (_Members, self).discard (ad (elem))
    self._effect (original)
    return result

  def pop (self):
    original = set (self)
    result = super (_Members, self).pop ()
    self._effect (original)
    return result

  def clear (self):
    original = set (self)
    super (_Members, self).clear ()
    self._effect (original)

  def __contains__ (self, element):
    return  super (_Members, self).__contains__ (ad (element))

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


class RootDSE (ADSimple):

  _properties = u"""configurationNamingContext
currentTime
defaultNamingContext
dnsHostName
domainControllerFunctionality
domainFunctionality
dsServiceName
forestFunctionality
highestCommittedUSN
isGlobalCatalogReady
isSynchronized
ldapServiceName
namingContexts
rootDomainNamingContext
schemaNamingContext
serverName
subschemaSubentry
supportedCapabilities
supportedControl
supportedLDAPPolicies
supportedLDAPVersion
supportedSASLMechanisms
  """.split ()

#~ ROOT_DSE = RootDSE (exc.wrapped (adsi.ADsGetObject, "LDAP://rootDSE"))

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
      except ActiveDirectoryError:
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
    from_ad, to_ad = converters.get (name, (None, None))
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
    return ad (self.com_object.Parent)
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
    name = name or exc.wrapped (win32api.GetUserName )
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

_CLASS_MAP = {
  u"group" : Group,
}
_WINNT_CLASS_MAP = {
  u"group" : WinNTGroup
}
def escaped_moniker (moniker):
  #
  # If the moniker *appears* to have been escaped
  # already, return it straight. This is obviously
  # fragile but seems to work for now.
  #
  if moniker.find (u"\\/") > -1:
    return moniker
  else:
    return moniker.replace (u"/", u"\\/")

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
AD_object = ad

def AD (server=None, username=None, password=None, use_gc=False):
  if use_gc:
    scheme = u"GC://"
  else:
    scheme = u"LDAP://"
  if server:
    root_moniker = scheme + server + u"/rootDSE"
  else:
    root_moniker = scheme + u"rootDSE"
  root_obj = exc.wrapped (adsi.ADsOpenObject, root_moniker, username, password, DEFAULT_BIND_FLAGS)
  default_naming_context = root_obj.Get (u"defaultNamingContext")
  moniker = scheme + default_naming_context
  obj = exc.wrapped (adsi.ADsOpenObject, moniker, username, password, DEFAULT_BIND_FLAGS)
  return ad (obj, username, password)


#
# Convenience functions for common needs
#
def find (name):
  return root ().find (name)

def find_user (name=None):
  return root ().find_user (name)

def find_computer (name=None):
  return root ().find_computer (name)

def find_group (name):
  return root ().find_group (name)

def find_ou (name):
  return root ().find_ou (name)

def find_public_folder (name):
  return root ().find_public_folder (name)

def search (*args, **kwargs):
  return root ().search (*args, **kwargs)

#
# root returns a cached object referring to the
#  root of the logged-on active directory tree.
#
_ad = None
def root (username=None, password=None):
  global _ad
  if _ad is None:
    _ad = AD (username=username, password=password)
  return _ad

def namespaces ():
  return ADBase (adsi.ADsGetObject (u"ADs:"), parse_schema=False)

def root_dse (username=None, password=None):
  return RootDSE (adsi.ADsOpenObject (u"LDAP://rootDSE", username, password, DEFAULT_BIND_FLAGS))



#
# Conversions
#
_PROPERTY_MAP = dict (
  accountExpires = converters.convert_to_datetime,
  auditingPolicy = converters.convert_to_hex,
  badPasswordTime = converters.convert_to_datetime,
  creationTime = converters.convert_to_datetime,
  dSASignature = converters.convert_to_hex,
  forceLogoff = converters.convert_to_datetime,
  #~ fSMORoleOwner = converters.convert_to_object (ad),
  groupType = converters.convert_to_flags (constants.GROUP_TYPES),
  isGlobalCatalogReady = converters.convert_to_boolean,
  isSynchronized = converters.convert_to_boolean,
  lastLogoff = converters.convert_to_datetime,
  lastLogon = converters.convert_to_datetime,
  lastLogonTimestamp = converters.convert_to_datetime,
  lockoutDuration = converters.convert_to_datetime,
  lockoutObservationWindow = converters.convert_to_datetime,
  lockoutTime = converters.convert_to_datetime,
  #~ manager = converters.convert_to_object (ad),
  #~ masteredBy = converters.convert_to_objects (ad),
  maxPwdAge = converters.convert_to_datetime,
  #~ member = converters.convert_to_objects (ad),
  #~ memberOf = converters.convert_to_objects (ad),
  minPwdAge = converters.convert_to_datetime,
  modifiedCount = converters.convert_to_datetime,
  modifiedCountAtLastProm = converters.convert_to_datetime,
  #~ msExchMailboxGuid = converters.convert_to_guid,
  #~ schemaIDGUID = converters.convert_to_guid,
  mSMQDigests = converters.convert_to_hex,
  mSMQSignCertificates = converters.convert_to_hex,
  objectClass = converters.convert_to_breadcrumbs,
  #~ objectGUID = converters.convert_to_guid,
  objectSid = converters.convert_to_sid,
  #~ publicDelegates = converters.convert_to_objects (ad),
  #~ publicDelegatesBL = converters.convert_to_objects (ad),
  pwdLastSet = converters.convert_to_datetime,
  replicationSignature = converters.convert_to_hex,
  replUpToDateVector = converters.convert_to_hex,
  repsFrom = converters.convert_to_hexes,
  repsTo = converters.convert_to_hex,
  sAMAccountType = converters.convert_to_enum (constants.SAM_ACCOUNT_TYPES),
  #~ subRefs = converters.convert_to_objects (ad),
  systemFlags = converters.convert_to_flags (constants.ADS_SYSTEMFLAG),
  userAccountControl = converters.convert_to_flags (constants.USER_ACCOUNT_CONTROL),
  #~ wellKnownObjects = converters.convert_to_objects (ad),
  whenCreated = converters.convert_pytime_to_datetime,
  whenChanged = converters.convert_pytime_to_datetime,
)
#~ _PROPERTY_MAP[u'msDs-masteredBy'] = converters.convert_to_objects (ad)

for k, v in _PROPERTY_MAP.items ():
  converters.register_converter (k, from_ad=v)

_PROPERTY_MAP_IN = dict (
  accountExpires = converters.convert_from_datetime,
  badPasswordTime = converters.convert_from_datetime,
  creationTime = converters.convert_from_datetime,
  dSASignature = converters.convert_from_hex,
  forceLogoff = converters.convert_from_datetime,
  fSMORoleOwner = converters.convert_from_object,
  groupType = converters.convert_from_flags (constants.GROUP_TYPES),
  lastLogoff = converters.convert_from_datetime,
  lastLogon = converters.convert_from_datetime,
  lastLogonTimestamp = converters.convert_from_datetime,
  lockoutDuration = converters.convert_from_datetime,
  lockoutObservationWindow = converters.convert_from_datetime,
  lockoutTime = converters.convert_from_datetime,
  masteredBy = converters.convert_from_objects,
  maxPwdAge = converters.convert_from_datetime,
  member = converters.convert_from_objects,
  memberOf = converters.convert_from_objects,
  minPwdAge = converters.convert_from_datetime,
  modifiedCount = converters.convert_from_datetime,
  modifiedCountAtLastProm = converters.convert_from_datetime,
  msExchMailboxGuid = converters.convert_from_guid,
  #~ objectGUID = converters.convert_from_guid,
  objectSid = converters.convert_from_sid,
  publicDelegates = converters.convert_from_objects,
  publicDelegatesBL = converters.convert_from_objects,
  pwdLastSet = converters.convert_from_datetime,
  replicationSignature = converters.convert_from_hex,
  replUpToDateVector = converters.convert_from_hex,
  repsFrom = converters.convert_from_hex,
  repsTo = converters.convert_from_hex,
  sAMAccountType = converters.convert_from_enum (constants.SAM_ACCOUNT_TYPES),
  subRefs = converters.convert_from_objects,
  userAccountControl = converters.convert_from_flags (constants.USER_ACCOUNT_CONTROL),
  wellKnownObjects = converters.convert_from_objects
)
_PROPERTY_MAP_IN['msDs-masteredBy'] = converters.convert_from_objects

for k, v in _PROPERTY_MAP_IN.items ():
  converters.register_converter (k, to_ad=v)

