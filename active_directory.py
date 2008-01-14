# -*- coding: iso-8859-1 -*-
"""active_directory - a lightweight wrapper around COM support
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

(c) Tim Golden <active-directory@timgolden.me.uk> October 2004
Licensed under the (GPL-compatible) MIT License:
http://www.opensource.org/licenses/mit-license.php

Many thanks, obviously to Mark Hammond for creating
the pywin32 extensions without which this wouldn't
have been possible.
"""
from __future__ import generators

__VERSION__ = "0.7"

import os, sys
import datetime

import win32api
from win32com.client import Dispatch, GetObject
import win32security

#
# Code contributed by Stian Søiland <stian@soiland.no>
#
def i32(x):
  """Converts a long (for instance 0x80005000L) to a signed 32-bit-int.

  Python2.4 will convert numbers >= 0x80005000 to large numbers
  instead of negative ints.    This is not what we want for
  typical win32 constants.

  Usage:
      >>> i32(0x80005000L)
      -2147363168
  """
  # x > 0x80000000L should be negative, such that:
  # i32(0x80000000L) -> -2147483648L
  # i32(0x80000001L) -> -2147483647L     etc.
  return (x&0x80000000L and -2*0x40000000 or 0) + int(x&0x7fffffff)

#
# For ease of presentation, ms-style constant lists are
# held as Enum objects, allowing access by number or
# by name, and by name-as-attribute. This means you can do, eg:
#
# print GROUP_TYPES[2]
# print GROUP_TYPES['GLOBAL_GROUP']
# print GROUP_TYPES.GLOBAL_GROUP
#
# The first is useful when displaying the contents
# of an AD object; the other two when you want a more
# readable piece of code, without magic numbers.
#
class Enum (object):

  def __init__ (self, **kwargs):
    self._name_map = {}
    self._number_map = {}
    for k, v in kwargs.items ():
      self._name_map[k] = i32 (v)
      self._number_map[i32 (v)] = k

  def __getitem__ (self, item):
    try:
      return self._name_map[item]
    except KeyError:
      return self._number_map[i32 (item)]
      
  def __getattr__ (self, attr):
    try:
      return self._name_map[attr]
    except KeyError:
      raise AttributeError

  def item_names (self):
    return self._name_map.items ()

  def item_numbers (self):
    return self._number_map.items ()

GROUP_TYPES = Enum (
  GLOBAL_GROUP = 0x00000002,
  DOMAIN_LOCAL_GROUP = 0x00000004,
  LOCAL_GROUP = 0x00000004,
  UNIVERSAL_GROUP = 0x00000008,
  SECURITY_ENABLED = 0x80000000
)

AUTHENTICATION_TYPES = Enum (
  SECURE_AUTHENTICATION = i32 (0x01),
  USE_ENCRYPTION = i32 (0x02),
  USE_SSL = i32 (0x02),
  READONLY_SERVER = i32 (0x04),
  PROMPT_CREDENTIALS = i32 (0x08),
  NO_AUTHENTICATION = i32 (0x10),
  FAST_BIND = i32 (0x20),
  USE_SIGNING = i32 (0x40),
  USE_SEALING = i32 (0x80),
  USE_DELEGATION = i32 (0x100),
  SERVER_BIND = i32 (0x200),
  AUTH_RESERVED = i32 (0x800000000)
)

SAM_ACCOUNT_TYPES = Enum (
  SAM_DOMAIN_OBJECT = 0x0 ,
  SAM_GROUP_OBJECT = 0x10000000 ,
  SAM_NON_SECURITY_GROUP_OBJECT = 0x10000001 ,
  SAM_ALIAS_OBJECT = 0x20000000 ,
  SAM_NON_SECURITY_ALIAS_OBJECT = 0x20000001 ,
  SAM_USER_OBJECT = 0x30000000 ,
  SAM_NORMAL_USER_ACCOUNT = 0x30000000 ,
  SAM_MACHINE_ACCOUNT = 0x30000001 ,
  SAM_TRUST_ACCOUNT = 0x30000002 ,
  SAM_APP_BASIC_GROUP = 0x40000000,
  SAM_APP_QUERY_GROUP = 0x40000001 ,
  SAM_ACCOUNT_TYPE_MAX = 0x7fffffff
)

USER_ACCOUNT_CONTROL = Enum (
  ADS_UF_SCRIPT = 0x00000001,
  ADS_UF_ACCOUNTDISABLE = 0x00000002,
  ADS_UF_HOMEDIR_REQUIRED = 0x00000008,
  ADS_UF_LOCKOUT = 0x00000010,
  ADS_UF_PASSWD_NOTREQD = 0x00000020,
  ADS_UF_PASSWD_CANT_CHANGE = 0x00000040,
  ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = 0x00000080,
  ADS_UF_TEMP_DUPLICATE_ACCOUNT = 0x00000100,
  ADS_UF_NORMAL_ACCOUNT = 0x00000200,
  ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = 0x00000800,
  ADS_UF_WORKSTATION_TRUST_ACCOUNT = 0x00001000,
  ADS_UF_SERVER_TRUST_ACCOUNT = 0x00002000,
  ADS_UF_DONT_EXPIRE_PASSWD = 0x00010000,
  ADS_UF_MNS_LOGON_ACCOUNT = 0x00020000,
  ADS_UF_SMARTCARD_REQUIRED = 0x00040000,
  ADS_UF_TRUSTED_FOR_DELEGATION = 0x00080000,
  ADS_UF_NOT_DELEGATED = 0x00100000,
  ADS_UF_USE_DES_KEY_ONLY = 0x00200000,
  ADS_UF_DONT_REQUIRE_PREAUTH = 0x00400000,
  ADS_UF_PASSWORD_EXPIRED = 0x00800000,
  ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 0x01000000
)

ENUMS = {
  "GROUP_TYPES" : GROUP_TYPES,
  "AUTHENTICATION_TYPES" : AUTHENTICATION_TYPES,
  "SAM_ACCOUNT_TYPES" : SAM_ACCOUNT_TYPES,
  "USER_ACCOUNT_CONTROL" : USER_ACCOUNT_CONTROL
}

def _set (obj, attribute, value):
  """Helper function to add an attribute directly into the instance
   dictionary, bypassing possible __getattr__ calls
  """
  obj.__dict__[attribute] = value

def _and (*args):
  """Helper function to return its parameters and-ed
   together and bracketed, ready for a SQL statement.

  eg,

    _and ("x=1", "y=2") => "(x=1 AND y=2)"
  """
  return u" AND ".join (args)

def _or (*args):
  """Helper function to return its parameters or-ed
   together and bracketed, ready for a SQL statement.

  eg,

    _or ("x=1", _and ("a=2", "b=3")) => "(x=1 OR (a=2 AND b=3))"
  """
  return u" OR ".join (args)

def _add_path (root_path, relative_path):
  """Add another level to an LDAP path.
  eg,

    _add_path ('LDAP://DC=gb,DC=vo,DC=local', "cn=Users")
      => "LDAP://cn=users,DC=gb,DC=vo,DC=local"
  """
  protocol = u"LDAP://"
  if relative_path.startswith (protocol):
    return relative_path

  if root_path.startswith (protocol):
    start_path = root_path[len (protocol):]
  else:
    start_path = root_path

  return protocol + relative_path + "," + start_path

def connection ():
  connection = Dispatch ("ADODB.Connection")
  connection.Provider = "ADsDSOObject"
  connection.Open ("Active Directory Provider")
  return connection

class ADO_record (object):
  """Simple wrapper around an ADO result set"""

  def __init__ (self, record):
    self.record = record
    self.fields = {}
    for i in range (record.Fields.Count):
      field = record.Fields.Item (i)
      self.fields[field.Name] = field

  def __getattr__ (self, name):
    """Allow access to field names by name rather than by Item (...)"""
    try:
      return self.fields[name]
    except KeyError:
      raise AttributeError

  def __str__ (self):
    """Return a readable presentation of the entire record"""
    s = []
    s.append (repr (self))
    s.append (u"{")
    for name, item in self.fields.items ():
      s.append (u"  %s = %s" % (name, item))
    s.append ("}")
    return u"\n".join (s)

def query (query_string, **command_properties):
  """Auxiliary function to serve as a quick-and-dirty
   wrapper round an ADO query
  """
  command = Dispatch ("ADODB.Command")
  command.ActiveConnection = connection ()
  #
  # Add any client-specified ADO command properties.
  # NB underscores in the keyword are replaced by spaces.
  #
  # Examples:
  #   "Cache_results" = False => Don't cache large result sets
  #   "Page_size" = 500 => Return batches of this size
  #   "Time Limit" = 30 => How many seconds should the search continue
  #
  for k, v in command_properties.items ():
    command.Properties (k.replace ("_", " ")).Value = v
  command.CommandText = query_string

  results = []
  recordset, result = command.Execute ()
  while not recordset.EOF:
    yield ADO_record (recordset)
    recordset.MoveNext ()

BASE_TIME = datetime.datetime (1601, 1, 1)
def ad_time_to_datetime (ad_time):
  hi, lo = i32 (ad_time.HighPart), i32 (ad_time.LowPart)
  ns100 = (hi << 32) + lo
  delta = datetime.timedelta (microseconds=ns100 / 10)
  return BASE_TIME + delta

def convert_to_object (item):
  if item is None: return None
  return AD_object (item)

def convert_to_objects (items):
  if items is None:
    return []
  else:
    if not isinstance (items, (tuple, list)):
      items = [items]
    return [AD_object (item) for item in items]

def convert_to_datetime (item):
  if item is None: return None
  return ad_time_to_datetime (item)

def convert_to_sid (item):
  if item is None: return None
  return win32security.SID (item)

def convert_to_guid (item):
  if item is None: return None
  guid = convert_to_hex (item)
  return u"{%s-%s-%s-%s-%s}" % (guid[:8], guid[8:12], guid[12:16], guid[16:20], guid[20:])

def convert_to_hex (item):
  if item is None: return None
  return "".join ([u"%x" % ord (i) for i in item])

def convert_to_enum (name):
  def _convert_to_enum (item):
    if item is None: return None
    return ENUMS[name][item]
  return _convert_to_enum

def convert_to_flags (enum_name):
  def _convert_to_flags (item):
    if item is None: return None
    item = i32 (item)
    enum = ENUMS[enum_name]
    return set (name for (bitmask, name) in enum.item_numbers () if item & bitmask)
  return _convert_to_flags

_PROPERTY_MAP = dict (
  accountExpires = convert_to_datetime,
  badPasswordTime = convert_to_datetime,
  creationTime = convert_to_datetime,
  dSASignature = convert_to_hex,
  forceLogoff = convert_to_datetime,
  fSMORoleOwner = convert_to_object,
  groupType = convert_to_flags ("GROUP_TYPES"),
  lastLogoff = convert_to_datetime,
  lastLogon = convert_to_datetime,
  lastLogonTimestamp = convert_to_datetime,
  lockoutDuration = convert_to_datetime,
  lockoutObservationWindow = convert_to_datetime,
  lockoutTime = convert_to_datetime,
  masteredBy = convert_to_objects,
  maxPwdAge = convert_to_datetime,
  member = convert_to_objects,
  memberOf = convert_to_objects,
  minPwdAge = convert_to_datetime,
  modifiedCount = convert_to_datetime,
  modifiedCountAtLastProm = convert_to_datetime,
  msExchMailboxGuid = convert_to_guid,
  objectGUID = convert_to_guid,
  objectSid = convert_to_sid,
  Parent = convert_to_object,
  publicDelegates = convert_to_objects,
  publicDelegatesBL = convert_to_objects,
  pwdLastSet = convert_to_datetime,
  replicationSignature = convert_to_hex,
  replUpToDateVector = convert_to_hex,
  repsFrom = convert_to_hex,
  repsTo = convert_to_hex,
  sAMAccountType = convert_to_enum ("SAM_ACCOUNT_TYPES"),
  subRefs = convert_to_objects,
  userAccountControl = convert_to_flags ("USER_ACCOUNT_CONTROL"),
  uSNChanged = convert_to_datetime,
  uSNCreated = convert_to_datetime,
  wellKnownObjects = convert_to_objects
)
_PROPERTY_MAP['msDs-masteredBy'] = convert_to_objects

class _AD_root (object):
  def __init__ (self, obj):
    _set (self, "com_object", obj)
    _set (self, "properties", {})
    for i in range (obj.PropertyCount):
      property = obj.Item (i)
      proprties[property.Name] = property.Value

class _AD_object (object):
  """Wrap an active-directory object for easier access
   to its properties and children. May be instantiated
   either directly from a COM object or from an ADs Path.

   eg,

     import active_directory
     users = AD_object (path="LDAP://cn=Users,DC=gb,DC=vo,DC=local")
  """

  def __init__ (self, obj):
    #
    # Be careful here with attribute assignment;
    #  __setattr__ & __getattr__ will fall over
    #  each other if you aren't.
    #
    _set (self, "com_object", obj)
    schema = GetObject (obj.Schema)
    _set (self, "properties", schema.MandatoryProperties + schema.OptionalProperties)
    _set (self, "is_container", schema.Container)

    self._property_map = _PROPERTY_MAP
    self._delegate_map = dict ()

  def __getitem__ (self, key):
    return getattr (self, key)

  def __getattr__ (self, name):
    #
    # Special-case find_... methods to search for
    # corresponding object types.
    #
    if name.startswith ("find_"):
      names = name[len ("find_"):].lower ().split ("_")
      first, rest = names[0], names[1:]
      object_class = "".join ([first] + [n.title () for n in rest])
      return self._find (object_class)
      
    if name.startswith ("search_"):
      names = name[len ("search_"):].lower ().split ("_")
      first, rest = names[0], names[1:]
      object_class = "".join ([first] + [n.title () for n in rest])
      return self._search (object_class)
    
    #
    # Allow access to object's properties as though normal
    #  Python instance properties. Some properties are accessed
    #  directly through the object, others by calling its Get
    #  method. Not clear why.
    #
    if name not in self._delegate_map:
      try:
        attr = getattr (self.com_object, name)
      except AttributeError:
        try:
          attr = self.com_object.Get (name)
        except:
          raise AttributeError

      converter = self._property_map.get (name)
      if converter:
        self._delegate_map[name] = converter (attr)
      else:
        self._delegate_map[name] = attr

    return self._delegate_map[name]

  def __setitem__ (self, key, value):
    setattr (self, key, value)

  def __setattr__ (self, name, value):
    #
    # Allow attribute access to the underlying object's
    #  fields.
    #
    if name in self.properties:
      self.com_object.Put (name, value)
      self.com_object.SetInfo ()
    else:
      _set (self, name, value)

  def as_string (self):
    return self.path ()

  def __str__ (self):
    return self.as_string ()

  def __repr__ (self):
    return u"<%s: %s>" % (self.__class__.__name__, self.as_string ())

  def __eq__ (self, other):
    return self.com_object.Guid == other.com_object.Guid
    
  def __hash__ (self):
    return hash (self.com_object.ADsPath)

  class AD_iterator:
    """ Inner class for wrapping iterated objects
    (This class and the __iter__ method supplied by
    Stian Søiland <stian@soiland.no>)
    """
    def __init__(self, com_object):
      self._iter = iter(com_object)
    def __iter__(self):
      return self
    def next(self):
      return AD_object(self._iter.next())

  def __iter__(self):
    return self.AD_iterator(self.com_object)
    
  def walk (self):
    """Analogous to os.walk, traverse this AD subtree, 
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

  def dump (self, ofile=sys.stdout):
    ofile.write (self.as_string () + u"\n")
    ofile.write (u"{\n")
    for name in self.properties:
      try:
        value = getattr (self, name)
      except:
        value = u"Unable to get value"
      if value:
        try:
          if isinstance (name, unicode):
            name = name.encode (sys.stdout.encoding)
          if isinstance (value, unicode):
            value = value.encode (sys.stdout.encoding)
          ofile.write ("  %s => %s\n" % (name, value))
        except UnicodeEncodeError:
          ofile.write ("  %s => %s\n" % (name, repr (value)))

    ofile.write (u"}\n")

  def set (self, **kwds):
    """Set a number of values at one time. Should be
     a little more efficient than assigning properties
     one after another.

    eg,

      import active_directory
      user = active_directory.find_user ("goldent")
      user.set (displayName = "Tim Golden", description="SQL Developer")
    """
    for k, v in kwds.items ():
      self.com_object.Put (k, v)
    self.com_object.SetInfo ()

  def path (self):
    return self.com_object.ADsPath

  def parent (self):
    """Find this object's parent"""
    return AD_object (path=self.com_object.Parent)

  def child (self, relative_path):
    """Return the relative child of this object. The relative_path
     is inserted into this object's AD path to make a coherent AD
     path for a child object.

    eg,

      import active_directory
      root = active_directory.root ()
      users = root.child ("cn=Users")

    """
    return AD_object (path=_add_path (self.path (), relative_path))

  def _find (self, object_class):
    """Helper function to allow general-purpose searching for
    objects of a class by calling a .find_xxx_yyy method.
    """
    def _find (name):
      for item in self.search (objectClass=object_class, name=name):
        return item
    return _find
  
  def _search (self, object_class):
    """Helper function to allow general-purpose searching for
    objects of a class by calling a .search_xxx_yyy method.
    """
    def _search (*args, **kwargs):
      return self.search (objectClass=object_class, *args, **kwargs)
    return _search

  def find (self, name):
    for item in self.search (name=name):
      return item
  
  def find_user (self, name=None):
    """Make a special case of (the common need of) finding a user
    either by username or by display name
    """
    name = name or win32api.GetUserName ()
    for user in self.search (u"sAMAccountName='%s' OR displayName='%s' OR cn='%s'" % (name, name, name), objectCategory=u'Person', objectClass=u'User'):
      return user

  def find_ou (self, name):
    """Convenient alias for find_organizational_unit"""
    return self.find_organizational_unit (name)
      
  def search (self, *args, **kwargs):
    """The key method which puts together its arguments to construct
    a valid AD search string, using AD-SQL (or whatever it's called)
    rather than the conventional LDAP syntax.
    
    Position args are AND-ed together and passed along verbatim
    Keyword args are AND-ed together as equi-filters
    The results are always wrapped as an _AD_object or one of
    its subclasses. No matter which class is returned, well-known
    attributes are converted according to a property map to more
    Pythonic types.
    """
    sql_string = []
    sql_string.append (u"SELECT *")
    sql_string.append (u"FROM '%s'" % self.path ())
    clauses = []
    if args:
      clauses.append (_and (*args))
    if kwargs:
      clauses.append (_and (*(u"%s='%s'" % (k, v) for (k, v) in kwargs.items ())))
    where_clause = _and (*clauses)
    if where_clause:
      sql_string.append (u"WHERE %s" % where_clause)

    for result in query (u"\n".join (sql_string), Page_size=50):
      yield AD_object (result.ADsPath.Value)

class _AD_user (_AD_object):
  def __init__ (self, *args, **kwargs):
    _AD_object.__init__ (self, *args, **kwargs)

class _AD_computer (_AD_object):
  def __init__ (self, *args, **kwargs):
    _AD_object.__init__ (self, *args, **kwargs)

class _AD_group (_AD_object):
  def __init__ (self, *args, **kwargs):
    _AD_object.__init__ (self, *args, **kwargs)

  def walk (self):
    """Override the usual .walk method by returning instead:
    
    group, groups, users
    """
    members = self.member or []
    groups = [m for m in members if m.Class == u'group']
    users = [m for m in members if m.Class == u'user']
    yield (self, groups, users)
    for group in groups:
      for result in group.walk ():
        yield result

class _AD_organisational_unit (_AD_object):
  def __init__ (self, *args, **kwargs):
    _AD_object.__init__ (self, *args, **kwargs)

class _AD_domain_dns (_AD_object):
  def __init__ (self, *args, **kwargs):
    _AD_object.__init__ (self, *args, **kwargs)
    
class _AD_public_folder (_AD_object):
  pass

_CLASS_MAP = {
  u"user" : _AD_user,
  u"computer" : _AD_computer,
  u"group" : _AD_group,
  u"organizationalUnit" : _AD_organisational_unit,
  u"domainDNS" : _AD_domain_dns,
  u"publicFolder" : _AD_public_folder
}
_CACHE = {}
def cached_AD_object (path, obj):
  try:
    return _CACHE[path]
  except KeyError:
    classed_obj = _CLASS_MAP.get (obj.Class, _AD_object) (obj)
    _CACHE[path] = classed_obj
    return classed_obj
    
def clear_cache ():
  _CACHE.clear ()

def escaped_moniker (moniker):
  #
  # If the moniker *appears* to have been escaped
  # already, return it straight. This is obviously
  # fragile but seems to work for now.
  #
  if "\\/" in moniker:
    return moniker
  else:
    return moniker.replace ("/", "\\/")

def AD_object (obj_or_path=None, path=""):
  """Factory function for suitably-classed Active Directory
  objects from an incoming path or object. NB The interface
  is now  intended to be:

    AD_object (obj_or_path)

  but for historical reasons will continue to support:

    AD_object (obj=None, path="")

  @param obj_or_path Either an COM AD object or the path to one. If
  the path doesn't start with "LDAP://" this will be prepended.

  @return An _AD_object or a subclass proxying for the AD object
  """
  scheme = "LDAP://"
  if path and not obj_or_path:
    obj_or_path = path
  try:
    if isinstance (obj_or_path, basestring):
      moniker = obj_or_path.lower ()
      if obj_or_path.upper ().startswith (scheme):
        moniker = obj_or_path[len (scheme):]
      else:
        moniker = obj_or_path
      moniker = escaped_moniker (moniker)
      return cached_AD_object (obj_or_path, GetObject ("LDAP://" + moniker))
    else:
      return cached_AD_object (obj_or_path.ADsPath, obj_or_path)
  except:
    raise
    #~ raise Exception, "Problem with path or object %s" % obj_or_path

def AD (server=None):
  default_naming_context = _root (server).Get ("defaultNamingContext")
  return AD_object (GetObject ("LDAP://%s" % default_naming_context))

def _root (server=None):
  if server:
    return GetObject ("LDAP://%s/rootDSE" % server)
  else:
    return GetObject ("LDAP://rootDSE")

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
def root ():
  global _ad
  if _ad is None:
    _ad = AD ()
  return _ad

def search_ex (query_string=""):
  """Search the Active Directory by specifying a complete
   query string. NB The results will *not* be AD_objects
   but rather ADO_objects which are queried for their fields.

   eg,

     import active_directory
     for user in active_directory.search_ex (\"""
       SELECT displayName
       FROM 'LDAP://DC=gb,DC=vo,DC=local'
       WHERE objectCategory = 'Person'
     \"""):
       print user.displayName
  """
  for result in query (query_string, Page_size=50):
    yield result
