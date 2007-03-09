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

+ The active directory object (AD_object) will determine its
   properties and allow you to access them as instance properties.

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
   search functions, two module-level functions to
   find a user and computer specifically and the search
   method on each AD_object. Usage is illustrated below:

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

   users = ad.root ().child ("cn=users")
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
 the pywin32 extensions.

9th Nov 2005 0.5   Added code from Stian Søiland to handle negative longs
                    coming from return codes. Also another patch from the
                    same source to correct iterator handling.
                   Changed license to MIT license (simpler and still 
                    GPL-compatible) 
12th May 2005 0.4  Added ADS_GROUP constants to support cookbook examples.
                   Added .dump method to AD_object to allow easy viewing
                    of all fields.
                   Allowed find_user / find_computer to have default values,
                    meaning the logged-on user and current machine.
                   Added license: PSF
20th Oct 2004 0.3  Added "Page Size" param to query to allow result
                    sets of > 1000.
                   Refactored search mechanisms to module-level and
                    switched to SQL queries.
19th Oct 2004 0.2  Added support for attribute assignment
                     (see AD_object.__setattr__)
                   Added module-level functions:
                     root - returns a default AD instance
                     search - calls root's search
                     find_user - returns first match for a user/fullname
                     find_computer - returns first match for a computer
                   Now runs under 2.2 (removed reference to basestring)
15th Oct 2004 0.1  Initial release by Tim Golden
"""
from __future__ import generators

__VERSION__ = "0.6"

import os, sys
import datetime
import win32api
import socket

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
  return " AND ".join (args)

def _or (*args):
  """Helper function to return its parameters or-ed
   together and bracketed, ready for a SQL statement.

  eg,

    _or ("x=1", _and ("a=2", "b=3")) => "(x=1 OR (a=2 AND b=3))"
  """
  return " OR ".join (args)

def _add_path (root_path, relative_path):
  """Add another level to an LDAP path.
  eg,

    _add_path ('LDAP://DC=gb,DC=vo,DC=local', "cn=Users")
      => "LDAP://cn=users,DC=gb,DC=vo,DC=local"
  """
  protocol = "LDAP://"
  if relative_path.startswith (protocol):
    return relative_path

  if root_path.startswith (protocol):
    start_path = root_path[len (protocol):]
  else:
    start_path = root_path

  return protocol + relative_path + "," + start_path

#
# Global cached ADO Connection object
#
_connection = None
def connection ():
  global _connection
  if _connection is None:
    _connection = Dispatch ("ADODB.Connection")
    _connection.Provider = "ADsDSOObject"
    _connection.Open ("Active Directory Provider")
  return _connection

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
    s.append ("{")
    for name, item in self.fields.items ():
      s.append ("  %s = %s" % (name, item))
    s.append ("}")
    return "\n".join (s)

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
  return AD_object (item)
  
def convert_to_objects (items):
  if not isinstance (items, (tuple, list)):
    items = [items]
  return [AD_object (item) for item in items]
  
def convert_to_datetime (item):
  return ad_time_to_datetime (item)
  
def convert_to_sid (item):
  return win32security.SID (item)
  
def convert_to_guid (item):
  guid = convert_to_hex (item)
  return "{%s-%s-%s-%s-%s}" % (guid[:8], guid[8:12], guid[12:16], guid[16:20], guid[20:])
    
def convert_to_hex (item):
  return "".join (["%x" % ord (i) for i in item])
  
def convert_to_enum (name):
  def _convert_to_enum (item):
    return ENUMS[name][item]
  return _convert_to_enum
  
def convert_to_flags (enum_name):
  def _convert_to_flags (item):
    item = i32 (item)
    enum = ENUMS[enum_name]
    return set (name for (bitmask, name) in enum.item_numbers () if item & bitmask)
  return _convert_to_flags

class _AD_object (object):
  """Wrap an active-directory object for easier access
   to its properties and children. May be instantiated
   either directly from a COM object or from an ADs Path.

   eg,

     import active_directory
     users = AD_object (path="LDAP://cn=Users,DC=gb,DC=vo,DC=local")     
  """

  def __init__ (self, obj=None, path=""):
    #
    # Be careful here with attribute assignment;
    #  __setattr__ & __getattr__ will fall over
    #  each other if you aren't.
    #
    if path:
      _set (self, "com_object", GetObject (path))
    else:
      _set (self, "com_object", obj)
    schema = GetObject (self.com_object.Schema)
    _set (self, "properties", schema.MandatoryProperties + schema.OptionalProperties)
    _set (self, "is_container", schema.Container)
    
    self._property_map = {}

  def __getattr__ (self, name):
    #
    # Allow access to object's properties as though normal
    #  Python instance properties. Some properties are accessed
    #  directly through the object, others by calling its Get
    #  method. Not clear why.
    #
    try:
      attr = getattr (self.com_object, name)
    except AttributeError:
      try:
        attr = self.com_object.Get (name)
      except:
        raise AttributeError
    
    converter = self._property_map.get (name)
    if converter:
      return converter (attr)
    else:
      return attr

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
    return "<%s: %s>" % (self.__class__.__name__, self.as_string ())
    
  def __eq__ (self, other):
    return self.com_object.Guid == other.com_object.Guid

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

  def dump (self, ofile=sys.stdout):
    ofile.write (self.as_string () + "\n")
    ofile.write ("{\n")
    for name in self.properties:
      value = getattr (self, name)
      if value:
        try:
          if isinstance (name, unicode):
            name = name.encode (sys.stdout.encoding)
          if isinstance (value, unicode):
            value = value.encode (sys.stdout.encoding)
          ofile.write ("  %s => %s\n" % (name, value))
        except UnicodeEncodeError:
          ofile.write ("  %s => %s\n" % (name, repr (value)))

    ofile.write ("}\n")

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

  def find_user (self, name=None):
    name = name or win32api.GetUserName ()
    for user in self.search ("objectCategory='Person'", "objectClass='User'", "sAMAccountName='%s' OR displayName='%s' OR cn='%s'" % (name, name, name)):
      return user
    
  def find_computer (self, name=None):
    name = name or socket.gethostname ()
    for computer in self.search ("objectCategory='Computer'", "cn='%s'" % name):
      return computer

  def find_group (self, name):
    for group in self.search ("objectCategory='group'", "cn='%s'" % name):
      return group

  def search (self, *args):
    sql_string = []
    sql_string.append ("SELECT *")
    sql_string.append ("FROM '%s'" % self.path ())
    where_clause = _and (*args)
    if where_clause:
      sql_string.append ("WHERE %s" % where_clause)

    for result in query ("\n".join (sql_string), Page_size=50):
      yield AD_object (path=result.ADsPath.Value)

class _AD_user (_AD_object):
  
  def __init__ (self, *args, **kwargs):
    _AD_object.__init__ (self, *args, **kwargs)
    self._property_map.update (dict (
      pwdLastSet = convert_to_datetime,
      memberOf = convert_to_objects,
      objectSid = convert_to_sid,
      accountExpires = convert_to_datetime,
      badPasswordTime = convert_to_datetime,
      lastLogoff = convert_to_datetime,
      lastLogon = convert_to_datetime,
      lastLogonTimestamp = convert_to_datetime,
      lockoutTime = convert_to_datetime,
      msExchMailboxGuid = convert_to_guid,
      objectGUID = convert_to_guid,
      publicDelegates = convert_to_objects,
      publicDelegatesBL = convert_to_objects,
      sAMAccountType = convert_to_enum ("SAM_ACCOUNT_TYPES"),
      userAccountControl = convert_to_flags ("USER_ACCOUNT_CONTROL"),
      uSNChanged = convert_to_datetime,
      uSNCreated = convert_to_datetime,
      replicationSignature = convert_to_hex
    ))

class _AD_computer (_AD_object):
  
  def __init__ (self, *args, **kwargs):
    _AD_object.__init__ (self, *args, **kwargs)
    self._property_map.update (dict (
      objectSid = convert_to_sid,
      accountExpires = convert_to_datetime,
      badPasswordTime = convert_to_datetime,
      lastLogoff = convert_to_datetime,
      lastLogon = convert_to_datetime,
      lastLogonTimestamp = convert_to_datetime,
      objectGUID = convert_to_guid,
      publicDelegates = convert_to_objects,
      publicDelegatesBL = convert_to_objects,
      pwdLastSet = convert_to_datetime,
      sAMAccountType = convert_to_enum ("SAM_ACCOUNT_TYPES"),
      userAccountControl = convert_to_flags ("USER_ACCOUNT_CONTROL"),
      uSNChanged = convert_to_datetime,
      uSNCreated = convert_to_datetime
    ))

class _AD_group (_AD_object):
  def __init__ (self, *args, **kwargs):
    _AD_object.__init__ (self, *args, **kwargs)
    self._property_map.update (dict (
      groupType = convert_to_flags ("GROUP_TYPES"),
      objectSid = convert_to_sid,
      member = convert_to_objects,
      memberOf = convert_to_objects,
      objectGUID = convert_to_guid,
      sAMAccountType = convert_to_enum ("SAM_ACCOUNT_TYPES"),
      uSNChanged = convert_to_datetime,
      uSNCreated = convert_to_datetime
    ))

_CLASS_MAP = {
  "user" : _AD_user,
  "computer" : _AD_computer,
  "group" : _AD_group
}
def AD_object (obj=None, path=""):
  if path and not obj:
    obj = path
  if isinstance (obj, basestring):
    if not obj.startswith ("LDAP://"):
      obj = "LDAP://" + obj
    obj = GetObject (obj)
  return _CLASS_MAP.get (obj.Class, _AD_object) (obj, path)

def AD (server=None):
  default_naming_context = _root (server).Get ("defaultNamingContext")
  return AD_object (GetObject ("LDAP://%s" % default_naming_context))

def _root (server=None):
  if server:
    return GetObject ("LDAP://%s/rootDSE" % server)
  else:
    return GetObject ("LDAP://rootDSE")

def find_user (name=None):
  return root ().find_user (name)
  
def find_computer (name=None):
  return root ().find_computer (name)
  
def find_group (name):
  return root ().find_group (name)

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

def search (*args, **kwargs):
  return root ().search (*args, **kwargs)

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

