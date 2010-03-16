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

* The active directory class (_AD_object or a subclass) will determine
  its properties and allow you to access them as instance properties::

     import active_directory
     goldent = active_directory.find_user ("goldent")
     print ad.displayName

* Any object returned by the AD object's operations is themselves
  wrapped as AD objects so you get the same benefits::

    import active_directory
    users = active_directory.root ().child ("cn=users")
    for user in users.search ("displayName='Tim*'"):
      print user.displayName

* To search the AD, there are two module-level general
  search functions, and module-level convenience functions
  to find a user, computer etc. Usage is illustrated below::

   import active_directory as ad

   for user in ad.search (
     "objectClass='User'",
     "displayName='Tim Golden' OR sAMAccountName='goldent'"
   ):
     #
     # This search returns an _AD_object
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

* Typical usage will be::

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
__VERSION__ = "1.0rc1"

import os, sys
import datetime
import re

import pythoncom
import pywintypes
import win32api
from win32com.client import Dispatch, GetObject
import win32security
from win32com import adsi
from win32com.adsi import adsicon

try:
  import collections
  SetBase = collections.MutableSet
except (ImportError, AttributeError):
  SetBase = object

class ActiveDirectoryError (Exception):
  """Base class for all AD Exceptions"""
  pass

class MemberAlreadyInGroupError (ActiveDirectoryError):
  pass

class MemberNotInGroupError (ActiveDirectoryError):
  pass

ERROR_DS_NO_SUCH_OBJECT = 0x80072030
ERROR_OBJECT_ALREADY_EXISTS = 0x80071392
ERROR_MEMBER_NOT_IN_ALIAS = 0x80070561
ERROR_MEMBER_IN_ALIAS = 0x80070562

def wrapper (winerror_map, default_exception):
  u"""Used by each module to map specific windows error codes onto
  Python exceptions. Always includes a default which is raised if
  no specific exception is found.
  """
  def _wrapped (function, *args, **kwargs):
    u"""Call a Windows API with parameters, and handle any
    exception raised either by mapping it to a module-specific
    one or by passing it back up the chain.
    """
    try:
      return function (*args, **kwargs)
    except pywintypes.com_error, (hresult_code, hresult_name, additional_info, parameter_in_error):
      exception_string = [u"%08X - %s" % (signed_to_unsigned (hresult_code), hresult_name)]
      if additional_info:
        wcode, source_of_error, error_description, whlp_file, whlp_context, scode = additional_info
        exception_string.append (u"  Error in: %s" % source_of_error)
        exception_string.append (u"  %08X - %s" % (signed_to_unsigned (scode), (error_description or "").strip ()))
      exception = winerror_map.get (hresult_code, default_exception)
      raise exception (hresult_code, hresult_name, "\n".join (exception_string))
    except pywintypes.error, (errno, errctx, errmsg):
      exception = winerror_map.get (errno, default_exception)
      raise exception (errno, errctx, errmsg)
    except (WindowsError, IOError), err:
      exception = winerror_map.get (err.errno, default_exception)
      if exception:
        raise exception (err.errno, "", err.strerror)
  return _wrapped

WINERROR_MAP = {
  ERROR_MEMBER_NOT_IN_ALIAS : MemberNotInGroupError,
  ERROR_MEMBER_IN_ALIAS : MemberAlreadyInGroupError
}
wrapped = wrapper (WINERROR_MAP, ActiveDirectoryError)


DEFAULT_BIND_FLAGS = adsicon.ADS_SECURE_AUTHENTICATION | adsicon.ADS_SERVER_BIND | adsicon.ADS_FAST_BIND

def delta_as_microseconds (delta) :
  return delta.days * 24* 3600 * 10**6 + delta.seconds * 10**6 + delta.microseconds

def signed_to_unsigned (signed):
  """Convert a (possibly signed) long to unsigned hex"""
  unsigned, = struct.unpack ("L", struct.pack ("l", signed))
  return unsigned

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

  def __repr__ (self):
    return repr (self._name_map)

  def __str__ (self):
    return str (self._name_map)

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

ADS_PROPERTY = Enum (
  CLEAR = 1,
  UPDATE = 2,
  APPEND = 3,
  DELETE = 4
)

ENUMS = {
  "GROUP_TYPES" : GROUP_TYPES,
  "AUTHENTICATION_TYPES" : AUTHENTICATION_TYPES,
  "SAM_ACCOUNT_TYPES" : SAM_ACCOUNT_TYPES,
  "USER_ACCOUNT_CONTROL" : USER_ACCOUNT_CONTROL,
  "ADS_PROPERTY" : ADS_PROPERTY
}

def _set (obj, attribute, value):
  """Helper function to add an attribute directly into the instance
   dictionary, bypassing possible __getattr__ calls
  """
  obj.__dict__[attribute] = value

def and_ (*args, **kwargs):
  return "&%s" % "".join (["(%s)" % s for s in args] + ["(%s=%s)" % (k, v) for (k, v) in kwargs.items ()])

def or_ (*args, **kwargs):
  return "|%s" % "".join (["(%s)" % s for s in args] + ["(%s=%s)" % (k, v) for (k, v) in kwargs.items ()])

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

class _Proxy (object):

  ESCAPED_CHARACTERS = dict ((special, r"\%02x" % ord (special)) for special in "*()\x00/")

  @classmethod
  def escaped_filter (cls, s):
    for original, escape in cls.ESCAPED_CHARACTERS.items ():
      s = s.replace (original, escape)
    return s

  def _munge (cls, other):
    if isinstance (other, Base):
      return other.dn

    if isinstance (other, datetime.datetime):
      return datetime_to_ad_time (other)

    other = unicode (other)
    if other.endswith (u"*"):
      other, suffix = other[:-1], other[-1]
    else:
      suffix = u""
    other = cls.escaped_filter (other)
    return other + suffix

  def __init__ (self, name):
    self._name = name

  def __unicode__ (self):
    return self._name

  def __eq__ (self, other):
    return u"%s=%s" % (self._name, self._munge (other))

  def __ne__ (self, other):
    return u"!%s=%s" % (self._name, self._munge (other))

  def __ge__ (self, other):
    return u"%s>=%s" % (self._name, self._munge (other))

  def __le__ (self, other):
    return u"%s<=%s" % (self._name, self._munge (other))

  def __and__ (self, other):
    return u"%s:1.2.840.113556.1.4.803:=%s" % (self._name, self._munge (other))

  def __or__ (self, other):
    return u"%s:1.2.840.113556.1.4.804:=%s" % (self._name, self._munge (other))

  def is_within (self, dn):
    return u"%s:1.2.840.113556.1.4.1941:=%s" % (self._name, self._munge (dn))

  def is_not_within (self, dn):
    return u"!%s:1.2.840.113556.1.4.1941:=%s" % (self._name, self._munge (dn))

class _Attributes (object):

  def __init__ (self):
    self._proxies = {}

  def __getattr__ (self, attr):
    return self._proxies.setdefault (attr, _Proxy (attr))

attr = _Attributes ()

def connect (username=None, password=None):
  """Return an ADODB connection, optionally authenticated by
  username & password.
  """
  connection = Dispatch ("ADODB.Connection")
  connection.Provider = "ADsDSOObject"
  if username:
    connection.Open ("Active Directory Provider", username, password)
  else:
    connection.Open ("Active Directory Provider")
  return connection

_command_properties = {
  "Page Size" : 500,
  "Asynchronous" : True
}
def query (query_string, connection=None, **command_properties):
  """Basic AD query, passing a raw query string straight through to an
  Active Directory, optionally using a (possibly pre-authenticated) connection
  or creating one on demand. command_properties may be specified which will be
  passed through to the ADO command with underscores replaced by spaces. Useful
  values include:

  =============== ==========================================================
  page_size       How many records to return in one go
  size_limit      Stop after returning this many records
  cache_results   Boolean: cache results; turn off if a large result
  time_limit      Stop returning records after this many seconds
  timeout         Stop waiting for the records to start after this many seconds
  asynchronous    Boolean: Start returning records immediately
  sort_on         field name to sort on
  =============== ==========================================================

  :param query_string: An AD query string in any acceptable format. See :func:`query_string`
                       for an easy way of producing this
  :param connection: (optional) An ADODB.Connection, as provided by :func:`connect`. If
                     this is supplied it will be used and not closed. If it is not supplied, a default connection
                     will be created, used and then closed.
  :param command_properties: A collection of keywords which will be passed through to the
                             ADO query as Properties.
  """
  command = Dispatch ("ADODB.Command")
  _connection = connection or connect ()
  command.ActiveConnection = _connection

  for k, v in _command_properties.items ():
    command.Properties (k.replace ("_", " ")).Value = v
  for k, v in command_properties.items ():
    command.Properties (k.replace ("_", " ")).Value = v
  command.CommandText = query_string

  results = []
  recordset, result = command.Execute ()
  while not recordset.EOF:
    yield dict ((field.Name, field.Value) for field in recordset.Fields)
    recordset.MoveNext ()

  if connection is None:
    _connection.Close ()

def query_string (base=None, filter="", attributes="*", scope="Subtree", range=None):
  """Easy way to produce a valid AD query string, with meaninful defaults. This
  is the first parameter to the :func:`query` function so the following will
  yield the display name of every user in the domain::

    import active_directory as ad

    for u in ad.query (
      ad.query_string (filter="(objectClass=User)", attributes="displayName")
    ):
      print u['displayName']

  :param base: An LDAP:// moniker representing the starting point of the search [domain root]
  :param filter: An AD filter string to limit the search [no filter]
  :param attributes: A comma-separated attributes string [* - ADsPath]
  :param scope: One of - Subtree, Base, OneLevel [Subtree]
  :param range: Limit the number of returns of multivalued attributes [no range]
  """
  if base is None:
    base = u"LDAP://" + GetObject (u"LDAP://rootDSE").Get (u"defaultNamingContext")
  if not filter.startswith ("("):
    filter = u"(%s)" % filter
  segments = [u"<%s>" % base, filter, attributes]
  if range:
    segments += [u"Range=%s-%s" % range]
  segments += [scope]
  return u";".join (segments)

def search_ex (query_string="", username=None, password=None):
  """FIXME: Historical version of :func:`query`"""
  return query (query_string, connection=connect (username, password))

BASE_TIME = datetime.datetime (1601, 1, 1)
def ad_time_to_datetime (ad_time):
  hi, lo = i32 (ad_time.HighPart), i32 (ad_time.LowPart)
  ns100 = (hi << 32) + lo
  delta = datetime.timedelta (microseconds=ns100 / 10)
  return BASE_TIME + delta

def datetime_to_ad_time (datetime):
  delta = datetime - BASE_TIME
  n_microseconds = delta.microseconds + (1000000 * delta.seconds) + (1000000 * 60 * 60 * 24 * delta.days)
  return 1 * n_microseconds

def pytime_to_datetime (pytime):
  return datetime.datetime.fromtimestamp (int (pytime))

def pytime_from_datetime (datetime):
  pass

def convert_to_object (item):
  if item is None: return None
  return ad (item)

def convert_to_objects (items):
  if items is None:
    return []
  else:
    if not isinstance (items, (tuple, list)):
      items = [items]
    return [ad (item) for item in items]

def convert_to_datetime (item):
  if item is None: return None
  return ad_time_to_datetime (item)

def convert_pytime_to_datetime (item):
  if item is None: return None
  return pytime_to_datetime (item)

def convert_to_sid (item):
  if item is None: return None
  return win32security.SID (item)

def convert_to_guid (item):
  if item is None: return None
  guid = convert_to_hex (item)
  return u"{%s-%s-%s-%s-%s}" % (guid[:8], guid[8:12], guid[12:16], guid[16:20], guid[20:])

def convert_to_hex (item):
  if item is None: return None
  return "".join ([u"%02x" % ord (i) for i in item])

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
    return set ([name for (bitmask, name) in enum.item_numbers () if item & bitmask])
  return _convert_to_flags

def ddict (**kwargs):
  return kwargs

_PROPERTY_MAP = ddict (
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
  manager = convert_to_object,
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
  wellKnownObjects = convert_to_objects,
  whenCreated = convert_pytime_to_datetime,
  whenChanged = convert_pytime_to_datetime,
)
_PROPERTY_MAP['msDs-masteredBy'] = convert_to_objects
_PROPERTY_MAP_OUT = _PROPERTY_MAP

def convert_from_object (item):
  if item is None: return None
  return item.com_object

def convert_from_objects (items):
  if items == []:
    return None
  else:
    return [obj.com_object for obj in items]

def convert_from_datetime (item):
  if item is None: return None
  try:
    return pytime_to_datetime (item)
  except:
    return ad_time_to_datetime (item)

def convert_from_sid (item):
  if item is None: return None
  return win32security.SID (item)

def convert_from_guid (item):
  if item is None: return None
  guid = convert_from_hex (item)
  return u"{%s-%s-%s-%s-%s}" % (guid[:8], guid[8:12], guid[12:16], guid[16:20], guid[20:])

def convert_from_hex (item):
  if item is None: return None
  return "".join ([u"%x" % ord (i) for i in item])

def convert_from_enum (name):
  def _convert_from_enum (item):
    if item is None: return None
    return ENUMS[name][item]
  return _convert_from_enum

def convert_from_flags (enum_name):
  def _convert_from_flags (item):
    if item is None: return None
    item = i32 (item)
    enum = ENUMS[enum_name]
    return set ([name for (bitmask, name) in enum.item_numbers () if item & bitmask])
  return _convert_from_flags

_PROPERTY_MAP_IN = ddict (
  accountExpires = convert_from_datetime,
  badPasswordTime = convert_from_datetime,
  creationTime = convert_from_datetime,
  dSASignature = convert_from_hex,
  forceLogoff = convert_from_datetime,
  fSMORoleOwner = convert_from_object,
  groupType = convert_from_flags ("GROUP_TYPES"),
  lastLogoff = convert_from_datetime,
  lastLogon = convert_from_datetime,
  lastLogonTimestamp = convert_from_datetime,
  lockoutDuration = convert_from_datetime,
  lockoutObservationWindow = convert_from_datetime,
  lockoutTime = convert_from_datetime,
  masteredBy = convert_from_objects,
  maxPwdAge = convert_from_datetime,
  member = convert_from_objects,
  memberOf = convert_from_objects,
  minPwdAge = convert_from_datetime,
  modifiedCount = convert_from_datetime,
  modifiedCountAtLastProm = convert_from_datetime,
  msExchMailboxGuid = convert_from_guid,
  objectGUID = convert_from_guid,
  objectSid = convert_from_sid,
  Parent = convert_from_object,
  publicDelegates = convert_from_objects,
  publicDelegatesBL = convert_from_objects,
  pwdLastSet = convert_from_datetime,
  replicationSignature = convert_from_hex,
  replUpToDateVector = convert_from_hex,
  repsFrom = convert_from_hex,
  repsTo = convert_from_hex,
  sAMAccountType = convert_from_enum ("SAM_ACCOUNT_TYPES"),
  subRefs = convert_from_objects,
  userAccountControl = convert_from_flags ("USER_ACCOUNT_CONTROL"),
  uSNChanged = convert_from_datetime,
  uSNCreated = convert_from_datetime,
  wellKnownObjects = convert_from_objects
)
_PROPERTY_MAP_IN['msDs-masteredBy'] = convert_from_objects

class _Members (set):

  def __init__ (self, group):
    super (_Members, self).__init__ (ad (i) for i in iter (group.com_object.members ()))
    self._group = group

  def _effect (self, original):
    group = self._group.com_object
    for member in (self - original):
      print "Adding", member
      #~ group.Add (member.AdsPath)
      wrapped (group.Add, member.AdsPath)
    for member in (original - self):
      print "Removing", member
      #~ group.Remove (member.AdsPath)
      wrapped (group.Remove, member.AdsPath)

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

class Base (object):
  """Wrap an active-directory object for easier access
   to its properties and children. May be instantiated
   either directly from a COM object or from an ADs Path.

   eg,

     import active_directory
     users = active_directory.ad ("LDAP://cn=Users,DC=gb,DC=vo,DC=local")
  """

  def __init__ (self, obj, username=None, password=None):
    #
    # Be careful here with attribute assignment;
    # __setattr__ & __getattr__ will fall over
    # each other if you aren't.
    #
    _set (self, "com_object", obj)
    schema = GetObject (obj.Schema)
    _set (self, "properties", getattr (schema, "MandatoryProperties", []) + getattr (schema, "OptionalProperties", []))
    self.is_container = getattr (schema, "Container", False)
    self.username = username
    self.password = password
    self.connection = connect (username=username, password=password)
    self.dn = getattr (self.com_object, "distinguishedName", self.com_object.name)
    self._property_map = _PROPERTY_MAP
    self._delegate_map = dict ()
    self._path = obj.AdsPath

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

    if name.startswith ("get_"):
      names = name[len ("get_"):].lower ().split ("_")
      first, rest = names[0], names[1:]
      object_class = "".join ([first] + [n.title () for n in rest])
      return self._get (object_class)

    #
    # Allow access to object's properties as though normal
    # Python instance properties. Some properties are accessed
    # directly through the object, others by calling its Get
    # method. Not clear why.
    #
    if name not in self._delegate_map:
      try:
        attr = getattr (self.com_object, name)
      except AttributeError:
        try:
          attr = self.com_object.Get (name)
        except:
          return super (Base, self).__getattr__ (name)

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
      super (Base, self).__setattr__ (name, value)

  def as_string (self):
    return self.path

  def __str__ (self):
    return self.as_string ()

  def __repr__ (self):
    return u"<%s: %s>" % (self.com_object.Class, self.dn)

  def __eq__ (self, other):
    return self.com_object.Guid == other.com_object.Guid

  def __hash__ (self):
    return hash (self.com_object.Guid)

  class AD_iterator:
    """ Inner class for wrapping iterated objects
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

  def _get_path (self):
    return self._path
  path = property (_get_path)

  def refresh (self):
    self.com_object.GetInfo ()

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

  def flat (self):
    for container, containers, items in self.walk ():
      for item in items:
        yield item

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

  def parent (self):
    """Find this object's parent"""
    return ad (self.com_object.Parent)

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

  def _get (self, object_class):
    """Helper function to allow general-purpose retrieval of a
    child object by class.
    """
    def _get (rdn):
      return self.get (object_class, rdn)
    return _get

  def find (self, name):
    for item in self.search (name=name):
      return item

  def find_user (self, name=None):
    """Make a special case of (the common need of) finding a user
    either by username or by display name
    """
    name = name or win32api.GetUserName ()
    filter = and_ (
      or_ (sAMAccountName=name, displayName=name, cn=name),
      sAMAccountType=SAM_ACCOUNT_TYPES.SAM_USER_OBJECT
    )
    for user in self.search (filter):
      return user

  def find_ou (self, name):
    """Convenient alias for find_organizational_unit"""
    return self.find_organizational_unit (name)

  def search (self, *args, **kwargs):
    filter = and_ (*args, **kwargs)
    query_string = "<%s>;(%s);distinguishedName;Subtree" % (self.ADsPath, filter)
    for result in query (query_string, connection=self.connection):
      yield ad (unicode (result['distinguishedName']), username=self.username, password=self.password)

  def get (self, object_class, relative_path):
    return ad (self.com_object.GetObject (object_class, relative_path))

  def new (self, object_class, sam_account_name, **kwargs):
    obj = self.com_object.Create (object_class, u"cn=%s" % sam_account_name)
    obj.Put ("sAMAccountName", sam_account_name)
    obj.SetInfo ()
    for name, value in kwargs.items ():
      obj.Put (name, value)
    obj.SetInfo ()
    return ad (obj)

class WinNT (Base):

  def __eq__ (self, other):
    return self.com_object.ADsPath.lower () == other.com_object.ADsPath.lower ()

  def __hash__ (self):
    return hash (self.com_object.ADsPath.lower ())

class Group (Base):

  def _get_x (self):
    return getattr (self, "_x", "Unknown")
  def _set_x (self, value):
    print "_set_x", value
    self._x = value
  x = property (_get_x, _set_x)

  def _get_members (self):
    return _Members (self)
  def _set_members (self, members):
    original = self.members
    new_members = set (ad (m) for m in members)
    print "original", original
    print "new members", new_members
    print "new_members - original", new_members - original
    for member in (new_members - original):
      print "Adding", member
      self.com_object.Add (member.AdsPath)
    print "original - new_members", original - new_members
    for member in (original - new_members):
      print "Removing", member
      self.com_object.Remove (member.AdsPath)
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
  if moniker.find ("\\/") > -1:
    return moniker
  else:
    return moniker.replace ("/", "\\/")

def ad (obj_or_path, username=None, password=None):
  """Factory function for suitably-classed Active Directory
  objects from an incoming path or object. NB The interface
  is now  intended to be:

    ad (obj_or_path)

  @param obj_or_path Either an COM AD object or the path to one. If
  the path doesn't start with "LDAP://" this will be prepended.

  @return An _AD_object or a subclass proxying for the AD object
  """
  matcher = re.compile ("(LDAP://|GC://|WinNT://)?(.*)")
  if isinstance (obj_or_path, Base):
    return obj_or_path
  elif isinstance (obj_or_path, basestring):
    scheme, dn = matcher.match (obj_or_path).groups ()
    if scheme is None: scheme = "LDAP://"
    if scheme == "WinNT://":
      moniker = dn
    else:
      moniker = escaped_moniker (dn)
    obj = adsi.ADsOpenObject (scheme + moniker, username, password)
  else:
    obj = obj_or_path
    scheme, dn = matcher.match (obj_or_path.AdsPath).groups ()

  if scheme == "WinNT://":
    class_map = _WINNT_CLASS_MAP.get (obj.Class.lower (), WinNT)
  else:
    class_map = _CLASS_MAP.get (obj.Class.lower (), Base)
  return class_map (obj)
AD_object = ad

def AD (server=None, username=None, password=None, use_gc=False):
  if use_gc:
    scheme = "GC://"
  else:
    scheme = "LDAP://"
  if server:
    root_moniker = scheme + server + "/rootDSE"
  else:
    root_moniker = scheme + "rootDSE"
  root_obj = adsi.ADsOpenObject (root_moniker, username, password, DEFAULT_BIND_FLAGS)
  default_naming_context = root_obj.Get ("defaultNamingContext")
  moniker = scheme + default_naming_context
  obj = adsi.ADsOpenObject (moniker, username, password, DEFAULT_BIND_FLAGS)
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
