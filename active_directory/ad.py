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
import logging

from win32com import adsi

from . import base
from . import constants
from . import core
from . import exc
from . import types
from . import utils

logger = logging.getLogger ("active_directory")
def enable_debugging ():
  logger.addHandler (logging.StreamHandler (sys.stdout))
  logger.setLevel (logging.DEBUG)

schema = types.Attributes ()

def search_ex (query_string=u"", username=None, password=None):
  u"""FIXME: Historical version of :func:`query`"""
  return core.query (query_string, connection=connect (username, password))

class RootDSE (base.ADSimple):

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

def AD (server=None, username=None, password=None, use_gc=False):
  if use_gc:
    scheme = u"GC://"
  else:
    scheme = u"LDAP://"
  if server:
    root_moniker = scheme + server + u"/rootDSE"
  else:
    root_moniker = scheme + u"rootDSE"
  root_obj = exc.wrapped (adsi.ADsOpenObject, root_moniker, username, password, constants.DEFAULT_BIND_FLAGS)
  default_naming_context = root_obj.Get (u"defaultNamingContext")
  moniker = scheme + default_naming_context
  obj = exc.wrapped (adsi.ADsOpenObject, moniker, username, password, constants.DEFAULT_BIND_FLAGS)
  return base.ad (obj, username, password)

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

def root_dse (username=None, password=None):
  return RootDSE (adsi.ADsOpenObject (u"LDAP://rootDSE", username, password, constants.DEFAULT_BIND_FLAGS))

#
# Register known attributes
#
_PROPERTY_MAP = dict (
  accountExpires = types.convert_to_datetime,
  auditingPolicy = types.convert_to_hex,
  badPasswordTime = types.convert_to_datetime,
  creationTime = types.convert_to_datetime,
  dSASignature = types.convert_to_hex,
  forceLogoff = types.convert_to_datetime,
  fSMORoleOwner = types.convert_to_object (base.ad),
  groupType = types.convert_to_flags (constants.GROUP_TYPES),
  isGlobalCatalogReady = types.convert_to_boolean,
  isSynchronized = types.convert_to_boolean,
  lastLogoff = types.convert_to_datetime,
  lastLogon = types.convert_to_datetime,
  lastLogonTimestamp = types.convert_to_datetime,
  lockoutDuration = types.convert_to_datetime,
  lockoutObservationWindow = types.convert_to_datetime,
  lockoutTime = types.convert_to_datetime,
  manager = types.convert_to_object (base.ad),
  masteredBy = types.convert_to_objects (base.ad),
  maxPwdAge = types.convert_to_datetime,
  member = types.convert_to_objects (base.ad),
  memberOf = types.convert_to_objects (base.ad),
  minPwdAge = types.convert_to_datetime,
  modifiedCount = types.convert_to_datetime,
  modifiedCountAtLastProm = types.convert_to_datetime,
  msExchMailboxGuid = types.convert_to_guid,
  schemaIDGUID = types.convert_to_guid,
  mSMQDigests = types.convert_to_hex,
  mSMQSignCertificates = types.convert_to_hex,
  objectClass = types.convert_to_breadcrumbs,
  objectGUID = types.convert_to_guid,
  objectSid = types.convert_to_sid,
  publicDelegates = types.convert_to_objects (base.ad),
  publicDelegatesBL = types.convert_to_objects (base.ad),
  pwdLastSet = types.convert_to_datetime,
  replicationSignature = types.convert_to_hex,
  replUpToDateVector = types.convert_to_hex,
  repsFrom = types.convert_to_hexes,
  repsTo = types.convert_to_hex,
  sAMAccountType = types.convert_to_enum (constants.SAM_ACCOUNT_TYPES),
  subRefs = types.convert_to_objects (base.ad),
  systemFlags = types.convert_to_flags (constants.ADS_SYSTEMFLAG),
  userAccountControl = types.convert_to_flags (constants.USER_ACCOUNT_CONTROL),
  wellKnownObjects = types.convert_to_objects (base.ad),
  whenCreated = types.convert_pytime_to_datetime,
  whenChanged = types.convert_pytime_to_datetime,
  showInAddressbook = types.convert_to_objects (base.ad),
)
_PROPERTY_MAP[u'msDs-masteredBy'] = types.convert_to_objects (base.ad)

for k, v in _PROPERTY_MAP.items ():
  types.register_converter (k, from_ad=v)

_PROPERTY_MAP_IN = dict (
  accountExpires = types.convert_from_datetime,
  badPasswordTime = types.convert_from_datetime,
  creationTime = types.convert_from_datetime,
  dSASignature = types.convert_from_hex,
  forceLogoff = types.convert_from_datetime,
  fSMORoleOwner = types.convert_from_object,
  groupType = types.convert_from_flags (constants.GROUP_TYPES),
  lastLogoff = types.convert_from_datetime,
  lastLogon = types.convert_from_datetime,
  lastLogonTimestamp = types.convert_from_datetime,
  lockoutDuration = types.convert_from_datetime,
  lockoutObservationWindow = types.convert_from_datetime,
  lockoutTime = types.convert_from_datetime,
  masteredBy = types.convert_from_objects,
  maxPwdAge = types.convert_from_datetime,
  member = types.convert_from_objects,
  memberOf = types.convert_from_objects,
  minPwdAge = types.convert_from_datetime,
  modifiedCount = types.convert_from_datetime,
  modifiedCountAtLastProm = types.convert_from_datetime,
  msExchMailboxGuid = types.convert_from_guid,
  objectGUID = types.convert_from_guid,
  objectSid = types.convert_from_sid,
  publicDelegates = types.convert_from_objects,
  publicDelegatesBL = types.convert_from_objects,
  pwdLastSet = types.convert_from_datetime,
  replicationSignature = types.convert_from_hex,
  replUpToDateVector = types.convert_from_hex,
  repsFrom = types.convert_from_hex,
  repsTo = types.convert_from_hex,
  sAMAccountType = types.convert_from_enum (constants.SAM_ACCOUNT_TYPES),
  subRefs = types.convert_from_objects,
  userAccountControl = types.convert_from_flags (constants.USER_ACCOUNT_CONTROL),
  wellKnownObjects = types.convert_from_objects
)
_PROPERTY_MAP_IN['msDs-masteredBy'] = types.convert_from_objects

for k, v in _PROPERTY_MAP_IN.items ():
  types.register_converter (k, to_ad=v)

