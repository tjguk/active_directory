import re

from win32com.adsi import adsi, adsicon

from . import adbase
from . import credentials
from . import constants
from . import core
from . import exc
from . import types
from . import utils

class ADReadOnly (exc.ActiveDirectoryError):
  pass

class ADObject (adbase.ADBase):
  u"""Wrap an active-directory object for easier access
   to its properties and children. May be instantiated
   either directly from a COM object or from an ADs Path.

   Every IADs-derived object has at least the following attributes:

   Name, Class, GUID, ADsPath, Parent, Schema

   eg,

     import active_directory as ad
     users = ad.ad ("LDAP://cn=Users,DC=gb,DC=vo,DC=local")
  """

  _converters = None

  def __init__ (self, obj, cred=None):
    adbase.ADBase.__init__ (self, obj, cred)
    core.attributes (self.properties, self.server, self.cred)

  def __getattr__ (self, name):
    #
    # Allow access to object's properties as though normal
    # Python instance properties. Some properties are accessed
    # directly through the object, others by calling its Get
    # method. Not clear why.
    #
    value = adbase.ADBase.__getattr__ (self, name)
    attr = core.attribute (name, self.server, self.cred)
    if attr:
      convert_from, _ = types.get_converters (name)
      if not convert_from:
        convert_from, _ = types.get_type_converters (attr.attributeSyntax)
      if convert_from:
        return convert_from (value)
      else:
        return value
    else:
      return value

  def __setattr__ (self, name, value):
    #
    # Allow attribute access to the underlying object's
    #  fields.
    #
    if name in self.properties:
      info = core.attributes[name]
      if info.systemOnly:
        raise ADReadOnlyError ("%s is read-only" % name)

      _, convert_to = types.get_converter (name)
      super (ADObject, self).__setattr__ (name, convert_to (value))
    else:
      super (ADObject, self).__setattr__ (name, value)

  def converters (cls):
    if cls._converters is None:
      cls._converters = types.Converters (cls)
    return cls._converters

class WinNT (ADObject):

  def __eq__ (self, other):
    return self.com_object.ADsPath.lower () == other.com_object.ADsPath.lower ()

  def __hash__ (self):
    return hash (self.com_object.ADsPath.lower ())

class _Members (set):

  def __init__ (self, group, name="members"):
    super (_Members, self).__init__ (ad (i) for i in iter (exc.wrapped (group.com_object.members)))
    self._group = group
    self._name = name

  def _effect (self, original):
    group = self._group.com_object
    for member in (self - original):
      exc.wrapped (group.Add, member.AdsPath)
    for member in (original - self):
      exc.wrapped (group.Remove, member.AdsPath)
    self._group._invalidate (self._name)

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

class Group (ADObject):

  def _get_members (self):
    return _Members (self)
  def _set_members (self, members):
    original = self.members
    new_members = set (ad (m) for m in members)
    for member in (new_members - original):
      exc.wrapped (self.com_object.Add, member.AdsPath)
    for member in (original - new_members):
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

def ad (obj_or_path, cred=None):
  u"""Factory function for suitably-classed Active Directory
  objects from an incoming path or object. NB The interface
  is now  intended to be:

    ad (obj_or_path)

  @param obj_or_path Either an COM AD object or the path to one. If
  the path doesn't start with "LDAP://" this will be prepended.

  @return An _AD_object or a subclass proxying for the AD object
  """
  if isinstance (obj_or_path, ADObject):
    return obj_or_path

  if isinstance (obj_or_path, basestring):
    scheme, server, dn = utils.parse_moniker (obj_or_path)
    obj_path = scheme + u"//" + (server or u"") + (dn or u"")
    obj = core.open_object (obj_path, cred)
  else:
    obj = obj_or_path
    scheme, server, dn = utils.parse_moniker (obj_or_path.AdsPath)

  if scheme == u"WinNT:":
    class_map = _WINNT_CLASS_MAP.get (obj.Class.lower (), WinNT)
  else:
    class_map = _CLASS_MAP.get (obj.Class.lower (), ADObject)
  return class_map (obj, cred)


#
# Register known attributes
#
#~ _PROPERTY_MAP = dict (
  #~ accountExpires = types.convert_to_datetime,
  #~ auditingPolicy = types.convert_to_hex,
  #~ badPasswordTime = types.convert_to_datetime,
  #~ creationTime = types.convert_to_datetime,
  #~ dSASignature = types.convert_to_hex,
  #~ forceLogoff = types.convert_to_datetime,
  #~ fSMORoleOwner = types.convert_to_object (adobject.ad),
  #~ groupType = types.convert_to_flags (constants.GROUP_TYPES),
  #~ isGlobalCatalogReady = types.convert_to_boolean,
  #~ isSynchronized = types.convert_to_boolean,
  #~ lastLogoff = types.convert_to_datetime,
  #~ lastLogon = types.convert_to_datetime,
  #~ lastLogonTimestamp = types.convert_to_datetime,
  #~ lockoutDuration = types.convert_to_datetime,
  #~ lockoutObservationWindow = types.convert_to_datetime,
  #~ lockoutTime = types.convert_to_datetime,
  #~ manager = types.convert_to_object (adobject.ad),
  #~ masteredBy = types.convert_to_objects (adobject.ad),
  #~ maxPwdAge = types.convert_to_datetime,
  #~ member = types.convert_to_objects (adobject.ad),
  #~ memberOf = types.convert_to_objects (adobject.ad),
  #~ minPwdAge = types.convert_to_datetime,
  #~ modifiedCount = types.convert_to_datetime,
  #~ modifiedCountAtLastProm = types.convert_to_datetime,
  #~ msExchMailboxGuid = types.convert_to_guid,
  #~ schemaIDGUID = types.convert_to_guid,
  #~ mSMQDigests = types.convert_to_hex,
  #~ mSMQSignCertificates = types.convert_to_hex,
  #~ objectClass = types.convert_to_breadcrumbs,
  #~ objectGUID = types.convert_to_guid,
  #~ objectSid = types.convert_to_sid,
  #~ publicDelegates = types.convert_to_objects (adobject.ad),
  #~ publicDelegatesBL = types.convert_to_objects (adobject.ad),
  #~ pwdLastSet = types.convert_to_datetime,
  #~ replicationSignature = types.convert_to_hex,
  #~ replUpToDateVector = types.convert_to_hex,
  #~ repsFrom = types.convert_to_hexes,
  #~ repsTo = types.convert_to_hex,
  #~ sAMAccountType = types.convert_to_enum (constants.SAM_ACCOUNT_TYPES),
  #~ subRefs = types.convert_to_objects (adobject.ad),
  #~ systemFlags = types.convert_to_flags (constants.ADS_SYSTEMFLAG),
  #~ userAccountControl = types.convert_to_flags (constants.USER_ACCOUNT_CONTROL),
  #~ wellKnownObjects = types.convert_to_objects (adobject.ad),
  #~ whenCreated = types.convert_pytime_to_datetime,
  #~ whenChanged = types.convert_pytime_to_datetime,
  #~ showInAddressbook = types.convert_to_objects (adobject.ad),
#~ )
#~ _PROPERTY_MAP[u'msDs-masteredBy'] = types.convert_to_objects (adobject.ad)

#~ for k, v in _PROPERTY_MAP.items ():
  #~ types.register_converter (k, from_ad=v)

#~ _PROPERTY_MAP_IN = dict (
  #~ accountExpires = types.convert_from_datetime,
  #~ badPasswordTime = types.convert_from_datetime,
  #~ creationTime = types.convert_from_datetime,
  #~ dSASignature = types.convert_from_hex,
  #~ forceLogoff = types.convert_from_datetime,
  #~ fSMORoleOwner = types.convert_from_object,
  #~ groupType = types.convert_from_flags (constants.GROUP_TYPES),
  #~ lastLogoff = types.convert_from_datetime,
  #~ lastLogon = types.convert_from_datetime,
  #~ lastLogonTimestamp = types.convert_from_datetime,
  #~ lockoutDuration = types.convert_from_datetime,
  #~ lockoutObservationWindow = types.convert_from_datetime,
  #~ lockoutTime = types.convert_from_datetime,
  #~ masteredBy = types.convert_from_objects,
  #~ maxPwdAge = types.convert_from_datetime,
  #~ member = types.convert_from_objects,
  #~ memberOf = types.convert_from_objects,
  #~ minPwdAge = types.convert_from_datetime,
  #~ modifiedCount = types.convert_from_datetime,
  #~ modifiedCountAtLastProm = types.convert_from_datetime,
  #~ msExchMailboxGuid = types.convert_from_guid,
  #~ objectGUID = types.convert_from_guid,
  #~ objectSid = types.convert_from_sid,
  #~ publicDelegates = types.convert_from_objects,
  #~ publicDelegatesBL = types.convert_from_objects,
  #~ pwdLastSet = types.convert_from_datetime,
  #~ replicationSignature = types.convert_from_hex,
  #~ replUpToDateVector = types.convert_from_hex,
  #~ repsFrom = types.convert_from_hex,
  #~ repsTo = types.convert_from_hex,
  #~ sAMAccountType = types.convert_from_enum (constants.SAM_ACCOUNT_TYPES),
  #~ subRefs = types.convert_from_objects,
  #~ userAccountControl = types.convert_from_flags (constants.USER_ACCOUNT_CONTROL),
  #~ wellKnownObjects = types.convert_from_objects
#~ )
#~ _PROPERTY_MAP_IN['msDs-masteredBy'] = types.convert_from_objects

#~ for k, v in _PROPERTY_MAP_IN.items ():
  #~ types.register_converter (k, to_ad=v)

"""
Attribute syntax ID	Active Directory syntax type	Equivalent ADSI syntax type
2.5.5.1	DN String	DN String
2.5.5.2	Object ID	CaseIgnore String
2.5.5.3	Case Sensitive String	CaseExact String
2.5.5.4	Case Ignored String	CaseIgnore String
2.5.5.5	Print Case String	Printable String
2.5.5.6	Numeric String	Numeric String
2.5.5.7	OR Name DNWithOctetString	Not Supported
2.5.5.8	Boolean	Boolean
2.5.5.9	Integer	Integer
2.5.5.10	Octet String	Octet String
2.5.5.11	Time	UTC Time
2.5.5.12	Unicode	Case Ignore String
2.5.5.13	Address	Not Supported
2.5.5.14	Distname-Address
2.5.5.15	NT Security Descriptor	IADsSecurityDescriptor
2.5.5.16	Large Integer	IADsLargeInteger
2.5.5.17	SID	Octet String
"""

types.register_type_converters ("2.5.5.1", types.dn_to_object (ADObject), types.object_to_dn)
