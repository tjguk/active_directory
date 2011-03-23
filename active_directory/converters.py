import datetime

import win32security

from . import core
from . import utils

#
# Converters
#
BASE_TIME = datetime.datetime (1601, 1, 1)
def ad_time_to_datetime (ad_time):
  hi, lo = utils.i32 (ad_time.HighPart), utils.i32 (ad_time.LowPart)
  ns100 = (hi << 32) + lo
  delta = datetime.timedelta (microseconds=ns100 / 10)
  return BASE_TIME + delta

def datetime_to_ad_time (datetime):
  return datetime.strftime ("%y%m%d%H%M%SZ")

def pytime_to_datetime (pytime):
  return datetime.datetime.fromtimestamp (int (pytime))

def pytime_from_datetime (datetime):
  raise NotImplementedError

def convert_to_object (factory):
  def _convert_to_object (item):
    if item is None: return None
    return factory (item)
  return _convert_to_object

def convert_to_objects (factory):
  def _convert_to_objects (items):
    if items is None:
      return []
    else:
      if not isinstance (items, (tuple, list)):
        items = [items]
      return [factory (item) for item in items]
  return _convert_to_objects

def convert_to_boolean (item):
  if item is None: return None
  return item == u"TRUE"

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
  return u"<%s>" % u"".join ([u"%02x" % ord (i) for i in item])

def convert_to_hexes (item):
  if item is None: return None
  return [convert_to_hex (i) for i in item]

def convert_to_enum (enum):
  def _convert_to_enum (item):
    if item is None: return None
    return enum[item]
  return _convert_to_enum

def convert_to_flags (enum):
  def _convert_to_flags (item):
    if item is None: return None
    item = i32 (item)
    return set ([name for (bitmask, name) in enum.item_numbers () if item & bitmask])
  return _convert_to_flags

def convert_to_breadcrumbs (item):
  return u" > ".join (item)

def convert_to_long (item):
  return (item.HighPart << 32) + item.LowPart

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
  return u"".join ([u"%x" % ord (i) for i in item])

def convert_from_enum (enum):
  def _convert_from_enum (item):
    if item is None: return None
    return enum[item]
  return _convert_from_enum

def convert_from_flags (enum):
  def _convert_from_flags (item):
    if item is None: return None
    item = i32 (item)
    return set ([name for (bitmask, name) in enum.item_numbers () if item & bitmask])
  return _convert_from_flags

converters = {}
def register_converter (attribute_name, from_ad=None, to_ad=None):
  from_to = converters.get (attribute_name, [None, None])
  if from_ad:
    from_to[0] = from_ad
  if to_ad:
    from_to[1] = to_ad
  converters[attribute_name] = from_to

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
TYPE_CONVERTERS = {
  "2.5.5.11" : ad_time_to_datetime,
  "2.5.5.16" : convert_to_long,
  "2.5.5.17" : convert_to_sid,
  "2.5.5.10" : convert_to_hex
}

types = {
  "guid" : (convert_to_guid, convert_from_guid),
  "time_as_long" : (convert_to_datetime, convert_from_datetime),
  "boolean" : (convert_to_boolean, None),
}

