# -*- coding: iso-8859-1 -*-
import datetime

import win32security

from . import constants
from . import core
from . import utils

#
# Converters
#
BASE_TIME = datetime.datetime (1601, 1, 1)
def ularge_to_datetime (ularge):
  hi, lo = utils.i32 (ularge.HighPart), utils.i32 (ularge.LowPart)
  ns100 = (hi << 32) + lo
  delta = datetime.timedelta (microseconds=ns100 / 10)
  return BASE_TIME + delta

def datetime_to_ularge (datetime):
  raise NotImplementedError

def pytime_to_datetime (pytime):
  return datetime.datetime.fromtimestamp (int (pytime))

def pytime_from_datetime (datetime):
  raise NotImplementedError

def dn_to_object (factory):
  def _convert_to_object (dn):
    if dn is None: return None
    return factory (dn)
  return _convert_to_object

def object_to_dn (obj):
  return obj.distinguishedName

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
    item = utils.i32 (item)
    return set ([name for (bitmask, name) in enum.item_numbers () if item & bitmask])
  return _convert_to_flags

def convert_to_breadcrumbs (item):
  return u" > ".join (item)

def ularge_to_long (ularge):
  return (item.HighPart << 32) + item.LowPart

def long_to_ularge (long):
  raise NotImplementedError

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

def octet_to_sid (octet):
  if octet is None: return None
  return win32security.SID (octet)

def sid_to_octet (sid):
  raise NotImplementedError

def convert_from_guid (item):
  if item is None: return None
  guid = convert_from_hex (item)
  return u"{%s-%s-%s-%s-%s}" % (guid[:8], guid[8:12], guid[12:16], guid[16:20], guid[20:])

def octet_to_hex (octet):
  if item is None: return None
  return u"".join ([u"%x" % ord (i) for i in item])

def hex_to_octet (hex):
  raise NotImplementedError

def convert_from_enum (enum):
  def _convert_from_enum (item):
    if item is None: return None
    return enum[item]
  return _convert_from_enum

def convert_from_flags (enum):
  def _convert_from_flags (item):
    if item is None: return None
    item = utils.i32 (item)
    return set ([name for (bitmask, name) in enum.item_numbers () if item & bitmask])
  return _convert_from_flags

class _Proxy (object):

  ESCAPED_CHARACTERS = dict ((special, ur"\%02x" % ord (special)) for special in u"*()\x00/")

  @classmethod
  def escaped_filter (cls, s):
    for original, escape in cls.ESCAPED_CHARACTERS.items ():
      s = s.replace (original, escape)
    return s

  @staticmethod
  def _munge (other):
    if hasattr (other, "dn"):
      return other.dn

    if isinstance (other, datetime.date):
      other = datetime.datetime (*other.timetuple ()[:7])
      # now drop through to datetime converter below

    if isinstance (other, datetime.datetime):
      return datetime_to_ad_time (other)

    other = unicode (other)
    if other.endswith (u"*"):
      other, suffix = other[:-1], other[-1]
    else:
      suffix = u""
    #~ other = cls.escaped_filter (other)
    return other + suffix

  def __init__ (self, name):
    self._name = name
    self._attribute = None

  def __unicode__ (self):
    return self._name

  def __repr__ (self):
    return u"<_Proxy for %s>" % self._name

  def __hash__ (self):
    return hash (self._name)

  def __getattr__ (self, attr):
    return getattr (self._attribute, attr)

  def __eq__ (self, other):
    return u"%s=%s" % (self._name, self._munge (other))

  def __ne__ (self, other):
    return u"!%s=%s" % (self._name, self._munge (other))

  def __gt__ (self, other):
    return u"!%s<=%s" % (self._name, self._munge (other))
    raise NotImplementedError (u"> Not implemented")

  def __ge__ (self, other):
    return u"%s>=%s" % (self._name, self._munge (other))

  def __lt__ (self, other):
    return u"!%s>=%s" % (self._name, self._munge (other))
    raise NotImplementedError (u"< Not implemented")

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

  def dump (self, *args, **kwargs):
    if self._attribute is None:
      self._attribute = attribute (self._name)
    self._attribute.dump (*args, **kwargs)

class Attributes (object):

  def __init__ (self):
    self._proxies = {}

  def __getattr__ (self, attr):
    return self[attr]

  def __getitem__ (self, item):
    if item not in self._proxies:
      self._proxies[item] = _Proxy (item)
    return self._proxies[item]

schema = Attributes ()

def no_conversion (value):
  return value

def get_converter (name):
  """
  * Look for a specific converter
  * Drop back to a type-inferred converter
  * Drop back to raw value
  """
  if name not in _converters:
    obj = None ## attribute (name)
    if obj and obj.attributeSyntax in TYPE_CONVERTERS:
      register_converter (name, from_ad=TYPE_CONVERTERS[obj.attributeSyntax])
    elif name.endswith ("GUID"):
      register_converter (name, from_ad=convert_to_guid)
  from_ad, to_ad = _converters.get (name, (None, None))
  return from_ad or no_conversion, to_ad or no_conversion

_converters = {}
def register_converters (attribute_name, from_ad=None, to_ad=None):
  from_to = _converters.get (attribute_name, [None, None])
  if from_ad:
    from_to[0] = from_ad
  if to_ad:
    from_to[1] = to_ad
  _converters[attribute_name] = from_to

_type_converters = {}
def register_type_converters (attribute_type, from_ad=None, to_ad=None):
  from_to = _type_converters.get (attribute_type, [None, None])
  if from_ad:
    from_to[0] = from_ad
  if to_ad:
    from_to[1] = to_ad
  _type_converters[attribute_type] = from_to

register_type_converters ("2.5.5.11", ularge_to_datetime, datetime_to_ularge)
register_type_converters ("2.5.5.16", ularge_to_long, long_to_ularge)
register_type_converters ("2.5.5.17", octet_to_sid, sid_to_octet)
register_type_converters ("2.5.5.10", octet_to_hex, hex_to_octet)
