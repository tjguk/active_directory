import datetime
import threading

import win32security

from . import constants
from . import core
from .log import logger
from . import utils
from . import support

_local = threading.local ()

#
# Converters come in two types: name converters and syntax converters
# Within each of those types there are getters, setters & searchers
# If a name converter exists this is used, otherwise a syntax converter
#
def _get_converters (type):
  return _local.__dict__.setdefault (type, {})

def register_converters (type, name, getter=None, setter=None, searcher=None):
  _converters = _get_converters (type)
  _getter, _setter, _searcher = _converters.get (name, (None, None, None))
  _converters[name] = (getter or _getter, setter or _setter, searcher or _searcher)

def _converters (type, name):
  return _get_converters (type).get (name, (None, None, None))

def _converter (name, offset):
  name_converter = converters ("name", name)[offset]
  if name_converter:
    return name_converter
  syntax = core.attribute (name).attributeSyntax
  syntax_converter = converters ("syntax", syntax)[offset]
  if syntax_converter:
    return syntax_converter
  return None

def getter (name):
  return _converter (name, 0)

def setter (name):
  return _converter (name, 1)

def searcher (name):
  return _converter (name, 2)

def converters (name):
  name_converters = _converters ("name", name)
  attribute = core.attribute (name)
  syntax = attribute.attributeSyntax
  syntax_converters = _converters ("syntax", syntax)
  return [(n or s) for (n, s) in zip (name_converters, syntax_converters)]

#
# Helper Functions
#
def list_getter (getter):
  def _list_getter (value):
    return [getter (v) for v in value]
  return _list_getter

#
# Generic converters
# These are general-purpose converters which can
# be used directly as name/syntax converters or
# indirectly via more complex converters below
#
BASE_TIME = datetime.datetime (1601, 1, 1)
DELTA0 = datetime.timedelta (0)
def interval_to_datetime (interval):
  hi = utils.signed_to_unsigned (interval.HighPart)
  lo = utils.signed_to_unsigned (interval.LowPart)
  ns100 = (hi << 32) + lo
  if ns100 in (0, 0x7FFFFFFFFFFFFFFF):
    return datetime.datetime.max
  delta = datetime.timedelta (microseconds=ns100 / 10)
  try:
    return BASE_TIME + delta
  except OverflowError:
    return datetime.datetime.max if delta > DELTA0 else datetime.datetime.min

def interval_to_timedelta (interval):
  #~ hi = utils.signed_to_unsigned (interval.HighPart)
  #~ lo = utils.signed_to_unsigned (interval.LowPart)
  ns100 = (interval.HighPart << 32) + interval.LowPart
  return datetime.timedelta (microseconds=-ns100 / 10)

def largeint_to_long (value):
  return (value.HighPart << 32) + value.LowPart

def pytime_to_datetime (pytime):
  return datetime.datetime.fromtimestamp (int (pytime))

def binary_string_to_tuple (value):
  return binary_to_guid (value.BinaryValue), value.DNString

def binary_to_guid (item):
  if item is None: return None
  guid = binary_to_hex (item)
  return u"{%s-%s-%s-%s-%s}" % (guid[:8], guid[8:12], guid[12:16], guid[16:20], guid[20:])

def binary_to_sid (item):
  if item is None: return None
  return win32security.SID (item)

def binary_to_hex (item):
  if item is None: return None
  return u"%s" % u"".join ([u"%02x" % ord (i) for i in item])


#
# Name Converters
#

#
# Register Syntax Converters
#
register_converters ("syntax", "2.5.5.7", getter=binary_string_to_tuple)
register_converters ("syntax", "2.5.5.10", getter=binary_to_guid)
register_converters ("syntax", "2.5.5.11", getter=pytime_to_datetime)
register_converters ("syntax", "2.5.5.16", getter=largeint_to_long)
register_converters ("syntax", "2.5.5.17", getter=binary_to_sid)

#
# Register Name Converters
#
register_converters ("name", "accountExpires", getter=interval_to_datetime)
register_converters ("name", "auditingPolicy", getter=binary_to_hex)
register_converters ("name", "badPasswordTime", getter=interval_to_datetime)
register_converters ("name", "creationTime", getter=interval_to_datetime)
register_converters ("name", "lastLogon", getter=interval_to_datetime)
register_converters ("name", "lockoutDuration", getter=interval_to_timedelta)
register_converters ("name", "lockoutObservationWindow", getter=interval_to_timedelta)
register_converters ("name", "maxPwdAge", getter=interval_to_datetime)
register_converters ("name", "pwdLastSet", getter=interval_to_datetime)
register_converters ("name", "wellKnownObjects", getter=list_getter (binary_string_to_tuple))
