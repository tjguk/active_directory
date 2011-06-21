import datetime
import threading

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

def converters (type, name):
  return _get_converters (type).get (name, (None, None, None))

def converter (name, offset):
  name_converter = converters ("name", name)[offset]
  if name_converter:
    return name_converter
  syntax = core.attribute (name).attributeSyntax
  syntax_converter = converters ("syntax", syntax)[offset]
  if syntax_converter:
    return syntax_converter
  return None

def getter (name):
  return converter (name, 0)

def setter (name):
  return converter (name, 1)

def searcher (name):
  return converter (name, 2)

#
# Generic converters
# These are general-purpose converters which can
# be used directly as name/syntax converters or
# indirectly via more complex converters below
#
BASE_TIME = datetime.datetime (1601, 1, 1)
DELTA0 = datetime.timedelta (0)
def interval_to_datetime (interval):
  delta = interval_to_timedelta (interval)
  try:
    return BASE_TIME - delta
  except OverflowError:
    return datetime.datetime.max if delta > DELTA0 else datetime.datetime.min

def interval_to_timedelta (interval):
  hi = utils.signed_to_unsigned (ularge.HighPart)
  lo = utils.signed_to_unsigned (ularge.LowPart)
  ns100 = (hi << 32) + lo
  return datetime.timedelta (microseconds=-ns100 / 10)

def largeint_to_long (value):
  return (value.HighPart << 32) + value.LowPart


#
# Syntax Converters
#
register_converters ("syntax", "2.5.5.16", getter=largeint_to_long)
