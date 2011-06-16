import re

import win32com.client
from win32com import adsi
from win32com.adsi import adsicon

from . import adbase
from . import credentials
from . import constants
from . import core
from . import exc
from . import types
from . import utils

class Descriptor (object):

  _descriptors = {}

  def __init__ (self, name):
     self.name = name
     self.getter = types.CONVERTERS.get (self.name, lambda x : x)

  def __get__ (self, obj, objtype=None):
    if obj is None:
      return self
    else:
      value = getattr (obj.com_object, self.name, None)
      if value is not None:
        return self.getter (value)

  def __set__ (self, obj, value):
    setattr (obj.com_object, self.name, value)

  def __delete__ (self, obj):
    raise AttributeError ("Can't delete %s" % self.name)

def descriptor (name):
  if name not in Descriptor._descriptors:
    Descriptor._descriptors[name] = Descriptor (name)
  return Descriptor._descriptors[name]

def _munged (name):
  return "_".join (name.split ("-"))

class ADMetaClass (type):

  def __new__ (meta, name, bases, dict):
    obj = dict.pop ("obj")
    if obj:
      schema = core.open_object (obj.Schema)
      dict['properties'] = schema.MandatoryProperties + schema.OptionalProperties
      for p in dict['properties']:
        dict[_munged (p)] = descriptor (_munged (p))
    return type.__new__ (meta, name, bases, dict)

class ADObject (adbase.ADBase):

  __metaclass__ = ADMetaClass
  klasses = {}
  obj = None

  @classmethod
  def from_obj (cls, obj):
    obj = adsi._get_good_ret (obj)
    klass = obj.Class.encode ("ascii")
    class_name = "%s" % klass[0].upper () + klass[1:]
    if class_name not in cls.klasses:
      cls.klasses[class_name] = type (class_name, (cls,), dict (obj=obj))
    return cls.klasses[class_name] (obj)

  def __getattr__ (self, attr):
    if hasattr (self, "com_object"):
      return getattr (self.com_object, attr)
    else:
      raise AttributeError (attr)

adobject = ADObject.factory
