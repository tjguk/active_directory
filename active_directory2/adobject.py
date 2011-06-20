import os, sys
import re

import win32com.client
from win32com import adsi
from win32com.adsi import adsicon

from . import adbase
from . import core
from . import exc
from .log import logger
from . import types
from . import utils

class Descriptor (object):

  _descriptors = {}

  def __init__ (self, name, attribute):
     self.name = name
     self.attribute = adbase.adbase (attribute)
     self.getter, self.setter = types.get_converters (self.name)

  def __get__ (self, obj, objtype=None):
    logger.debug ("%s.__get__: obj=%s, objtype=%s", self.name, obj, objtype)
    if obj is None:
      return self
    else:
      value = getattr (obj.com_object, self.name, None)
      if value is None or self.getter is None:
        return value
      return self.getter (value)

  def __set__ (self, obj, value):
    logger.debug ("%s.__set__: obj=%s, value=%s", self.name, obj, value)
    if self.attribute.systemOnly:
      raise AttributeError ("Attribute %s is read-only" % self.name)
    if self.setter is None:
      setattr (obj.com_object, self.name, value)
    else:
      setattr (obj.com_object, self.name, self.setter (value))
    obj.com_object.SetInfo ()

  def __delete__ (self, obj):
    raise NotImplementedError

  def __repr__ (self):
    return "<%s: %s>" % (self.__class__.__name__, self.name)

  def dump (self, ofile=sys.stdout):
    self.attribute.dump (ofile=ofile)

def descriptor (name, attribute):
  if name not in Descriptor._descriptors:
    Descriptor._descriptors[name] = Descriptor (name, attribute)
  return Descriptor._descriptors[name]

def _munged (name):
  return "_".join (name.split ("-"))

class ADMetaClass (type):

  def __new__ (meta, name, bases, dict):
    obj = dict.pop ("obj")
    cred = dict.pop ("cred")
    if obj:
      scheme, server, dn = utils.parse_moniker (obj.ADsPath)
      server = server.rstrip ("/")
      schema = core.open_object (obj.Schema)
      dict['properties'] = schema.MandatoryProperties + schema.OptionalProperties
      core.attributes (dict['properties'], server)
      for p in dict['properties']:
        dict[_munged (p)] = descriptor (p, core.attribute (p, server))
    return type.__new__ (meta, name, bases, dict)

class ADObject (adbase.ADBase):

  __metaclass__ = ADMetaClass
  klasses = {}
  obj = None
  cred = None
  schema_obj = {}

  @classmethod
  def from_obj (cls, obj, cred=None):
    obj = adsi._get_good_ret (obj)
    klass = obj.Class.encode ("ascii")
    class_name = "%s" % klass[0].upper () + klass[1:]
    if class_name not in cls.klasses:
      cls.klasses[class_name] = type (class_name, (cls,), dict (obj=obj, cred=cred))
    return cls.klasses[class_name] (obj)

  def __getattr__ (self, attr):
    if hasattr (self, "com_object"):
      return getattr (self.com_object, attr)
    else:
      raise AttributeError (attr)

  def __setattr__ (self, attr, value):
    logger.debug ("ADObject.__setattr__: attr=%s, value=%s", attr, value)
    super (adbase.ADBase, self).__setattr__ (attr, value)

adobject = ADObject.factory
