# -*- coding: iso-8859-1 -*-
import os, sys
import re

import win32api
import win32com.client
from win32com.adsi import adsi, adsicon

from . import core
from . import constants
from . import credentials
from . import exc
from . import types
from . import utils

class NotAContainerError (exc.ActiveDirectoryError):
  pass

class ADContainer (object):

  def __init__ (self, ad_com_object):
    try:
      self.container = exc.wrapped (ad_com_object.QueryInterface, adsi.IID_IADsContainer)
    except exc.ActiveDirectoryError, (error_code, _, _):
      if error_code == exc.E_NOINTERFACE:
        raise NotAContainerError

  def __iter__ (self):
    enumerator = exc.wrapped (adsi.ADsBuildEnumerator, self.container)
    while True:
      items = exc.wrapped (adsi.ADsEnumerateNext, enumerator, 10)
      if items:
        for item in items:
          yield exc.wrapped (item.QueryInterface, adsi.IID_IADs)
      else:
        break

class ADBase (object):
  """A slender wrapper around an AD COM object which facilitates getting,
  setting and clearing an object's attributes plus pretty-printing to stdout.
  It does no validation of the names passed and an no conversions of the
  values. It can be used alone (most easily via the :func:`ADBase` function
  which takes an AD path and returns an ADBase object). It also provides the
  basis for the other AD classes below.
  """

   #
   # For speed, hardcode the known properties of the IADs class
   #
  _properties = ["ADsPath", "Class", "GUID", "Name", "Parent", "Schema"]
  _class_properties = {}
  _class_containers = {}
  _property_schemas = {}

  def __init__ (self, obj, cred=None):
    com_object = obj.QueryInterface (adsi.IID_IADs)
    utils._set (self, "com_object", com_object)
    utils._set (self, "cred", cred)
    utils._set (self, "path", self.com_object.ADsPath)
    cls = exc.wrapped (getattr, com_object, "Class")
    utils._set (self, "cls", cls)
    if cls not in self.__class__._class_properties:
      schema_path = exc.wrapped (getattr, com_object, "Schema")
      schema_obj = core.open_object (schema_path, cred=cred)
      self.__class__._class_properties[cls] = \
        exc.wrapped (getattr, schema_obj, "mandatoryProperties") + \
        exc.wrapped (getattr, schema_obj, "optionalProperties")
      self.__class__._class_containers[cls] = exc.wrapped (getattr, schema_obj, "container")

  def _put (self, name, value):
    operation = constants.ADS_PROPERTY.CLEAR if value is None else constants.ADS_PROPERTY.UPDATE
    exc.wrapped (self.com_object.PutEx, operation, name, value)

  def __getattr__ (self, name):
    try:
      return exc.wrapped (getattr, self.com_object, name)
    except (AttributeError, NotImplementedError):
      try:
        return exc.wrapped (self.com_object.Get, name)
      except NotImplementedError:
        raise AttributeError

  def __setattr__ (self, name, value):
    self._put (name, value)
    exc.wrapped (self.com_object.SetInfo)

  def __delattr__ (self, name):
    self._put (name, None)
    exc.wrapped (self.com_object.SetInfo)

  def __repr__ (self):
    return "<%s: %s>" % (self.__class__.__name__, self.as_string ())

  def __str__ (self):
    return self.as_string ()

  def __iter__(self):
    try:
      for item in ADContainer (self.com_object):
        yield self.__class__ (item, self.cred)
    except NotAContainerError:
      raise TypeError ("%r is not iterable" % self)

  @classmethod
  def from_path (cls, path, cred=None):
    return cls (core.open_object (path, cred))

  query = core.dquery

  def search (self, filter):
    for result in self.query (filter, ['ADsPath']):
      yield self.__class__ (core.open_object (result['ADsPath'][0], cred=self.cred))

  def walk (self, level=0):
    subordinates = list (self)
    yield (
      level,
      self,
      (s for s in subordinates if self.__class__._class_containers[s.cls]),
      (s for s in subordinates if not self.__class__._class_containers[s.cls])
    )
    for subordinate in (s for s in subordinates if self.__class__._class_containers[s.cls]):
      for walked in subordinate.walk (level+1):
        yield walked

  def as_string (self):
    return self.path

  def dump (self, ofile=sys.stdout):
    ofile.write (self.as_string () + u"\n")
    ofile.write ("{\n")
    for property in self.__class__._class_properties[self.cls]:
      value = exc.wrapped (getattr, self, property, None)
      if value:
        ofile.write ("  %s => %r\n" % (property, value))
    ofile.write ("}\n")

def adbase (obj_or_path=None, cred=None):
  if obj_or_path is None:
    return ADBase (core.root_obj (), cred=cred)
  elif isinstance (obj_or_path, ADBase):
    return obj_or_path
  else:
    return ADBase.from_path (obj_or_path, cred)
