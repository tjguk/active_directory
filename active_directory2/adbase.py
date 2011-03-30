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
  _property_schemas = {}

  def __init__ (self, obj, cred=None):
    com_object = obj.QueryInterface (adsi.IID_IADs)
    utils._set (self, "com_object", com_object)
    utils._set (self, "path", self.com_object.ADsPath)
    cls = exc.wrapped (getattr, com_object, "Class")
    #~ cls = exc.wrapped (com_object.Get, "Class")
    utils._set (self, "cls", cls)
    if cls not in self.__class__._class_properties:
      schema_path = exc.wrapped (getattr, com_object, "Schema")
      schema_obj = core.open_object (schema_path, cred=cred)
      properties = exc.wrapped (getattr, schema_obj, "mandatoryProperties") + exc.wrapped (getattr, schema_obj, "optionalProperties")
      self.__class__._class_properties[cls] = properties

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
        yield self.__class__ (item)
    except NotAContainerError:
      raise TypeError ("%r is not iterable" % self)

  @classmethod
  def from_path (cls, path, cred=None):
    return cls (core.open_object (path, cred))

  def query (self, filter, attributes=None, flags=0):
    SEARCH_PREFERENCES = {
      adsicon.ADS_SEARCHPREF_PAGESIZE : 1000,
      adsicon.ADS_SEARCHPREF_SEARCH_SCOPE : adsicon.ADS_SCOPE_SUBTREE,
    }
    directory_search = exc.wrapped (self.com_object.QueryInterface, adsi.IID_IDirectorySearch)
    directory_search.SetSearchPreference ([(k, (v,)) for k, v in SEARCH_PREFERENCES.items ()])
    hSearch = directory_search.ExecuteSearch (filter, attributes)
    try:
      hResult = directory_search.GetFirstRow (hSearch)
      while hResult == 0:
        results = dict ()
        while True:
          attr = exc.wrapped (directory_search.GetNextColumnName, hSearch)
          if attr is None:
            break
          _, _, attr_values = exc.wrapped (directory_search.GetColumn, hSearch, attr)
          results[attr] = [value for (value, _) in attr_values]
        yield results
        hResult = directory_search.GetNextRow (hSearch)
    finally:
      directory_search.AbandonSearch (hSearch)
      directory_search.CloseSearchHandle (hSearch)

  def search (self, filter, cred=None):
    for result in self.query (filter, ['ADsPath']):
      yield self.__class__ (core.open_object (result['ADsPath'][0], cred=cred))

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

def adbase (obj_or_path, cred=None):
  if isinstance (obj_or_path, ADBase):
    return obj_or_path
  else:
    return ADBase.from_path (obj_or_path, cred)
