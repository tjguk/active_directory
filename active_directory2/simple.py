# -*- coding: iso-8859-1 -*-
import os, sys
import re

import win32api
from win32com import adsi

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

class ADSimple (object):
  """A slender wrapper around an AD COM object which facilitates getting,
  setting and clearing an object's attributes plus pretty-printing to stdout.
  It does no validation of the names passed and an no conversions of the
  values. It can be used alone (most easily via the :func:`adsimple` function
  which takes an AD path and returns an ADSimple object). It also provides the
  basis for the other AD classes below.
  """

   #
   # For speed, hardcode the known properties of the IADs class
   #
  _properties = ["ADsPath", "Class", "GUID", "Name", "Parent", "Schema"]

  def __init__ (self, obj):
    utils._set (self, "com_object", obj.QueryInterface (adsi.IID_IADs))
    utils._set (self, "properties", self._properties)
    utils._set (self, "path", obj.ADsPath)

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
  def from_path (cls, path, cred=credentials.Passthrough):
    cred = credentials.credentials (cred)
    return cls (adsi.ADsOpenObject (path, cred.username, cred.password, cred.authentication_type))

  def as_string (self):
    return self.path

  def munge_attribute_for_dump (self, name, value):
    if value is None:
      return ""
    if isinstance (value, unicode):
      value = value.encode ("ascii", "backslashreplace")
    return str (value)

  def dump (self, ofile=sys.stdout):
    ofile.write (self.as_string () + u"\n")
    ofile.write ("{\n")
    for name in self.properties:
      ofile.write ("  %s => " % name)
      try:
        value = getattr (self, name)
      except AttributeError:
        value = "Unable to get value"
      ofile.write ("%s\n" % self.munge_attribute_for_dump (name, value))
    ofile.write ("}\n")

def adsimple (obj_or_path, cred=credentials.Passthrough):
  cred = credentials.credentials (cred)
  if isinstance (obj_or_path, ADSimple):
    return obj_or_path
  else:
    return ADSimple.from_path (obj_or_path, cred)

