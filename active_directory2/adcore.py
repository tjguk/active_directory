# -*- coding: iso-8859-1 -*-
import os, sys

import win32com.client
from win32com import adsi

from . import core
from . import constants
from . import exc
from . import utils

class NotAContainerError (exc.ActiveDirectoryError):
  pass

class _ADContainer (object):
  ur"""A support object which takes an existing AD COM object
  which implements the IADsContainer interface and provides
  a corresponding iterator.

  It is not expected to be called by user code (although it
  can be). It is the basis of the :meth:`ADBase.__iter__` method
  of :class:`ADBase` and its subclasses.
  """

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
          yield item
      else:
        break

class ADCore (object):

  _properties = ["ADsPath", "Class", "GUID", "Name", "Parent", "Schema"]

  def __init__ (self, obj, cred=None):
    utils._set (self, "com_object", adsi._get_good_ret (obj))
    utils._set (self, "cred", cred)

  @staticmethod
  def _munged_attribute (name):
    #
    # AD names are either camelCase or hyphen-separated, never underscored
    # Since Python identifiers can't include hypens but can include underscores,
    # translate underscores to hyphens.
    #
    return u"-".join (name.rstrip ("_").split (u"_"))

  @staticmethod
  def _unmunged_attribute (name):
    return "_".join (name.split ("-"))

  def __repr__ (self):
    return u"<%s: %s>" % (self.__class__.__name__, self.as_string ())

  def __str__ (self):
    return self.as_string ()

  def __getattr__ (self, name):
    try:
      return exc.wrapped (self.com_object.Get, name)
    except AttributeError:
      try:
        return exc.wrapped (self.com_object.Get, self._munged_attribute (name))
      except AttributeError:
        return exc.wrapped (getattr, self.com_object, name)

  def _put (self, name, value):
    #
    # The only way to clear an AD value is by using the
    # extended PutEx mechanism.
    #
    if value is None:
      exc.wrapped (self.com_object.PutEx, constants.ADS_PROPERTY.CLEAR, name, None)
    else:
      exc.wrapped (self.com_object.Put, name, value)

  def __setattr__ (self, name, value):
    logger.debug ("name=%s, value=%s", name, value)
    munged_name = self._munged_attribute (name)
    self._put (munged_name, value)
    exc.wrapped (self.com_object.SetInfo)

  def __eq__ (self, other):
    return self.com_object.ADsPath == other.com_object.ADsPath

  def __hash__ (self):
    return hash (self.com_object.ADsPath)

  def __iter__(self):
    try:
      for item in _ADContainer (self.com_object):
        yield self.__class__.factory (item)
    except NotAContainerError:
      raise TypeError ("%r is not iterable" % self)

  @classmethod
  def from_path (cls, path, cred=None):
    ur"""Create an object of this class from an AD path

    :param obj_or_path: a valid LDAP moniker
    :returns: a :class:`ADBase` object
    """
    return cls (core.open_object (path, cred=cred))

  @classmethod
  def from_obj (cls, obj, cred=None):
    return cls (obj, cred=cred)

  @classmethod
  def factory (cls, obj_or_path=None, cred=None):
    ur"""Return an :class:`ADBase` object corresponding to `obj_or_path`.

    * If `obj_or_path` is an existing instance of this class, return it
    * If `obj_or_path` is a Python COM object, return an instance of this class which wraps it
    * If `obj_or_path` has a `com_object` attribute return an instance of this class which wraps it
    * Otherwise, assume that `obj_or_path` is an LDAP path and return the
      corresponding instance of this class

    :param obj_or_path: an existing instance of this or a related class, a Python COM object or an LDAP moniker
    :returns: an instance of this class
    """
    if isinstance (obj_or_path, cls):
      return obj_or_path
    elif isinstance (obj_or_path, (adsi.IDispatchType, win32com.client.CDispatch)):
      return cls.from_obj (obj_or_path, cred=cred)
    elif hasattr (obj_or_path, "com_object"):
      return cls.from_obj (obj_or_path.com_object, cred=cred)
    else:
      return cls.from_path (obj_or_path, cred=cred)

  def as_string (self):
    return self.com_object.ADsPath

  def dump (self, ofile=sys.stdout):
    ur"""Pretty-print the contents of this object, starting with the
    AD class definition, and followed by the attributes of this particular
    instance.

    :param ofile: the open file to write output to [`sys.stdout`]
    """
    def munged (value):
      if isinstance (value, unicode):
        value = value.encode ("ascii", "backslashreplace")
      return value
    ofile.write (self.as_string () + u"\n")
    ofile.write ("[\n")
    for property in self.__class__._properties:
      try:
        value = exc.wrapped (getattr, self, property, None)
      except NotImplementedError:
        value = None
      if not value in (None, ''):
        ofile.write ("  %s => %r\n" % (unicode (property).encode ("ascii", "backslashreplace"), munged (value)))
    ofile.write ("]\n")

class RootDSE (ADCore):

  _properties = ADCore._properties + u"""configurationNamingContext
currentTime
defaultNamingContext
dnsHostName
domainControllerFunctionality
domainFunctionality
dsServiceName
forestFunctionality
highestCommittedUSN
isGlobalCatalogReady
isSynchronized
ldapServiceName
namingContexts
rootDomainNamingContext
schemaNamingContext
serverName
subschemaSubentry
supportedCapabilities
supportedControl
supportedLDAPPolicies
supportedLDAPVersion
supportedSASLMechanisms
  """.split ()

adcore = ADCore.factory

def namespaces ():
  return ADCore (core.namespaces ())

def root_dse (server=None):
  return RootDSE (core.root_dse (server=server))

def attribute (*args, **kwargs):
  return adcore (core.attribute (*args, **kwargs))