# -*- coding: iso-8859-1 -*-
ur"""Core functionality behind the active_directory2 functionality.
The functions in this module either return strings or Python COM
objects representing the underlying ADSI COM objects. These will
be wrapped by :class:`adbase.ADBase` and other classes to give extended
functionality, but they can be useful on their own.

The functions in this module are built upon by all the other modules,
but can sensibly be used in their own right by users of the active_directory2
package. For efficiency, most of them cache their results so that, eg, a
lookup for an attribute's schema doesn't need to bind to the directory
every time. (Although the bind itself should be cached by AD behind the scenes).
"""
import re
import threading

import win32com.client
from win32com import adsi
from win32com.adsi import adsicon

from . import constants
from . import credentials
from . import exc
from .log import logger
from . import support
from . import utils

def namespaces ():
  ur"""Return the ADs: namespaces object. This can only be accessed via
  a GetObject call, and can't be authenticated against.
  """
  return win32com.client.GetObject ("ADs:")

_base_monikers = {}
def _base_moniker (server=None, scheme="LDAP:"):
  ur"""Form a moniker from a server and scheme, returning a cached hit if available.

  :param server: A valid server name or `None` for a a serverless moniker
  :param scheme: A valid AD scheme; typically LDAP: but could be GC: or WinNT:
  :return: A string of the form LDAP://<server>/ where the server segment might be missing
  """
  logger.debug ("_base_moniker: server=%s, scheme=%s", server, scheme)
  if (server, scheme) not in _base_monikers:
    if server:
      _base_monikers[server, scheme] = scheme + "//" + server + "/"
    else:
      _base_monikers[server, scheme] = scheme + "//"
  return _base_monikers[server, scheme]

_root_dses = {}
def root_dse (server=None, scheme="LDAP:"):
  ur"""Return the object representing the RootDSE for a domain, optionally
  specified by a server and a scheme (typically LDAP: or GC:).

  :param server: A specific server whose rootDSE is to be found [None - serverless]
  :param scheme: Typically LDAP: or GC: [LDAP:]
  :returns: The COM Object corresponding to the RootDSE for the server or domain
  """
  logger.debug ("root_dse: server=%s, scheme=%s", server, scheme)
  if (server, scheme) not in _root_dses:
    _root_dses[server, scheme] = exc.wrapped (
      win32com.client.GetObject,
      _base_moniker (server, scheme) + "rootDSE"
    )
  return _root_dses[server, scheme]

_root_monikers = {}
def root_moniker (server=None, scheme="LDAP:"):
  ur"""Return the moniker representing the domain specified by a server and
  a scheme (typically LDAP: or GC:). If you need the corresponding object,
  use :func:`root_obj`.

  :param server: A specific server whose rootDSE is to be found [none - any server]
  :param scheme: Typically LDAP: or GC: [LDAP:]
  :returns: The moniker corresponding to the domain
  """
  if (server, scheme) not in _root_monikers:
    dse = root_dse (server, scheme)
    _root_monikers[server, scheme] = \
      _base_moniker (server, scheme) + exc.wrapped (dse.Get, "defaultNamingContext")
  return _root_monikers[server, scheme]

_root_objs = {}
def root_obj (server=None, scheme="LDAP:", cred=None):
  ur"""Return the COM object representing the domain specified by a server and
  a scheme (typically LDAP: or GC:), optionally authenticated. If you only
  need the moniker, use :func:`root_moniker`.

  :param server: A specific server whose rootDSE is to be found [none - any server]
  :param scheme: Typically LDAP: or GC: [LDAP:]
  :param cred: anything accepted by :func:`credentials.credentials`
  :returns: The COM object corresponding to the domain
  """
  if server not in _root_objs:
    _root_objs[server] = open_object (root_moniker (server, scheme), cred=cred)
  return _root_objs[server]

def _partition_obj (partition, server=None, cred=None):
  return open_object (
    _base_moniker (server) + exc.wrapped (root_dse (server).Get, partition),
    cred=cred
  )

_schema_objs = {}
def schema_obj (server=None, cred=None):
  ur"""Return the COM object representing the schema for the domain specified
  by a server, optionally authenticated.

  :param server: A specific server whose rootDSE is to be found [none - any server]
  :param cred: anything accepted by :func:`credentials.credentials`
  :returns: The COM object corresponding to the domain schema
  """
  if server not in _schema_objs:
    _schema_objs[server] = _partition_obj ("schemaNamingContext", server, cred)
  return _schema_objs[server]

_configuration_objs = {}
def configuration_obj (server=None, cred=None):
  ur"""Return the COM object representing the configuration for the domain specified
  by a server, optionally authenticated.

  :param server: A specific server whose rootDSE is to be found [none - any server]
  :param cred: anything accepted by :func:`credentials.credentials`
  :returns: The COM object corresponding to the domain configuration
  """
  if server not in _configuration_objs:
    _configuration_objs[server] = _partition_obj ("configurationNamingContext", server, cred)
  return _configuration_objs[server]

_class_schemas = {}
def class_schema (class_name, server=None, cred=None):
  ur""":returns: the name of the schema for a particular AD Class
  """
  if class_name not in _class_schemas:
    _class_schemas[class_name] = open_object (_base_moniker (server) + "schema/%s" % class_name, cred=cred)
  return _class_schemas[class_name]

_attributes = {}
def attributes (names="*", server=None, cred=None):
  ur"""Return an iteration of name, obj pairs representing all the attributes named.
  The dict contains: lDAPDisplayName, instanceType, oMObjectClass, oMSyntax, attributeId, isSingleValued

  :param names: A list of names for attributes to be returned [all attributes]
  :param server: A specific server whose rootDSE is to be found [none - any server]
  :param cred: anything accepted by :func:`credentials.credentials`
  :returns: An iteration of (name, info)
  """
  schema = schema_obj (server, cred)
  unknown_names = set (names) - set (_attributes)
  if unknown_names:
    filter = support.and_ (
      "objectCategory=attributeSchema",
      support.or_ (*["lDAPDisplayName=%s" % name for name in unknown_names])
    )
    for row in query (schema, filter, ['lDAPDisplayName', 'ADsPath']):
      _attributes[row['lDAPDisplayName'][0]] = open_object (row['ADsPath'][0], cred=cred)

  if names == "*":
    names = iter (_attributes)

  for name in names:
    try:
      yield name, _attributes[name]
    except KeyError:
      raise exc.AttributeNotFound (name)

def attribute (name, server=None, cred=None):
  ur"""Return the first attribute corresponding to `name` from :func:`attributes`.

  :param name: The name of an attribute whose data is to be returned
  :param server: A specific server whose rootDSE is to be found [`None` - any server]
  :param cred: anything accepted by :func:`credentials.credentials`

  :returns: `name`, `info` per :func:`attributes` for the named attribute
  """
  for name, attr in attributes ([name], server=server, cred=cred):
    return attr

def query (obj, filter, attributes=None, flags=constants.ADS_SEARCHPREF.Unset):
  ur"""Run an LDAP query specified by `filter` against the AD object `obj`.
  This query is at the heart of the search functionality in this package.
  It can be called directly either from this module or from any of the
  higher-level AD objects such as :class:`adbase.ADBase` which expose
  it as a method.

  The result is an iteration of dictionaries mapping attribute names
  to a list of values. No attempt is made here to determine which
  values are single and which are multivalued.

  Typical usage:

  :param obj: An ADSI object which implements the IDirectorySearch interface
  :param filter: A valid ADSI/LDAP filter string
  :param attributes: A list of attributes (AD fields) to return. None => All
  :param flags: A combination of :data:`constants.ADS_SEARCHPREF` values
  :returns: iterator over a dictionary mapping attribute names to lists of values
  """
  SEARCH_PREFERENCES = {
    adsicon.ADS_SEARCHPREF_PAGESIZE : 1000,
    adsicon.ADS_SEARCHPREF_SEARCH_SCOPE : adsicon.ADS_SCOPE_SUBTREE,
  }
  directory_search = exc.wrapped (obj.QueryInterface, adsi.IID_IDirectorySearch)
  exc.wrapped (directory_search.SetSearchPreference, [(k, (v,)) for k, v in SEARCH_PREFERENCES.items ()])
  #
  # Simple queries may not be bracketed; add surrounding brackets if necessary
  #
  if filter and not re.match (r"\([^)]+\)", filter):
    filter = u"(%s)" % filter
  hSearch = exc.wrapped (directory_search.ExecuteSearch, filter, attributes)
  try:
    hResult = exc.wrapped (directory_search.GetFirstRow, hSearch)
    while hResult == 0:
      results = {}
      while True:
        attr = exc.wrapped (directory_search.GetNextColumnName, hSearch)
        if attr is None:
          break
        _, _, attr_values = exc.wrapped (directory_search.GetColumn, hSearch, attr)
        results[attr] = [value for (value, type) in attr_values]
      yield results
      hResult = exc.wrapped (directory_search.GetNextRow, hSearch)
  finally:
    exc.wrapped (directory_search.AbandonSearch, hSearch)
    exc.wrapped (directory_search.CloseSearchHandle, hSearch)

def open_object (moniker, cred=None, flags=constants.AUTHENTICATION_TYPES.DEFAULT):
  ur"""Open an AD object represented by `moniker`, optionally authenticated. You
  will not normally call this yourself: it is used internally by the AD objects.

  :param moniker: A complete AD moniker representing an AD object
  :param cred: anything accepted by :func:`credentials.credentials`
  :param flags: optional :data:`constants.AUTHENTICATION_TYPES` flags. The credentials
                will set the appropriate flags for authentication, and server binding will be used
                if the moniker is server-based.
  :returns: a COM object corresponding to `moniker` and authenticated according to `cred`

  This function is at the heart of authenticated access to AD offered by this package.
  The credentials work as follows:

  * `cred` is passed to :func:`credentials.credentials` for initial processing
  * If `cred` is now a :class:`credentials.Credentials` object, this is used for authentication
  * Otherwise `moniker` is parsed to determine the (optional) server name and the cache is
    checked for corresponding credentials.
  * If no cached credentials are found, passthrough authentication is assumed.

  This will normally do what you expect. The default (passthrough) is far and away
  the most common. Specific credentials can either be passed in, eg, as a tuple,
  or can be held in the credentials cache and inferred from the server::

    from active_directory2 import core, credentials

    me = core.open_object ("LDAP://cn=Tim Golden,dc=goldent,dc=local")
    me = core.open_object ("LDAP://cn=Tim Golden,dc=goldent,dc=local", cred=("goldent\\tim", "pa55w0rd"))
    with credentials.credentials (("goldent\\tim", "5ecret", "testing")):
      me = core.open_object (core.root_moniker ())
    me = core.open_object ("LDAP://testing/dc=test,dc=local", cred=credentials.Anonymous)
  """
  scheme, server, dn = utils.parse_moniker (moniker)
  cred = credentials.credentials (cred)
  if cred is None:
    cred = credentials.Credentials.cache.get (server.rstrip ("/"))
  if cred is None:
    cred = credentials.Passthrough
  return exc.wrapped (
    adsi.ADsOpenObject,
    moniker,
    cred.username,
    cred.password,
    flags | (constants.AUTHENTICATION_TYPES.SERVER_BIND if server else 0) | cred.authentication_type
  )
