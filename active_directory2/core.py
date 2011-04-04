# -*- coding: iso-8859-1 -*-
import re

import win32com.client
from win32com import adsi
from win32com.adsi import adsicon

from . import constants
from . import credentials
from . import exc
from .log import logger
from . import utils

def and_ (*args, **kwargs):
  """Combine its arguments together as a valid LDAP AND-search. Positional
  arguments are taken to be strings already in the correct format (eg
  'displayName=tim*') while keyword arguments will be converted into
  an equals condition for the names and values::

    from active_directory.core import and_

    print and_ (
      "whenCreated>=2010-01-01",
      displayName="tim*", objectCategory="person"
    )

    # &(whenCreated>=2010-01-01)(displayName=tim*)(objectCategory=person)
  """
  return u"&%s" % "".join ([u"(%s)" % s for s in args] + [u"(%s=%s)" % (k, v) for (k, v) in kwargs.items ()])

def or_ (*args, **kwargs):
  """Combine its arguments together as a valid LDAP OR-search. Positional
  arguments are taken to be strings already in the correct format (eg
  'displayName=tim*') while keyword arguments will be converted into
  an equals condition for the names and values::

    from active_directory.core import or_

    print or_ (
      "whenCreated>=2010-01-01",
      objectCategory="person"
    )

    # |(whenCreated>=2010-01-01)(objectCategory=person)
  """
  return u"|%s" % u"".join ([u"(%s)" % s for s in args] + [u"(%s=%s)" % (k, v) for (k, v) in kwargs.items ()])

def connect (
  cred=None,
  #~ is_password_encrypted=False,
  adsi_flags=constants.AUTHENTICATION_TYPES.DEFAULT
):
  u"""Return an ADODB connection, optionally authenticated by cred

  :param cred: anything accepted by :func:`credentials.credentials`
  :param adsi_flags: any combination of :data:`constants.AUTHENTICATION_TYPES`
  :returns: an ADO connection, optionally authenticated by `cred`
  """
  cred = credentials.credentials (cred)
  if cred is None:
    cred = credentials.Passthrough
  connection = exc.wrapped (win32com.client.Dispatch, u"ADODB.Connection")
  connection.Provider = u"ADsDSOObject"
  if cred.username:
    connection.Properties ("User ID").Value = cred.username
  if cred.password:
    connection.Properties ("Password").Value = cred.password
  #~ connection.Properties ("Encrypt Password").Value = is_password_encrypted
  connection.Properties ("ADSI Flag").Value = adsi_flags | cred.authentication_type
  exc.wrapped (connection.Open, u"Active Directory Provider")
  return connection

#
# If Page Size is unset (system default is 0) then a maximum of 1000
# records will be returned from any query before an error is raised
# by the AD provider. Therefore we default to 500 to give a reasonable
# default. This can still be overridden at query level.
#
_command_properties = {
  u"Page Size" : 500,
  u"Asynchronous" : True
}
def query (query_string, connection=None, **command_properties):
  u"""Basic AD query, passing a raw query string straight through to an
  Active Directory, optionally using a (possibly pre-authenticated) connection
  or creating one on demand. command_properties may be specified which will be
  passed through to the ADO command with underscores replaced by spaces. Useful
  values include:

  =============== ==========================================================
  page_size       How many records to return in one go
  size_limit      Stop after returning this many records
  cache_results   Boolean: cache results; turn off if a large result
  time_limit      Stop returning records after this many seconds
  timeout         Stop waiting for the records to start after this many seconds
  asynchronous    Boolean: Start returning records immediately
  sort_on         field name to sort on
  =============== ==========================================================

  :param query_string: An AD query string in any acceptable format. See :func:`query_string`
                       for an easy way of producing this
  :param connection: (optional) An ADODB.Connection, as provided by :func:`connect`. If
                     this is supplied it will be used and not closed. If it is not supplied,
                     a default connection will be created, used and then closed.
  :param command_properties: A collection of keywords which will be passed through to the
                             ADO query as Properties.
  """
  command = exc.wrapped (win32com.client.Dispatch, u"ADODB.Command")
  _connection = connection or connect ()
  command.ActiveConnection = _connection

  for k, v in _command_properties.items ():
    command.Properties (k.replace (u"_", u" ")).Value = v
  for k, v in command_properties.items ():
    command.Properties (k.replace (u"_", u" ")).Value = v
  command.CommandText = query_string

  results = []
  recordset, result = exc.wrapped (command.Execute)
  while not recordset.EOF:
    yield dict ((field.Name, field.Value) for field in recordset.Fields)
    exc.wrapped (recordset.MoveNext)

  if connection is None:
    exc.wrapped (_connection.Close)

def query_string (filter="", base=None, attributes=[u"ADsPath"], scope=u"Subtree", range=None):
  u"""Easy way to produce a valid AD query string, with meaningful defaults. This
  is the first parameter to the :func:`query` function so the following will
  yield the display name of every user in the domain::

    import active_directory as ad

    qs = ad.query_string (filter="(objectClass=User)", attributes=["displayName"])
    for u in ad.query (qs):
      print u['displayName']

  :param filter: An AD filter string to limit the search [no filter]. The :func:`or_` and :func:`and_`
                 functions provide an easy way to produce a valid filter, optionally combined with the
                 schema class.
  :param base: An LDAP:// moniker representing the starting point of the search [domain root]
  :param attributes: Iterable of attribute names [ADsPath]
  :param scope: One of - Subtree, Base, OneLevel. Subtree (the default) is the most common and does
                the search you expect. OneLevel enumerates the children of the base item. Base
                checks for the existence of the object itself. [Subtree].
  :param range: Limit the number of returns of multivalued attributes [no range]
  """
  if base is None:
    base = core.root_moniker ()
  if filter and not re.match (r"\([^)]+\)", filter):
    filter = u"(%s)" % filter
  segments = [u"<%s>" % base, filter, ",".join (attributes)]
  if range:
    segments += [u"Range=%s-%s" % range]
  segments += [scope]
  return u";".join (segments)

_base_monikers = {}
def _base_moniker (server=None, scheme="LDAP:"):
  if (server, scheme) not in _base_monikers:
    if server:
      _base_monikers[server, scheme] = scheme + "//" + server + "/"
    else:
      _base_monikers[server, scheme] = scheme + "//"
  return _base_monikers[server, scheme]

_root_dses = {}
def root_dse (server=None, scheme="LDAP:"):
  u"""Return the object representing the RootDSE for a domain, optionally
  specified by a server and a scheme (typically LDAP: or GC:).

  :param server: A specific server whose rootDSE is to be found [none - any server]
  :param scheme: Typically LDAP: or GC: [LDAP:]
  :returns: The COM Object corresponding to the RootDSE for the server or domain
  """
  if (server, scheme) not in _root_dses:
    _root_dses[server, scheme] = exc.wrapped (
      win32com.client.GetObject,
      _base_moniker (server, scheme) + "rootDSE"
    )
  return _root_dses[server, scheme]

_root_monikers = {}
def root_moniker (server=None, scheme="LDAP:"):
  u"""Return the moniker representing the domain specified by a server and
  a scheme (typically LDAP: or GC:). If you need the corresponding object,
  use :func:`root_obj`.

  :param server: A specific server whose rootDSE is to be found [none - any server]
  :param scheme: Typically LDAP: or GC: [LDAP:]
  :returns: The moniker corresponding to the domain
  """
  if (server, scheme) not in _root_monikers:
    dse = root_dse (server, scheme)
    _root_monikers[server, scheme] = _base_moniker (server, scheme) + dse.Get ("defaultNamingContext")
  return _root_monikers[server, scheme]

_root_objs = {}
def root_obj (server=None, scheme="LDAP:", cred=None):
  u"""Return the COM object representing the domain specified by a server and
  a scheme (typically LDAP: or GC:), optionally authenticated. If you only
  need the moniker, use :func:`root_moniker`.

  :param server: A specific server whose rootDSE is to be found [none - any server]
  :param scheme: Typically LDAP: or GC: [LDAP:]
  :param cred: anything accepted by :func:`credentials.credentials`
  :returns: The COM object corresponding to the domain
  """
  return open_object (root_moniker (server, scheme), cred=cred)

_schema_objs = {}
def schema_obj (server=None, cred=None):
  u"""Return the COM object representing the schema for the domain specified
  by a server, optionally authenticated.

  :param server: A specific server whose rootDSE is to be found [none - any server]
  :param cred: anything accepted by :func:`credentials.credentials`
  :returns: The COM object corresponding to the domain schema
  """
  if server not in _schema_objs:
    dse = root_dse (server)
    _schema_objs[server] = open_object (
      _base_moniker (server) + dse.Get ("schemaNamingContext"),
      cred=cred
    )
  return _schema_objs[server]

_attributes = {}
_attribute_info = ['lDAPDisplayName', 'instanceType', 'oMObjectClass', 'oMSyntax', 'attributeId', 'isSingleValued']
def attribute_info (names=["*"], server=None, cred=None):
  u"""Return an iteration of name, dict pairs representing all the attributes named.
  The dict contains: lDAPDisplayName, instanceType, oMObjectClass, oMSyntax, attributeId, isSingleValued

  :param names: A list of names for attributes to be returned [all attributes]
  :param server: A specific server whose rootDSE is to be found [none - any server]
  :param cred: anything accepted by :func:`credentials.credentials`
  :returns: An iteration of tuples containing (name, info)
  """
  schema = schema_obj (server, cred)
  unknown_names = set (names) - set (_attributes)
  if unknown_names:
    filter = or_ (*["lDAPDisplayName=%s" % name for name in unknown_names])
    for row in dquery (schema, filter, _attribute_info):
      _attributes[row['lDAPDisplayName'][0]] = dict ((k, v[0]) for (k, v) in row.items ())

  if names == ['*']:
    names = iter (_attributes)
  for name in names:
    yield name, _attributes[name]

def dquery (obj, filter, attributes=None, flags=0):
  SEARCH_PREFERENCES = {
    adsicon.ADS_SEARCHPREF_PAGESIZE : 1000,
    adsicon.ADS_SEARCHPREF_SEARCH_SCOPE : adsicon.ADS_SCOPE_SUBTREE,
  }
  directory_search = exc.wrapped (obj.QueryInterface, adsi.IID_IDirectorySearch)
  directory_search.SetSearchPreference ([(k, (v,)) for k, v in SEARCH_PREFERENCES.items ()])
  if filter and not re.match (r"\([^)]+\)", filter):
    filter = u"(%s)" % filter
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

def open_object (moniker, cred=None, flags=constants.AUTHENTICATION_TYPES.DEFAULT):
  """Open an AD object represented by `moniker`, optionally authenticated. You
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
  * `moniker` is parsed to determine the (optional) server name and the cache is checked
    for corresponding credentials.
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
    flags | (constants.AUTHENTICATION_TYPES.SERVER_BIND if server else 0) | cred.authentication_type,
    adsi.IID_IADs
  )
