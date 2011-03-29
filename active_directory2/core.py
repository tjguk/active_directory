# -*- coding: iso-8859-1 -*-
import win32com.client
from win32com import adsi
from win32com.adsi import adsicon

from . import constants
from . import credentials
from . import exc

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
  cred=credentials.Passthrough,
  #~ is_password_encrypted=False,
  adsi_flags=0
):
  u"""Return an ADODB connection, optionally authenticated by
  username & password.
  """
  cred = credentials.credentials (cred)
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
  print "query_string:", query_string
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
    base = u"LDAP://" + exc.wrapped (adsi.ADsGetObject, "LDAP://rootDSE").Get (u"defaultNamingContext")
  if filter and not filter.startswith ("("):
    filter = u"(%s)" % filter
  segments = [u"<%s>" % base, filter, ",".join (attributes)]
  if range:
    segments += [u"Range=%s-%s" % range]
  segments += [scope]
  return u";".join (segments)
