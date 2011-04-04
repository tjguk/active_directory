import os, sys
import pythoncom
from win32com.adsi import adsi, adsicon

SEARCH_PREFERENCES = {
  adsicon.ADS_SEARCHPREF_PAGESIZE : 1000,
  adsicon.ADS_SEARCHPREF_SEARCH_SCOPE : adsicon.ADS_SCOPE_SUBTREE,
}

class Result (dict):
  def __getattr__ (self, attr):
    return self[attr]

ESCAPED_CHARACTERS = dict ((special, r"\%02x" % ord (special)) for special in "*()\x00/")
def escaped (s):
  for original, escape in ESCAPED_CHARACTERS.items ():
    s = s.replace (original, escape)
  return s

def ldap_moniker (root=None, server=None, username=None, password=None):
  if root is None:
    root = adsi.ADsOpenObject (
      ldap_moniker ("rootDSE", server),
      username, password,
      adsicon.ADS_SECURE_AUTHENTICATION | adsicon.ADS_SERVER_BIND | adsicon.ADS_FAST_BIND,
      adsi.IID_IADs
    ).Get ("defaultNamingContext")
  if server:
    return "LDAP://%s/%s" % (server, root)
  else:
    return "LDAP://%s" % root

def connect (server=None, root=None, username=None, password=None):
    return adsi.ADsOpenObject (
      ldap_moniker (root, server, username, password),
      username, password,
      adsicon.ADS_SECURE_AUTHENTICATION | adsicon.ADS_SERVER_BIND | adsicon.ADS_FAST_BIND,
      adsi.IID_IADs
    )


def search (filter, columns=["distinguishedName"], root=None, server=None, username=None, password=None):

  def get_column_value (hSearch, column):
    #
    # FIXME: Need a more general-purpose way of determining which
    # fields are indeed lists. Either a factory function or a
    # peek at the schema.
    #
    CONVERT_TO_LIST = set (['memberOf', "member", "proxyAddresses"])
    try:
      column_name, column_type, column_values = directory_search.GetColumn (hSearch, column)
      if column_name in CONVERT_TO_LIST:
        return list (value for value, type in column_values)
      else:
        for value, type in column_values:
          return value
    except adsi.error as details:
      if details.args[0] == adsicon.E_ADS_COLUMN_NOT_SET:
        return None
      else:
        raise

  pythoncom.CoInitialize ()
  try:
    directory_search = adsi.ADsOpenObject (
      ldap_moniker (root, server, username, password),
      username, password,
      adsicon.ADS_SECURE_AUTHENTICATION | adsicon.ADS_SERVER_BIND | adsicon.ADS_FAST_BIND,
      adsi.IID_IDirectorySearch
    )
    directory_search.SetSearchPreference ([(k, (v,)) for k, v in SEARCH_PREFERENCES.items ()])

    hSearch = directory_search.ExecuteSearch (filter, columns)
    try:
      hResult = directory_search.GetFirstRow (hSearch)
      while hResult == 0:
        yield Result ((column, get_column_value (hSearch, column)) for column in columns)
        hResult = directory_search.GetNextRow (hSearch)
    finally:
      directory_search.AbandonSearch (hSearch)
      directory_search.CloseSearchHandle (hSearch)

  finally:
    pythoncom.CoUninitialize ()

def _and (*args):
  return "(&%s)" % "".join ("(%s)" % s for s in args)

def _or (*args):
  return "(|%s)" % "".join ("(%s)" % s for s in args)

def find_user (name, root_path=None, server=None, username=None, password=None):
  name = escaped (name)
  for user in search (
    _and (
      "objectClass=user",
      "objectCategory=person",
      _or (
        "sAMAccountName=" + name,
        "displayName=" + name,
        "cn=" + name
      )
    ),
    ["distinguishedName", "sAMAccountName", "displayName", "memberOf", "physicalDeliveryOfficeName", "title", "telephoneNumber", "homePhone", "proxyAddresses"],
    root_path,
    server,
    username,
    password
  ):
    return user

def find_group (name, root_path=None, server=None, username=None, password=None):
  name = escaped (name)
  for group in search (
    _and ("objectClass=group", _or ("sAMAccountName=" + name, "displayName=" + name, "cn=" + name)),
    ["distinguishedName", "sAMAccountName", "displayName", "member"],
    root_path,
    server,
    username,
    password
  ):
    return group

def find_active_users (root=None, server=None, username=None, password=None):
  return search (
    filter=_and (
      "objectClass=user",
      "objectCategory=person",
      "!memberOf=CN=non intranet,OU=IT Other,OU=IT,OU=Camden,DC=gb,DC=vo,DC=local",
      "!userAccountControl:1.2.840.113556.1.4.803:=2",
      "displayName=*"
    ),
    columns=[
      "distinguishedName",
      "sAMAccountName",
      "displayName",
      "memberOf",
      "physicalDeliveryOfficeName",
      "title",
      "telephoneNumber",
      "homePhone",
      "department",
      "proxyAddresses",
      "mobile",
      "scriptPath",
    ],
    root=None, server=None, username=None, password=None
  )
##
## To support existing IT Support code
##
active_users = find_active_users

def find_all_users (root=None, server=None, username=None, password=None):
  return search (
    filter=_and ("objectClass=user", "objectCategory=person", "displayName=*"),
    columns=["userAccountControl", "distinguishedName", "sAMAccountName", "displayName", "memberOf", "physicalDeliveryOfficeName", "title", "telephoneNumber", "homePhone", "department", "proxyAddresses"],
    root=None, server=None, username=None, password=None
  )
