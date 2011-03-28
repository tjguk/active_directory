"""The ADSimple class wraps an ADSI com object with very simple attribute
read and write functionality plus a method to dump the contents looking
like a dictionary. It does very little validation of the path or the
attributes and does no conversion of the values returned.
"""
from active_directory import core
qs = core.query_string (
  filter = core.and_ ("objectClass=computer", "OperatingSystem=Windows Server*"),
  attributes=['cn', 'OperatingSystem', 'OperatingSystemVersion']
)
for computer in core.query (qs, sort_on="cn"):
  print "%(cn)s: %(OperatingSystem)s [%(OperatingSystemVersion)s]" % computer

print

"""This example illustrates every option in the query string builder. It
uses one query to pick out one (arbitrary) OU and then searches only
that OU, using it as the base for the query string and specifying no
subtree searching. The distinguishedName and whenCreated are returned
"""
from active_directory import core
for ou in core.query (
  core.query_string ("objectCategory=organizationalUnit"),
  page_size=1
):
  base = ou['ADsPath']
  break

query_string = core.query_string (
  base=base,
  filter="cn=*",
  attributes=['distinguishedName', 'whenCreated'],
  scope="OneLevel"
)
print "Querying:", query_string
for obj in core.query (query_string):
  print "%(distinguishedName)s created on %(whenCreated)s" % obj

print

"""By default, a call to query will create an ad-hoc connection which
will then be closed. This example shows how to pass an existing
connection into the query function. You might want to do this if
you needed to authenticate the connection or if you were performing
a number of queries and wanted to avoid the overhead of repeated
connection-building.
"""
import win32api
import win32con
import win32cred
from active_directory import core

username, domain = win32cred.CredUIParseUserName (win32api.GetUserNameEx (win32con.NameDnsDomain))
username, password, save = win32cred.CredUIPromptForCredentials (
  domain, 0, username,
  None, True, 0, None
)
connection = core.connect (username, password)
for obj in core.query (
  core.query_string ("objectCategory=organizationalUnit"),
  connection=connection
):
  print obj

print

"""When you create a specific connection you can pass various flags
through as the adsi_flags parameter. The possible flags are documented
on MSDN: http://msdn.microsoft.com/en-us/library/aa772247%28v=vs.85%29.aspx
and are presented via the AUTHENTICATION_TYPES enum from the constants
module. For example, using a fast bind can speed up simple operations.
"""
from active_directory import core
from active_directory.constants import AUTHENTICATION_TYPES

connection = core.connect (adsi_flags=AUTHENTICATION_TYPES.FAST_BIND)
for obj in core.query (
  core.query_string ("objectCategory=organizationalUnit"),
  connection=connection
):
  print obj

print
