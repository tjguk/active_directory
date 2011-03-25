from active_directory import core

"""Find the dn, Operating System & OS Version of all computers
running some kind of Windows Server. Sort the output by the
computer's name (its cn in AD).

core.query_string is a convenience to produce an LDAP search
string. It provides useful defaults for all fields so here
we're not supplying a base for the query (which will be
the root of our AD) nor a scope (which will be subtree).

core.query creates an ad-hoc connection to AD and issues the
query string you supply. You can specify as keyword arguments
any properties which the ADO connection support. The ADO
flags are space-separated titlecase words; the Python equivalents
are underscore_delimited lowercase.
"""
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

