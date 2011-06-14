active_directory2 - a Python interface to Active Directory
==========================================================

Active Directory is an LDAP-style database managing all the objects in a Windows
domain structure: users, computers, shares, printers, the domains themselves, &c.
This package presents a Python interface to Active Directory via the ADSI API.

Some effort has been made to ensure the package is useful at the interpreter as
much as in a running program. The :mod:`ad` module exposes convenient functions
for day-to-day use which make use of the lower-level modules::

  from active_directory2 import ad

  fred = ad.find_user ("Fred Smith")
  fax_users = ad.find_group ("Fax Users")
  if fax_users.dn not in fred.memberOf:
    print "Fred is not a fax user"

The lowest level of the package is the :mod:`core`
module which exposes some of the basic ADSI operations such as accessing an AD
object with optional credentials or querying the root of the domain.All of its
functions return strings of pywin32 COM objects.::

  from active_directory2 import core

  root = core.root_obj (cred=("tim", "Pa55w0rd"))
  for result in core.query (root, "(displayName=Tim Golden)"):
    print result

The :mod:`adbase` module builds
on `core` and exposes an Python class for every underlying AD objects. This class
wrapper is still fairly thin but does make certain operations slightly more intuitive
for the Python programmer::

  from active_directory2 import core, adbase, support, constants

  adroot = adbase.adbase (core.root_obj ())
  archive_ou = adroot.find_ou ("archive")

  for disabled_account in adroot.search (
    support.band (
      "userAccountControl",
      constants.USER_ACCOUNT_CONTROL.ACCOUNTDISABLE
    ),
    objectCategory="person",
    objectClass="user"
  ):
    disabled_account.move (archive_ou)