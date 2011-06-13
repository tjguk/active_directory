__all__ = ['and_', 'or_', 'not_', 'band', 'bor']

def and_ (*args, **kwargs):
  ur"""Combine its arguments together as a valid LDAP AND-search. Positional
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
  params = [u"(%s)" % s for s in args] + [u"(%s=%s)" % (k, v) for (k, v) in kwargs.items ()]
  if len (params) < 2:
    return "".join (params)
  else:
    return u"&%s" % "".join (params)

def or_ (*args, **kwargs):
  ur"""Combine its arguments together as a valid LDAP OR-search. Positional
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
  params = [u"(%s)" % s for s in args] + [u"(%s=%s)" % (k, v) for (k, v) in kwargs.items ()]
  if len (params) < 2:
    return "".join (params)
  else:
    return u"|%s" % u"".join (params)

def band (name, value):
  ur"""Perform bitwise-and matching between an AD field and a numeric
  value. A typical use would be to check whether an account is disabled::

    from active_directory2 import ad, constants, support

    for disabled_account in ad.query (
      support.band (
        "userAccountControl",
        constants.USER_ACCOUNT_CONTROL.ACCOUNTDISABLE
      )
    ):
      print "%s (%s)" % (disabled_account.displayName, disabled_account.sAMAccountName)
  """
  return u"%s:1.2.840.113556.1.4.803:=%s" % (name, value)

def bor (name, value):
  ur"""Perform bitwise-or matching between an AD field and a numeric value.
  """
  return u"%s:1.2.840.113556.1.4.804:=%s" % (name, value)

def not_ (*args, **kwargs):
  ur"""Return the logically negated form of the one expression which can be either
  a preformed expression in a positional argument or a keyword arg which will be
  treated as an equality check::

    from active_directory2 import support

    account_disabled = support.band ("userAccountControl", constants.USER_ACCOUNT_CONTROL.ACCOUNTDISABLE)
    print support.not_ (account_disabled)
    print support.not_ (sAMAccountName="tim")
  """
  if len (args) > 1:
    raise TypeError ("Can only specify one expression for not")
  if len (kwargs) > 1:
    raise TypeError ("Can only specify one keyword arg for not")
  if args and kwargs:
    raise TypeError ("Can only specify arg or kwargs")
  if args:
    expression = args[0]
  else:
    for k, v in kwargs.items ():
      expression = "%s=%s" % (k, v)
  return u"!(%s)" % expression

def within (name, dn):
  ur"""Return the LDAP string expression for efficiently searching up in a
  hierarchy to discover a parent. This is typically used to check for
  membership in a group through no matter how many levels::

    from active_directory2 import ad, support

    domain_admins = ad.find_group ("Domain Admins")
    print support.within ("memberOf", domain_admins.distinguishedName)
  """
  return u"%s:1.2.840.113556.1.4.1941:=%s" % (name, dn)

def searchable_sid (sid):
  ur"""Return a string version of the binary Sid which can be used
  for searching, eg to find a user account by well-known Sid.

  :param sid: a Sid object which exposes a buffer whose bytes consist of
              the binary version of the Sid
  :returns: A string consisting of the hexadecimal version of the `sid` buffer
  """
  return u"".join (u"\\%02x" % ord (x) for x in buffer (sid))

def rdn (dn0, dn1):
  ur"""Return a relative distinguished name which related dn1 to dn0::

    from active_directory2 import support
    print support.rdn ("OU=tim,DC=westpark,DC=local", "CN=Group01,OU=tim,DC=westpark,DC=local")
  """
  #
  # We will assume that dn0 is the parent (shorter) dn
  #
  if len (dn0) > len (dn1):
    dn0, dn1 = dn1, dn0

  if not dn1.endswith (dn0):
    raise RuntimeError ("The distinguished names are not related")

  return dn1[:-len (dn0)-1]
