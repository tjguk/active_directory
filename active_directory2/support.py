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
  return u"%s:1.2.840.113556.1.4.803:=%s" % (name, value)

def bor (name, value):
  return u"%s:1.2.840.113556.1.4.804:=%s" % (name, value)

def not_ (expression):
  return u"!%s" % expression

def within (name, dn):
  return u"%s:1.2.840.113556.1.4.1941:=%s" % (self._name, dn)
