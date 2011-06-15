import netrc
import win32cred

from . import constants
from . import exc
from .log import logger

class CredentialsError (exc.ActiveDirectoryError):
  pass

class NetrcNotFoundError (CredentialsError):
  pass

class InvalidCredentialsError (CredentialsError):
  pass

class CredentialsAlreadyCachedError (CredentialsError):
  pass

class CredentialsCache (object):

  #
  # TODO: This needs to use the secure cacheing mechanism
  #

  def __init__ (self):
    self._cache = {}

  def __repr__ (self):
    return "<%s: %s>" % (self.__class__.__name__, list (self._cache) or "Empty")

  def __str__ (self):
    return str (self._cache)

  def push (self, cred):
    cred = credentials (cred)
    self._cache.setdefault (cred.server, []).append (cred)

  def pop (self, server):
    return self._cache[server].pop ()

  def get (self, server, default=None):
    creds = self._cache.setdefault (server, [])
    if creds:
      return creds[-1]
    else:
      return default

  def __iter__ (self):
    return ((k, v[-1]) for (k, v) in self._cache.items ())

  def clear (self):
    self._cache.clear ()

class Credentials (object):

  ANONYMOUS = 0
  SIMPLE = 1
  PASSTHROUGH = 2

  cache = CredentialsCache ()

  def __init__ (self, username, password, server=None, type=SIMPLE):
    self.username = username
    self.password = password
    self.server = server
    self.type = type
    if type == self.ANONYMOUS:
      self.authentication_type = constants.AUTHENTICATION_TYPES.NO_AUTHENTICATION
    else:
      self.authentication_type = constants.AUTHENTICATION_TYPES.SECURE_AUTHENTICATION

  def __repr__ (self):
    return "<%s: %r %r on %s>" % (self.__class__.__name__, self.username, self.password, self.server)

  def __enter__ (self):
    self.__class__.cache.push (self)
    return self

  def __exit__ (self, *args):
    self.__class__.cache.pop (self.server)

  @classmethod
  def from_netrc (cls, host, netrc_filepath=None):
    auth = netrc.netrc (netrc_filepath).authenticators (host)
    if auth:
      login, _, password = auth
      return cls (login, password)
    else:
      raise NetrcNotFoundError ("No entry for %s in netrc" % host)

  @classmethod
  def from_cache (cls, target):
    raise NotImplementedError

Passthrough = Credentials (None, None, Credentials.PASSTHROUGH)
Anonymous = Credentials (None, None, Credentials.ANONYMOUS)

def credentials (cred):
  if cred is None:
    return None
  elif isinstance (cred, Credentials):
    return cred
  else:
    try:
      return Credentials (*cred)
    except (ValueError, TypeError):
      raise InvalidCredentialsError ("Credentials must be a Credentials object or (username, password[, server])")

cache = Credentials.cache

def push (cred):
  cred = credentials (cred)
  Credentials.cache.push (cred)

def pop (server):
  return Credentials.cache.pop (server)
