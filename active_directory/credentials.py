import netrc
import win32cred

from . import constants
from . import exc

class NetrcNotFound (exc.ActiveDirectoryError):
  pass

class Credentials (object):

  ANONYMOUS = 0
  SIMPLE = 1
  PASSTHROUGH = 2

  def __init__ (self, username, password, type=SIMPLE):
    self.username = username
    self.password = password
    self.type = type
    if type == self.ANONYMOUS:
      self.authentication_type = constants.AUTHENTICATION_TYPES.NO_AUTHENTICATION
    else:
      self.authentication_type = constants.AUTHENTICATION_TYPES.SECURE_AUTHENTICATION

  def __repr__ (self):
    return "<%s: %r %r>" % (self.__class__.__name__, self.username, self.password)

  def save_to_cache (self, target):
    raise NotImplementedError

  @classmethod
  def from_netrc (cls, host, netrc_filepath=None):
    auth = netrc.netrc (netrc_filepath).authenticators (host)
    if auth:
      login, _, password = auth
      return cls (login, password)
    else:
      raise NetrcNotFound ("No entry for %s in netrc" % host)

  @classmethod
  def from_cache (cls, target):
    raise NotImplementedError

Passthrough = Credentials (None, None, Credentials.PASSTHROUGH)
Anonymous = Credentials (None, None, Credentials.ANONYMOUS)
