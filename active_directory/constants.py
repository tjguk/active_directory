# -*- coding: iso-8859-1 -*-
from . import utils

#
# For ease of presentation, ms-style constant lists are
# held as Enum objects, allowing access by number or
# by name, and by name-as-attribute. This means you can do, eg:
#
# print GROUP_TYPES[2]
# print GROUP_TYPES['GLOBAL']
# print GROUP_TYPES.GLOBAL
#
# The first is useful when displaying the contents
# of an AD object; the other two when you want a more
# readable piece of code, without magic numbers.
#
class Enum (object):

  def __init__ (self, **kwargs):
    self._name_map = {}
    self._number_map = {}
    for k, v in kwargs.items ():
      self._name_map[k] = utils.i32 (v)
      self._number_map[utils.i32 (v)] = k

  def __getitem__ (self, item):
    try:
      return self._name_map[item]
    except KeyError:
      return self._number_map[utils.i32 (item)]

  def __getattr__ (self, attr):
    try:
      return self._name_map[attr]
    except KeyError:
      raise AttributeError

  def __repr__ (self):
    return repr (self._name_map)

  def __str__ (self):
    return str (self._name_map)

  def item_names (self):
    return self._name_map.items ()

  def item_numbers (self):
    return self._number_map.items ()

ADS_SYSTEMFLAG = Enum (
  DISALLOW_DELETE             = 0x80000000,
  CONFIG_ALLOW_RENAME         = 0x40000000,
  CONFIG_ALLOW_MOVE           = 0x20000000,
  CONFIG_ALLOW_LIMITED_MOVE   = 0x10000000,
  DOMAIN_DISALLOW_RENAME      = 0x08000000,
  DOMAIN_DISALLOW_MOVE        = 0x04000000,
  CR_NTDS_NC                  = 0x00000001,
  CR_NTDS_DOMAIN              = 0x00000002,
  ATTR_NOT_REPLICATED         = 0x00000001,
  ATTR_IS_CONSTRUCTED         = 0x00000004
)

GROUP_TYPES = Enum (
  GLOBAL = 0x00000002,
  DOMAIN_LOCAL = 0x00000004,
  LOCAL = 0x00000004,
  UNIVERSAL = 0x00000008,
  SECURITY_ENABLED = 0x80000000
)

AUTHENTICATION_TYPES = Enum (
  SECURE_AUTHENTICATION = utils.i32 (0x01),
  USE_ENCRYPTION = utils.i32 (0x02),
  USE_SSL = utils.i32 (0x02),
  READONLY_SERVER = utils.i32 (0x04),
  PROMPT_CREDENTIALS = utils.i32 (0x08),
  NO_AUTHENTICATION = utils.i32 (0x10),
  FAST_BIND = utils.i32 (0x20),
  USE_SIGNING = utils.i32 (0x40),
  USE_SEALING = utils.i32 (0x80),
  USE_DELEGATION = utils.i32 (0x100),
  SERVER_BIND = utils.i32 (0x200),
  AUTH_RESERVED = utils.i32 (0x800000000)
)

SAM_ACCOUNT_TYPES = Enum (
  DOMAIN_OBJECT = 0x0 ,
  GROUP_OBJECT = 0x10000000 ,
  NON_SECURITY_GROUP_OBJECT = 0x10000001 ,
  ALIAS_OBJECT = 0x20000000 ,
  NON_SECURITY_ALIAS_OBJECT = 0x20000001 ,
  USER_OBJECT = 0x30000000 ,
  NORMAL_USER_ACCOUNT = 0x30000000 ,
  MACHINE_ACCOUNT = 0x30000001 ,
  TRUST_ACCOUNT = 0x30000002 ,
  APP_BASIC_GROUP = 0x40000000,
  APP_QUERY_GROUP = 0x40000001 ,
  ACCOUNT_TYPE_MAX = 0x7fffffff
)

USER_ACCOUNT_CONTROL = Enum (
  SCRIPT = 0x00000001,
  ACCOUNTDISABLE = 0x00000002,
  HOMEDIR_REQUIRED = 0x00000008,
  LOCKOUT = 0x00000010,
  PASSWD_NOTREQD = 0x00000020,
  PASSWD_CANT_CHANGE = 0x00000040,
  ENCRYPTED_TEXT_PASSWORD_ALLOWED = 0x00000080,
  TEMP_DUPLICATE_ACCOUNT = 0x00000100,
  NORMAL_ACCOUNT = 0x00000200,
  INTERDOMAIN_TRUST_ACCOUNT = 0x00000800,
  WORKSTATION_TRUST_ACCOUNT = 0x00001000,
  SERVER_TRUST_ACCOUNT = 0x00002000,
  DONT_EXPIRE_PASSWD = 0x00010000,
  MNS_LOGON_ACCOUNT = 0x00020000,
  SMARTCARD_REQUIRED = 0x00040000,
  TRUSTED_FOR_DELEGATION = 0x00080000,
  NOT_DELEGATED = 0x00100000,
  USE_DES_KEY_ONLY = 0x00200000,
  DONT_REQUIRE_PREAUTH = 0x00400000,
  PASSWORD_EXPIRED = 0x00800000,
  TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 0x01000000
)

ADS_PROPERTY = Enum (
  CLEAR = 1,
  UPDATE = 2,
  APPEND = 3,
  DELETE = 4
)
