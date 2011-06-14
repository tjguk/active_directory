import os, sys
import fnmatch
import unittest as unittest0
try:
  unittest0.skipUnless
  unittest0.skip
except AttributeError:
  import unittest2 as unittest
else:
  unittest = unittest0
del unittest0

from win32com import adsi

from active_directory2 import core, adbase
from active_directory2.tests import utils
from active_directory2.tests import config

DATA = [
  ("organizationalUnit", "OU=adtest", [
    ("group", "CN=Group01", []),
    ("user", "CN=user01", []),
    ("user", "CN=user02", []),
    ("organizationalUnit", "OU=test2", [
      ("group", "CN=Group01", []),
      ("user", "CN=user02", []),
      ("user", "CN=user03", []),
    ]),
    ("organizationalUnit", "OU=test3", []),
  ])
]

def build_from_data (root, data):
  for type, name, subdata in data:
    newroot = root.Create (type, name)
    newroot.SetInfo ()
    if subdata:
      build_from_data (newroot, subdata)
  return newroot

def find_pattern (type_pattern="*", name_pattern="*"):
  def _find_pattern (type_pattern="*", name_pattern="*", data=DATA, path=None):
    if path is None:
      path = []
    for type, rdn, subdata in data:
      if fnmatch.fnmatch (type, type_pattern) and fnmatch.fnmatch (rdn, name_pattern):
        return type, rdn, ",".join (reversed (path + [rdn]))
      else:
        result = _find_pattern (type_pattern, name_pattern, subdata, path + [rdn])
        if result:
          return result

  result = _find_pattern (type_pattern, name_pattern)
  if result is None:
    raise RuntimeError ("Couldn't match %s and %s" % (type_pattern, name_pattern))
  else:
    return result

class Base (unittest.TestCase):

  def setUp (self):
    self.root = core.root_obj (server=config.server, cred=config.cred)
    self._ou = self.root.GetObject ("organizationalUnit", config.test_base)
    self.ou = build_from_data (self._ou, DATA)
    self.addCleanup (self._remove_ou)

  def _remove_ou (self):
    self.ou.QueryInterface (adsi.IID_IADsDeleteOps).DeleteObject (0)
