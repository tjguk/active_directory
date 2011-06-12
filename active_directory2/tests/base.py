import os, sys
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
    ("user", "CN=user03", []),
    ("organizationalUnit", "OU=test2", [
      ("group", "CN=Group01", []),
      ("user", "CN=user01", []),
      ("user", "CN=user02", []),
      ("user", "CN=user03", []),
    ])
  ])
]

def build_from_data (root, data):
  for type, name, subdata in data:
    newroot = root.Create (type, name)
    newroot.SetInfo ()
    if subdata:
      build_from_data (newroot, subdata)
  return newroot

class Base (unittest.TestCase):

  def setUp (self):
    self.root = core.root_obj (server=config.server, cred=config.cred)
    self._ou = self.root.GetObject ("organizationalUnit", config.test_base)
    self.ou = build_from_data (self._ou, DATA)
    self.addCleanup (self._remove_ou)

    #~ self.ou = self._ou.Create ("organizationalUnit", "ou=adtest")
    #~ self.ou.SetInfo ()
    #~ self.addCleanup (self._remove_ou)

    #~ self.ou.Create ("group", "cn=Group01").SetInfo ()
    #~ for i in range (1, 4):
      #~ u = self.ou.Create ("user", "cn=User%02d" % i)
      #~ u.displayName = "User %d" % i
      #~ u.givenName = "U%d" % i
      #~ u.SetInfo ()

    #~ ou2 = self.ou.Create ("organizationalUnit", "ou=test2")
    #~ ou2.SetInfo ()
    #~ ou2.Create ("group", "cn=Group01").SetInfo ()
    #~ for i in range (1, 4):
      #~ ou2.Create ("user", "cn=User%02d" % i).SetInfo ()

  def _remove_ou (self):
    self.ou.QueryInterface (adsi.IID_IADsDeleteOps).DeleteObject (0)
