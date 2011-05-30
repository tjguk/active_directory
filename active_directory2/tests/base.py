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

class Base (unittest.TestCase):

  def setUp (self):
    self.root = core.root_obj (server=config.server, cred=config.cred)
    self.ou = self.root.Create ("organizationalUnit", "ou=%s" % config.ou)
    self.ou.SetInfo ()

    self.ou.Create ("group", "cn=Group01").SetInfo ()
    for i in range (1, 10):
      u = self.ou.Create ("user", "cn=User%02d" % i)
      u.displayName = "User %d" % i
      u.givenName = "U%d" % i
      u.SetInfo ()

    ou2 = self.ou.Create ("organizationalUnit", "ou=test2")
    ou2.SetInfo ()
    ou2.Create ("group", "cn=Group01").SetInfo ()
    for i in range (11, 20):
      ou2.Create ("user", "cn=User%02d" % i).SetInfo ()

  def tearDown (self):
    self.ou.QueryInterface (adsi.IID_IADsDeleteOps).DeleteObject (0)
