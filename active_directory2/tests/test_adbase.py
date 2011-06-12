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
import tempfile
import uuid

import pythoncom
from win32com import adsi

from active_directory2 import core, adbase
from active_directory2.tests import base, utils
from active_directory2.tests import config

class TestFactory (base.Base):

  def test_existing_adbase (self):
    ou = adbase.ADBase (self.ou)
    self.assertIsInstance (ou, adbase.ADBase)
    self.assertIs (ou, adbase.adbase (ou))

  def test_python_com (self):
    ou = adbase.adbase (self.ou)
    self.assertIsInstance (ou, adbase.ADBase)
    self.assertEquals (self.ou.ADsPath, ou.com_object.ADsPath)

  def test_moniker (self):
    ou = adbase.adbase (self.ou.ADsPath, cred=config.cred)
    self.assertIsInstance (ou, adbase.ADBase)
    self.assertEquals (self.ou.ADsPath, ou.ADsPath)

class TestADBase (base.Base):

  def test_init (self):
    ou = adbase.ADBase (self.ou)
    self.assertEquals (self.ou.ADsPath, ou.com_object.ADsPath)
    self.assertEquals (self.ou.ADsPath, ou.path)
    self.assertEquals (self.ou.Class, ou.cls)
    self.assertEquals (ou.schema.Class, "Class")

  def test_getattr (self):
    ou = adbase.ADBase (self.ou)
    self.assertEquals (ou.distinguishedName, self.ou.Get ("distinguishedName"))

  def test_setattr (self):
    ou = adbase.ADBase (self.ou)
    guid = str (uuid.uuid1 ())
    ou.displayName = guid
    self.assertEquals (self.ou.displayName, guid)

  def test_setattr_to_none (self):
    ou = adbase.ADBase (self.ou)
    ou.displayName = None
    self.assertEquals (self.ou.displayName, None)

  @unittest.skip ("Skip until we can find a property with a dash")
  def test_underscore_to_hyphen (self):
    self.assertTrue (True)

  def test_item_identifier (self):
    cases = {
      "user" : "cn",
      "group" : "cn",
      "organizationalUnit" : "ou"
    }
    ou = adbase.adbase (self.ou, config.cred)
    for cls, ident in cases.items ():
      self.assertEquals ("%s=XX" % (ident), ou._item_identifier (cls, "XX"))

  def test_getitem (self):
    ou = adbase.ADBase (self.ou, cred=config.cred)
    user01 = ou['user', 'cn=User01'].objectGuid
    self.assertEquals (user01, ou.GetObject ("user", "cn=User01").objectGuid)

  def test_getitem_without_qualifier (self):
    ou = adbase.ADBase (self.ou, cred=config.cred)
    self.assertEquals (ou['user', 'User01'].objectGuid, ou.GetObject ("user", "cn=User01").objectGuid)

  def test_setitem (self):
    ou = adbase.ADBase (self.ou, cred=config.cred)
    ou['user', 'cn=User99'] = {}
    self.assertEquals (ou['user', 'cn=User99'].objectGuid, ou.GetObject ("user", "cn=User99").objectGuid)

  def test_setitem_without_qualifier (self):
    ou = adbase.ADBase (self.ou, cred=config.cred)
    ou['user', 'User98'] = {}
    self.assertEquals (ou['user', 'cn=User98'].objectGuid, ou.GetObject ("user", "cn=User98").objectGuid)

  def test_delitem (self):
    ou = adbase.ADBase (self.ou, cred=config.cred)
    del ou['user', 'cn=User01']
    self.assertNotIn (("user", "CN=User01"), [(i.Class, i.Name) for i in self.ou])

  def test_delitem_without_qualifier (self):
    ou = adbase.ADBase (self.ou, cred=config.cred)
    del ou['user', 'User01']
    self.assertNotIn (("user", "CN=User01"), [(i.Class, i.Name) for i in self.ou])

  def test_equality (self):
    ou1 = adbase.ADBase (self.ou, cred=config.cred)
    ou2 = adbase.ADBase (self.ou, cred=config.cred)
    self.assertEquals (ou1, ou2)

  def test_identity (self):
    ou1 = adbase.ADBase (self.ou, cred=config.cred)
    ou2 = adbase.ADBase (self.ou, cred=config.cred)
    self.assertEquals (len (set ([ou1, ou2])), 1)

  def test_from_path (self):
    ou1 = adbase.adbase (self.ou, cred=config.cred)
    ou2 = adbase.ADBase.from_path (self.ou.ADsPath, cred=config.cred)
    self.assertEquals (ou1, ou2)

  def test_dump (self):
    #
    # sanity test only
    #
    with tempfile.TemporaryFile () as ofile:
      adbase.adbase (self.ou, cred=config.cred).dump (ofile=ofile)
      ofile.seek (0)
      data = ofile.read ()
      self.assertIn ("ADsPath => %s" % self.ou.ADsPath.encode ("ascii"), data)

  def test_set (self):
    user1 = adbase.adbase (self.ou, cred=config.cred)['user', 'User01']
    x = str (uuid.uuid1 ())
    self.assertNotEquals (user1.displayName, x)
    self.assertNotEquals (user1.givenName, x)
    user1.set (displayName=x, givenName=x)
    self.assertEquals (user1.displayName, x)
    self.assertEquals (user1.givenName, x)

  def test_delete (self):
    ou = adbase.adbase (self.ou, cred=config.cred)
