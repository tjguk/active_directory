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
    self.assertEquals (self.ou.Get ("displayName"), guid)

  def test_delattr (self):
    ou = adbase.ADBase (self.ou)
    del ou.displayName
    with self.assertRaises (pythoncom.com_error):
      self.ou.Get ("displayName")

  @unittest.skip ("Skip until we can find a property with a dash")
  def test_underscore_to_hyphen (self):
    self.assertTrue (True)

  def test_getitem (self):
    ou = adbase.ADBase (self.ou, cred=config.cred)
    self.assertEquals (ou['user', 'cn=User01'].objectGuid, ou.GetObject ("user", "cn=User01").objectGuid)

  def test_getitem_without_qualifier (self):
    ou = adbase.ADBase (self.ou, cred=config.cred)
    self.assertEquals (ou['user', 'User01'].objectGuid, ou.GetObject ("user", "cn=User01").objectGuid)
