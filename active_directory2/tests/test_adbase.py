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

class Base (base.Base):

  def setUp (self):
    base.Base.setUp (self)
    self.ou0 = self.ou
    self.ou = adbase.ADBase (self.ou)

class TestFactory (Base):

  def test_existing_adbase (self):
    self.assertIsInstance (self.ou, adbase.ADBase)
    self.assertIs (self.ou, adbase.adbase (self.ou))

  def test_python_com (self):
    self.assertIsInstance (self.ou, adbase.ADBase)
    self.assertEquals (self.ou0.ADsPath, self.ou.com_object.ADsPath)

  def test_moniker (self):
    self.assertIsInstance (self.ou, adbase.ADBase)
    self.assertEquals (self.ou0.ADsPath, self.ou.ADsPath)

class TestADBase (Base):

  def test_init (self):
    self.assertEquals (self.ou0.ADsPath, self.ou.com_object.ADsPath)
    self.assertEquals (self.ou0.ADsPath, self.ou.path)
    self.assertEquals (self.ou0.Class, self.ou.cls)
    self.assertEquals (self.ou.schema.Class, "Class")

  def test_getattr (self):
    self.assertEquals (self.ou.distinguishedName, self.ou0.Get ("distinguishedName"))

  def test_setattr (self):
    guid = str (uuid.uuid1 ())
    self.ou.displayName = guid
    self.assertEquals (self.ou0.displayName, guid)

  def test_setattr_to_none (self):
    self.ou.displayName = None
    self.assertEquals (self.ou0.displayName, None)

  #~ @unittest.skip ("Skip until we can find a property with a dash")
  def test_underscore_to_hyphen (self):
    self.assertEquals ("abc", adbase.ADBase._munged_attribute ("abc"))
    self.assertEquals ("abc", adbase.ADBase._munged_attribute ("abc_"))
    self.assertEquals ("abc-def", adbase.ADBase._munged_attribute ("abc_def"))

  def test_getitem (self):
    user01 = self.ou['cn=User01'].objectGuid
    self.assertEquals (user01, self.ou.GetObject ("user", "cn=User01").objectGuid)

  def test_setitem (self):
    self.ou['cn=User99'] = {"Class" : "user"}
    self.assertEquals (self.ou['cn=User99'].objectGuid, self.ou.GetObject ("user", "cn=User99").objectGuid)

  def test_delitem (self):
    del self.ou['cn=User01']
    self.assertNotIn (("user", "CN=User01"), [(i.Class, i.Name) for i in self.ou0])

  def test_equality (self):
    ou2 = adbase.ADBase (self.ou0)
    self.assertEquals (self.ou, ou2)

  def test_identity (self):
    ou2 = adbase.ADBase (self.ou0)
    self.assertEquals (len (set ([self.ou, ou2])), 1)

  def test_from_path (self):
    ou2 = adbase.ADBase.from_path (self.ou0.ADsPath, cred=config.cred)
    self.assertEquals (self.ou, ou2)

  def test_dump (self):
    #
    # sanity test only
    #
    with tempfile.TemporaryFile () as ofile:
      self.ou.dump (ofile=ofile)
      ofile.seek (0)
      data = ofile.read ()
      self.assertIn ("ADsPath => %r" % self.ou0.ADsPath.encode ("ascii"), data)

  def test_set (self):
    _, username, _ = base.find_pattern (type_pattern="user")
    user1 = self.ou[username]
    x = str (uuid.uuid1 ())
    self.assertNotEquals (user1.displayName, x)
    self.assertNotEquals (user1.givenName, x)
    user1.set (displayName=x, givenName=x)
    self.assertEquals (user1.displayName, x)
    self.assertEquals (user1.givenName, x)

  def test_get (self):
    #
    # There doesn't seem to be a reliable way to test this.
    # At the very least, do a sanity check to make sure
    # it's not failing.
    #
    user1 = self.ou.find (objectCategory="person")
    new_name = str (uuid.uuid1 ())
    user1.displayName = new_name
    self.assertEquals (new_name, user1.get ("displayName"))

  def test_delete (self):
    ou = self.ou.find ("!distinguishedName=%s" % self.ou.distinguishedName, objectCategory="organizationalUnit")
    self.assertIsNot (ou, None)
    dn = ou.distinguishedName
    ou.delete ()
    self.assertIs (self.ou.find (distinguishedName=dn), None)

  def test_move (self):
    ous = list (self.ou.search (
      "!distinguishedName=%s" % self.ou.distinguishedName,
      objectCategory="organizationalUnit"
    ))
    ou1, ou2 = ous[:2]
    u1 = ou1.find (objectCategory="person")
    u1_guid = u1.objectGuid
    ou1.move (u1.Name, ou2)
    u2 = ou2.find (cn=u1.cn)
    self.assertEquals (u1_guid, u2.objectGuid)

  def test_rename (self):
    ou = self.ou.find (
    "!distinguishedName=%s" % self.ou.distinguishedName,
      objectCategory="organizationalUnit"
    )
    self.assertIsNot (ou, None)
    name = str (uuid.uuid1 ())
    u1 = ou.find (objectCategory="person")
    u1_guid = u1.objectGuid
    ou.rename (u1.Name, "cn=%s" % name)
    u2 = ou.find (cn=name)
    self.assertEquals (u1_guid, u2.objectGuid)

class TestSearch (Base):

  def setUp (self):
    Base.setUp (self)
    self.type, self.name, self.path = base.find_pattern ()

  def test_no_filter (self):
    with self.assertRaises (adbase.NoFilterError):
      self.ou.search ().next ()

  def test_args_only (self):
    searcher = self.ou.search (self.name)
    self.assertIn (self.name, [i.Name for i in searcher])

  def test_kwargs_only (self):
    searcher = self.ou.search (objectCategory=self.type)
    self.assertIn (self.name, [i.Name for i in searcher])

  def test_args_and_kwargs (self):
    searcher = self.ou.search (self.name, objectCategory=self.type)
    self.assertEquals ([self.name], [i.name for i in searcher])

  def test_find_no_filter (self):
    with self.assertRaises (adbase.NoFilterError):
      self.ou.find ().next ()

  def test_find_args_only (self):
    self.assertEquals (self.name, self.ou.find (self.name).Name)

  def test_find_kwargs_only (self):
    self.assertEquals (self.name, self.ou.find (objectCategory=self.type).Name)

  def test_find_args_and_kwargs (self):
    self.assertEquals (self.name, self.ou.find (self.name, objectCategory=self.type).Name)

  def test_find_user (self):
    _, username, _ = base.find_pattern (type_pattern="user")
    qualifier, _, name = username.partition ("=")
    user = self.ou.find_user (name)
    self.assertTrue (user)
    self.assertEquals (username, user.Name)

if __name__ == '__main__':
  unittest.main (exit=sys.stdout.isatty)
  raw_input ()