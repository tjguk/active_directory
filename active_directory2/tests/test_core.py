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

import win32com.client

from active_directory2 import core, credentials
from active_directory2.tests import config

com_object = win32com.client.CDispatch

class TestBaseMoniker (unittest.TestCase):

  def test_defaults (self):
    self.assertEquals (core._base_moniker (), "LDAP://")

  def test_server (self):
    self.assertEquals (core._base_moniker (server="server"), "LDAP://server/")

  def test_scheme (self):
    self.assertEquals (core._base_moniker (scheme="GC:"), "GC://")

  def test_server_and_scheme (self):
    self.assertEquals (core._base_moniker (server="server", scheme="GC:"), "GC://server/")

  def test_cacheing (self):
    #
    # Use a ridiculously long server name to defeat Python string interning
    #
    server = "s" * 1024
    m1 = core._base_moniker (server=server)
    m2 = core._base_moniker (server=server)
    self.assertIs (m1, m2)

class TestRootDSE (unittest.TestCase):

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_defaults (self):
    obj = core.root_dse ()
    self.assertIsInstance (obj, com_object)
    self.assertEquals (obj.ADsPath, "LDAP://rootDSE")

  def test_server (self):
    obj = core.root_dse (server=config.server)
    self.assertIsInstance (obj, com_object)
    self.assertEquals (obj.ADsPath, "LDAP://%s/rootDSE" % config.server)

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_scheme (self):
    obj = core.root_dse (scheme="GC:")
    self.assertIsInstance (obj, com_object)
    self.assertEquals (obj.ADsPath, "GC://rootDSE")

  def test_server_and_scheme (self):
    obj = core.root_dse (server=config.server, scheme="GC:")
    self.assertIsInstance (obj, com_object)
    self.assertEquals (obj.ADsPath, "GC://%s/rootDSE" % config.server)

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_cacheing (self):
    obj1 = core.root_dse ()
    obj2 = core.root_dse ()
    self.assertIs (obj1, obj2)

class TestRootMoniker (unittest.TestCase):

  def _expected (self, server=None, scheme="LDAP:"):
    return scheme + "//" + ((server + "/") if server else "") + config.domain_dn

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_defaults (self):
    self.assertEquals (core.root_moniker (), self._expected ())

  def test_server (self):
    self.assertEquals (core.root_moniker (server=config.server), self._expected (server=config.server))

  def test_server_and_scheme (self):
    self.assertEquals (core.root_moniker (server=config.server, scheme="GC:"), self._expected (server=config.server, scheme="GC:"))

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_scheme (self):
    self.assertEquals (core.root_moniker (scheme="GC:"), self._expected (scheme="GC:"))

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_cacheing (self):
    self.assertIs (core.root_moniker (), core.root_moniker ())

class TestRootObj (unittest.TestCase):

  def _test (self, *args, **kwargs):
    print "About to _test with", args, kwargs
    root_obj = core.root_obj (cred=config.cred, *args, **kwargs)
    self.assertIsInstance (root_obj, com_object)
    self.assertEquals (root_obj.ADsPath, core.root_moniker (*args, **kwargs))

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_defaults (self):
    self._test ()

  def test_server (self):
    self._test (server=config.server)

  def test_server_and_scheme (self):
    self._test (server=config.server, scheme="GC:")

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_scheme (self):
    self._test (scheme="GC:")

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_cacheing (self):
    self.assertIs (core.root_obj (cred=config.cred), core.root_obj (cred=config.cred))

class TestSchemaObj (unittest.TestCase):

  def _expected (self, server=None):
    return "LDAP://" + ((server + "/") if server else "") + "CN=Schema,CN=Configuration," + config.domain_dn

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_defaults (self):
    schema_obj = core.schema_obj (cred=config.cred)
    self.assertIsInstance (schema_obj, com_object)
    self.assertEquals (schema_obj.ADsPath, self._expected ())

  def test_server (self):
    schema_obj = core.schema_obj (config.server, cred=config.cred)
    self.assertIsInstance (schema_obj, com_object)
    self.assertEquals (schema_obj.ADsPath, self._expected (config.server))

class TestClassSchema (unittest.TestCase):

  def _expected (self, class_name, server=None):
    #
    # The abstract schema is a special, serverless object
    #
    return "LDAP://" + class_name + ",schema"

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_class_with_defaults (self):
    class_schema = core.class_schema ("user")
    self.assertIsInstance (class_schema, com_object)
    self.assertEquals (class_schema.ADsPath, self._expected ("user"))

  def test_class_with_server (self):
    class_schema = core.class_schema ("user", server=config.server, cred=config.cred)
    self.assertIsInstance (class_schema, com_object)
    self.assertEquals (class_schema.ADsPath, self._expected ("user", server=config.server))

class TestAttributes (unittest.TestCase):

  def setUp (self):
    self.all_attributes = set (i.ldapDisplayName for i in core.schema_obj () if i.Class == "attributeSchema")

  @unittest.skipUnless (config.can_run_serverless, "Serverless testing not enabled")
  def test_defaults (self):
    attributes = core.attributes ()
    self.assertSetEqual (self.all_attributes, set (name for name, _ in attributes))
    self.assertTrue (all (i.Class == "attributeSchema" for i in attributes))

  def test_server (self):
    attributes = core.attributes (server=config.server)
    self.assertSetEqual (self.all_attributes, set (name for name, _ in attributes))
    self.assertTrue (all (i.Class == "attributeSchema" for i in attributes))
