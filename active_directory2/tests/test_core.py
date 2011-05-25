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

from active_directory2 import core
from active_directory2.tests import utils

dc = utils.get_config ("general", "dc")
domain_dn = utils.get_config ("general", "domain_dn")

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

  com_object = win32com.client.CDispatch

  def test_defaults (self):
    obj = core.root_dse ()
    self.assertIsInstance (obj, self.com_object)
    self.assertEquals (obj.ADsPath, "LDAP://rootDSE")

  @unittest.skipUnless (dc, "No DC in setup.ini [general]")
  def test_server (self):
    obj = core.root_dse (server=dc)
    self.assertIsInstance (obj, self.com_object)
    self.assertEquals (obj.ADsPath, "LDAP://%s/rootDSE" % dc)

  def test_scheme (self):
    obj = core.root_dse (scheme="GC:")
    self.assertIsInstance (obj, self.com_object)
    self.assertEquals (obj.ADsPath, "GC://rootDSE")

  @unittest.skipUnless (dc, "No DC in setup.ini [general]")
  def test_server_and_scheme (self):
    obj = core.root_dse (server=dc, scheme="GC:")
    self.assertIsInstance (obj, self.com_object)
    self.assertEquals (obj.ADsPath, "GC://%s/rootDSE" % dc)

  def test_cacheing (self):
    obj1 = core.root_dse ()
    obj2 = core.root_dse ()
    self.assertIs (obj1, obj2)

class TestRootMoniker (unittest.TestCase):

  def _expected (self, server=None, scheme="LDAP:"):
    return scheme + "//" + ((server + "/") if server else "") + domain_dn

  @unittest.skipUnless (domain_dn, "No domain_dn in setup.ini [general]")
  def test_defaults (self):
    self.assertEquals (core.root_moniker (), self._expected ())

  @unittest.skipUnless (domain_dn, "No domain_dn in setup.ini [general]")
  @unittest.skipUnless (dc, "No dc in setup.ini [general]")
  def test_server (self):
    self.assertEquals (core.root_moniker (server=dc), self._expected (server=dc))

  @unittest.skipUnless (domain_dn, "No domain_dn in setup.ini [general]")
  @unittest.skipUnless (dc, "No dc in setup.ini [general]")
  def test_server_and_scheme (self):
    self.assertEquals (core.root_moniker (server=dc, scheme="GC:"), self._expected (server=dc, scheme="GC:"))

  @unittest.skipUnless (domain_dn, "No domain_dn in setup.ini [general]")
  def test_scheme (self):
    self.assertEquals (core.root_moniker (scheme="GC:"), self._expected (scheme="GC:"))

  @unittest.skipUnless (domain_dn, "No domain_dn in setup.ini [general]")
  def test_cacheing (self):
    self.assertIs (core.root_moniker (), core.root_moniker ())
