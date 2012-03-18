# -*- coding: UTF-8 -*-
import os, sys
try:
  import ConfigParser
except ImportError:
  import configparser as ConfigParser
import filecmp
import random
import shutil
import tempfile
import time
import unittest
try:
  from StringIO import StringIO
except ImportError:
  from io import StringIO

import win32com.client
import win32net

import active_directory

ini = ConfigParser.ConfigParser ()
ini.read ("testing_config.ini")

def get_config (section, item, function=ConfigParser.ConfigParser.get):
  if ini.has_option (section, item):
    return function (ini, section, item)
  else:
    return None

try:
  dc = win32net.NetGetDCName ().strip ("\\")
except win32net.error:
  dc = None
is_inside_domain = bool (dc)

server = get_config ("general", "server")
if not server and not dc:
  raise RuntimeError ("No server supplied and no DC found")

username = get_config ("general", "username")
password = get_config ("general", "password")
domain_dn = win32com.client.GetObject ("LDAP://rootDSE").Get ("rootDomainNamingContext")
test_base = get_config ("general", "test_base")
if not test_base:
  raise RuntimeError ("test_base must be supplied")

class ActiveDirectoryTestCase (unittest.TestCase):

  uid = "".join ([random.choice ("abcdef1234567890") for i in range (10)])
  ou_id = "ou-" + uid
  user_id = "user-" + uid
  group_id = "group-" + uid
  computer_id = "computer-" + uid

  #
  # Set up (and later tear down) an OU with a single user in it.
  # Both are named __class__.uid which is a random collection of
  # ten letters & digits.
  #
  def setUp (self):
    self.base_ou = active_directory.AD_object (test_base)
    self.ou = self.base_ou.Create ("organizationalUnit", "ou=%s" % self.ou_id)
    self.ou.SetInfo ()
    self.user = self.ou.Create ("user", "cn=%s" % self.user_id)
    self.user.displayName = "£9.99"
    self.user.SetInfo ()
    self.group = self.ou.Create ("group", "cn=%s" % self.group_id)
    self.group.SetInfo ()
    self.computer= self.ou.Create ("computer", "cn=%s" % self.computer_id)
    self.computer.SetInfo ()

  def tearDown (self):
    self.ou.Delete ("group", "cn=%s" % self.group_id)
    self.ou.Delete ("user", "cn=%s" % self.user_id)
    self.ou.Delete ("computer", "cn=%s" % self.computer_id)
    self.base_ou.Delete ("organizationalUnit", "ou=%s"  % self.ou_id)

  def assertEqualCI (self, s1, s2, *args, **kwargs):
    self.assertEqual (s1.lower (), s2.lower (), *args, **kwargs)

  def assertIs (self, item1, item2, *args, **kwargs):
    self.assertTrue (item1 is item2, *args, **kwargs)

  def assertIsInstance (self, item, klass, *args, **kwargs):
    self.assertTrue (isinstance (item, klass), *args, **kwargs)

  def assertADEqual (self, item1, item2, *args, **kwargs):
    self.assertEqual (item1.GUID, item2.GUID, "%s is not the same as %s" % (item1.ADsPath, item2.ADsPath))

class TestConvenienceFunctions (ActiveDirectoryTestCase):

  def test_find (self):
    self.assertADEqual (active_directory.find (self.user_id), self.user)

  def test_find_user (self):
    self.assertADEqual (active_directory.find_user (self.user_id), self.user)

  def test_find_group (self):
    self.assertADEqual (active_directory.find_group (self.group_id), self.group)

  def test_find_ou (self):
    self.assertADEqual (active_directory.find_ou (self.ou_id), self.ou)

  def test_find_computer (self):
    self.assertADEqual (active_directory.find_computer (self.computer_id), self.computer)

if __name__ == '__main__':
  unittest.main ()
