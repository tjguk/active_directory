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

ini = ConfigParser.ConfigParser()
ini.read("testing_config.ini")

def get_config(section, item, function=ConfigParser.ConfigParser.get):
    if ini.has_option(section, item):
        return function(ini, section, item)
    else:
        return None

try:
    dc = win32net.NetGetDCName().strip("\\")
except win32net.error:
    dc = None
is_inside_domain = bool(dc)

server = get_config("general", "server")
if not server and not dc:
    raise RuntimeError("No server supplied and no DC found")

username = get_config("general", "username")
password = get_config("general", "password")
if server:
    moniker = "LDAP://%s/rootDSE" % server
else:
    moniker = "LDAP://rootDSE"
domain_dn = win32com.client.GetObject(moniker).Get("rootDomainNamingContext")
test_base = get_config("general", "test_base")
if not test_base:
    raise RuntimeError("test_base must be supplied")

class ActiveDirectoryTestCase(unittest.TestCase):

    #
    # Set up(and later tear down) an OU with a few objects in it of different classes.
    # These are named <class>-<uid> where uid is a random collection of
    # ten letters & digits.
    #
    def setUp(self):
        uid = "".join([random.choice("abcdef1234567890") for i in range(10)])
        self.ou_id = "ou-" + uid
        self.user_id = "user-" + uid
        self.group_id = "group-" + uid
        self.computer_id = "computer-" + uid

        self.base_ou = active_directory.AD_object(test_base)
        self.ou = self.base_ou.Create("organizationalUnit", "ou=%s" % self.ou_id)
        self.ou.SetInfo()
        self.user = self.ou.Create("user", "cn=%s" % self.user_id)
        self.user.displayName = "£9.99" # non-ASCII
        self.user.SetInfo()
        self.group = self.ou.Create("group", "cn=%s" % self.group_id)
        self.group.SetInfo()
        self.computer = self.ou.Create("computer", "cn=%s" % self.computer_id)
        self.computer.SetInfo()
        self.dns = set(item.distinguishedName for item in(self.ou, self.user, self.group, self.computer))

    def tearDown(self):
        self.ou.Delete("group", "cn=%s" % self.group_id)
        self.ou.Delete("user", "cn=%s" % self.user_id)
        self.ou.Delete("computer", "cn=%s" % self.computer_id)
        self.base_ou.Delete("organizationalUnit", "ou=%s" % self.ou_id)

    def assertEqualCI(self, s1, s2, *args, **kwargs):
        self.assertEqual(s1.lower(), s2.lower(), *args, **kwargs)

    def assertIs(self, item1, item2, *args, **kwargs):
        self.assertTrue(item1 is item2, *args, **kwargs)

    def assertIsInstance(self, item, klass, *args, **kwargs):
        self.assertTrue(isinstance(item, klass), *args, **kwargs)

    def assertADEqual(self, item1, item2, *args, **kwargs):
        self.assertEqualCI(item1.GUID, item2.GUID, *args, **kwargs)

if is_inside_domain:
    class TestConvenienceFunctions(ActiveDirectoryTestCase):

        def test_find(self):
            self.assertADEqual(active_directory.find(self.user_id), self.user)

        def test_find_user(self):
            self.assertADEqual(active_directory.find_user(self.user_id), self.user)

        def test_find_group(self):
            self.assertADEqual(active_directory.find_group(self.group_id), self.group)

        def test_find_ou(self):
            self.assertADEqual(active_directory.find_ou(self.ou_id), self.ou)

        def test_find_computer(self):
            self.assertADEqual(active_directory.find_computer(self.computer_id), self.computer)

        def test_find(self):
            self.assertADEqual(active_directory.find(self.user_id), self.user)

        def test_search(self):
            for computer in active_directory.search("cn='%s'" % self.computer_id, objectClass="Computer"):
                self.assertADEqual(computer, self.computer)
                break
            else:
                raise RuntimeError("Computer not found")

        def test_search_ex_sql(self):
            dns = set(
                item.distinguishedName.Value \
                    for item in active_directory.search_ex("SELECT distinguishedName FROM '%s'" % self.ou.ADsPath)
            )
            self.assertEqual(dns, self.dns)

        def test_search_ex_ldap(self):
            dns = set(
                item.distinguishedName.Value \
                    for item in active_directory.search_ex("<%s>;;distinguishedName;Subtree" % self.ou.ADsPath)
            )
            self.assertEqual(dns, self.dns)

        def test_AD(self):
            self.assertEqual(active_directory.AD().distinguishedName, domain_dn)

if server:
    class TestServerBased(ActiveDirectoryTestCase):

        def setUp(self):
            ActiveDirectoryTestCase.setUp(self)
            self.base = active_directory.AD(server, username, password)

        def tearDown(self):
            self.base = None
            ActiveDirectoryTestCase.tearDown(self)

        def test_find(self):
            self.assertADEqual(self.base.find(self.user_id), self.user)

        def test_find_user(self):
            self.assertADEqual(self.base.find(self.user_id), self.user)

        def test_find_group(self):
            self.assertADEqual(self.base.find_group(self.group_id), self.group)

        def test_find_ou(self):
            self.assertADEqual(self.base.find_ou(self.ou_id), self.ou)

        def test_find_computer(self):
            self.assertADEqual(self.base.find_computer(self.computer_id), self.computer)

        def test_find(self):
            self.assertADEqual(self.base.find(self.user_id), self.user)

        def test_search(self):
            for computer in self.base.search("cn='%s'" % self.computer_id, objectClass="Computer"):
                self.assertADEqual(computer, self.computer)
                break
            else:
                raise RuntimeError("Computer not found")

if is_inside_domain:
    class TestDomainBased(ActiveDirectoryTestCase):

        def setUp(self):
            ActiveDirectoryTestCase.setUp(self)
            self.base = active_directory.AD(None, username, password)

        def tearDown(self):
            self.base = None
            ActiveDirectoryTestCase.tearDown(self)

        def test_find(self):
            self.assertADEqual(self.base.find(self.user_id), self.user)

        def test_find_user(self):
            self.assertADEqual(self.base.find(self.user_id), self.user)

        def test_find_group(self):
            self.assertADEqual(self.base.find_group(self.group_id), self.group)

        def test_find_ou(self):
            self.assertADEqual(self.base.find_ou(self.ou_id), self.ou)

        def test_find_computer(self):
            self.assertADEqual(self.base.find_computer(self.computer_id), self.computer)

        def test_find(self):
            self.assertADEqual(self.base.find(self.user_id), self.user)

        def test_search(self):
            for computer in self.base.search("cn='%s'" % self.computer_id, objectClass="Computer"):
                self.assertADEqual(computer, self.computer)
                break
            else:
                raise RuntimeError("Computer not found")


if False:
    class TestRelativePath(unittest.TestCase):

        def test_shorter_is_error(self):
            p2 = active_directory.Path ([1, 2])
            self.assertRaises(active_directory.PathTooShortError, active_directory.Path.relative_to, [1, 2], [1])

        def test_disjoint_is_error(self):
            self.assertRaises(active_directory.PathDisjointError, active_directory.Path.relative_to, [1], [2])

        def test_equal_is_empty(self):
            expected = []
            answer = relative_to([1], [1])
            self.assertEqual(answer, expected)

        def test_true_relative(self):
            expected = [1, 2]
            answer = relative_to([3, 4], [1, 2, 3, 4])
            self.assertEqual(answer, expected)


if __name__ == '__main__':
    unittest.main()
