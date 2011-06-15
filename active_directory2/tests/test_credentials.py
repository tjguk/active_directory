import os, sys
import subprocess
import unittest as unittest0
try:
  unittest0.skipUnless
  unittest0.skip
except AttributeError:
  import unittest2 as unittest
else:
  unittest = unittest0
del unittest0

from active_directory2 import credentials
from active_directory2.tests import base, utils

class Base (base.Base):
  pass

class TestCredentials (Base):

  def setUp (self):
    pass

  def tearDown (self):
    pass

class TestCredentialsFactory (Base):

  def test_none (self):
    self.assertIs (None, credentials.credentials (None))

  def test_credentials_object (self):
    cred = credentials.Credentials ("username", "password")
    self.assertIs (cred, credentials.credentials (cred))

  def test_credentials_invalid_tuple (self):
    with self.assertRaises (credentials.InvalidCredentialsError):
      credentials.credentials ((1, 2, 3, 4))

if __name__ == '__main__':
  unittest.main (exit=sys.stdout.isatty)
  raw_input ()