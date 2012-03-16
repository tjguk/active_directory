import os, sys
import filecmp
import shutil
import tempfile
import time
import unittest

try:
  from StringIO import StringIO
except ImportError:
  from io import StringIO

import active_directory

class ActiveDirectoryTestCase (unittest.TestCase):

  def assertEqualCI (self, s1, s2, *args, **kwargs):
    self.assertEqual (s1.lower (), s2.lower (), *args, **kwargs)

  def assertIs (self, item1, item2, *args, **kwargs):
    self.assertTrue (item1 is item2, *args, **kwargs)

  def assertIsInstance (self, item, klass, *args, **kwargs):
    self.assertTrue (isinstance (item, klass), *args, **kwargs)

if __name__ == '__main__':
  unittest.main ()
