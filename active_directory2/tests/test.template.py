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

from active_directory2 import core
from active_directory2.tests import base, config, utils

class Base (base.base):
  pass

class TestTemplate (Base):

  def setUp (self):
    pass

  def tearDown (self):
    pass

if __name__ == '__main__':
  unittest.main (exit=sys.stdout.isatty)
  raw_input ()