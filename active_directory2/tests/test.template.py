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

from active_directory2 import core
from active_directory2.tests import utils
from active_directory2.tests import config

class TestTemplate (unittest.TestCase):

  def setUp (self):
    pass

  def tearDown (self):
    pass

