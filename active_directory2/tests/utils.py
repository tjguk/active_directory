import os, sys
import ConfigParser

import win32security


def authenticated_options ():



SETUP = dict (
  test_authenticated="Yes",
  test_passthrough="No",
  test_anonymous="Yes"
)
def setup_values ():
  config = ConfigParser.ConfigParser (SETUP)
  config.read ("setup.ini")
  setup = {}
  setup['test_authenticated'] = config.getboolean ("setup", "test_authenticated")
  setup['authenticated_server'] = config.get ("setup", "authenticated_server")
  setup['authenticated_password'] = config.get ("setup", "authenticated_password")
  setup['authenticated_username'] = config.get ("setup", "authenticated_username")
  setup['test_passthrough'] = config.getboolean ("setup", "test_passthrough")
  setup['test_anonymous'] = config.getboolean ("setup", "test_anonymous")
  return setup
