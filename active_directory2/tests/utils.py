import os, sys
import ConfigParser

import win32security

config = ConfigParser.ConfigParser ()
config.read ("setup.ini")

def get_config (section, item):
  if config.has_option (section, item):
    return config.get (section, item)
  else:
    return None
