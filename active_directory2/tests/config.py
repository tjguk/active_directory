import os, sys
import ConfigParser

import win32com.client
import win32net

ini = ConfigParser.ConfigParser ()
ini.read ("config.ini")

def get_config (section, item, function=ConfigParser.ConfigParser.get):
  if ini.has_option (section, item):
    return function (ini, section, item)
  else:
    return None

try:
  dc = win32net.NetGetAnyDCName ().strip ("\\")
except win32net.error:
  dc = None
is_inside_domain = bool (dc)

server = get_config ("general", "server")
if not server and not dc:
  raise RuntimeError ("No server supplied and no DC found")

username = get_config ("general", "username")
password = get_config ("general", "password")
cred = (username, password, server)
if server:
  domain_dn = win32com.client.GetObject ("LDAP://%s/rootDSE" % server).Get ("rootDomainNamingContext")
else:
  domain_dn = win32com.client.GetObject ("LDAP://rootDSE").Get ("rootDomainNamingContext")
test_base = get_config ("general", "test_base")
if not test_base:
  raise RuntimeError ("test_base must be supplied")
