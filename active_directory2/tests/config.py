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
  can_run_serverless = False
else:
  can_run_serverless = True

server = get_config ("general", "server") or dc
if not server:
  raise RuntimeError ("No server supplied and no DC found")

username = get_config ("general", "username")
password = get_config ("general", "password")
cred = (username, password, server)
#~ domain_dn = get_config ("general", "domain_dn")
#~ if not domain_dn:
domain_dn = win32com.client.GetObject ("LDAP://%s/rootDSE" % server).Get ("rootDomainNamingContext")
