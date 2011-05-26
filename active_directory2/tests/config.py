import os, sys
import ConfigParser

import win32com.client

ini = ConfigParser.ConfigParser ()
ini.read ("config.ini")

def get_config (section, item, function=ConfigParser.ConfigParser.get):
  if ini.has_option (section, item):
    return function (ini, section, item)
  else:
    return None

test_serverless = get_config ("setup", "test_serverless", ConfigParser.ConfigParser.getboolean)
server = get_config ("general", "server")
if not server:
  try:
    server = win32net.NetGetAnyDCName ()
  except win32net.error:
    raise RuntimeError ("No server supplied and no DC found")
username = get_config ("general", "username")
password = get_config ("general", "password")
cred = (username, password, server)
domain_dn = get_config ("general", "domain_dn")
if not domain_dn:
  domain_dn = win32com.client.GetObject ("LDAP://%s/rootDSE" % server).Get ("rootDomainNamingContext")
