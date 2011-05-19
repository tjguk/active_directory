# -*- coding: iso-8859-1 -*-
from ctypes.wintypes import *
from ctypes import windll, wintypes
kernel32 = windll.kernel32
import ctypes
import datetime
import re
import struct

import win32api

def delta_as_microseconds (delta) :
  return delta.days * 24* 3600 * 10**6 + delta.seconds * 10**6 + delta.microseconds

def signed_to_unsigned (signed):
  u"""Convert a (possibly signed) long to unsigned hex"""
  unsigned, = struct.unpack ("L", struct.pack ("l", signed))
  return unsigned

def _set (obj, attribute, value):
  u"""Helper function to add an attribute directly into the instance
   dictionary, bypassing possible __getattr__ calls
  """
  obj.__dict__[attribute] = value

#
# Code contributed by Stian Søiland <stian@soiland.no>
#
def i32(x):
  u"""Converts a long (for instance 0x80005000L) to a signed 32-bit-int.

  Python2.4 will convert numbers >= 0x80005000 to large numbers
  instead of negative ints.    This is not what we want for
  typical win32 constants.

  Usage:
      >>> i32(0x80005000L)
      -2147363168
  """
  # x > 0x80000000L should be negative, such that:
  # i32(0x80000000L) -> -2147483648L
  # i32(0x80000001L) -> -2147483647L     etc.
  return (x&0x80000000L and -2*0x40000000 or 0) + int(x&0x7fffffff)

class FILETIME (ctypes.Structure):
  _fields_ = [
    ("dwLowDateTime", DWORD),
    ("dwHighDateTime", DWORD),
  ]

class SYSTEMTIME (ctypes.Structure):
  _fields_ = [
    ("wYear", WORD),
    ("wMonth", WORD),
    ("wDayOfWeek", WORD),
    ("wDay", WORD),
    ("wHour", WORD),
    ("wMinute", WORD),
    ("wSecond", WORD),
    ("wMilliseconds", WORD),
  ]

def error (exception, context="", message=""):
  errno = win32api.GetLastError ()
  message = message or win32api.FormatMessageW (errno)
  raise exception (errno, context, message)

def file_time_to_system_time (ularge):
  filetime = FILETIME (ularge.LowPart, ularge.HighPart)
  systemtime = SYSTEMTIME ()
  if kernel32.FileTimeToSystemTime (ctypes.pointer (filetime), ctypes.byref (systemtime)) == 0:
    error (WindowsError)
  return datetime.datetime (
    systemtime.wYear, systemtime.wMonth, systemtime.wDay,
    systemtime.wHour, systemtime.wMinute, systemtime.wSecond,
    systemtime.wMilliseconds * 1000
  )

def escaped_moniker (moniker):
  #
  # If the moniker *appears* to have been escaped
  # already, return it straight. This is obviously
  # fragile but seems to work for now.
  #
  if moniker.find (u"\\/") > -1:
    return moniker
  else:
    return moniker.replace (u"/", u"\\/")

def parse_moniker (moniker):
  scheme, server, dn = re.match ("([^:]+://)([A-za-z0-9-_]+/)?(.*)", moniker).groups ()
  if scheme is None:
    scheme = u"LDAP://"
  if scheme != u"WinNT:":
    dn = escaped_moniker (dn)
  return scheme or u"", (server.rstrip (u"/") + u"/") if server else u"", dn or u""
