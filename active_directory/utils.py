# -*- coding: iso-8859-1 -*-
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

