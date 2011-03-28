# -*- coding: iso-8859-1 -*-
import pywintypes

from . import utils

class ActiveDirectoryError (Exception):
  u"""Base class for all AD Exceptions"""
  pass

class MemberAlreadyInGroupError (ActiveDirectoryError):
  pass

class MemberNotInGroupError (ActiveDirectoryError):
  pass

class BadPathnameError (ActiveDirectoryError):
  pass

class AttributeNotFound (ActiveDirectoryError):
  pass

def wrapper (winerror_map, default_exception):
  u"""Used by each module to map specific windows error codes onto
  Python exceptions. Always includes a default which is raised if
  no specific exception is found.
  """
  def _wrapped (function, *args, **kwargs):
    u"""Call a Windows API with parameters, and handle any
    exception raised either by mapping it to a module-specific
    one or by passing it back up the chain.
    """
    try:
      return function (*args, **kwargs)
    except pywintypes.com_error, (hresult_code, hresult_name, additional_info, parameter_in_error):
      hresult_code = utils.signed_to_unsigned (hresult_code)
      exception_string = [u"%08X - %s" % (hresult_code, hresult_name)]
      if additional_info:
        wcode, source_of_error, error_description, whlp_file, whlp_context, scode = additional_info
        scode = utils.signed_to_unsigned (scode)
        exception_string.append (u"  Error in: %s" % source_of_error)
        exception_string.append (u"  %08X - %s" % (scode, (error_description or "").strip ()))
      else:
        scode = None
      exception = winerror_map.get (hresult_code, winerror_map.get (scode, default_exception))
      raise exception (hresult_code, hresult_name, u"\n".join (exception_string))
    except pywintypes.error, (errno, errctx, errmsg):
      exception = winerror_map.get (errno, default_exception)
      raise exception (errno, errctx, errmsg)
    except (WindowsError, IOError), err:
      exception = winerror_map.get (err.errno, default_exception)
      if exception:
        raise exception (err.errno, u"", err.strerror)
  return _wrapped

ERROR_DS_NO_SUCH_OBJECT = 0x80072030
ERROR_OBJECT_ALREADY_EXISTS = 0x80071392
ERROR_MEMBER_NOT_IN_ALIAS = 0x80070561
ERROR_MEMBER_IN_ALIAS = 0x80070562
E_ADS_BAD_PATHNAME = 0x80005000
ERROR_NOT_IMPLEMENTED = 0x80004001
E_ADS_PROPERTY_NOT_FOUND = 0x8000500D
E_ADS_PROPERTY_NOT_SUPPORTED = 0x80005006
E_ADS_PROPERTY_INVALID = 0x80005007

WINERROR_MAP = {
  ERROR_MEMBER_NOT_IN_ALIAS : MemberNotInGroupError,
  ERROR_MEMBER_IN_ALIAS : MemberAlreadyInGroupError,
  E_ADS_BAD_PATHNAME : BadPathnameError,
  ERROR_NOT_IMPLEMENTED : NotImplementedError,
  E_ADS_PROPERTY_NOT_FOUND : AttributeError,
  E_ADS_PROPERTY_NOT_SUPPORTED : AttributeError,
  E_ADS_PROPERTY_INVALID : AttributeError,
}
wrapped = wrapper (WINERROR_MAP, ActiveDirectoryError)

