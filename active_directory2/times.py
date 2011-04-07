from ctypes.wintypes import *
from ctypes import windll, wintypes
import ctypes
import win32api
import win32file

kernel32 = windll.kernel32

def error (exception, context="", message=""):
  errno = win32api.GetLastError ()
  message = message or win32api.FormatMessageW (errno)
  raise exception (errno, context, message)

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

def file_time_to_system_time (ularge):
  filetime = FILETIME (ularge.HighPart, ularge.LowPart)
  systemtime = SYSTEMTIME ()
  if kernel32.FileTimeToSystemTime (ctypes.pointer (filetime), ctypes.byref (systemtime)) == 0:
    error (RuntimeError)
  return systemtime

if __name__ == '__main__':
  class ULARGE (object):
    def __init__ (self, hi, lo):
      self.HighPart = hi
      self.LowPart = lo

  ctime = FILETIME ()
  atime = FILETIME ()
  wtime = FILETIME ()
  handle = HANDLE (int (fs.handle (sys.executable)))
  if not kernel32.GetFileTime (handle, ctypes.byref (ctime), ctypes.byref (atime), ctypes.byref (wtime))
    error (RuntimeError)

  ularge = ULARGE (2147483647, -1)
  st = file_time_to_system_time (ularge)
  print st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond
