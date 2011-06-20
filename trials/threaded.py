import os, sys
import logging
import Queue
import threading

from active_directory2 import ad, credentials

def f (ident, filter, queue):
  with credentials.credentials (("tim@westpark.local", "password", ident)):
    for item in ad.AD (server=ident).search (filter):
      queue.put ((ident, item))
    queue.put ((ident, None))

if __name__ == '__main__':
  q = Queue.Queue ()
  t1 = threading.Thread (target=f, args=('holst', 'objectClass=user', q))
  t2 = threading.Thread (target=f, args=('holst', 'objectClass=group', q))
  t1.start ()
  t2.start ()
  incomplete = dict (users=True, groups=True)
  while any (incomplete.values ()):
    ident, dn = q.get ()
    if dn is None:
      incomplete[ident] = False
    else:
      print ident, ":", dn.ADsPath.encode (sys.stdout.encoding, "backslashreplace")