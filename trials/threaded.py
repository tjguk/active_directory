import os, sys
import logging
import Queue
import threading

from active_directory2 import ad

def f (ident, filter, queue):
  for item in ad.AD (server=ident).search (filter):
    queue.put ((ident, item))
  queue.put ((ident, None))

if __name__ == '__main__':
  q = Queue.Queue ()
  t1 = threading.Thread (target=f, args=('svr-dc1', 'objectClass=user', q))
  t2 = threading.Thread (target=f, args=('svr-dc2', 'objectClass=group', q))
  t1.start ()
  t2.start ()
  incomplete = dict (users=True, groups=True)
  while any (incomplete.values ()):
    ident, dn = q.get ()
    if dn is None:
      incomplete[ident] = False
    else:
      print ident, ":", dn.ADsPath.encode (sys.stdout.encoding, "backslashreplace")