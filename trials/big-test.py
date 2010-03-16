import os, sys

import active_directory as ad

def main ():
  big_ou = ad.root ().find_ou ("BigOU") or ad.root ().new_ou ("BigOU")
  big_group = big_ou.find_group ("BigGroup") or big_ou.new_group ("BigGroup")
  for i in range (4000):
    username = "user%04d" % i
    user = big_ou.find_user (username)
    if not user:
      print username
      user = big_ou.new ("user", username)
      big_group.com_object.Add (user.ADsPath)

if __name__ == '__main__':
  main (*sys.argv[1:])
