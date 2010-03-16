import os, sys

import active_directory as ad

def main ():
  big_ou = ad.find_ou ("BigOU") or ad.root ().new_ou ("BigOU")
  big_group = big_ou.find_group ("BigGroup") or big_ou.new ("group", "BigGroup")

if __name__ == '__main__':
  main (*sys.argv[1:])
