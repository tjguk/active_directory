import os, sys

import active_directory as ad

def main ():
  big_ou = ad.find_ou ("BigOU")
  if big_ou is None:
    big_ou = ad.root ().new_ou ("BigOU")

if __name__ == '__main__':
  main (*sys.argv[1:])
