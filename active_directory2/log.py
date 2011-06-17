import os, sys
import logging

formatter = logging.Formatter ("[%(levelname)s] %(module)s.%(funcName)s: %(message)s")
logger = logging.getLogger ("active_directory2")
logger.setLevel (logging.DEBUG)

stderr_handler = logging.StreamHandler (sys.stderr)
stderr_handler.setLevel (logging.WARN)
stderr_handler.setFormatter (formatter)
logger.addHandler (stderr_handler)

debug_handler = logging.FileHandler ("active_directory2.debug.log")
debug_handler.setLevel (logging.DEBUG)
debug_handler.setFormatter (formatter)
logger.addHandler (debug_handler)
