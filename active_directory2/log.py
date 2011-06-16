import os, sys
import logging

formatter = logging.Formatter ("%(levelname)s: %(module)s.%(funcName)s - %(message)s")
logger = logging.getLogger ("active_directory2")
logger.setLevel (logging.DEBUG)
stderr_handler = logging.StreamHandler (sys.stderr)
stderr_handler.setLevel (logging.DEBUG)
stderr_handler.setFormatter (formatter)
logger.addHandler (stderr_handler)
