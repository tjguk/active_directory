import os, sys
import logging

logger = logging.getLogger ("active_directory2")
logger.setLevel (logging.DEBUG)
stderr_handler = logging.StreamHandler (sys.stderr)
stderr_handler.setLevel (logging.DEBUG)
logger.addHandler (stderr_handler)
