:mod:`active_directory` -- Windows Active Directory Management
==============================================================

..  automodule:: active_directory

Exceptions
----------

All COM-related exceptions are wrapped in :exc:`ActiveDirectoryError` or one of its
subclasses. Therefore you can safely trap :exc:`ActiveDirectoryError` as a root exception.

..  autoexception:: ActiveDirectoryError

Support Classes & Functions
---------------------------

..  autofunction:: connect
..  autofunction:: query
..  autofunction:: query_string

Main Entry Points
-----------------

..  autofunction:: ad

..  autofunction:: find_user
..  autofunction:: find_group
..  autofunction:: find_ou
..  autofunction:: search

