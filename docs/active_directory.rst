:mod:`active_directory2` -- Windows Active Directory Management
===============================================================

..  automodule:: active_directory2.core
    :members:

..  automodule:: active_directory2.adbase
    :members:

..  autoclass:: active_directory2.adbase.ADBase
    :members:

..  autoclass:: active_directory2.adbase.ADContainer
    :members:

Exceptions
----------

All COM-related exceptions are wrapped in :exc:`active_directory2.exc.ActiveDirectoryError` or one of its
subclasses. Therefore you can safely trap :exc:`active_directory2.exc.ActiveDirectoryError` as a root exception.

..  autoexception:: active_directory2.exc.ActiveDirectoryError
