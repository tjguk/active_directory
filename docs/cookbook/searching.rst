.. currentmodule:: active_directory
.. highlight:: python
   :linenothreshold: 1

Managing shortcuts
==================

.. _find-a-user:

Find a user
-----------

Find a user in the default AD

..  literalinclude:: searching/find_user.py

Discussion
~~~~~~~~~~
The :func:`find_user` convenience function uses a cached AD root and
returns the first object which is a User and a Person (without
which computers accounts would be returned). The search uses
the built-in Ambiguous Name Resolution so all likely name fields
are matched. To search against a specific field, eg displayName,
use the general :func:`find` function.

The returned item is an :class:`AD_Object`. If no matching user is
found, `None` is returned.