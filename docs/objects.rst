..  module:: active_directory

Objects
=======

..  py:class:: _AD_object

    An object which represents an ADSI instance of an Active Directory object.
    The object gives attribute access to the AD object's attributes, casting
    them where possible to Python-friendly datatypes (eg datetime instances).
    As a convenience, the :meth:`dump` method displays all the attributes of
    this object in a dictionary-like output.

    Where permitted, the attribute values can also be read. At present, the
    equivalent casting is not carried out, so datetime values must be written
    back as, in some cases, the number of 100ns intervals since Jan 1601.
    Since there is something of an overhead in setting attributes individually,
    the :meth:`set` method allows a number of attributes to be set at one go.

    Objects of this type are not expected to be instantiated directly (although
    it is possible and is occasionally a convenient workaround). Instead, they
    are returned from searches on objects higher up in the tree, or from a
    call to the module-level :func:`AD` function which returns the root of
    a domain tree.

    Equality & hashing are implemented on top of the AD object's GUID. This
    is then guaranteed to work even if an object has been moved (where its
    path would then have been changed).

    Where iteration is possible (ie for an underlying AD container) the object
    can be iterated over in the usual ways. A couple of extended iteration patterns
    are also provided: :meth:`walk` mimics Python's own os.walk functionality,
    yielding `container`, `containers`, `items` from the point at which it's
    called; :meth:`flat` returns all the items from the walk.

    The underlying COM object is available as the :attr:`com_object` attribute.
    The list of possible properties, mandatory & optional, is available as the
    :attr:`properties` attribute.

    The `find` and `search` functions described in :doc:`searching` are available
    here as instance methods, searching from this object downwards. Indeed, the
    module-level functions are implemented by constructing a domain root object
    and calling its `search` or `find` methods.
