"""By default, the active_directory package assumes that you're
using standard Windows authentication to bind to AD. But there
are two other options: anonymous authentication; and user/password
sign-on. If you need either of those, you'll have to create a
credentials object and pass it around to any function which needs
to bind or query AD.
"""

"""Query AD as an anonymous user
"""
from active_directory import core, credentials

qs = core.query_string (base="LDAP://sibelius")
for i in core.query (query_string=qs, connection=core.connect (cred=credentials.Anonymous)):
  print i

