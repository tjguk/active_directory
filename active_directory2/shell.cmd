python -i -c "from active_directory2 import core, credentials, adobject; credentials.push (('tim@westpark.local', 'password', 'sibelius')); root = adobject.ADObject (core.root_obj ('sibelius'))"
