python -i -c "from active_directory2 import core, credentials, ad; credentials.push (('tim@westpark.local', 'password', 'holst')); root = ad.adobject(core.root_obj ('holst'))"
