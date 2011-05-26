python -i -c "from active_directory2 import core, credentials, adbase, ad; credentials.push (('tim@westpark.local', 'password', 'holst')); root = adbase.ADBase (core.root_obj ('holst'))"
