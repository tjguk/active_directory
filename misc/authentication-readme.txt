* ADsOpenObject and OpenDSObject are the same
* ADsGetObject and GetObject are the same
* Connections are cached so long as the server & username are the same (and various flags)
* Don't cache the password - use a NULL password on every invocation after the first
