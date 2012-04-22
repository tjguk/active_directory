"""
In principle, having created a connection with username & password,
any future connection to the same server / domain with the same
same flags can simply use the username with a NULL password. This
does not, however, appear to work.
"""
import os, sys
import unittest

from win32com.adsi import adsi, adsicon


class Test(unittest.TestCase):

    server = "holst"
    domain = "westpark.local"
    username = "fred@%s" % domain
    password = "Pa$$w0rd"

    def open_object(self, path, use_password=True):
        return adsi.ADsOpenObject(
            path,
            self.username,
            self.password if use_password else None,
            adsicon.ADS_SECURE_AUTHENTICATION,
            adsi.IID_IADs
        )

    def test_with_server(self):
        root = adsi.ADsGetObject("LDAP://%s/rootDSE" % self.server, adsi.IID_IADs)
        domain_rdn = root.Get("defaultNamingContext")
        domain0 = self.open_object("LDAP://%s/%s" % (self.server, domain_rdn), use_password=True)
        domain1 = self.open_object("LDAP://%s/%s" % (self.server, domain_rdn), use_password=False)
        self.assertEquals(domain0.objectGuid, domain1.objectGuid)

    def test_with_domain(self):
        root = adsi.ADsGetObject("LDAP://%s/rootDSE" % self.domain, adsi.IID_IADs)
        domain_rdn = root.Get("defaultNamingContext")
        domain0 = self.open_object("LDAP://%s/%s" % (self.server, domain_rdn), use_password=True)
        domain1 = self.open_object("LDAP://%s/%s" % (self.server, domain_rdn), use_password=False)
        self.assertEquals(domain0.objectGuid, domain1.objectGuid)

    def test_with_serverless(self):
        root = adsi.ADsGetObject("LDAP://rootDSE", adsi.IID_IADs)
        domain_rdn = root.Get("defaultNamingContext")
        domain0 = self.open_object("LDAP://%s" % (domain_rdn), use_password=True)
        domain1 = self.open_object("LDAP://%s" % (domain_rdn), use_password=False)
        self.assertEquals(domain0.objectGuid, domain1.objectGuid)

if __name__ == '__main__':
    unittest.main()