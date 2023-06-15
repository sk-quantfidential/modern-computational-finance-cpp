# modern-computational-finance-cpp

## Signing XLL

### Create test certificate

https://learn.microsoft.com/en-us/windows-hardware/drivers/install/creating-test-certificates

Create only a MakeCert test certificate to sign all driver packages on a development computer.

```shell
makecert -r -pe -ss PrivateCertStore -n CN=Quantfidential.com(Test) -eku 1.3.6.1.5.5.7.3.3 Quantfidential.cer
```

* The -r option creates a self-signed certificate with the same issuer and subject name.
* The -pe option specifies that the private key that is associated with the certificate can be exported.
* The -ss option specifies the name of the certificate store that contains the test certificate (PrivateCertStore).
* The -n CN= option specifies the name of the certificate, Quantfidential.com(Test). This name is used with the SignTool tool to identify the certificate.
* The EKU option inserts a list of one or more comma-separated, enhanced key usage object identifiers (OIDs) into the certificate. For example, -eku 1.3.6.1.5.5.7.3.2 inserts the client authentication OID. For definitions of allowable OIDs, see the Wincrypt.h file in CryptoAPI 2.0.
* QuantfidentialTest.cer is the file name that contains a copy of the test certificate, Quantfidential.com(Test). The certificate file is used to add the certificate to the Trusted Root Certification Authorities certificate store and the Trusted Publishers certificate store.

### Install Test Certificate to Verify Signature

To install a test-signed driver package on a test computer, allow the test computer to verify the signature by install the test certificate in the test computer's Trusted Root Certification Authorities (CA) certificate store.

Add the CA certificate - once - to the Trusted Root Certification Authorities certificate store to verify the signature of all drivers or driver packages digitally signed with the certificate.

``` shell
certmgr /add QuantfidentialTest.cer /s /r localMachine root
```

* The /add option specifies that the certificate in the ContosoTest.cer file is to be added to the specified certificate store.
* The /s option specifies that the certificate is to be added to a system store.
* The /r option specifies the system store location, which is either currentUser or localMachine.
* Root specifies the name of the destination store for the local computer, which is either root to specify the Trusted Root Certification Authorities certificate store or trustedpublisher to specify the Trusted Publishers certificate store.

