Certificate Automation in Windows servers

This project contains various scripts that monitor the expiry of SSL certificate, CSR creation, installing and assigning the certificates to websites running on IIS along with root and intermediate.

Requirements:

Copy the entire project along with directory structure.

Tested platforms:

* Works with Windows server 2012,2012 R2.
* Dedicated script for Windows Server 2008 R2.
* Supported certificate extensions are cer, crt, p7b and pfx.

Description:

Monitoring SSL certificate expiry
* Monitors the expiry of the certificates used by the websites running on IIS.
* Send a customized email along with various details like Site name as per IIS, DNS name, Expiry date and Issuer in tabular form.

Creating CSR
* Execute generate_csr.ps1 script to create CSR based on the user inputs like common name(URL) organization name etc and copies the generated CSR to clipboard.
* Also saves the generated CSR to CSRFOlder.

Installing certificate
* Copies the root and intermediate certificate to their respective directories.
* Copy the generated SSL certificate to the personal directory (supported extensions cer, crt, p7b and pfx)
* Execute the install_cert.ps1 script to install the certificates.
* Script also prompts for user input to assign it to the websites running on IIS.

