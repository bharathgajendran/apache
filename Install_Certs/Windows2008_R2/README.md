Certificate Automation in Windows servers

This project automates the entire certificate installation and assign it to the websites running in IIS and it has script to create CSR.

Requirements:

Copy the entire project along with directory structure.

Tested platforms:

* Works with Windows server 2012,2012 R2.
* Works with Windows 8.1.
* Supported certificate extensions are cer, crt, p7b and pfx.

Description:

Creating CSR
* Execute generate_csr.ps1 script to create CSR based on the user inputs like common name(URL) oraganization name etc and copies the generated CSR to clipboard.
* Also saves the generated CSR to CSRFOlder.

Installing certificate
* Copies the root and intermediate certiifcate to their respective directories.
* Copy the generated SSL certificate to the personal directory (supported extensions cer, crt, p7b and pfx)
* Execute the install_cert.ps1 script to install the certificates.
* Script also prompts for user input to assign it to the websites running on IIS.

