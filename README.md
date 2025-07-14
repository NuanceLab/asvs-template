Easy to fill ASVS worksheets, color coded to help keep track of implemented changes. The expected use case is to provide a checklist-type spreadsheets from the ASVS.  

The script to parse the OWASP Application Security Verification Standard (ASVS) JSON format to xlsx, is derived from the work of cxosmo (https://github.com/cxosmo) in the repository https://github.com/cxosmo/asvs-to-xlsx, modified to work with the new schema implemented in version 5.0.0 of ASVS with a few formatting changes. Optional flag available (`-c`) for appending custom columns. 

*Tested on [ASVS 5.0.0](https://github.com/OWASP/ASVS/releases/download/v5.0.0_release/OWASP_Application_Security_Verification_Standard_5.0.0_en.json) - standard (non-flat)(non-legacy) version only. Script may require updating if the ASVS JSON schema changes!*
