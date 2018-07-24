# TrustedRecords
A Powershell script to search for TrustedRecords registry and parse the results.

# TrustRecords:
When a user click "Enable Content" or "Enable Editing" after opening a document files, the Office application will create a registry record containing some information about this file, so when the user open the document later it will not show the enabling button.

# How to use:

Run the Powershell script

>> CheckMacro.ps1 .\results.csv

The output in results.csv will give the following information:
. File Path
. Registry Path
. Status
. Creation Time
. User SID
. Username
. File Exists
