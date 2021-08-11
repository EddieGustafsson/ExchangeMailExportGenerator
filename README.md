# ExchangeMailExportGenerator
Python script that generates New-MailboxExportRequest commands from Office 365 PST mapping file 

## Running the script
Run script with: `py .\ExchangeMailExportGenerator.py <Input> <Output> <Server name>`

**Parameters:**

Input = "K:\\path\\to\\PstImportMappingFile.csv"

Output = "C:\\temp\\testcsv.txt"

Server name = "Host-EX02"

## Recommended content
- [Overview of importing your organization's PST files](https://docs.microsoft.com/en-us/microsoft-365/compliance/importing-pst-files-to-office-365)
- [Use network upload to import your organization's PST files to Microsoft 365](https://docs.microsoft.com/en-us/microsoft-365/compliance/use-network-upload-to-import-pst-files)
