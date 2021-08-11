import csv
import sys

#
#   Script that generates New-MailboxExportRequest commands from Office 365 PST mapping file 
# 
#   Run script with: py .\ExchangeMailExportGenerator.py "K:\\path\\to\\PstImportMappingFile.csv" "C:\\temp\\testcsv.txt" "Host-EX02"
#
#   input = "K:\\path\\to\\PstImportMappingFile.csv"
#   output = "C:\\temp\\testcsv.txt"
#   server name = "Host-EX02"
#

try:

    # Assings the input arguments 
    input = sys.argv[1]
    output = sys.argv[2]
    serverName = sys.argv[3]
    # Creates the output file
    outputFile = open(output, "w")

    # Opens the PST mapping file
    with open(input, newline='') as f:

        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(f.readline())

        # Reads the csv file using the defined delimiter
        reader = csv.reader(f, delimiter=dialect.delimiter)

        # For each csv row write the export command to the output file
        for row in reader:
            outputFile.write(f'New-MailboxExportRequest -Mailbox "{row[3]}" -FilePath "\\\{serverName}\Export\{row[2]}"\n')
    
    # Saves and closes the output file
    outputFile.close()
    # Prints success message in console
    print(f"Successfully created file with Exchange export commands!\r\nFile location: {output}")

except Exception as ex:
    print(ex)
