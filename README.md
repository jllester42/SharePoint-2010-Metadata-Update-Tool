SharePoint-2010-Metadata-Update-Tool
====================================

Powershell Script used to manage metadata for a Sharepoint 2010 server


File Name: metadataTool
Type:	   Windows Powershell Script Version 1.0 (.ps1)
Execution Type: To be run as a Sharepoint scheduled task

Set-up / initialization:
------------------------
The code begins execution by establishing both a SQL and Sharepoint server
connection. Once these connections have been made we begin initializing 
variables that will be used throughout the code.

		
Creating Files:
---------------
The BCP utility will be called, as a sql query, to temporary recreate the files
that are being stored as BLOBs. These files will be saved in a local folder on 
the server computer so they can be accessed when the code enters the extraction
process.

	Functions used by this process: createFilesFromBLOB


Extraction Process:
-------------------
We read through each file generated in the previous step and begin extracting the
metadata from each file/folder. The metadata is initially stored in a basic hash
table; we then use the values from the hash table to generate an xml file which
stores all the extracted metadata (the hash table is deleted after this process).
Finally we create a csv file, whose contents can be inserted into a sql database 
table, from the xml file. This process is finally completed when we add the contents 
of the xml file to a sql database table.

	Functions used by this process: extractMetadata, createCSV, buildTable


Update Process:
---------------
We begin by creating 2 datasets that are filled with information from both the
extracted metadata table and the sharepoint metadata table. The information 
stored in the databases then goes through a comparison process; if all 
comparisons are true we update the specified information with the sharepoint
metadata database with the equivalent information from the extracted metadata
table.

	Functions used by this process: IsNullOrEmpty, updateSharePointMetadata, readInDatabaseValues.


Termination:
------------
Before terminating the code removes all files that might contain sensitive 
information. This includes: deleting all files recreated using blob data, 
deleting the xml and csv files that store the extracted metadata information,
and deleting the contents of the extracted metadata table. Once complete,
the code closes both the sql and sharepoint server connections and terminates.

	Functions used by this process: deleteData
	
