# UPDATE FILE, FOLDER, AND DATABASE VALUES/NAMES AS NEEDED.

Add-PSSnapin "Microsoft.SharePoint.PowerShell"

# Sharepoint SQL Connection:
# --------------------------
$SPsqlServer = "portal.vtdev.vt.edu"
$SPSqlTable = "AllDocs"
$blobTable = "AllDocStreams"
$SPsqlConnection = New-Object Microsoft.SQLServer.Management.SMO.Server
$SPsqlConnection.ConnectionString = "Server = $SPsqlServer; Integrated Security=SSPI;"
$SPsqlConnection.Open()


# SQL Connection:
# ---------------

# CHANGE TO REFLECT PERMANANT VARIABLE NAMES AND FILE PATHS
$SQLServer = "vtdevsql01"
$SQLDBName = "JLTest"
$SqlTable = "metadataTable"
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName;  User ID=; Password=;"
$SqlConnection.Open()


# Location of the files, recreated from blob information, on the server (host) computer 
# MUST include "\" at end of file path.
$fileFolder = "C:\test\"

# The location and file name of the created XML file used to temporarily represent extracted metadata information.
$XMLFile = "C:\scripts\outputFile.xml"

# The location/name of the CSV file, that is used to store extracted metadata information.
$CSVFile = "C:\scripts\outputFile.csv"

# The format file used to recreate the files from blob information.
$formatFile = "C:\scripts\formatFile.fmt"


# Determines if a string is null or empty.
function IsNullOrEmpty($str) 
{
	if ($str) {
		return $false
	} 
	else {
		return $true
	}
}


# Creates a folder of file that have been reconstructed using BLOB information, the
# metadata will be extracted from these files.
function createFilesFromBLOB()
{
	# Fills an array with blob IDs to be used when reconstructing files.
    $idQuery = "(SELECT Id FROM $SPSqlTable)"
	$idCmd = New-Object System.Data.SqlClient.SqlCommand
	$idCmd.CommandText = $idQuery
	$idCmd.Connection = $SPsqlConnection
	$idAdapter = new-object system.data.sqlclient.sqldataadapter
	$idAdapter.SelectCommand = $idCmd
    $idTable = new-object system.data.datatable
    $idAdapter.Fill($idTable) | out-null
    $idArray = @($idTable | select -ExpandProperty ID)

	# Fills an array with blob file names to be used when reconstructing files.
	$nameQuery = "(SELECT LeafName FROM $SPSqlTable)"
	$nameCmd = New-Object System.Data.SqlClient.SqlCommand
	$nameCmd.CommandText = $nameQuery
	$nameCmd.Connection = $SPsqlConnection
	$nameAdapter = new-object system.data.sqlclient.sqldataadapter
	$nameAdapter.SelectCommand = $nameCmd
    $nameTable = new-object system.data.datatable
    $nameAdapter.Fill($nameTable) | out-null
    $nameArray = @($nameTable | select -ExpandProperty Name)
	
	$totalCount = $nameArray.length
	
	New-Item $formatFile -type file -force | out-null 
	
	# Creates a format file to be used when reconstructing files.
	$formatQuery = "BCP $blobTable format nul -f '$formatFile' -T"
	$formatCmd = New-Object System.Data.SqlClient.SqlCommand
	$formatCmd.CommandText = $formatQuery
	$formatCmd.Connection = $SPsqlConnection
	$formatCmd.ExecuteNonQuery()
	
	# Reconstructs the files represented as blobs
	# Count begins at 0 due to array indexing
	for ($count = 0; $count -lt $totalCount; $count++) {	
	
		$saveTo = $fileFolder + $nameArray[$count]
		
		$blobQuery = "BCP 'SELECT RbsId FROM $blobTable where Id = $idArray[$count]' QUERYOUT '$saveTo' -T -f '$formatFile' -S"
					
		$blobCmd2 = New-Object System.Data.SqlClient.SqlCommand
		$blobCmd2.CommandText = $blobQuery
		$blobCmd2.Connection = $SPsqlConnection
		$blobCmd2.ExecuteNonQuery()
	}
}
# Calls the createFilesFromBLOB function
createFilesFromBLOB 


# Extracts the metadata from a file, saves the information in a hash table, and
# then uses the hash table to create an xml file.
function extractMetadata($location)
{
	# Searches each folder for files to extract metadata from.
	ForEach($sFolder in $location ) {

		$currentIndex = 143
		$objShell = New-Object -ComObject Shell.Application
		$objFolder = $objShell.namespace($sFolder)

		# We extract the metadata for each file within a folder.
		ForEach ($file in $objFolder.Items()) {
		
			# The ContentCreated metadata is extracted from the file and added to the hash table.
			if ($currentIndex -eq 143)
			{
				$hash += @{ `

					$($objFolder.getDetailsOf($objFolder.Items, $currentIndex))  =`
					$($objFolder.getDetailsOf($file, $currentIndex))
				}
			}
			
			$currentIndex = 145;
			
			# The DateLastSaved metadata is extracted from the file and added to the hash table.
			if ($currentIndex -eq 145)
			{
				$hash += @{ `

					$($objFolder.getDetailsOf($objFolder.Items, $currentIndex))  =`
					$($objFolder.getDetailsOf($file, $currentIndex))
				}
			}
			
			$currentIndex = 155;
			
			# The Filename metadata is extracted from the file and added to the hash table.
			if ($currentIndex -eq 155)
			{
				$hash += @{ `

					$($objFolder.getDetailsOf($objFolder.Items, $currentIndex))  =`
					$($objFolder.getDetailsOf($file, $currentIndex))
				}
			}
			
			$xmlCpn = $xmlDoc.CreateElement("File")
			
			# Creates an xml element using the values from the hash table generated earlier.
			foreach($entry in $hash.keys)
			{ 
				#Removes the spaces from the key values so they can be used as xml element names.
				#Also capilalizes the first letter of each word for improved readability.
				$entryModified = [Regex]::Replace($entry, '\b(\w)', { param($m) $m.Value.ToUpper() });
				$entryModified = $entryModified -replace " ",""

				$xmlElt = $xmlDoc.CreateElement($entryModified)
				$xmlText = $xmlDoc.CreateTextNode($hash[$entry])
				$null = $xmlElt.AppendChild($xmlText)
				$null = $xmlCpn.AppendChild($xmlElt)

				$null = $xmlDoc.LastChild.AppendChild($xmlCpn);
			}
			
			$hash.clear()
			$currentIndex = 143;
		}
	}
}


# Recursively creats a file with references to all other folders contained in the original
# data-path; this includes sub-folders of folders.
$folder = Get-ChildItem -Path $fileFolder -Recurse | where {$_.PsIsContainer} | % { $_.FullName } 


# Create XML root
[xml]$xmlDoc = New-Object system.Xml.XmlDocument
$xmlDoc.LoadXml("<?xml version=`"1.0`" encoding=`"utf-8`"?><Root></Root>")


# Calls the extractMetadata function for the given file path
extractMetadata $fileFolder


# Calls the extractMetadata function for any subfolders within the given path
# given that they are not null or empty.
if ( !(IsNullOrEmpty $folder)) {
    extractMetadata $folder
}


# Once the XML file has been fully constructed, we save the file so it can be accessed/read.
$xmlDoc.Save($XMLFile)


# Converts the XML file to a CSV file
# Unicode encoding is used in place of UTF8 because the BULK INSERT SQL command cannot properly
# process csv files with UTF8 encoding.
function createCSV()
{
	# Reads in the XML file from memory and converts to a CSV file.
	[xml]$inputFile = Get-Content $XMLFile
	# export xml as csv
	$inputFile.Root.ChildNodes | Export-Csv $CSVFile -NoTypeInformation -Delimiter:';' -Encoding:Unicode 

	# File Modifications:
	# Removes the "" from the csv file
	(gc $CSVFile) -replace('"','') | Out-File $CSVFile -Force

	# Removes the first line of the csv file that contains the headings for each field.
	$LinesCount = $(get-content $CSVFile).Count
	get-content $CSVFile |
		select -Last $($LinesCount-1) | 
		set-content "$file-temp"
	move "$file-temp" $CSVFile -Force

	# Removes the ? from the time/date fields, the ? modifier is added in the move process above. 
	(gc $CSVFile) -replace("\?","") | Out-File $CSVFile -Force -Encoding:Unicode
}
# Calls the createCSV function.
createCSV


# Builds a new metadata table using the CSV file generated above, this table will
# overwrite the previous table so that the data fields contain the correct values
function buildTable()
{
	$SqlQuery = "USE $SQLDBName
				DROP TABLE $SqlTable
				CREATE TABLE $SqlTable
				(	
					DateLastSaved DATETIME,	
					Filename VARCHAR(255),	
					ContentCreated DATETIME
				)
				BULK
				INSERT $SqlTable
				FROM $CSVFile
				WITH
				(
					DataFileType = 'widechar',
					KEEPNULLS,
					FIELDTERMINATOR = ';',
					ROWTERMINATOR = '\n'
				)"

	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	$SqlCmd.CommandText = $SqlQuery
	$SqlCmd.Connection = $SqlConnection
	$SqlCmd.ExecuteNonQuery()
}
# Calls the buildTable function
buildTable


# Sharepoint Metadata Query to determine number of table entries:
# ---------------------------------------------------------------
$SPSqlQuery = "SELECT * from $SPSqlTable"
$SPSqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SPSqlCmd.CommandText = $SPSqlQuery
$SPSqlCmd.Connection = $SPsqlConnection
$SPSqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SPSqlAdapter.SelectCommand = $SPSqlCmd
$SPDataSet = New-Object System.Data.DataSet
$SProwNumber = $SPSqlAdapter.Fill($SPDataSet)


# SQL Metadata Query to determine number of table entries:
# ---------------------------------------------------------
$SqlQuery2 = "SELECT * from $SqlTable"
$SqlCmd2 = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd2.CommandText = $SqlQuery2
$SqlCmd2.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd2
$DataSet = New-Object System.Data.DataSet
$rowNumber = $SqlAdapter.Fill($DataSet)


# Updates the metadata values stored in the sharepoint database
# change connection to Sharepoint connection?
function updateSharePointMetadata($Filename, $DatelastSaved, $ContentCreated, $SPFilename, $SPDatelastSaved, $SPContentCreated)
{
	# We only attempt updating metadata for matching files
	if( $SPFilename -eq $Filename)
	{
		# We attempt to update the datelastsaved metadata only if the extracted metadata does not match
		# the original metadata and the extracted value isn't null.
		if ( ($DatelastSaved -ne $SPDatelastSaved) -and !(IsNullOrEmpty $DatelastSaved) )
		{
			write-host "Updating DateLastSaved for" $Filename
			$SqlQuery3 = "update $SPSqlTable set DateLastSaved = '$DatelastSaved' where Filename='$SPFilename' "
			$SqlCmd3 = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd3.CommandText = $SqlQuery3
			$SqlCmd3.Connection = $SqlConnection
			$SqlCmd3.ExecuteNonQuery() | out-null
		}
	
		# We attempt to update the contentcreated metadata only if the extracted metadata does not match
		# the original metadata and the extracted value isn't null.
		if ( ($ContentCreated -ne $SPContentCreated) -and !(IsNullOrEmpty $ContentCreated)  )
		{
			write-host "Updating ContentCreated for" $Filename 
			$SqlQuery4 = "update $SPSqlTable set ContentCreated = '$ContentCreated' where Filename='$SPFilename' "
			$SqlCmd4 = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd4.CommandText = $SqlQuery4
			$SqlCmd4.Connection = $SqlConnection
			$SqlCmd4.ExecuteNonQuery() | out-null
		}
	}
}


# Reads in the Database values and sets the parameters for updateSharePointMetadata accordingly.
function readInDatabaseValues{

	if ($rowNumber -eq $SProwNumber)
	{
		for ($index = 0; $index -lt $SProwNumber; $index++)
		{
			$i = $index;
			$j = $index;
		
			$Filename = $null
			$DateLastSaved = $null
			$ContentCreated = $null
			
			$SPFilename = $null
			$SPDateLastSaved = $null
			$SPContentCreated = $null
		
			if ( !(IsNullOrEmpty "$($DataSet.Tables[0].Rows[$j][1])"))
			{		
				$Filename = "$($DataSet.Tables[0].Rows[$j][1])"
		
				$DateLastSaved = "$($DataSet.Tables[0].Rows[$j][0])"
				
				$ContentCreated = "$($DataSet.Tables[0].Rows[$j][2])"
			}

			if ( !(IsNullOrEmpty "$($SPDataSet.Tables[0].Rows[$i][3])"))
			{
				$SPFilename = "$($SPDataSet.Tables[0].Rows[$i][3])"
	
				$SPDateLastSaved = "$($SPDataSet.Tables[0].Rows[$i][23])"
			
				$SPContentCreated = "$($SPDataSet.Tables[0].Rows[$i][22])"
			}
						
			updateSharePointMetadata $Filename $DatelastSaved $ContentCreated $SPFilename $SPDatelastSaved $SPContentCreated
		}	
	}
}
# Calls the readInDatabaseValues function
readInDatabaseValues


# Deletes the files created from the blob data, deletes the CSV and XML files,
# and deletes the contents of the metadata table.
function deleteData()
{
	# Remove comment notaton (#) once local testing has been completed so that local files are not deleted.
	# Remove-Item ($fileFolder + '*')
	Remove-Item $XMLFile;
	Remove-Item $CSVFile;
	Remove-Item $formatFile;
	
	$emptyQuery = "TRUNCATE TABLE $SqlTable"
	$emptyCmd = New-Object System.Data.SqlClient.SqlCommand
	$emptyCmd.CommandText = $emptyQuery
	$emptyCmd.Connection = $SqlConnection
	$emptyCmd.ExecuteNonQuery() | out-null
}
# Calls the deleteData function
deleteData


# Closes the server connection established at the beginning of the code.
$SqlConnection.Close()
$SPsqlConnection.Close()
