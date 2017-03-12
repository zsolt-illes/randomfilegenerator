[cmdletbinding()]
param 
(
    [Parameter(Mandatory=$true)][string]$DictionaryFile,
	[Parameter(Mandatory=$true)][int]$Columns,
    [Parameter(Mandatory=$true)][int]$Rows,
	[Parameter(Mandatory=$true)][int]$FileCount,
	[Parameter(Mandatory=$true)][string]$DestinationFolder,
	[Parameter(Mandatory=$true)][string]$FilePrefix,
	[Parameter(Mandatory=$false)][string]$PersonListFile,
	[Parameter(Mandatory=$false)][string]$DatesFile,
    [Parameter(Mandatory=$false)][string]$TemplateXLSXFile,
    [Parameter(Mandatory=$false)][switch]$NoProgressBar
)

$MainProgressVector = 0

function GenerateTXTforXLSX
{
    [cmdletbinding()]
    param 
    (
	    [Parameter(Mandatory=$true)][int]$Columns,
        [Parameter(Mandatory=$true)][int]$Rows,
        [Parameter(Mandatory=$true)][string]$TempTXT
    )

    # Set the progress bar.
    If (!$NoProgressBar)
    {
        Write-Progress -Id 1 -Activity 'Generating text...' -PercentComplete 0
    }

    # Making sure the TempTXT file does not exist
    If (Test-Path $DestinationFolder -PathType Leaf)
    {
        Try
        {
            Remove-Item $TempTXT -Force -ErrorAction SilentlyContinue
        }
        Catch
        {
            Write-Host 'There was a Temp.txt file in the destination folder: ' -ForegroundColor Red -NoNewline
            Write-Host $DestinationFolder -ForegroundColor Cyan
            Write-Host 'The script tried to delete it, but the file is locked.' -ForegroundColor Red
            Write-Host 'The script halted.' -ForegroundColor Yellow
            Break
        }
    }
    $FileStream = [System.IO.StreamWriter] $TempTXT

    # Construct table file
    $GUID = ([guid]::NewGuid()).GUID
    $FileStream.Write($GUID)
    

    For ($Column=1;$Column -le $Columns; $Column++)
    {
        $RandomWordLine = Get-Random -Minimum 1 -Maximum $DictionaryFileRows
		$RandomWord = $DictionaryFileContent[$RandomWordLine]
        $FileStream.Write("`t$RandomWord")
    }

    For($Row = 1; $Row -le $Rows; $Row++)
    {
        If (!$NoProgressBar)
        {
            Write-Progress -Id 1 -Activity 'Generating text...' -PercentComplete (($Row/$Rows)*100)
        }
        $FileStream.Write("`r`n") # New line
        $RandomWordLine = Get-Random -Minimum 1 -Maximum $DictionaryFileRows
		$RandomWord = $DictionaryFileContent[$RandomWordLine]
        $fileStream.Write("$RandomWord`t") # Row first column
        For ($Column=1;$Column -le $Columns; $Column++)
        {
            $NumberTXT = [string](Get-Random -Minimum 1 -Maximum 10000) + "`t"
            $FileStream.Write($NumberTXT) # The actual numbers in the table
        }
    }

    $FileStream.Close()
}

function ConvertTXTtoXLSX
{
    [cmdletbinding()]
    param 
    (
        [Parameter(Mandatory=$true)][string]$TempTXT,
        [Parameter(Mandatory=$true)][string]$ExcelFile
    )

    # IF we have a template file, we paste into that.
    If ($TemplateXLSXFile)
    {
        # First, let's clear whatever there is in the file
        $Cleared = $ExcelSheet.UsedRange.Clear()
        # Then open the TXT file
        Try
        {
            $TempDocument = $ExcelApplication.Workbooks.Open($TempTXT,$null,$true)
            $TempSheet = ($TempDocument.Sheets)[1]
        }
        Catch
        {
            Write-Host 'Could not open the TXT file: ' -ForegroundColor Red -NoNewline
            Write-Host $TempTXT -ForegroundColor Cyan
        }
        # Copy the used cells
        $Copied = $TempSheet.UsedRange.Copy()
        If (!$Copied)
        {
            Write-Host 'Could not copy the content of the ' -ForegroundColor Red -NoNewline
            Write-Host $TempTXT -ForegroundColor Cyan -NoNewline
            Write-Host ' file.'
        }

        # Then paste it into our Sheet.
        $Pasted = $ExcelSheet.Range("A1").PasteSpecial()
        If (!$Pasted)
        {
            Write-Host 'Could not paste into the template file.' -ForegroundColor Red
        }

        # Save it as the Excel file
        Try
        {
            $ExcelDocument.SaveAs($ExcelFile, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault)
        }
        Catch
        {
            Write-Host 'Could not save the Excel file: ' -ForegroundColor Red -NoNewline
            Write-Host $ExcelFile -ForegroundColor Cyan
        }

        # And finally close the template file and remove it.
        $TempDocument.Close()
        Remove-Item $TempTXT -Force -ErrorAction SilentlyContinue

    }
    #If not, then we just opent the TXT and save it as an XLS
    Else
    {
        $ExcelWorkBook = $ExcelApplication.Workbooks.Open($TempTXT)
        $ExcelWorkBook.SaveAs($ExcelFile, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault)
        $ExcelWorkBook.Saved = $true
        $ExcelWorkBook.Close()
        # And delete the Template TXT file as we do not need that anymore.
        Remove-Item $TempTXT -Force -ErrorAction SilentlyContinue
    }

}


# Loading and creating the model XML file
Function EnsureCoreXML
{
	[cmdletbinding()]
	param
	(
		[Parameter(Mandatory=$true)][string]$WordFilePath
	)

    $Script:CoreXMLCreated=$false
	$CoreXMLPath = "$DestinationFolder\XML\core.xml"

	# Fist test if we already have the core XML in the memory

	# No, so we have to check if it exist already. If yes, we load it.
	$CoreXMLExist = Test-Path $CoreXMLPath -ErrorAction SilentlyContinue
	If ($CoreXMLExist)
	{
		# It does exist, so we try to load it.
		Try
		{
			[xml]$Script:CoreXML = Get-Content $CoreXMLPath -ErrorAction SilentlyContinue
		}
		Catch
		{
			Write-Host 'A core.xml file exist on the ' -ForegroundColor Red -NoNewline
			Write-Host $CoreXMLPath -ForegroundColor Cyan -NoNewline
			Write-Host ' location, but it could not be read.' -ForegroundColor Red
			Write-Host 'Cannot update the document properties.'
			Break
		}
	}
	Else
	{
		# It doesn't exist, so we create one by exporting it from the word file.
	    # First we create a copy of the DOCX file to a ZIP.
	    $ZIPFilePath = $WordFilePath.Replace('xlsx','zip')
	    Try
	    {
		    Copy-Item -Path $WordFilePath -Destination ($ZIPFilePath) -ErrorAction SilentlyContinue
	    }
	    Catch
	    {
		    Write-Host 'Could not copy the Word file: ' -ForegroundColor Red -NoNewline
		    Write-Host $WordFilePath -ForegroundColor Cyan -NoNewline
		    Write-Host ' to ZIP file: ' -ForegroundColor Red -NoNewline
		    Write-Host $($ZIPFilePath)
		    Break
	    }

	    # Then unpack the core.xml
	    Try
	    {
		    $zipfile = [System.IO.Compression.ZipFile]::Open( $ZIPFilePath, 'Read' )
	    }
	    Catch
	    {
		    Write-Host 'Cannot open the zip file: ' -ForegroundColor Red -NoNewline
		    Write-Host $ZIPFilePath -ForegroundColor Cyan
            Write-Host 'The script halted.' -ForegroundColor Yellow
		    Break
	    }
            
        # Extract the content of the core.xml file from the ZIP
        Try
        {
            $ZippedCoreXMLFile = [System.IO.StreamReader]($zipfile.Entries | Where-Object { $_.FullName -match 'docProps/core.xml' }).Open()
		    [xml]$Script:CoreXML = $ZippedCoreXMLFile.ReadToEnd()
        }
        Catch
        {
            Write-Host 'Could not extract the core.xml file from the ' -ForegroundColor Red -NoNewline
            Write-Host $ZIPFilePath -ForegroundColor Cyan -NoNewline
            Write-Host ' file.' -ForegroundColor Red
            Write-Host 'The script halted.' -ForegroundColor Yellow
            Break
        }

        # Here's a little trick that we have to do, because the COM Object generated
        # core.xml does not contain the creator property. If it does not, we have to
        # add it for further use
        If(!($Script:CoreXML.coreproperties.creator))
        {
            # So it seems it is not there, so we add it
            $CreatorNameSpace = 'http://purl.org/dc/elements/1.1'
            $CreatorElement = $Script:CoreXML.CreateElement('dc','creator',$CreatorNameSpace)
            $null = $Script:CoreXML.DocumentElement.AppendChild($CreatorElement)
            # We have to get rid of the namespace definition, otherwise it'll generate an error.
            $CreatorNameSpaceToNull = ' xmlns:dc="' + $CreatorNameSpace + '"'
            $Script:CoreXML = $Script:CoreXML.OuterXml.Replace($CreatorNameSpaceToNull,'')
        }


        # And write it for later use
        Try
        {
            $null = mkdir "$DestinationFolder\XML\" -ErrorAction SilentlyContinue
            $Script:CoreXML.Save($CoreXMLPath)
            $Script:CoreXMLCreated = $true
        }
        Catch
        {
            Write-Host 'Could not create the ' -ForegroundColor Red -NoNewline
            Write-Host "$DestinationFolder\XML\core.xml" -ForegroundColor Cyan -NoNewline
            Write-Host ' file.' -ForegroundColor Red
            Write-Host 'The script halted.' -ForegroundColor Yellow
            Break
        }

        # And the cleanup
        $ZippedCoreXMLFile.Close()
        $ZippedCoreXMLFile.Dispose()
        $zipfile.Dispose()

        Remove-Item -Path $ZIPFilePath -Force -ErrorAction SilentlyContinue
	}
}

# This is a function to update the creator and lastModifiedBy fields of the document
Function UpdateDocXMLProps
{
	[cmdletbinding()]
	param
	(
		[Parameter(Mandatory=$true)][string]$XLSXFilePath
	)

    If (!$Script:CoreXML)
	{
        EnsureCoreXML $XLSXFilePath
    }

    # Check if we have a Persons List.

    If ($PersonListFile)
    {
        # We do, so we update the Creator field
        $RandomPerson = Get-Random -Minimum 0 -Maximum $script:PersonCount
        $RandomPerson = $PersonList[$RandomPerson]
        $Script:CoreXML.coreProperties.creator = $RandomPerson.ToString()
        # And the lastModifiedBy field
        $RandomPerson = Get-Random -Minimum 0 -Maximum ($script:PersonCount-1)
        $RandomPerson = $PersonList[$RandomPerson]
        $Script:CoreXML.coreProperties.lastModifiedBy = $RandomPerson.ToString()
    }

    # Check if we have a Dates List.
    If ($DatesList)
    {
        # We do, so we update the Created field
        $RandomDateNr = Get-Random -Minimum 0 -Maximum $DatesCount
        $RandomDate = $DatesList[$RandomDateNr]
        $Script:CoreXML.coreProperties.created.'#text' = $RandomDate.ToString()
        # And the Modified field
        # Here we need a trick, as we have to make sure the last modifed
        # date is not earlier than the date creation. (This is why we ordered the list.)
        $RandomDateNr = Get-Random -Minimum $RandomDateNr -Maximum $DatesCount
        $RandomDate = $DatesList[$RandomDateNr]
        $Script:CoreXML.coreProperties.modified.'#text' = $RandomDate.ToString()
    }

    # Now that we have the parameters set, we update the Word file with this XML
    Try
    {
        $ExcelFile = [System.IO.Compression.ZipFile]::Open( $XLSXFilePath, 'Update' )
        $ZippedCoreXMLFile = [System.IO.StreamWriter]($ExcelFile.Entries | Where-Object { $_.FullName -match 'docProps/core.xml' }).Open()
        $ZippedCoreXMLFile.BaseStream.SetLength(0)
        $ZippedCoreXMLFile.Write($Script:CoreXML.OuterXml)
        $ZippedCoreXMLFile.Flush()
        $ZippedCoreXMLFile.Close()
        $ExcelFile.Dispose()
    }
    Catch
    {
        Write-Host 'There was an error updating the core properties of: ' -ForegroundColor Red -NoNewline
        Write-Host $XLSXFilePath -ForegroundColor Cyan
    }
}



####################################################
# This is where the main part of the script starts #
####################################################
$StartTime = Get-Date
Cls

# Loading the Word Interop assembly
Write-Host 'Loading assemblies...'
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Loading assemblies' -PercentComplete 0
}
Try
{
    # For operating with Word
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel
    # For operating ZIP files
    $null = [System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')
    Write-Host '... Assemblies loaded.' -ForegroundColor Green
}
Catch
{
    Write-Host 'Could not load assemblies.' -ForegroundColor Red
    Write-Host 'The script halted.' -ForegroundColor Yellow
    Break
}

# Making sure the Destination folder is not having a training backslash
Write-Host 'Ensuring destination folder.'
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Ensuring destination folder' -PercentComplete 0
}
$DestinationFolder = $DestinationFolder.TrimEnd('\')
If (!((Test-Path $DestinationFolder -PathType Container) -and [System.IO.Path]::IsPathRooted($DestinationFolder)))
{
    Try
    {
        Write-Host 'The destination directory (' -NoNewline
        Write-Host $DestinationFolder -ForegroundColor Cyan -NoNewline
        Write-Host ') was not found. Trying to create it...' -NoNewline
        $result = mkdir -Path $DestinationFolder -Force
        Write-Host '... The directory is created.' -ForegroundColor Green
    }
    Catch
    {
        Write-Host 'The specified output directory (' -ForegroundColor Red -NoNewline
        Write-Host $DestinationFolder -ForegroundColor Cyan -NoNewline
        Write-Host ') could not be found, accessed or created.' -ForegroundColor Red
        Write-Host 'Please fix the issue and try gain.' -ForegroundColor Yellow
        Break
    }
}
Write-Host '... Folder confirmed.' -ForegroundColor Green

# Checking if the Template Excel file exist
Write-Host 'Looking for template Excel file.'
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Looking for template Excel file' -PercentComplete 0
}
If($TemplateXLSXFile)
{
    <#
    If ($TemplateXLSXFile)
    {
        Write-Host '############' -ForegroundColor Red
        Write-Host '# Warning! #' -ForegroundColor Red
        Write-Host '############' -ForegroundColor Red
        Write-Host 'The TemplateXLSXFile parameter is defined. The script is using the Windows Clipboard to paste information into the file.' -ForegroundColor Yellow
        Write-Host 'If you use the computer while the files are being generated, ' -ForegroundColor Yellow
        Write-Host 'you might end up with invalid information in your random ' -ForegroundColor Yellow
        Write-Host 'Excel files. ' -ForegroundColor Yellow
        Write-Host 'Do you want to continue? (y/n)'
        $ShouldProceed = Read-Host -Prompt 'Do you want to proceed?'
	    If (!($ShouldProceed -eq 'y') -or !($ShouldProceed -eq 'Y'))
	    {
		    Break
        }
    }
    #>

    If (!((Test-Path $TemplateXLSXFile) -and ([System.IO.Path]::IsPathRooted($TemplateXLSXFile))))
    {
        Write-Host 'The TemplateXLSXFile parameter was defined with the value: ' -ForegroundColor Red -NoNewline
        Write-Host $TemplateXLSXFile -ForegroundColor Cyan -NoNewline
        Write-Host ' but could not be found, or is using relative path. Please try again.' -ForegroundColor Red
        Write-Host 'The script Halted.' -ForegroundColor Yellow
        Break
    }
    Else
    {
        Write-Host '... The file is present.' -ForegroundColor Green
    }
}

# Loading the dictionary file
Write-Host 'Loading the dictionary file: ' -NoNewline
Write-Host $DictionaryFile -ForegroundColor Cyan
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Loading Dictionary File' -PercentComplete 0
}
If (Test-Path $DictionaryFile -ErrorAction SilentlyContinue)
{
    Try
    {
	    $DictionaryFileContent = Get-Content $DictionaryFile -ErrorAction SilentlyContinue
        Write-Host '... Content loaded.' -ForegroundColor Green
        Write-Host
    }
    Catch
    {
        Write-Host 'Cannot access the specified dictionary file: ' -ForegroundColor Red -NoNewline
        Write-Host $DictionaryFile
        Write-Host 'Please try again.'
        Break
    }
}
Else
{
	Write-Host 'The dictionary file specified (' -NoNewline -ForegroundColor Red
	Write-Host $DictionaryFile -NoNewline -ForegroundColor Cyan
	Write-Host ') cannot be found.' -ForegroundColor Red
	Write-Host 'Correct the missing information and try again.' -ForegroundColor Yellow
	Break
}

# Get the lines in the dictionary file. 
#We're going to use this to choose the random words from the file.
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Parsing Dictionary file' -PercentComplete 0
}
$DictionaryFileRows = ($DictionaryFileContent | Measure-Object -Line).Lines
Write-Host 'The dictionary file contains ' -NoNewline
Write-Host $DictionaryFileRows -ForegroundColor Cyan -NoNewline
Write-Host ' rows.'

# Quick sanity check if the dictionary file is not line separeted
If($DictionaryFileRows -eq 1)
{
	Write-Host 'The dictionary file specified (' -NoNewline -ForegroundColor Red
	Write-Host $DictionaryFile -NoNewline -ForegroundColor Cyan
	Write-Host ') is not line separated (probably a CSV), or contains only one entry.' -ForegroundColor Red
	Write-Host 'This script can only operate with a line separeated dictionary file with more than one entry.' -ForegroundColor Red
    Write-Host 'Please try again with a different file.' -ForegroundColor Yellow
	Break
}

# Trying to load the PersonListFile
Write-Host 'Looking for a person list file.' 
$ProgressVector = 0
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Preparing persons list' -PercentComplete 0
}
If($PersonListFile)
{
    If (!$NoProgressBar)
    {
        Write-Progress -Id 1 -Activity "Processing the list of persons" -PercentComplete 0
    }
    Write-Host 'Loading the PersonListFile: ' -NoNewline
    Write-Host $PersonListFile -ForegroundColor Cyan
    Try
    {
        [array]$script:PersonList = Get-Content $PersonListFile -ErrorAction SilentlyContinue
        Write-Host '... Content loaded.' -ForegroundColor Green
    }
    Catch
    {
        Write-Host 'Could not load the specified PersonListFile: ' -ForegroundColor Red -NoNewline
        Write-Host $PersonListFile -ForegroundColor Cyan
        Write-Host 'Please try again.' -ForegroundColor Yellow
        Break
    }

    Write-Host
    Write-Host 'Validating entries in the file...'
    # Quick sanity check on the content
    [int]$NotValidPersons = 0
    [int]$ProgressVector = 0
    $EmailRegex = '^(([\w-]+\.)+[\w-]+|([a-zA-Z]{1}|[\w-]{2,}))@(([\w-]+\.)+[\w-]+|([a-zA-Z]{1}|[\w-]{2,}))'
    ForEach($Person in $PersonList)
    {
        If (!$NoProgressBar)
        {
            $ProgressVector++
            Write-Progress -Id 1 -Activity "Processing the list of persons" -PercentComplete ($ProgressVector / ($($PersonList.Count))*100) -ParentId 0
        }
        If (![regex]::IsMatch($Person, $EmailRegex))
        {
            # This does not seem as an e-mail address.
            $NotValidPersons++
        }
    }
    # Then we validate
    If ($NotValidPersons -gt 0)
    {
        Write-Host 'The person list file specified (' -ForegroundColor Red -NoNewline
        Write-Host $PersonListFile -ForegroundColor Cyan -NoNewline
        Write-Host ' contains ' -ForegroundColor Red -NoNewline
        Write-Host $NotValidPersons -ForegroundColor Cyan -NoNewline
        Write-Host ' entries that do not seem to be a valid email.'
        $ShouldProceed = Read-Host -Prompt 'Do you want to proceed?'
        If (!($ShouldProceed -eq 'y') -or !($ShouldProceed -eq 'Y'))
        {
            Break
        }
    }
    Else
    {
        Write-Host '... All entries seems simantically valid.' -ForegroundColor Green
        Write-Host
    }

    # And we do he count
    $script:PersonCount = $PersonList.Count

    If (!$NoProgressBar)
    {
        Write-Progress -Id 1 -Activity 'Completed.' -Completed
    }
}

#Trying to load the DatesFile
Write-Host 'Looking for a dates list file.'
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Processing dates list' -PercentComplete 0
}
$ProgressVector = 0
If($DatesFile)
{
    If (!$NoProgressBar)
    {
        Write-Progress -Id 1 -Activity "Processing the list of dates" -PercentComplete 0
    }
    Write-Host 'Loading the Dates list file: ' -NoNewline
    Write-Host $DatesFile -ForegroundColor Cyan
    Try
    {
        $DatesListFileContent = Get-Content $DatesFile -ErrorAction SilentlyContinue
        Write-Host '... Content loaded.' -ForegroundColor Green
    }
    Catch
    {
        Write-Host 'Could not load the specified DatesListFile: ' -ForegroundColor Red -NoNewline
        Write-Host $DatesFile -ForegroundColor Cyan
        Write-Host 'Please try again.' -ForegroundColor Yellow
        Break
    }

    Write-Host 'Parsing the dates'
    # Parsing and Quick sanity check on the content
    [int]$NotValidDates = 0
    [int]$ProgressVector = 0
    $script:DatesList = @()
    ForEach($Date in $DatesListFileContent)
    {
        $ProgressVector++
        If (!$NoProgressBar)
        {
            Write-Progress -Id 1 -Activity "Processing the list of dates" -PercentComplete ($ProgressVector / ($($DatesListFileContent.Count))*100) -ParentId 0
        }
        Try
        {
            # If it seems a valid date, we add it to the array.
            $script:DatesList += Get-Date $Date -Format s
        }
        Catch
        {
            # If not, we increment the counter.
            $NotValidDates++
        }
    }
    # Then we validate
    If ($NotValidDates -gt 0)
    {
        Write-Host 'The dates list file specified (' -ForegroundColor Red -NoNewline
        Write-Host $DatesFile -ForegroundColor Cyan -NoNewline
        Write-Host ' contains ' -ForegroundColor Red -NoNewline
        Write-Host $NotValidDates -ForegroundColor Cyan -NoNewline
        Write-Host ' out of ' -ForegroundColor Red -NoNewline
        Write-Host $DatesFile.Count -ForegroundColor Cyan -NoNewline
        Write-Host ' entries that do not seem to be a valid date.'
        Write-Host 'The invalid entries have been filtered out, which leaves the valid dates count to: ' -ForegroundColor Red -NoNewline
        Write-Host $DatesList.Count -ForegroundColor Cyan
        
        If ($dateslist.Count -gt 0)
        {
            $ShouldProceed = Read-Host -Prompt 'Do you want to proceed?'
            If (!($ShouldProceed -eq 'y') -or !($ShouldProceed -eq 'Y'))
            {
                Break
            }
        }
        Else
        {
            Write-Host 'The script cannot continue.' -ForegroundColor Yellow
            Break
        }
    }
    Else
    {
        Write-Host '... All dates are valid.' -ForegroundColor Green
        Write-Host
    }

    # And we do the count
    $script:DatesCount = $DatesList.Count
    # Finally we sort, which will come handy later
    $script:DatesList = $script:DatesList | Sort-Object

    If (!$NoProgressBar)
    {
        Write-Progress -Id 1 -Activity 'Completed' -Completed
    }
}

# Generating files
#Open Excel Application
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation "Creating Excel COM Object" -PercentComplete 0
}
Write-Host 'Creating MS Excel COM object...'
Try
{
    $ExcelApplication = New-Object -ComObject "Excel.Application"
    $ExcelApplication.DisplayAlerts = $false
    Write-Host '... Object created.' -ForegroundColor Green
}
Catch
{
    Write-Host 'Could not create Excel COM Object.' -ForegroundColor Red
    Write-Host 'The script halted.' -ForegroundColor Yellow
}

# If available, use the one in the template
If ($TemplateXLSXFile)
{
    Write-Host 'Opening template document...'
    Try
    {
        $ExcelDocument = $ExcelApplication.Workbooks.Open($TemplateXLSXFile)
        $ExcelSheet = ($ExcelDocument.Sheets)[1]
        Write-Host '... Document opened.' -ForegroundColor Green
    }
    Catch
    {
        Write-Host 'Could not open the provided Template Excel document: ' -ForegroundColor -NoNewline
        Write-Host $TemplateXLSXFile -ForegroundColor Cyan
        Write-Host 'The script halted.' -ForegroundColor Yellow
        Break
    }
}

# Generate the files.
Write-Host 'Generating XLS Files.'
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'File generation' -CurrentOperation "Creating Excel files" -PercentComplete (($MainProgressVector / ($FileCount*2))*100)
}
$TempTXT = $DestinationFolder + '\temp.txt'
For($i=1; $i -le $FileCount;$i++)
{
    $MainProgressVector++
    If (!$NoProgressBar)
    {
        Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'TXT File Generation' -CurrentOperation "$i files generated..." -PercentComplete (($MainProgressVector / ($FileCount*2))*100)
    }

    # First we generate some text
    Try
    {
	    GenerateTXTforXLSX -Columns $Columns -Rows $Rows -TempTXT $TempTXT
    }
    Catch
    {
        Write-Host 'There was an error with the string generation: ' -ForegroundColor Red -NoNewline
        Write-Host $i -ForegroundColor Cyan
    }

    # Then we convert it
    $ExcelFilePath = "$DestinationFolder\$FilePrefix" + "_$i" + ".xlsx"
    ConvertTXTtoXLSX -TempTXT $TempTXT -ExcelFile $ExcelFilePath
}
Write-Host '... Finished generating files.' -ForegroundColor Green


# Then we see if we need ot update the core properties
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Updating core properties' -CurrentOperation '' -PercentComplete (($MainProgressVector / ($FileCount*2))*100)
}
If ($PersonListFile -or $DatesFile)
{
    Write-Host
    Write-Host 'Updating document properties...'
    $ExcelFiles = Get-ChildItem $DestinationFolder -File -Filter "*.xlsx"
    ForEach($ExcelFile in $ExcelFiles)
    {
        $MainProgressVector++
        Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Updating core properties' -CurrentOperation '' -PercentComplete (($MainProgressVector / ($FileCount*2))*100)
        $ExcelFilePath = $ExcelFile.FullName
        UpdateDocXMLProps -XLSXFilePath $ExcelFilePath
    }
    Write-Host '...Done.' -ForegroundColor Green
}



# Close Excel Application and clean up.
Write-Host 'Cleaning up'
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Cleanup' -CurrentOperation '' -PercentComplete 100
}
Write-Host 'Closing MS Excel COM object...'
Try
{
    $ExcelApplication.Workbooks.Close()
    $ExcelApplication.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelApplication)
    $ExcelApplication = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
Catch
{
    Write-Host 'Could not close the MS Excel COM object.' -ForegroundColor Red
}

$EndTime = Get-Date
$TimeTaken = $EndTime - $StartTime

# Dump some statistics on the screen.
Write-Host 'The script took ' -NoNewline
Write-Host $TimeTaken.Days -ForegroundColor Green -NoNewline
Write-Host ' days, ' -NoNewline 
Write-Host $TimeTaken.Hours -ForegroundColor Green -NoNewline
Write-Host " hours, " -NoNewline
Write-Host $TimeTaken.Minutes -ForegroundColor Green -NoNewline
Write-Host " minutes, " -NoNewline
Write-Host $TimeTaken.Seconds -ForegroundColor Green -NoNewline
Write-Host " seconds and " -NoNewline
Write-Host $TimeTaken.Milliseconds -ForegroundColor Green -NoNewline
Write-Host " milliseconds to run."
Write-Host 'It created ' -NoNewline
Write-Host $SumFiles -ForegroundColor Green -NoNewline 
Write-Host ' files.' 
