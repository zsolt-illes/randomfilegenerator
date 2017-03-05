[cmdletbinding()]
param 
(
    [Parameter(Mandatory=$true)][string]$DictionaryFile,
	[Parameter(Mandatory=$true)][int]$WordCount,
	[Parameter(Mandatory=$true)][int]$FileCount,
	[Parameter(Mandatory=$true)][string]$DestinationFolder,
	[Parameter(Mandatory=$true)][string]$FilePrefix,
	[Parameter(Mandatory=$false)][string]$SeparatorWeightFile,
	[Parameter(Mandatory=$false)][string]$PersonListFile,
	[Parameter(Mandatory=$false)][string]$DatesFile,
    [Parameter(Mandatory=$false)][string]$TemplateDOCXFile,
    [Parameter(Mandatory=$false)][switch]$NoProgressBar
)

$MainProgressVector = 0

# A function to generate the TXT Files
function GenerateTXT
{
	[cmdletbinding()]
	param 
	(
		[Parameter(Mandatory=$true)][int]$WordCount
	)

    # Have a string to store our text
    [string]$LongString = ''

	# We start with a GUID so that every file would be unique.
	$GUID = ([guid]::NewGuid()).ToString()
    $GUIDToDump = $GUID.Replace('-','_') + ' '
    $LongString = $LongString + $GUIDToDump

	# Then we generate as many words as requested.
	For($i=1; $i -le $WordCount; $i++)
	{
        If (!$NoProgressBar)
        {
            Write-Progress -Id 1 -Activity 'Generating text...' -PercentComplete (($i/$WordCount)*100)
        }
		$RandomWordLine = Get-Random -Minimum 1 -Maximum $DictionaryFileRows
		$RandomWord = $DictionaryFileContent[$RandomWordLine]
        If ($Separators.Count -eq 1)
        {
            $RandomSeparator = $Separators[0]
        }
        Else
        {
            $RandomSeparator = $Separators[(Get-Random -Minimum 0 -Maximum $($Separators.Count))]
        }
        $RandomWord = $RandomWord + $RandomSeparator
		$LongString = $LongString + $RandomWord
	}

    Return $LongString

}

# A function to generate docx from the TXT file
function ConvertTXTToDocX
{
	[cmdletbinding()]
	param
	(
		[Parameter(Mandatory=$true)][string]$DocxFile,
        [Parameter(Mandatory=$true)][string]$StringToInsert
	)

	$WordSelection = $WordApplication.Selection
    $WordSelection.Select()
    $null = $WordSelection.Delete()
	$WordSelection.InsertAfter($StringToInsert)

	$WordDocument.SaveAs([REF]"$DocxFile")

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
	    $ZIPFilePath = $WordFilePath.Replace('docx','zip')
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
		[Parameter(Mandatory=$true)][string]$WordFilePath
	)

    If (!$Script:CoreXML)
	{
        EnsureCoreXML $WordFilePath
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
        $WordFile = [System.IO.Compression.ZipFile]::Open( $WordFilePath, 'Update' )
        $ZippedCoreXMLFile = [System.IO.StreamWriter]($WordFile.Entries | Where-Object { $_.FullName -match 'docProps/core.xml' }).Open()
        $ZippedCoreXMLFile.BaseStream.SetLength(0)
        $ZippedCoreXMLFile.Write($Script:CoreXML.OuterXml)
        $ZippedCoreXMLFile.Flush()
        $ZippedCoreXMLFile.Close()
        $WordFile.Dispose()
    }
    Catch
    {
        Write-Host 'There was an error updating the core properties of: ' -ForegroundColor Red -NoNewline
        Write-Host $WordFilePath -ForegroundColor Cyan
    }
}

####################################################
# This is where the main part of the script starts #
####################################################
$StartTime = Get-Date
Cls

# Loading the Word Interop assembly
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Loading assemblies' -PercentComplete 0
}
# For operating with Word
Add-Type -AssemblyName Microsoft.Office.Interop.Word
# For operating ZIP files
$null = [System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')

# Making sure the Destination folder is not having a training backslash
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

# Checking if the Template Word file exist
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Ensuring template Word file' -PercentComplete 0
}
If($TemplateDOCXFile)
{
    If (!((Test-Path $TemplateDOCXFile) -and ([System.IO.Path]::IsPathRooted($TemplateDOCXFile))))
    {
        Write-Host 'The TemplateDOCXFile parameter was defined with the value: ' -ForegroundColor Red -NoNewline
        Write-Host $TemplateDOCXFile -ForegroundColor Cyan -NoNewline
        Write-Host ' but could not be found, or is using relative path. Please try again.' -ForegroundColor Red
        Write-Host 'The script Halted.' -ForegroundColor Yellow
        Break
    }
}

# Loading the dictionary file
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Loading Dictionary File' -PercentComplete 0
}
Write-Host 'Loading the dictionary file: ' -NoNewline
Write-Host $DictionaryFile -ForegroundColor Cyan
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

# Trying to load the Separator Weight file
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Preparing Separators' -PercentComplete 0
}
$Separators = @()
If (($SeparatorWeightFile) -and (Test-Path $SeparatorWeightFile -ErrorAction SilentlyContinue))
{
    Write-Host 'Loading the dictionary file: ' -NoNewline
    Write-Host $SeparatorWeightFile -ForegroundColor Cyan

    Try
    {
	    $Separators = Get-Content $SeparatorWeightFile
        Write-Host '... The file loaded.' -ForegroundColor Green
        Write-Host
    }
    Catch
    {
        Write-Host 'Cannot access the specified separator weight file: ' -ForegroundColor Red -NoNewline
        Write-Host $SeparatorWeightFile
        Write-Host 'Please try again.'
        Break
    }
}
Else
{
	$Separators += ' '
    Write-Host 'No separator weight file was specified. Words will be separated by space.' -ForegroundColor Gray
    Write-Host
}

# Trying to load the PersonListFile
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
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Processing dates list' -PercentComplete 0
}
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
#Open Word Application
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation "Creating Word COM Object" -PercentComplete (($ProgressVector / ($FileCount*2))*100)
}

Write-Host 'Creating MS Word COM object...'
Try
{
    $WordApplication = New-Object -ComObject "Word.Application"
    Write-Host '... Object created.' -ForegroundColor Green
}
Catch
{
    Write-Host 'Could not create Word COM Object.' -ForegroundColor Red
    Write-Host 'The script halted.' -ForegroundColor Yellow
}

# Either create a new document or use the one in the template
Write-Host 'Creating document...'
If ($TemplateDOCXFile)
{
    $WordDocument = $WordApplication.Documents.Open($TemplateDOCXFile)
}
Else
{
    $WordDocument = $WordApplication.Documents.Add()
}


$ProgressVector = 0
For($i=1; $i -le $FileCount;$i++)
{
    $ProgressVector++
    If (!$NoProgressBar)
    {
        Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'TXT File Generation' -CurrentOperation "$i files generated..." -PercentComplete (($ProgressVector / ($FileCount*2))*100)
    }

    Try
    {
	    $StringToDocx = GenerateTXT -WordCount $WordCount
    }
    Catch
    {
        Write-Host 'There was an error with the string generation: ' -ForegroundColor Red -NoNewline
        Write-Host $i -ForegroundColor Cyan
    }

    $destinationFile = "$DestinationFolder\$FilePrefix" + "_$i" + '.docx'

    ConvertTXTToDocX -DocxFile $destinationFile -StringToInsert $StringToDocx

}
Write-Host
Write-Host '... Files created.' -ForegroundColor Green


# Close Word Application anc clean up.
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Cleanup' -CurrentOperation '' -PercentComplete (($ProgressVector / ($FileCount*2))*100)
}
Write-Host 'Closing MS Word COM object...'
Try
{
    $WordApplication.Quit() 
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordApplication)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
Catch
{
    Write-Host 'Could not close the MS Word COM object.' -ForegroundColor Red
}


# Then we see if we need ot update the core properties
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Updating core properties' -CurrentOperation '' -PercentComplete (($ProgressVector / ($FileCount*2))*100)
}
If ($PersonListFile -or $DatesFile)
{
    Write-Host
    Write-Host 'Updating document properties...'
    $WordFiles = Get-ChildItem $DestinationFolder -File -Filter "*.docx"
    ForEach($WordFile in $WordFiles)
    {
        $ProgressVector++
        Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Updating core properties' -CurrentOperation '' -PercentComplete (($ProgressVector / ($FileCount*2))*100)
        $WordFilePath = $WordFile.FullName
        UpdateDocXMLProps -WordFilePath $WordFilePath
    }
    Write-Host '...Done.' -ForegroundColor Green
}


# And finaly some cleanup, in case we used the CoreXML file
If ($script:CoreXMLCreated)
{
    Remove-Item -Path "$DestinationFolder\XML" -Recurse -Force -ErrorAction SilentlyContinue
}

$EndTime = Get-Date
$TimeTaken = $EndTime - $StartTime

Write-Host
Write-Host 'The script finished.'
Write-Host 'It took ' -NoNewline
Write-Host $TimeTaken.Days -ForegroundColor Cyan -NoNewline
Write-Host ' days, ' -NoNewline
Write-Host $TimeTaken.Hours -ForegroundColor Cyan -NoNewline
Write-Host ' hours, ' -NoNewline
Write-Host $TimeTaken.Minutes -ForegroundColor Cyan -NoNewline
Write-Host ' minutes, and ' -NoNewline
Write-Host $TimeTaken.Seconds -ForegroundColor Cyan -NoNewline
Write-Host ' seconds to create ' -NoNewline
Write-Host $FileCount -ForegroundColor Cyan -NoNewline
Write-Host ' file(s).'