[cmdletbinding()]
param 
(
    [Parameter(Mandatory=$true)][string]$DictionaryFile,
	[Parameter(Mandatory=$true)][int]$WordCount,
	[Parameter(Mandatory=$true)][int]$FileCount,
	[Parameter(Mandatory=$true)][string]$DestinationFolder,
	[Parameter(Mandatory=$true)][string]$FilePrefix,
	[Parameter(Mandatory=$false)][string]$SeparatorWeightFile,
    [Parameter(Mandatory=$false)][switch]$OnlyLoadFunctions
)

# A function to generate the TXT Files
function GenerateTXT
{
	[cmdletbinding()]
	param 
	(
		[Parameter(Mandatory=$true)][string]$TXTFile,
		[Parameter(Mandatory=$true)][int]$WordCount
	)

	# Open a file stream to dump stuff into the file directly
	$FileStream = New-Object System.IO.FileStream($TXTFile,[System.IO.FileMode]::CreateNew)
    $StreamWriter = New-Object System.IO.StreamWriter($FileStream,[System.Text.Encoding]::ASCII,128)

	# We start with a GUID so that every file would be unique.
	$GUID = ([guid]::NewGuid()).ToString()
    $GUIDToDump = $GUID.Replace('-','_') + ' '
	$StreamWriter.Write($GUIDToDump)

	# Then we generate as many words as requested.
	For($i=1; $i -le $WordCount; $i++)
	{
		$RandomWordLine = Get-Random -Minimum 1 -Maximum $DictionaryFileRows
		$RandomWord = $DictionaryFileContent[$RandomWordLine]
        If ($Separators.Count -eq 1)
        {
            $RandomSeparator = $Separators[0]
        }
        Else
        {
            $RandomSeparator = $Separators[(Get-Random -Minimum 0 -Maximum ($($Separators.Count) -1))]
        }
        $RandomWord = $RandomWord + $RandomSeparator
		$StreamWriter.Write($RandomWord)
	}

	# And finally close the file stream.
	$StreamWriter.Close()
    $FileStream.Close()
}

# A function to generate docx from the TXT file
function ConvertTXTToDocX
{
	[cmdletbinding()]
	param
	(
		[Parameter(Mandatory=$true)][string]$TXTFile
	)

	$WordDocument = $WordApplication.Documents.Add()
	$WordSelection = $WordApplication.Selection
	$WordSelection.InsertFile($TXTFile)

	$WordFilePath = [System.IO.Path]::GetDirectoryName($TXTFile)
	$WordFileName = [System.IO.Path]::GetFileNameWithoutExtension($TXTFile) + '.docx'
	$WordFilePath = "$WordFilePath\$WordFileName"

	$WordDocument.SaveAs([REF]"$WordFilePath")
	$WordDocument.Close()

    Remove-Item $TXTFile

}

####################################################
# This is where the main part of the script starts #
####################################################
$StartTime = Get-Date

If (!$OnlyLoadFunctions)
{
    # Loading the Word Interop assembly
    Add-Type -AssemblyName Microsoft.Office.Interop.Word

    # Loading the dictionary file
    If (Test-Path $DictionaryFile)
    {
	    $DictionaryFileContent = Get-Content $DictionaryFile
    }
    Else
    {
	    Write-Host 'The dictionary file specified (' -NoNewline -ForegroundColor Red
	    Write-Host $DictionaryFile -NoNewline -ForegroundColor Cyan
	    Write-Host ') cannot be found.' -ForegroundColor Red
	    Write-Host 'Correct the missing information and try again.' -ForegroundColor Yellow
	    Break
    }

    # Trying to load the Separator Weight file
    $Separators = @()
    If (($SeparatorWeightFile) -and (Test-Path $SeparatorWeightFile -ErrorAction SilentlyContinue))
    {
	    $Separators = Get-Content $SeparatorWeightFile
    }
    Else
    {
	    $Separators += ' '
    }

    # Get the lines in the dictionary file. 
    #We're going to use this to choose the random words from the file.
    $DictionaryFileRows = ($DictionaryFileContent | Measure-Object -Line).Lines

    # Quick sanity check if the dictionary file is not line separeted
    If($DictionaryFileRows -eq 1)
    {
	    Write-Host 'The dictionary file specified (' -NoNewline -ForegroundColor Red
	    Write-Host $DictionaryFile -NoNewline -ForegroundColor Cyan
	    Write-Host ') is not line separated. Probably it is a CSV.' -ForegroundColor Red
	    Write-Host 'This script can only operate with line separeated dictionary file.' -ForegroundColor Yellow
	    Break
    }

    #Open Word Application
    $WordApplication = New-Object -ComObject "Word.Application"

    # Making sure the Destination folder is not having a training backslash
    $DestinationFolder = $DestinationFolder.TrimEnd('\')
    If (!(Test-Path $DestinationFolder))
    {
        $result = mkdir -Path $DestinationFolder -Force
    }


    # Generating files
    For($i=1; $i -le $FileCount;$i++)
    {
	    $destinationFile = "$DestinationFolder\$FilePrefix" + "_$i" + '.txt'
	    GenerateTXT -TXTFile $destinationFile -WordCount $WordCount
    }

    # Convert all TXT files to DocX
    $TXTFiles = Get-ChildItem $DestinationFolder -Filter "*.txt"



    ForEach($TXTFile in $TXTFiles)
    {
	    ConvertTXTToDocX $($TXTFile.FullName)
    }

    # Close Word Application anc clean up.
    $WordApplication.Quit() 
	[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordApplication)
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
}

$EndTime = Get-Date
$TimeTaken = $EndTime - $StartTime

$TimeTaken