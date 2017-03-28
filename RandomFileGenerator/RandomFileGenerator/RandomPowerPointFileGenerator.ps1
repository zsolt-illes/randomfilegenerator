[cmdletbinding()]
param 
(
    [Parameter(Mandatory=$true)][string]$DictionaryFile,
    [Parameter(Mandatory=$true)][int]$SlideCount,
	[Parameter(Mandatory=$true)][int]$WordsOnSlide,
	[Parameter(Mandatory=$true)][int]$FileCount,
    [Parameter(Mandatory=$true)][string]$PicturesFolder,
	[Parameter(Mandatory=$true)][string]$DestinationFolder,
	[Parameter(Mandatory=$true)][string]$FilePrefix,
	[Parameter(Mandatory=$false)][string]$PersonListFile,
	[Parameter(Mandatory=$false)][string]$DatesFile,
    [Parameter(Mandatory=$false)][string]$TemplatePPTXFile,
    [Parameter(Mandatory=$false)][switch]$NoProgressBar
)

# Loading and creating the model XML file
Function EnsureCoreXML
{
	[cmdletbinding()]
	param
	(
		[Parameter(Mandatory=$true)][string]$PPTFilePath
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
	    $ZIPFilePath = $PPTFilePath.Replace('pptx','zip')
	    Try
	    {
		    Copy-Item -Path $PPTFilePath -Destination ($ZIPFilePath) -ErrorAction SilentlyContinue
	    }
	    Catch
	    {
		    Write-Host 'Could not copy the Word file: ' -ForegroundColor Red -NoNewline
		    Write-Host $PPTFilePath -ForegroundColor Cyan -NoNewline
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
		[Parameter(Mandatory=$true)][string]$PPTFilePath
	)

    If (!$Script:CoreXML)
	{
        EnsureCoreXML $PPTFilePath
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
        $PPTFile = [System.IO.Compression.ZipFile]::Open( $PPTFilePath, 'Update' )
        $ZippedCoreXMLFile = [System.IO.StreamWriter]($PPTFile.Entries | Where-Object { $_.FullName -match 'docProps/core.xml' }).Open()
        $ZippedCoreXMLFile.BaseStream.SetLength(0)
        $ZippedCoreXMLFile.Write($Script:CoreXML.OuterXml)
        $ZippedCoreXMLFile.Flush()
        $ZippedCoreXMLFile.Close()
        $PPTFile.Dispose()
    }
    Catch
    {
        Write-Host 'There was an error updating the core properties of: ' -ForegroundColor Red -NoNewline
        Write-Host $PPTFilePath -ForegroundColor Cyan
    }
}

####################################################
# This is where the main part of the script starts #
####################################################
$StartTime = Get-Date
cls
$MainProgressVector = 0

$PicturesLeft = 600
$PicturesTop = 200
$PicturesMaxWidth = 290
$PicturesmaxHeight = 280

Cls

# Quick Warning if the Template File is used
If ($TemplatePPTXFile)
{
    Write-Host 'WARNING! ' -ForegroundColor Yellow -NoNewline
    Write-Host 'The TemplatePPTXFile attribute has been used.' -ForegroundColor Red
    Write-Host 'The script assumes that the first slide of the presentation has a Title and a Text field.'
    Write-Host 'If it does not, errors will occur.'
    Write-Host 'Do you want to proceed?' -ForegroundColor Yellow
    $ShouldProceed = Read-Host
    If (!($ShouldProceed -eq 'y') -or !($ShouldProceed -eq 'Y'))
    {
        Break
    }
}

# Loading the Interop assemblies
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Loading assemblies' -PercentComplete 0
}
# For operating with PowerPoint
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
# For operating with Images
Add-Type -AssemblyName System.Windows.Forms
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
Write-Host '... Folder confirmed.' -ForegroundColor Green

# Checking if the Template PowerPoint file exist
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Ensuring template PowerPoint file' -PercentComplete 0
}
If($TemplatePPTXFile)
{
    If (!((Test-Path $TemplatePPTXFile) -and ([System.IO.Path]::IsPathRooted($TemplatePPTXFile))))
    {
        Write-Host 'The TemplatePPTXFile parameter was defined with the value: ' -ForegroundColor Red -NoNewline
        Write-Host $TemplatePPTXFile -ForegroundColor Cyan -NoNewline
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

# Parsing the picture directory
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation 'Processing pictures directory' -PercentComplete $MainProgressVector
}
$NonPictureFiles = 0
If ($PicturesFolder)
{
    Write-Host 'Fetching pictures from the directory: ' -NoNewline
    Write-Host $PicturesFolder -ForegroundColor Cyan
    $Pictures = @()
    If (Test-Path $PicturesFolder) 
    {
        $PictureFiles = Get-ChildItem $PicturesFolder
        ForEach ($PictureFile in $PictureFiles)
        {
            Try
            {
                $Picture = [System.Drawing.Image]::FromFile($PictureFile.FullName)
                $PictureObject = New-Object -TypeName 'System.Object'
                $PictureObject | Add-Member -Name 'FilePath' -Type NoteProperty -Value $PictureFile.FullName
                # Here comes the trick with the picture sizes, we need to convert pixels into points
                $PictureWidth = [math]::Round(($Picture.Width / $Picture.HorizontalResolution * 72),4)
                $PictureHeight = [math]::Round(($Picture.Height / $Picture.VerticalResolution * 72),4)

                # Here we do some the transformation so that we would need to do it every time.
                # you can modify this according to your business needs.
                If ($PictureWidth -gt $PicturesMaxWidth)
                {
                    $Multiplier = $PicturesMaxWidth / $PictureWidth
                    $PictureHeight = $PictureHeight * $Multiplier
                }
                If ($PictureHeight -gt $PicturesmaxHeight)
                {
                    $Multiplier = $PicturesmaxHeight = $PictureHeight
                    $PictureWidth = $PictureWidth * $Multiplier
                }
                # The above will not do anything if the picture size fits into the limit (See the beginning of the Main section.)
                
                $PictureObject | Add-Member -Name 'Width' -Type NoteProperty -Value $PictureWidth
                $PictureObject | Add-Member -Name 'Height' -Type NoteProperty -Value $PictureHeight
                $Pictures += $PictureObject
            }
            Catch
            {
                $NonPictureFiles++
            }
        }

        $PicturesCount = $pictures.Count

        If($PicturesCount -eq $NonPictureFiles)
        {
            Write-Host 'The directory specified in the PicturesFolder variable (' -ForegroundColor Red -NoNewline
            Write-Host $PicturesFolder -ForegroundColor Cyan -NoNewline
            Write-Host ') does not contain any valid picture files.' -ForegroundColor Red -NoNewline
            Write-Host 'The script halted.' -ForegroundColor Yellow
            Break
        }
    }
    Else
    {
        Write-Host 'The directory specified in the PicturesFolder variable (' -ForegroundColor Red -NoNewline
        Write-Host $PicturesFolder -ForegroundColor Cyan -NoNewline
        Write-Host ') is not valid.' -ForegroundColor Red -NoNewline
        Write-Host 'The script halted.' -ForegroundColor Yellow
        Break
    }

    Write-Host '... Finished fetching the pictures.' -ForegroundColor Green
}

# Generating files
#Open PowerPoint Application
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Preparation' -CurrentOperation "Creating PowerPoint COM Object" -PercentComplete (($MainProgressVector / ($FileCount*2))*100)
}

Write-Host 'Creating MS PowerPoint COM object...'
Try
{
    $PPTApplication = New-Object -ComObject "PowerPoint.Application"
    # Some msoTriStates
    $msoTrue = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $msoFalse = [Microsoft.Office.Core.MsoTriState]::msoFalse
    Write-Host '... Object created.' -ForegroundColor Green
}
Catch
{
    Write-Host 'Could not create PowerPoint COM Object.' -ForegroundColor Red
    Write-Host 'The script halted.' -ForegroundColor Yellow
}

# Either create a new presentation or use the one in the template
Write-Host 'Creating presentation...'
If ($TemplatePPTXFile)
{
    $PPTDocument = $PPTApplication.Presentations.Open($TemplatePPTXFile, $msoFalse) # Use this if you want to hide the PPT App
    #$PPTDocument = $PPTApplication.Presentations.Open($TemplatePPTXFile) # Use this if you want to see the PPT App
    $PPTSlideLayout = $PPTDocument.Slides[1].Layout
}
Else
{
    $PPTDocument = $PPTApplication.Presentations.Add($msoFalse) # Use this if you want to hide the PPT App
    #$PPTDocument = $PPTApplication.Presentations.Add() # Use this if you want to see the PPT App
    $PPTSlideLayout = [microsoft.office.interop.powerpoint.ppSlideLayout]::ppLayoutText
}
Write-Host '... Done.' -ForegroundColor Green


# Generating files.
Write-Host
Write-Host 'Generating files...' -NoNewline

# Creating the Animations array:
$SlideAnimationEffects = 257, 258, 513, 769, 770, 1025, 1026, 1281, 1282, 1283, 1284, 1285, 1286, 1287, 1288, 1537, 1793, 2049, 2050, 2051, 2052, 2053, 2054, 2055, 2056, 2305, 2306, 2561, 2562, 2563, 2564, 2565, 2566, 2567, 2568, 2817, 2818, 2819, 2820, 3073, 3074, 3585, 3586, 3587, 3588, 3844, 3845, 3846, 3847, 3848, 3849, 3850, 3851, 3852, 3853, 3854, 3855, 3856, 3857, 3858, 3859, 3860, 3861
$TextAnimationEffects = 257, 258, 513, 769, 770, 1025, 1026, 1281, 1282, 1283, 1284, 1285, 1286, 1287, 1288, 1537, 1793, 2049, 2050, 2051, 2052, 2053, 2054, 2055, 2056, 2305, 2306, 2561, 2562, 2563, 2564, 2565, 2566, 2567, 2568, 2817, 2818, 2819, 2820, 3073, 3074, 3585, 3586, 3587, 3588, 3844

For($FileCounter=1; $FileCounter -le $FileCount; $FileCounter++)
{
    Write-Host '.' -NoNewline
    $MainProgressVector++
    If (!$NoProgressBar)
    {
        Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Creating PowerPoint files' -CurrentOperation "Creating file: $FileCounter" -PercentComplete (($MainProgressVector / ($FileCount*2))*100)
    }
    # Let's clear the existing slides. (On first iteration it is obsolete, but doesn't hurt.)
    If ($PPTDocument.Slides.Count -gt 0)
    {
        $PPTDocument.Slides | % {$_.Delete()}
    }

    # Generate the slides
    For($SlideCounter = 1; $SlideCounter -le $SlideCount; $SlideCounter++)
    {
        # First we add a new slide
        $Slide = $PPTDocument.Slides.Add($SlideCounter,$PPTSlideLayout)

        # Add a random animation to it
        $RandomAnimation = ($SlideAnimationEffects | Get-Random)
        $Slide.SlideShowTransition.EntryEffect = $RandomAnimation 


        # Then fill in the Title
        $TitleText = ''
        $TitleWordCount = Get-Random(1..3)
        For ($i=1; $i -le $TitleWordCount; $i++)
        {
            $RandomWordLine = Get-Random -Minimum 1 -Maximum $DictionaryFileRows
            $RandomWord = $DictionaryFileContent[$RandomWordLine] + ' '
            $TitleText += $RandomWord
        }
        $slide.Shapes.Title.TextFrame.TextRange.Text = $TitleText
        $slide.Shapes.Title.AnimationSettings.TextLevelEffect = 1
        $RandomAnimation = ($TextAnimationEffects | Get-Random)
        $slide.Shapes.Title.AnimationSettings.EntryEffect = $RandomAnimation
        $slide.Shapes.Title.AnimationSettings.AdvanceMode = 2
        $slide.Shapes.Title.AnimationSettings.Animate = $msoTrue

        # Add the rest of the text
        $SlideText = ''
        for ($i=1; $i -le $WordsOnSlide; $i++)
        {
            $RandomWordLine = Get-Random -Minimum 1 -Maximum $DictionaryFileRows
            $RandomWord = $DictionaryFileContent[$RandomWordLine] + "`r`n"
            $SlideText += $RandomWord
        }
        $slide.Shapes.Item(2).TextFrame2.Column.Number = 3 # This is where you might wish to change the number of the item if you have the text field on a different index
        $slide.Shapes.Item(2).TextFrame.TextRange.Text = $SlideText # This is where you might wish to change the number of the item if you have the text field on a different index
        $slide.Shapes.Item(2).AnimationSettings.TextLevelEffect = 1
        $RandomAnimation = ($TextAnimationEffects | Get-Random)
        $slide.Shapes.Item(2).AnimationSettings.EntryEffect = $RandomAnimation
        $slide.Shapes.Item(2).AnimationSettings.AdvanceMode = 2
        $slide.Shapes.Item(2).AnimationSettings.Animate = $msoTrue


        # Add a picture randomly
        If(Get-Random($true,$false))
        {
            If ($PicturesCount -eq 1)
            {
                $RandomPicture = $Pictures[0]
            }
            Else
            {
                $RandomPictureNumber = Get-Random -Minimum 0 -Maximum ($PicturesCount - 1)
                $RandomPicture = $Pictures[$RandomPictureNumber]
            }

            $pic = $slide.Shapes.AddPicture(($RandomPicture.FilePath),$msoFalse, $msoTrue,$PicturesLeft,$PicturesTop, $RandomPicture.Width, $RandomPicture.Height)
            $pic.ZOrder([Microsoft.Office.Core.MsoZOrderCmd]::msoSendToBack) # Just putting it behind the text
        }

    }

    # Save the file
    $PPTFileName = "$DestinationFolder\$FilePrefix" + '_' + $FileCounter + '.pptx'
    $PPTDocument.SaveAs($PPTFileName)

    # And finally clear the slides to make sure we can create a new set of them.
    $PPTDocument.slides | %{$_.Delete()}
}
Write-Host
Write-Host '... Files created.' -ForegroundColor Green

# Close the PowerPoint application
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Cleanup' -CurrentOperation '' -PercentComplete (($MainProgressVector / ($FileCount*2))*100)
}
Write-Host 'Closing MS PowerPoint COM object...'
Try
{
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    $PPTDocument.Saved = $msoTrue
    $PPTDocument.Close()
    $null = [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($PPTDocument)
    $PPTDocument = $null
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    $PPTApplication.quit()
    $null = [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($PPTApplication)
    $PPTApplication = $null
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
}
Catch
{
    Write-Host 'Could not close the MS PowerPoint COM object.' -ForegroundColor Red
}

# Then we see if we need ot update the core properties
If (!$NoProgressBar)
{
    Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Updating core properties' -CurrentOperation '' -PercentComplete (($MainProgressVector / ($FileCount*2))*100)
}
If ($PersonListFile -or $DatesFile)
{
    Write-Host
    Write-Host 'Updating document properties...'
    $PPTFiles = Get-ChildItem $DestinationFolder -File -Filter "*.pptx"
    ForEach($PPTFile in $PPTFiles)
    {
        $ProgressVector++
        Write-Progress -Id 0 -Activity "Generating $FileCount files" -Status 'Updating core properties' -CurrentOperation '' -PercentComplete (($MainProgressVector / ($FileCount*2))*100)
        $PPTFilePath = $PPTFile.FullName
        UpdateDocXMLProps -PPTFilePath $PPTFilePath
    }
    Write-Host '...Done.' -ForegroundColor Green
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