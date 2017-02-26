####################################################################
# This script is to generate n number of files for stress testing. #
# It is using GUIDs to generate the files of different sizes.      #
####################################################################
[cmdletbinding()]
param 
(
    [parameter(mandatory=$true)][string]$DestinationFolder,
    [parameter(mandatory=$false)][switch]$CreateFoldersForFiles,
    [parameter(mandatory=$true)][string]$FilePrefix,
    [parameter(mandatory=$false)][int]$_5kFiles,
    [parameter(mandatory=$false)][int]$_10kFiles,
    [parameter(mandatory=$false)][int]$_20kFiles,
    [parameter(mandatory=$false)][int]$_50kFiles,
    [parameter(mandatory=$false)][int]$_100kFiles,
    [parameter(mandatory=$false)][int]$_200kFiles,
    [parameter(mandatory=$false)][int]$_500kFiles,
    [parameter(mandatory=$false)][int]$_1MBFiles,
    [parameter(mandatory=$false)][int]$_2MBFiles,
    [parameter(mandatory=$false)][int]$_5MBFiles,
    [parameter(mandatory=$false)][int]$_10MBFiles,
    [parameter(mandatory=$false)][int]$_20MBFiles,
    [parameter(mandatory=$false)][int]$_50MBFiles,
    [parameter(mandatory=$false)][int]$_100MBFiles,
    [parameter(mandatory=$false)][int]$_200MBFiles,
    [parameter(mandatory=$false)][int]$_500MBFiles,
    [parameter(mandatory=$false)][int]$_1GBFiles,
    [parameter(mandatory=$false)][int]$CustomSize,
    [parameter(mandatory=$false)][int]$CustomCount,
    [parameter(mandatory=$false)][switch]$NoGUIDs,
    [parameter(mandatory=$false)][switch]$NoProgressBar
)

# A little math...
# 1 GUID = 36 characters, 37 with separating space (this is what we use)
#5KB file = 5*1024 characters = 138 GUIDs + 14 characters.
#10KB file = 10*1024 characters = 276 GUIDs + 28 characters.
#20KB file = 20*1024 characters = 553 GUIDs + 19 characters.
#50KB file = 50*1024 characters = 1383 GUIDs + 29 characters.
#100KB file = 100*1024 characters = 2767 GUIDs + 21 characters.
#200KB file = 200*1024 characters = 5535 GUIDs + 5 characters.
#500KB file = 500*1024 characters = 13837 GUIDs + 31 characters.
#1MB file = 1*1048576 characters = 28339 GUIDs + 33 characters.
#2MB file = 2*1048576 characters = 56679 GUIDs + 29 characters.
#5MB file = 5*1048576 characters = 141699 GUIDs + 17 characters.
#10MB file = 10*1048576 characters = 283398 GUIDs + 34 characters.
#20MB file = 20*1048576 characters = 566797 GUIDs + 31 characters.
#50MB file = 50*1048576 characters = 1416994 GUIDs + 22 characters.
#100MB file = 100*1048576 characters = 2833989 GUIDs + 7 characters.
#200MB file = 200*1048576 characters = 5667978 GUIDs + 14 characters.
#500MB file = 500*1048576 characters = 14169945 GUIDs + 35 characters.
#1GB file = 1*1073741824 characters = 29020049 GUIDs + 11 characters.



# Function to check if we have enough free space on the destination drive.
# We do not want to write the disk full...
function CheckFreeSpace
{
    [cmdletbinding()]
    param
    (
        [parameter(mandatory=$true)][string]$DestinationFolder,
        [parameter(mandatory=$true)][double]$FilesFullSize
    )

    # Get the Disk information
    If($DestinationFolder.StartsWith('\\'))
    {
        $Disk = Get-WmiObject Win32_LogicalDisk | ?{($_.ProviderName) -and ($DestinationFolder -like "*$($_.ProviderName)*")}
    }
    Else
    {
        $DestinationDrive = $DestinationFolder | Split-Path -Qualifier
        $Disk = Get-WmiObject Win32_LogicalDisk -Filter $("DeviceID='" + $DestinationDrive + "'")
    }

    # Prepare some variables
    $DiskFreeSpaceFormatted = "{0:N0}" -f $disk.FreeSpace
    $FreeSpaceAfter = $Disk.FreeSpace - $FilesFullSize
    $FreeSpaceAfterFormatted = "{0:N0}" -f $FreeSpaceAfter
    # Needed for crash dump
    $PhysicalMemory = (((Get-CimInstance -ClassName "cim_physicalmemory").Capacity) | Measure-Object -Sum).Sum

    Write-Host 'The free space on the destination drive: ' -NoNewline
    Write-Host $DiskFreeSpaceFormatted -ForegroundColor Cyan -NoNewline
    Write-Host ' bytes.'
    Write-Host 'The free space on the destination drive after the script ran: ' -NoNewline
    Write-Host $FreeSpaceAfterFormatted -ForegroundColor Cyan -NoNewline
    Write-Host ' bytes.'

    # We have to be extra cautious with the system drive 
    If (($DestinationDrive -eq $env:SystemDrive) -and ($FreeSpaceAfter -le $PhysicalMemory))
    {
        Write-Host 'System drive is used and the free space after the files created is less than the Physical Memory of the machine' -ForegroundColor Red
        Write-Host 'To protect your environment, the script will not continue.' -ForegroundColor Red
        Write-Host 'Please choose a different destination.' -ForegroundColor Yellow
        Break
    }
    ElseIf ($FreeSpaceAfter -le 0)
    {
        Write-Host 'The destination drive does not have enough free space to generate the files.' -ForegroundColor Red
        Write-Host 'To protect your environment, the script will not continue.' -ForegroundColor Red
        Write-Host 'Please choose a different destination.' -ForegroundColor Yellow
        Break
    }

}

# A function to calculate the sum size of the files requested
function CalculateSumSize
{
    $FullSize = 0
    
    $FullSize += $_5kFiles * 5 * 1kb
    $FullSize += $_10kFiles * 10 * 1kb
    $FullSize += $_20kFiles * 20 * 1kb
    $FullSize += $_50kFiles * 50 * 1kb
    $FullSize += $_100kFiles * 100 * 1kb
    $FullSize += $_200kFiles * 200 * 1kb
    $FullSize += $_500kFiles * 500 * 1kb
    $FullSize += $_1MBFiles * 1 * 1mb
    $FullSize += $_2MBFiles * 2 * 1mb
    $FullSize += $_5MBFiles * 5 * 1mb
    $FullSize += $_10MBFiles * 10 * 1mb
    $FullSize += $_20MBFiles * 20 * 1mb
    $FullSize += $_50MBFiles * 50 * 1mb
    $FullSize += $_100MBFiles * 100 * 1mb
    $FullSize += $_200MBFiles * 200 * 1mb
    $FullSize += $_500MBFiles * 500 * 1mb
    $FullSize += $_1GBFiles * 1 * 1gb

    Return $FullSize
}


# Just a function we're going to call to generate each file
function GenerateFile
{
    [cmdletbinding()]
    param
    (
        [parameter(mandatory=$true)][string]$FullFilePath,
        [parameter(mandatory=$true)][int]$GuidsToGenerate,
        [parameter(mandatory=$true)][int]$RandomCharNum
    )

    # Open a File Stream for better performance
    $FileStream = New-Object System.IO.FileStream($FullFilePath,[System.IO.FileMode]::CreateNew)
    $StreamWriter = New-Object System.IO.StreamWriter($FileStream,[System.Text.Encoding]::ASCII,128)

    # Just a pointer for the actual file creation
    $FileContentPointer = 1

    For ($i=1; $i -le $GuidsToGenerate; $i++)
    {
        # This is the progress bar for the actual file
        If(!$NoProgressBar)
        {
            Write-Progress -Id 3 -ParentId 2 -Activity "Generating file: $FullFilePath" -Status '...' -PercentComplete ($FileContentPointer/$GuidsToGenerate*100)
        }
        If($NoGUIDs)
        {
            $StringToDump = -join ((65..90) + (97..122) | Get-Random -Count 37 | %{[char]$_})
        }
        Else
        {
            $GUID = ([guid]::NewGuid()).ToString()
            $StringToDump = $GUID.Replace('-','_') + ' '
        }

        $StreamWriter.Write($StringToDump)
        $FileContentPointer++
    }

    $RandomChars = -join ((65..90) + (97..122) | Get-Random -Count $RandomCharNum | %{[char]$_})

    $StreamWriter.Close()
    $FileStream.Close()
}

# This function is to calculate the number of GUIDs that fit into the random size files
function GetGUIDCount
{
    [cmdletbinding()]
    param
    (
        [parameter(mandatory=$true)][int]$CustomFileSize
    )

    $GUIDs = [Math]::Truncate($CustomFileSize / 37)
    $Remainder = $CustomFileSize - ($GUIDs * 37)


    $GUIDCount = New-Object System.Object
    $GUIDCount | Add-Member -Type NoteProperty -Name 'GUIDs' -Value $GUIDs
    $GUIDCount | Add-Member -Type NoteProperty -Name 'Remainder' -Value $Remainder

    return $GUIDCount
}

# Main part of the script
# We're going to iterate through each size and call the GenerateFile function
$StartTime = Get-Date
Clear-Host
$Error.Clear()

# We need this, because the progress bar will take the top of the window
For ($i=1; $i -le 15; $i++)
{
    Write-Host
}

#Calculate the size of all the files combined.
$FilesFullSize = CalculateSumSize
$FilesFullSizeFormatted = "{0:N0}" -f $FilesFullSize
Write-Host 'The total size of the files requested = ' -NoNewline
Write-Host $FilesFullSizeFormatted -ForegroundColor Cyan -NoNewline
Write-Host ' bytes.'

$DestinationFolder = $DestinationFolder.TrimEnd('\')

# Check if we actually fit into the drive
CheckFreeSpace -DestinationFolder $DestinationFolder -FilesFullSize $FilesFullSize

# Create a variable for statistics
$SumFiles = 0

# Now iterate through each parameter.
$ScriptName = $PSCmdlet.MyInvocation.InvocationName
$ParameterList = (Get-Command -Name $ScriptName).Parameters

ForEach ($Parameter in $ParameterList)
{
    $ParamMatrix = Get-Variable -Name $Parameter.Values.Name -ErrorAction SilentlyContinue
}

# Just a pointer for a progress bar
$ParameterPointer = 1

ForEach ($Parameter in $ParamMatrix)
{
    # This is the progress bar for the parameters
    If(!$NoProgressBar)
    {
        Write-Progress -Id 1 -Activity 'Processing parameters' -Status "Parameter: $($Parameter.Name)" -PercentComplete ($ParameterPointer/$($ParamMatrix.Count)*100)
    }

    # We have to execute only those parameters where the value is greater than 0
    If ((($Parameter.Name -like "_*") -and ($Parameter.Value -gt 0))-or (($Parameter.Name -eq 'CustomSize') -and $CustomCount -gt 0))
    {
        # Increment the sum number of files variable
        $SumFiles += $parameter.Value

        # Create a folder if necessary
        If ($CreateFoldersForFiles)
        {
            $DestinationFolderForFile = "$DestinationFolder\$($Parameter.Name)"
            mkdir -Path $DestinationFolderForFile -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Else
        {
            $DestinationFolderForFile = $DestinationFolder
        }
        
        # This is where we use the math from the beginning of the script
        Switch ($($Parameter.Name))
        {
            '_5kFiles'
                {
                    $GuidsToGenerate = 138
                    $RandomCharNum = 14
                }
            '_10kFiles'
                {
                    $GuidsToGenerate = 276
                    $RandomCharNum = 28
                }
            '_20kFiles'
                {
                    $GuidsToGenerate = 553
                    $RandomCharNum = 19
                }
            '_50kFiles'
                {
                    $GuidsToGenerate = 1383
                    $RandomCharNum = 29
                }
            '_100kFiles'
                {
                    $GuidsToGenerate = 2767
                    $RandomCharNum = 21
                }
            '_200kFiles'
                {
                    $GuidsToGenerate = 5535
                    $RandomCharNum = 5
                }
            '_500kFiles'
                {
                    $GuidsToGenerate = 13837
                    $RandomCharNum = 31
                }
            '_1MBFiles'
                {
                    $GuidsToGenerate = 28339
                    $RandomCharNum = 33
                }
            '_2MBFiles'
                {
                    $GuidsToGenerate = 56679
                    $RandomCharNum = 29
                }
            '_5MBFiles'
                {
                    $GuidsToGenerate = 141699
                    $RandomCharNum = 17
                }
            '_10MBFiles'
                {
                    $GuidsToGenerate = 283398
                    $RandomCharNum = 34
                }
            '_20MBFiles'
                {
                    $GuidsToGenerate = 566797
                    $RandomCharNum = 31
                }
            '_50MBFiles'
                {
                    $GuidsToGenerate = 1416994
                    $RandomCharNum = 22
                }
            '_100MBFiles'
                {
                    $GuidsToGenerate = 2833989
                    $RandomCharNum = 7
                }
            '_200MBFiles'
                {
                    $GuidsToGenerate = 5667978
                    $RandomCharNum = 14
                }
            '_500MBFiles'
                {
                    $GuidsToGenerate = 14169945
                    $RandomCharNum = 35
                }
            '_1GBFiles'
                {
                    $GuidsToGenerate = 29020049
                    $RandomCharNum = 11
                }
            'CustomSize'
                {
                    $GUIDCount = GetGUIDCount $CustomSize
                    $GuidsToGenerate = $GUIDCount.GUIDs
                    $RandomCharNum = $GUIDCount.Remainder
                }
        }

        # Just a pointer for the actual files
        $FileGeneratorPointer = 1

        If ($Parameter.Name -eq 'CustomSize')
        {
            $FilesToGenerate = $CustomCount
        }
        Else
        {
            $FilesToGenerate = $Parameter.Value
        }

        # Now make an iteration on the number of files was requested
        For ($i=1; $i -le $FilesToGenerate; $i++)
        {
            # This is the progress bar for the file generation
            If(!$NoProgressBar)
            {
                Write-Progress -Id 2 -ParentId 1 -Activity 'Generating files...' -Status "FileCount: $i of $($Parameter.Value)" -PercentComplete ($FileGeneratorPointer/$($Parameter.Value)*100)
            }

            If ($Parameter.Name -eq 'CustomSize')
            {
                $FileName = "$FilePrefix" + '_' + $($Parameter.Name) + "_" + "$i.txt"
            }
            Else
            {
                $FileName = "$FilePrefix" + $($Parameter.Name) + "_" + "$i.txt"
            }
            $FullFilePath = "$DestinationFolderForFile\$FileName"
            GenerateFile -FullFilePath $FullFilePath -GuidsToGenerate $GuidsToGenerate -RandomCharNum $RandomCharNum

            $FileGeneratorPointer++
        }

    }

    $ParameterPointer++
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
Write-Host " minutes and " -NoNewline
Write-Host $TimeTaken.Milliseconds -ForegroundColor Green -NoNewline
Write-Host " milliseconds to run."
Write-Host 'It created ' -NoNewline
Write-Host $SumFiles -ForegroundColor Green -NoNewline 
Write-Host ' files.' 
