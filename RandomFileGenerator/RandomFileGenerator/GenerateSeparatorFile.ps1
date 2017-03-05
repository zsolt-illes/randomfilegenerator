[cmdletbinding()]
param 
(
    [Parameter(Mandatory=$true)][string]$SeparatorFile,
	[Parameter(Mandatory=$false)][decimal]$Space=87.58,
	[Parameter(Mandatory=$false)][decimal]$FullStop=3.75,
	[Parameter(Mandatory=$false)][decimal]$Comma=3.97,
	[Parameter(Mandatory=$false)][decimal]$Colon=0.24,
	[Parameter(Mandatory=$false)][decimal]$SemiColon=0.18,
	[Parameter(Mandatory=$false)][decimal]$Apostrophe=1.66,
    [Parameter(Mandatory=$false)][decimal]$Exclamation=0.04,
	[Parameter(Mandatory=$false)][decimal]$Question=0.11,
	[Parameter(Mandatory=$false)][decimal]$DoubleQuote=0.7,
	[Parameter(Mandatory=$false)][decimal]$Hyphen=1.76,
	[Parameter(Mandatory=$false)][decimal]$Underscore=0.01
)

############################
# Main script starts here. #
############################

# Get the parameter list
$ScriptName = $PSCmdlet.MyInvocation.InvocationName
$ParameterList = (Get-Command -Name $ScriptName).Parameters

ForEach ($Parameter in $ParameterList)
{
    $ParamMatrix = Get-Variable -Name $Parameter.Values.Name -ErrorAction SilentlyContinue
}

# Check if the provided separator ratio adds up a hundred.
[decimal]$SumOfSeparators = ($ParamMatrix.value | ?{$_.Gettype().Name -eq 'Decimal'} | Measure-Object -Sum).Sum
If ($SumOfSeparators -ne 100)
{
    Write-Host "The sum of separator values ($SumOfSeparators) is not equal to 100." -ForegroundColor Red
    Write-Host 'Please run the script again with the correct numbers.' -ForegroundColor Yellow
    Break
}

Try
{
    # Open a File Stream for better performance
    $FileStream = New-Object System.IO.FileStream($SeparatorFile,[System.IO.FileMode]::CreateNew)
    $StreamWriter = New-Object System.IO.StreamWriter($FileStream,[System.Text.Encoding]::ASCII,128)

    ForEach ($Parameter in $ParamMatrix)
    {
	    $Separator = ''

        If ((($Parameter.Value).GetType().Name -ne 'String') -and ($Parameter.Value -gt 0))
        {
            Switch ($($Parameter.Name))
            {
                'Space'
                    {
                        $Separator = ' '
                    }
	    		'FullStop'
                    {
                        $Separator = '.'
                    }
			    'Comma'
                    {
                        $Separator = ','
                    }
			    'Colon'
                    {
                        $Separator = ':'
                    }
			    'SemiColon'
                    {
                        $Separator = ';'
                    }
			    'Apostrophe'
                    {
                        $Separator = "'"
                    }
			    'Exclamation'
                    {
                        $Separator = '!'
                    }
			    'Question'
                    {
                        $Separator = '?'
                    }
			    'DoubleQuote'
                    {
                        $Separator = '"'
                    }
			    'Hyphen'
                    {
                        $Separator = '-'
                    }
			    'UnderScore'
                    {
                        $Separator = '_'
                    }
		    }

            For ($i=1; $i -le $($Parameter.Value)*100; $i++)
		    {
                $StreamWriter.WriteLine($Separator)
			    #Add-Content -Path $SeparatorFile -Value $Separator -Force
    	    }
        }        
    }

    $StreamWriter.Close()
    $FileStream.Close()

    Write-Host 'The separator list file (' -NoNewline -ForegroundColor Green
    Write-Host $SeparatorFile -NoNewline
    Write-Host ') was created successfully.' -ForegroundColor Green
}
Catch
{
    Write-Host 'Could not write to the output file: ' -ForegroundColor Red -NoNewline
    Write-Host $SeparatorFile -ForegroundColor Cyan
    Write-Host 'Please make sure the file is not locked and the path is available.' -ForegroundColor Yellow
    Break
}