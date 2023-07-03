<#
Copies an excel file to a CSV file.

Has two optional parameters that can be used when invoking the script:
InputPath: Path to excel file that will be copied. Must be fully qualified path including file name and extension.
OutputPath: Path to export CSV file. Must be fully qualified path including file name and extension.
#> 

# params
param (
    [string]$inputPath,
    [string]$outputPath
)

# functions
function Initialize-ColorScheme
{
    $script:successColor = "Green"
    $script:infoColor = "DarkCyan"
    $script:warningColor = "Yellow"
    $script:failColor = "Red"    
}

function Show-Introduction
{
    Write-Host "This script converts excel files to CSV." -ForegroundColor $infoColor
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule $moduleName
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor $infoColor
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required." -ForegroundColor $infoColor
        $confirmInstall = Read-Host -Prompt "Would you like to install the module? (y/n)"
    }
    while ($confirmInstall -inotmatch "^\s*y\s*$") # regex matches a y but allows spaces
}

function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Write-Host ("Please run script with admin privileges.`n" +
        "1. Open Powershell as admin.`n" +
        "2. CD into script directory.`n" +
        "3. Run .\scriptname`n") -ForegroundColor $failColor
        Read-Host "Press Enter to exit"
        exit
    }
}

function Validate-InputPath($path)
{
    $trimmedPath = $path.Trim('"')

    if (-not(Test-FileHasExtension -FileName $trimmedPath -Extension ".XLSX"))
    {
        Write-Warning "File is not an XLSX."
        return $false
    }

    $fileExists = Test-Path $trimmedPath
    if (-not($fileExists))
    {
        Write-Warning "File does not exist."
        return $false
    }

    return $true
}

function Test-FileHasExtension([string]$fileName, [string]$extension)
{
    $actualExtension = [System.IO.Path]::GetExtension($fileName)
    return $actualExtension -ieq $extension.Trim()
}

function Prompt-InputPath
{
    do
    {
        $inputPath = Read-Host "Enter path to excel file"
        $inputPath = $inputPath.Trim('"')
        $validPath = Validate-InputPath $inputPath
    }
    while (-not($validPath))

    return $inputPath
}

function Validate-OutputPath($path)
{
    $trimmedPath = $path.Trim('"')

    if (-not(Test-FileHasExtension -FileName $trimmedPath -Extension ".CSV"))
    {
        Write-Warning "Path must end in .csv."
        return $false
    }

    $rooted = [System.IO.Path]::IsPathRooted($trimmedPath)
    if (-not($rooted))
    {
        Write-Warning "Path is invalid."
        return $false
    }

    return $true
}

function Import-ExcelFile($path)
{
    Write-Host "Importing excel file..." -ForegroundColor $infoColor

    return Import-Excel -Path $path
}

function Export-CsvFile
{
    [CmdletBinding()]

    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        $inputObject,
        [Parameter(Mandatory)]
        $outputPath
    )   

    begin {
        Write-Host "Exporting csv file..." -ForegroundColor $infoColor
    }

    process {
        Export-CSV -InputObject $inputObject -Path $outputPath -NoTypeInformation -Append
    }
}

function Get-OutputPath($inputPath)
{
    $inputPathWithoutExtension = Get-FilePathWithoutExtension $inputPath
    return "$inputPathWithoutExtension.csv"
}

function Get-FilePathWithoutExtension($path)
{
    $folder = Split-Path -Path $path -Parent
    $fileName = Split-Path -Path $path -Leaf
    $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    return "$folder\$baseFileName"
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module "ImportExcel"

if ($inputPath)
{
    $inputPath = $inputPath.Trim('"')
    $pathValid = Validate-InputPath $inputPath
    if (-not($pathValid))
    {
        $inputPath = Prompt-InputPath
    }
}

if ($outputPath)
{
    $outputPath = $outputPath.Trim('"')
    $validPath = Validate-OutputPath $outputPath
    if (-not($validPath))
    {
        $outputPath = $null
    }
}

if ($inputPath)
{
    $excelFile = Import-ExcelFile -Path $inputPath

    if ($outputPath)
    {        
        $excelFile | Export-CsvFile -OutputPath $outputPath     
    }
    else
    {
        $outputPath = Get-OutputPath $inputPath
        $excelFile | Export-CsvFile -OutputPath $outputPath
    }
}
else
{
    $inputPath = Prompt-InputPath
    $excelFile = Import-ExcelFile -Path $inputPath

    if ($outputPath)
    {
        $excelFile | Export-CsvFile -OutputPath $outputPath
    }
    else
    {
        $outputPath = Get-OutputPath $inputPath
        $excelFile | Export-CsvFile -OutputPath $outputPath
    }
}

Write-Host "Finished copying to $outputPath" -ForegroundColor $successColor
Write-Host "Press Enter to exit"