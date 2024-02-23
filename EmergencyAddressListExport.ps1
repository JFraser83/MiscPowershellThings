# Check if Microsoft Teams PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name "MicrosoftTeams")) {
    # Install Microsoft Teams PowerShell module
    Install-Module -Name "MicrosoftTeams" -Force
}


# Prompt for filename and validate location is valid
$filename = Read-Host "Enter the filename"
$isValidLocation = Test-Path (Split-Path -Path $filename -Parent)
if (-not $isValidLocation) {
    Write-Host "Invalid location. Please enter a valid filename."
    return
}

Get-CsOnlineLisLocation | Select-Object Location,CompanyName,HouseNumber,HouseNumberSuffix,StreetName,StreetSuffix,City,PostalCode,StateOrProvince,CountryOrRegion,Description | Export-CSV -Path $filename



