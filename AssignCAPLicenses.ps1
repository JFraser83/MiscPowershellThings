# Check if Microsoft Teams PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name "MicrosoftTeams")) {
    # Install Microsoft Teams PowerShell module
    Install-Module -Name "MicrosoftTeams" -Force
}

# Check if Graph PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph")) {
    # Install MSOnline PowerShell module
    Install-Module -Name "Microsft.Graph" -Force
}





Import-Module MicrosoftTeams
Import-Module Microsoft.Graph

$LogDate = get-date -f dd-MM-yyyy_HHmmffff

#Location of CSV File that will include the UPN and phone # 
$ImportPath = "C:\scripts\TestPhoneUserAssignment.csv"

$ExportPath = "C:\scripts\PhoneNumberAssignment_$LogDate.csv"

$Users = Import-CSV $importPath 

Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All"

#>
foreach ($User in $Users) {

    $UPN = $User.UserPrincipalName 
    
    try
    {
         # Loop until the user $upn is created
         do {
            try {
                $user = Get-CsOnlineUser -Identity $upn -ErrorAction Stop
                $upnCreated = $true
                }
            catch {
                Write-Host "User $upn is not created yet. Retrying..."
                Start-Sleep -Seconds 30
                }       
        } until ($upnCreated)
        Update-MgUser -UserId $UPN -UsageLocation "US"
        $PhoneSysVirtualSku = Get-MgSubscribedSku -All | Where-Object { $_.SkuPartNumber -eq 'MCOCAP' }

        Set-MgUserLicense -UserId $UPN -AddLicenses @{SkuId = $PhoneSysVirtualSku.SkuId} -Removelicenses @()
        
        $Result = "The user $UPN license is updated"
        Write-Host $Result  -ForegroundColor Cyan 

    }
    catch
    {
        $Result = "$UPN has an error"
        Write-Host $Result -ForegroundColor Red 
    }
    Finally
    {
        $User | Add-Member -MemberType NoteProperty -Name "Result" -Value $Result -Force 
    }


}
$Users | Export-CSV -Encoding UTF8 $ExportPath -NoTypeInformation 
