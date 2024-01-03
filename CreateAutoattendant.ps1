
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

# Check if Exchange Online PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name "ExchangeOnlineManagement")) {
    # Install Exchange Online PowerShell module
    Install-Module -Name "ExchangeOnlineManagement" -Force
}

#Auto Attendant: ce933385-9390-45d1-9512-c8d228074e07
#Call Queue: 11cd3e2e-fccb-42ad-ad00-878b93575e07


# Connect to Microsoft Teams PowerShell module
Connect-MicrosoftTeams 
Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All" -NoWelcome
Connect-ExchangeOnline 
#1. Create Group - This is the dial scope for the AA Add members 

#2. Create the resource account and assign the license 

#3. Create the AA 

#4. Create the CQ 

#5 After hours call flow 

# Prompt the user for the office location
do {
    $officelocation = Read-Host "Enter the office location name"
    $validLocation = [string]::IsNullOrEmpty($officelocation) -eq $false
    if (-not $validLocation) {
        Write-Host "Invalid value. Please try again."
    }
} until ($validLocation)

# Prompt the user for the consultant ID
do {
    $consultantID = Read-Host "Enter the consultant ID"
    $validConsultantID = [string]::IsNullOrEmpty($consultantID) -eq $false
    if (-not $validConsultantID) {
        Write-Host "Invalid consultant ID. Please try again."
    }
} until ($validConsultantID)

# Prompt the user for phone number input in E164 format
#Add validation to this to check if the number is already in use and must be a service number.
do {
    $phoneNumber = Read-Host "Enter the phone number in E164 format (e.g. +1234567890)"
    $validPhoneNumber = $phoneNumber -match '^\+\d{1,15}$'
    if (-not $validPhoneNumber) {
        Write-Host "Invalid phone number. Please try again."
    }
} until ($validPhoneNumber)

$PhoneNumberString = $phoneNumber.trimstart("+")

# Prompt the user for timezone input
do {
    $timezone = Read-Host "Enter the timezone (e.g. Pacific Standard Time)"
    $validTimezone = Get-TimeZone -Name $timezone -ErrorAction SilentlyContinue
    if (-not $validTimezone) {
        Write-Host "Invalid timezone. Please try again."
    }
} until ($validTimezone)

# Prompt the user for a valid language
do {
    $language = Read-Host "Enter the language (English/French)"
    $validLanguage = $language -eq "English" -or $language -eq "French"
    if (-not $validLanguage) {
        Write-Host "Invalid language. Please try again."
    }
} until ($validLanguage)


$TimedRedirectAA = $officeLocation + "-" + $consultantID + "-" + $PhoneNumberString + "-TRAA"
$CallQueueName = $officeLocation + "-" + $consultantID + "-" + $PhoneNumberString + "-CQ"
$AAName = $officeLocffation + "-" + $consultantID + "-" + $PhoneNumberString + "-AA"       
$DialDirectoryName = $officeLocation + "-" + $consultantID + "-" + $PhoneNumberString + "-DD"     
$CallQueueMembership = $officeLocation + "-" + $consultantID + "-" + $PhoneNumberString + "-CQ"
$upnSuffix = (Get-MgDomain | Where-Object {$_.isDefault}).Id

# Create Microsoft 365 group named $DialDirectoryName


$existingGroup = Get-MgGroup -Filter "DisplayName eq '$DialDirectoryName'"
if (-not $existingGroup) {
    try {
        New-MgGroup -DisplayName $DialDirectoryName -MailEnabled -MailNickname $DialDirectoryName -SecurityEnabled -GroupTypes Unified
    } catch {
        Write-Host "An error occurred: $_"
    }
}
#Create a Team for the call queue  
# Create a Team for the call queue membership
try {
    $teamID = New-Team -DisplayName $CallQueueMembership -Visibility "private"
    Write-Host "Team named $CallQueueMembership is created successfully" -ForegroundColor Green
} catch {
    Write-Host "An error occurred: $_"
}
try {
    $channelID = Get-TeamChannel -GroupID $teamID.GroupID | Where-Object {$_.DisplayName -eq "General"} | Select-Object -ExpandProperty Id
    Write-Host "ChannelID is $channelID"
}
catch {
    Write-Host "An error occurred: $_"
}


# Create AutoAttendant Resource Account 
   $upn = $PhoneNumberString + "RA@" + $upnSuffix 
    #Check if UPN Already Exists if not create it 
    try {
            $user = Get-CsOnlineUser -Identity $upn -ErrorAction Stop
        
        }
    catch {
            New-CsOnlineApplicationInstance -UserPrincipalName $upn -DisplayName $displayName -ApplicationID "ce933385-9390-45d1-9512-c8d228074e07"    
        }       
    $upnCreated = $false
    
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
        Update-MgUser -UserId $upn -UsageLocation "CA"
        $PhoneSysVirtualSku = Get-MgSubscribedSku -All | Where-Object { $_.SkuPartNumber -eq 'PHONESYSTEM_VIRTUALUSER' }
        Set-MgUserLicense -UserId $upn -AddLicenses @{SkuId = $PhoneSysVirtualSku.SkuId} -Removelicenses @()
    

     #Create Call Queue
try {
    $teamSupportID = $teamID.GroupID

    New-CsCallQueue -Name $CallQueueName -AgentAlertTime 15 -AllowOptOut $false -ChannelID $channelID -DistributionLists $teamSupportID -OverflowAction SharedVoicemail -EnableOverflowSharedVoicemailTranscription $true -TimeoutAction SharedVoicemail -TimeoutActionTarget $teamSupportID -TimeoutThreshold 2700 -TimeoutSharedVoicemailTextToSpeechPrompt "We're sorry to have kept you waiting and are now transferring your call to voicemail." -EnableTimeoutSharedVoicemailTranscription $true -RoutingMethod LongestIdle -ConferenceMode $true -LanguageID "en-US" -ErrorAction Stop
    Write-Host "Call Queue $CallQueueName created successfully" -ForegroundColor Green 
}
catch {
    Write-Host "An error occurred: $_"
}


#Create AutoAttendant 

#$timerangeMoFr = New-CsOnlineTimeRange -Start 08:30 -end 17:00

#$afterHoursSchedule = New-CsOnlineSchedule -Name "After Hours Schedule" -WeeklyRecurrentSchedule -MondayHours @($timerangeMoFr) -TuesdayHours @($timerangeMoFr) -WednesdayHours @($timerangeMoFr) -ThursdayHours @($timerangeMoFr) -FridayHours @($timerangeMoFr)  -Complement

#$openallOptions = New-CsAutoAttendantMenuOption -Action TransfercallToTarget -DtfmResponse Automatic -CallTarget (Get-CsOnlineUser $CallQueueName).Identity

#New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone2 -CallTarget $openHoursMenuOption2Entity
#$salesGroupID = Find-CsGroup -SearchQuery "Sales" | % { $_.Id }
#$supportGroupID = Find-CsGroup -SearchQuery "Support" | % { $_.Id }
#$dialScope = New-CsAutoAttendantDialScope -GroupScope -GroupIds @($salesGroupID, $supportGroupID)

