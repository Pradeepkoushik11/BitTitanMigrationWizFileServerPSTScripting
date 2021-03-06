<#
.SYNOPSIS
    Script to prepare the execution of a BitTitan FileServer Home Directory to OneDrive For Business migration. 
    
.DESCRIPTION
    Script to prepare the execution of a BitTitan FileServer Home Directory to OneDrive For Business migration. It will ask for the
    - Working directory: where all PowerShell scripts will be automatically downloaded from GitHub.
    - BitTItan workgroup
    - BitTitan customer
    - Source 'AzureFileSystem' endpoint: to create the MigrationWiz projects
    - Destination 'OneDriveProAPI' endpoint : to create the MigrationWiz projects
    - Azure Storage Primary Access Key: for UploaderWiz to create blob container and upload the home directories
    - Folder path to the FileServer Home Directory root    
    
.NOTES
    Author          Pablo Galan Sabugo <pablogalanscripts@gmail.com> 
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>

#######################################################################################################################
#                  HELPER FUNCTIONS                          
#######################################################################################################################
Function Import-PowerShellModules {
    if (!(((Get-Module -Name "MSOnline") -ne $null) -or ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null))) {
        Write-Host
        $msg = "INFO: MSOnline PowerShell module not installed."
        Write-Host $msg     
        $msg = "INFO: Installing MSOnline PowerShell module."
        Write-Host $msg

        Sleep 5
    
        try {
            Install-Module -Name MSOnline -force -ErrorAction Stop
        }
        catch {
            $msg = "ERROR: Failed to install MSOnline module. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Write-Host
            $msg = "ACTION: Run this script 'As administrator' to intall the MSOnline module."
            Write-Host -ForegroundColor Yellow $msg
            Exit
        }
        Import-Module MSOnline
    }
}
function Import-MigrationWizPowerShellModule {
    
    Write-host
    $version = (get-host | select-object Version).version
    Write-host -ForegroundColor Magenta "PowerShell version installed: $($version.Major).$($version.Minor).$($version.Build).$($version.Revision)"

    if ($version.Major -lt 5) {
        Write-host -ForegroundColor Yellow "ACTION: Install at least PowerShell 5.0. Script aborted."
        Write-host

        Exit
    }

    if (( $null -ne (Get-Module -Name "BitTitanPowerShell")) -or ( $null -ne (Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue))) {
        return
    }

    $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
    $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll", "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
    foreach ($moduleLocation in $moduleLocations) {
        if (Test-Path $moduleLocation) {
            Import-Module -Name $moduleLocation
            return
        }
    }
    
    $msg = "INFO: BitTitan PowerShell SDK not installed."
    Write-Host -ForegroundColor Red $msg 

    Write-Host
    $msg = "ACTION: Install BitTitan PowerShell SDK 'bittitanpowershellsetup.msi' downloaded from 'https://www.bittitan.com'."
    Write-Host -ForegroundColor Yellow $msg

    Start-Sleep 5

    $url = "https://www.bittitan.com/downloads/bittitanpowershellsetup.msi " 
    $result = Start-Process $url
    Exit

}

# Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory = $true)] [string]$workingDir,
        [parameter(Mandatory = $true)] [string]$logDir
    )
    if ( !(Test-Path -Path $global:btWorkingDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $global:btWorkingDir -Force -ErrorAction Stop
            $msg = "SUCCESS: Folder '$($global:btWorkingDir)' for CSV files has been created."
            Write-Host -ForegroundColor Green $msg
        }
        catch {
            $msg = "ERROR: Failed to create '$global:btWorkingDir'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
        }
    }
    if ( !(Test-Path -Path $logDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop      

            $msg = "SUCCESS: Folder '$($logDir)' for log files has been created."
            Write-Host -ForegroundColor Green $msg 
        }
        catch {
            $msg = "ERROR: Failed to create log directory '$($logDir)'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
        } 
    }
}

# Function to write information to the Log File
Function Log-Write {
    param
    (
        [Parameter(Mandatory = $true)]    [string]$Message
    )
    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
    Add-Content -Path $script:logFile -Value $lineItem
}

Function Get-FileName {
    param 
    (      
        [parameter(Mandatory = $false)] [String]$initialDirectory,
        [parameter(Mandatory = $false)] [String]$DefaultColumnName,
        [parameter(Mandatory = $false)] [String]$Extension
    )

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    if ($extension -eq "csv") {
        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    }
    elseif ($extension -eq "json") {
        $OpenFileDialog.filter = "JSON (*.json)| *.json"
    }
    $OpenFileDialog.ShowDialog() | Out-Null
    $script:inputFile = $OpenFileDialog.filename

    if ($OpenFileDialog.filename -eq "") {

        if ($defaultColumnName -eq "PrimarySmtpAddress") {
            # create new import file
            $inputFileName = "FilteredPrimarySmtpAddress-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
            $script:inputFile = "$initialDirectory\$inputFileName"

            $csv = "PrimarySmtpAddress`r`n"
            $file = New-Item -Path $initialDirectory -name $inputFileName -ItemType file -value $csv -force 

            $msg = "SUCCESS: Empty CSV file '$script:inputFile' created."
            Write-Host -ForegroundColor Green  $msg
                
            $msg = "WARNING: Populate the CSV file with the source 'PrimarySmtpAddress' you want to process in a single column and save it as`r`n         File Type: 'CSV (Comma Delimited) (.csv)'`r`n         File Name: '$script:inputFile'."
            Write-Host -ForegroundColor Yellow $msg
        }
        elseif ($defaultColumnName -eq "MailNickName") {
            # create new import file
            $inputFileName = "FilteredTeamMailNickName-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
            $script:inputFile = "$initialDirectory\$inputFileName"

            $csv = "MailNickName`r`n"
            $file = New-Item -Path $initialDirectory -name $inputFileName -ItemType file -value $csv -force 

            $msg = "SUCCESS: Empty CSV file '$script:inputFile' created."
            Write-Host -ForegroundColor Green  $msg
                
            $msg = "WARNING: Populate the CSV file with the source 'MailNickName' you want to process in a single column and save it as`r`n         File Type: 'CSV (Comma Delimited) (.csv)'`r`n         File Name: '$script:inputFile'."
            Write-Host -ForegroundColor Yellow $msg  
        }
        elseif ($defaultColumnName -eq "SourceEmailAddress,DestinationEmailAddress") {
            # create new import file
            $inputFileName = "UserMapping-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
            $script:inputFile = "$initialDirectory\$inputFileName"

            $csv = "SourceEmailAddress,DestinationEmailAddress`r`n"
            $file = New-Item -Path $initialDirectory -name $inputFileName -ItemType file -value $csv -force 

            $msg = "SUCCESS: Empty CSV file '$script:inputFile' created."
            Write-Host -ForegroundColor Green  $msg
                
            $msg = "WARNING: Populate the CSV file with the 'SourceEmailAddress', 'DestinationEmailAddress' columns and save it as`r`n         File Type: 'CSV (Comma Delimited) (.csv)'`r`n         File Name: '$script:inputFile'."
            Write-Host -ForegroundColor Yellow $msg  
        }
        else {
            Return $false
        }            

        # open file for editing
        Start-Process $file 

        do {
            $confirm = (Read-Host -prompt "Are you done editing the import CSV file?  [Y]es or [N]o")
            if ($confirm -eq "Y") {
                $importConfirm = $true
            }

            if ($confirm -eq "N") {
                $importConfirm = $false
                try {
                    #Open the CSV file for editing
                    Start-Process -FilePath $script:inputFile
                }
                catch {
                    $msg = "ERROR: Failed to open '$script:inputFile' CSV file. Script aborted."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $_.Exception.Message
                }
            }
        }
        while (-not $importConfirm)
            
        $msg = "SUCCESS: CSV file '$script:inputFile' saved."
        Write-Host -ForegroundColor Green  $msg

        Return $true
    }
    else {
        Write-Host
        $msg = "SUCCESS: $($Extension.ToUpper()) file '$($OpenFileDialog.filename)' selected."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
        Return $true
    }
}

Function Get-Directory($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null    
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.ShowDialog() | Out-Null

    if ($FolderBrowser.SelectedPath -ne "") {
        $workingDir = $FolderBrowser.SelectedPath               
    }
    Write-Host -ForegroundColor Gray  "INFO: Directory '$workingDir' selected."
}

# Function to wait for the user to press any key to continue
Function WaitForKeyPress {
    $msg = "ACTION: If you have edited and saved the CSV file then press any key to continue. Press 'Ctrl + C' to exit." 
    Write-Host $msg
    Log-Write -Message $msg
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}

Function Import-CSV_UserMapping {

    $result = Get-FileName $global:btWorkingDir -DefaultColumnName "SourceEmailAddress,DestinationEmailAddress" -Extension "csv"

    if ($result) {
        ##Import the CSV file
        Try {
            $script:emailAddressMappingCSVFile = @(Import-Csv $script:inputFile | Where-Object { $_.PSObject.Properties.Value -ne "" } )
        }
        Catch [Exception] {
            Write-Host -ForegroundColor Red "ERROR: Failed to import the CSV file '$script:inputFile'."
            Write-Host -ForegroundColor Red $_.Exception.Message
            Exit
        }

        #Check if CSV is formated properly
        If ($script:emailAddressMappingCSVFile.SourceEmailAddress -eq $null -or $script:emailAddressMappingCSVFile.DestinationEmailAddress -eq $null) {
            Write-Host -ForegroundColor Red "ERROR: The CSV file format is invalid. It must have 2 columns: 'SourceEmailAddress' and 'DestinationEmailAddress' columns."
            Exit 
        }

        #Load existing advanced options
        $ADVOPTString += $Connector.AdvancedOptions
        $ADVOPTString += "`n"

        $count = 0

        #Processing CSV into string
        Write-Host "         INFO: Applying UserMapping from CSV File:"

        $script:emailAddressMappingCSVFile | ForEach-Object {

            $sourceAddress = $_.SourceEmailAddress
            $destinationAddress = $_.DestinationEmailAddress

            $userMapping = "UserMapping=`"$sourceAddress->$destinationAddress`""

            $count += 1

            Write-Host -ForegroundColor Green "         SUCCESS: UserPrincipalName mapping $sourceAddress->$destinationAddress found." 
                   
            $allUserMappings += $userMapping
            $allUserMappings += "`n"

        }

        Write-Host -ForegroundColor Green "         SUCCESS: CSV file '$script:inputFile' succesfully processed. $count recipient mappings applied."

        Return $allUserMappings
    }
}

Function isNumeric($x) {
    $x2 = 0
    $isNum = [System.Int32]::TryParse($x, [ref]$x2)
    return $isNum
}

# Function to query destination email addresses
Function Apply-EmailAddressMapping {
    do {
        $confirm = (Read-Host -prompt "Are you migrating to the same UserPrincipalName prefixes?  [Y]es or [N]o")
    } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    if ($confirm.ToLower() -eq "n") {
        
        Return $true         
    }
    else {
        Return $false 
    }
}

#######################################################################################################################
#                    BITTITAN
#######################################################################################################################

# Function to authenticate to BitTitan SDK
Function Connect-BitTitan {
    #[CmdletBinding()]

    if ((Get-Module PackageManagement)) { 
        #Install Packages/Modules for Windows Credential Manager if required
        If (!(Get-PackageProvider -Name 'NuGet')) {
            Install-PackageProvider -Name NuGet -Force
        }
        If ((Get-Module PowerShellGet) -and !(Get-Module -ListAvailable -Name 'CredentialManager')) {
            Install-Module CredentialManager -Force
            $useCredentialManager = $true
        } 
        else { 
            Import-Module CredentialManager
            $useCredentialManager = $true
        }

        if ($useCredentialManager ) {
            # Authenticate
            $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'
        }
    }
    else {
        $useCredentialManager = $false
    }
    
    if (!$script:creds) {
        $credentials = (Get-Credential -Message "Enter BitTitan credentials")
        if (!$credentials) {
            $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Exit
        }

        if ($useCredentialManager) {
            New-StoredCredential -Target 'https://migrationwiz.bittitan.com' -Persist 'LocalMachine' -Credentials $credentials | Out-Null
            
            $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' stored in Windows Credential Manager."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg

            $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'

            $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' retrieved from Windows Credential Manager."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg
        }
        else {
            $script:creds = $credentials
        }
    }
    else {
        $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' retrieved from Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }

    try { 
        # Get a ticket and set it as default
        $script:ticket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction Stop
        # Get a MW ticket
        $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction Stop 
    }
    catch {

        $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
        $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll", "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
        foreach ($moduleLocation in $moduleLocations) {
            if (Test-Path $moduleLocation) {
                Import-Module -Name $moduleLocation

                # Get a ticket and set it as default
                $script:ticket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction SilentlyContinue
                # Get a MW ticket
                $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction SilentlyContinue 

                if (!$script:ticket -or !$script:mwTicket) {
                    $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                    Exit
                }
                else {
                    $msg = "SUCCESS: Connected to BitTitan."
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg
                }

                return
            }
        }

        $msg = "ACTION: Install BitTitan PowerShell SDK 'bittitanpowershellsetup.msi' downloaded from 'https://www.bittitan.com' and execute the script from there."
        Write-Host -ForegroundColor Yellow $msg
        Write-Host

        Start-Sleep 5

        $url = "https://www.bittitan.com/downloads/bittitanpowershellsetup.msi " 
        $result = Start-Process $url

        Exit
    }  

    if (!$script:ticket -or !$script:mwTicket) {
        $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Exit
    }
    else {
        $msg = "SUCCESS: Connected to BitTitan."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
}
Function Connect-Office365 {
    #Install Packages/Modules for Windows Credential Manager if required
    If (!(Get-PackageProvider -Name 'NuGet')) {
        Install-PackageProvider -Name NuGet -Force
    }
    If (!(Get-Module -ListAvailable -Name 'CredentialManager')) {
        Install-Module CredentialManager -Force
    } 
    else { 
        Import-Module CredentialManager
    }

    # Authenticate
    $global:btO365Credentials = Get-StoredCredential -Target 'MigrationWizFileServerMigration'
    
    if (!$global:btO365Credentials) {
        $credentials = (Get-Credential -Message "Enter Office 365 credentials")
        if (!$credentials) {
            $msg = "ERROR: Failed to authenticate with Office 365. Please enter valid Office 365 Credentials. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Exit
        }
        New-StoredCredential -Target 'MigrationWizFileServerMigration' -Persist 'LocalMachine' -Credentials $credentials | Out-Null
        
        $msg = "SUCCESS: Office 365 credentials for target 'MigrationWizFileServerMigration' stored in Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg

        $global:btO365Credentials = Get-StoredCredential -Target 'MigrationWizFileServerMigration'

        $msg = "SUCCESS: Office 365 credentials for target 'MigrationWizFileServerMigration' retrieved from Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
    else {
        $msg = "SUCCESS: Office 365 credentials for target 'MigrationWizFileServerMigration' retrieved from Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
}

# Function to create a connector under a customer
Function Create-MW_Connector {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId,
        [parameter(Mandatory = $true)] [String]$ProjectName,
        [parameter(Mandatory = $true)] [String]$ProjectType,
        [parameter(Mandatory = $true)] [String]$importType,
        [parameter(Mandatory = $true)] [String]$exportType,   
        [parameter(Mandatory = $true)] [guid]$exportEndpointId,
        [parameter(Mandatory = $true)] [guid]$importEndpointId,  
        [parameter(Mandatory = $true)] [object]$exportConfiguration,
        [parameter(Mandatory = $true)] [object]$importConfiguration,
        [parameter(Mandatory = $false)] [String]$advancedOptions,   
        [parameter(Mandatory = $false)] [String]$folderFilter = "",
        [parameter(Mandatory = $false)] [String]$maximumSimultaneousMigrations = 100,
        [parameter(Mandatory = $false)] [String]$MaxLicensesToConsume = 10,
        [parameter(Mandatory = $false)] [int64]$MaximumDataTransferRate,
        [parameter(Mandatory = $false)] [String]$Flags,
        [parameter(Mandatory = $false)] [String]$ZoneRequirement,
        [parameter(Mandatory = $false)] [Boolean]$updateConnector   
        
    )
    try {
        $connector = @(Get-MW_MailboxConnector -ticket $script:MwTicket `
                -UserId $script:MwTicket.UserId `
                -OrganizationId $customerOrganizationId `
                -Name "$ProjectName" `
                -ErrorAction SilentlyContinue
            #-SelectedExportEndpointId $exportEndpointId `
            #-SelectedImportEndpointId $importEndpointId `        
            #-ProjectType $ProjectType `
            #-ExportType $exportType `
            #-ImportType $importType `

        ) 

        if ($connector.Count -eq 1) {
            $msg = "WARNING: Connector '$($connector.Name)' already exists with the same configuration." 
            write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg 

            if ($updateConnector) {
                $connector = Set-MW_MailboxConnector -ticket $script:MwTicket `
                    -MailboxConnector $connector `
                    -Name $ProjectName `
                    -ExportType $exportType `
                    -ImportType $importType `
                    -SelectedExportEndpointId $exportEndpointId `
                    -SelectedImportEndpointId $importEndpointId `
                    -ExportConfiguration $exportConfiguration `
                    -ImportConfiguration $importConfiguration `
                    -AdvancedOptions $advancedOptions `
                    -FolderFilter $folderFilter `
                    -MaximumDataTransferRate ([int]::MaxValue) `
                    -MaximumDataTransferRateDuration 600000 `
                    -MaximumSimultaneousMigrations $maximumSimultaneousMigrations `
                    -PurgePeriod 180 `
                    -MaximumItemFailures 1000 `
                    -ZoneRequirement $ZoneRequirement `
                    -MaxLicensesToConsume $MaxLicensesToConsume  
                #-Flags $Flags `

                $msg = "SUCCESS: Connector '$($connector.Name)' updated." 
                write-Host -ForegroundColor Blue $msg
                Log-Write -Message $msg 

                return $connector.Id
            }
            else { 
                return $connector.Id 
            }
        }
        elseif ($connector.Count -gt 1) {
            $msg = "WARNING: $($connector.Count) connectors '$ProjectName' already exist with the same configuration." 
            write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg 

            return $null

        }
        else {
            try { 
                $connector = Add-MW_MailboxConnector -ticket $script:MwTicket `
                    -UserId $script:MwTicket.UserId `
                    -OrganizationId $customerOrganizationId `
                    -Name $ProjectName `
                    -ProjectType $ProjectType `
                    -ExportType $exportType `
                    -ImportType $importType `
                    -SelectedExportEndpointId $exportEndpointId `
                    -SelectedImportEndpointId $importEndpointId `
                    -ExportConfiguration $exportConfiguration `
                    -ImportConfiguration $importConfiguration `
                    -AdvancedOptions $advancedOptions `
                    -FolderFilter $folderFilter `
                    -MaximumDataTransferRate ([int]::MaxValue) `
                    -MaximumDataTransferRateDuration 600000 `
                    -MaximumSimultaneousMigrations $maximumSimultaneousMigrations `
                    -PurgePeriod 180 `
                    -MaximumItemFailures 1000 `
                    -ZoneRequirement $ZoneRequirement `
                    -MaxLicensesToConsume $MaxLicensesToConsume  
                #-Flags $Flags `

                $msg = "SUCCESS: Connector '$($connector.Name)' created." 
                write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                return $connector.Id
            }
            catch {
                $msg = "ERROR: Failed to create mailbox connector '$($connector.Name)'."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message  
            }
        }
    }
    catch {
        $msg = "ERROR: Failed to get mailbox connector '$($connector.Name)'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message  
    }

}

# Function to create an endpoint under a customer. Configuration Table in https://www.bittitan.com/doc/powershell.html#PagePowerShellmspcmd%20
Function Create-MSPC_Endpoint {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId,
        [parameter(Mandatory = $false)] [String]$endpointType,
        [parameter(Mandatory = $false)] [String]$endpointName,
        [parameter(Mandatory = $false)] [object]$endpointConfiguration,
        [parameter(Mandatory = $false)] [String]$exportOrImport,
        [parameter(Mandatory = $false)] [Boolean]$updateEndpoint
    )

    $script:CustomerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId
    
    if ($endpointType -eq "AzureFileSystem") {
        
        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")

            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key ").trim()
            }while ($secretKey -eq "")

            $msg = "INFO: Azure storage account access key is '$secretKey'."
            Write-Host $msg
            Log-Write -Message $msg 

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $azureFileSystemConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $azureAccountName; #Azure Storage Account Name        
                "AccessKey"                    = $secretKey; #Azure Storage Account SecretKey         
                "ContainerName"                = $ContainerName #Container Name
            }
        }
        else {
            $azureFileSystemConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername; #Azure Storage Account Name        
                "AccessKey"                    = $endpointConfiguration.AccessKey; #Azure Storage Account SecretKey         
                "ContainerName"                = $endpointConfiguration.ContainerName #Container Name
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $azureFileSystemConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $azureFileSystemConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1
        }    
    }
    elseif ($endpointType -eq "AzureSubscription") {
           
        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($secretKey -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            do {
                $azureSubscriptionID = (Read-Host -prompt "Please enter the Azure subscription ID").trim()
            }while ($azureSubscriptionID -eq "")

            $msg = "INFO: Azure subscription ID is '$azureSubscriptionID'."
            Write-Host $msg
            Log-Write -Message $msg 

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureSubscriptionConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $adminUsername;     
                "AdministrativePassword"       = $adminPassword;         
                "SubscriptionID"               = $azureSubscriptionID
            }
        }
        else {
            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureSubscriptionConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;  
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword;    
                "SubscriptionID"               = $endpointConfiguration.SubscriptionID 
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1
        }   
    }
    elseif ($endpointType -eq "BoxStorage") {
        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $boxStorageConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.BoxStorageConfiguration' -Property @{  
            }
        }
        else {
            $boxStorageConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.BoxStorageConfiguration' -Property @{  
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $boxStorageConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $boxStorageConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1
        }  
    }
    elseif ($endpointType -eq "DropBox") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $dropBoxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.DropBoxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativePassword"       = ""
            }
        }
        else {
            $dropBoxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.DropBoxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativePassword"       = ""
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $dropBoxConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $dropBoxConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1              
        }
    }      
    elseif ($endpointType -eq "Gmail") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $domains = (Read-Host -prompt "Please enter the domain or domains (separated by comma)").trim()
            }while ($domains -eq "")
        
            $msg = "INFO: Domain(s) is (are) '$domains'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $GoogleMailboxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleMailboxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "Domains"                      = $Domains;
                "ContactHandling"              = 'MigrateSuggestedContacts';
            }
        }
        else {
            $adminUsername = $endpointConfiguration.AdministrativeUsername
            $GoogleMailboxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleMailboxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "Domains"                      = $endpointConfiguration.Domains;
                "ContactHandling"              = 'MigrateSuggestedContacts';
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleMailboxConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $GoogleMailboxConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1              
        }
    }
    elseif ($endpointType -eq "GSuite") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                Write-host -NoNewline "Please enter the file path to the Google service account credentials using JSON file:"
                
                $result = Get-FileName $global:btWorkingDir -Extension "json"

                #Read CSV file
                try {
                    $script:jsonFileContent = get-content $script:inputFile -raw
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'."
                    Write-Host -ForegroundColor Red  $msg
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message
                    Return -1    
                } 
            }while ($script:jsonFileContent -eq "")
        
            $msg = "INFO: The file path to the JSON file is '$script:inputFile'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################
          
            $GoogleMailboxConfiguration = New-BT_GSuiteConfiguration -AdministrativeUsername $adminUsername `
                -CredentialsFileName $script:inputFile `
                -Credentials $script:jsonFileContent   

        }
        else {
            $adminUsername = $endpointConfiguration.AdministrativeUsername
            do {
                Write-host -NoNewline "Please enter the file path to the Google service account credentials using JSON file:"

                $result = Get-FileName $global:btWorkingDir -Extension "json"

                #Read CSV file
                try {
                    $script:jsonFileContent = get-content $script:inputFile -raw
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'."
                    Write-Host -ForegroundColor Red  $msg
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message
                    Return -1 
                } 
            }while ($script:jsonFileContent -eq "")
        
            $msg = "INFO: The file path to the JSON file is '$script:inputFile'."
            Write-Host $msg
            Log-Write -Message $msg 
            $GoogleMailboxConfiguration = New-BT_GSuiteConfiguration -AdministrativeUsername $adminUsername `
                -CredentialsFileName $script:inputFile `
                -Credentials $script:jsonFileContent   
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleMailboxConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $GoogleMailboxConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1              
        }
    }
    elseif ($endpointType -eq "GoogleSharedDrive") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                Write-host -NoNewline "Please enter the file path to the Google service account credentials using JSON file:"
                
                $result = Get-FileName $global:btWorkingDir -Extension "json" 

                #Read CSV file
                try {
                    $script:jsonFileContent = get-content $script:inputFile -raw
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'."
                    Write-Host -ForegroundColor Red  $msg
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message
                    Return -1    
                } 
            }while ($script:jsonFileContent -eq "")
        
            $msg = "INFO: The file path to the JSON file is '$script:inputFile'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $script:selectMigrateSharedDriveMembership = $true
                $confirm = (Read-Host -prompt "Do you want to migrate Shared Drive Membership?  [Y]es or [N]o")
                if ($confirm.ToLower() -eq "y") {
                    $MigrateSharedDriveMembership = $true
                }
                if ($confirm.ToLower() -eq "n") {
                    $MigrateSharedDriveMembership = $false
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
            
            $msg = "INFO: MigrateSharedDriveMembership '$MigrateSharedDriveMembership'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################
          
            $fileName = Split-Path $script:inputFile -Leaf 
            $GoogleSharedDriveConfiguration = New-BT_GoogleSharedDriveConfiguration -AdministrativeUsername $adminUsername `
                -CredentialsFileName $fileName `
                -Credentials $script:jsonFileContent `
                -MigrateSharedDriveMembership $MigrateSharedDriveMembership    

        }
        else {
            $adminUsername = $endpointConfiguration.AdministrativeUsername
            do {
                Write-host -NoNewline "Please enter the file path to the Google service account credentials using JSON file:"

                $result = Get-FileName $global:btWorkingDir -Extension "json"

                #Read CSV file
                try {
                    $script:jsonFileContent = get-content $script:inputFile -raw
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'."
                    Write-Host -ForegroundColor Red  $msg
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message
                    Return -1 
                } 
            }while ($script:jsonFileContent -eq "")
        
            $msg = "INFO: The file path to the JSON file is '$script:inputFile'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $script:selectMigrateSharedDriveMembership = $true
                $confirm = (Read-Host -prompt "Do you want to migrate Shared Drive Membership?  [Y]es or [N]o")
                if ($confirm.ToLower() -eq "y") {
                    $MigrateSharedDriveMembership = $true
                }
                if ($confirm.ToLower() -eq "n") {
                    $MigrateSharedDriveMembership = $false
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
        
            $msg = "INFO: MigrateSharedDriveMembership '$MigrateSharedDriveMembership'."
            Write-Host $msg
            Log-Write -Message $msg 

            $fileName = Split-Path $script:inputFile -Leaf 
            $GoogleSharedDriveConfiguration = New-BT_GoogleSharedDriveConfiguration -AdministrativeUsername $adminUsername `
                -CredentialsFileName $fileName `
                -Credentials $script:jsonFileContent `
                -MigrateSharedDriveMembership $MigrateSharedDriveMembership   
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleSharedDriveConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $GoogleSharedDriveConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1              
        }
    }
    elseif ($endpointType -eq "GoogleDrive") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $domains = (Read-Host -prompt "Please enter the domain or domains (separated by comma)").trim()
            }while ($domains -eq "")
        
            $msg = "INFO: Domain(s) is (are) '$domains'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $GoogleDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleDriveConfiguration' -Property @{              
                "AdminEmailAddress" = $adminUsername;
                "Domains"           = $Domains;
            }
        }
        else {
            $GoogleDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleDriveConfiguration' -Property @{              
                "AdminEmailAddress" = $endpointConfiguration.AdminEmailAddress;
                "Domains"           = $endpointConfiguration.Domains;
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleDriveConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $GoogleDriveConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1              
        }
    }
    elseif ($endpointType -eq "ExchangeServer") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $url = (Read-Host -prompt "Please enter the Exchange Server 2003+ URL").trim()
            }while ($url -eq "")
        
            $msg = "INFO: Exchange Server 2003+ URL is '$url'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "Url"                          = $url
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "Url"                          = $url
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $exchangeOnline2Configuration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $exchangeOnline2Configuration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1             
        }

    }
    elseif ($endpointType -eq "ExchangeOnline2") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $exchangeOnline2Configuration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $exchangeOnline2Configuration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1             
        }

    }    
    elseif ($endpointType -eq "ExchangeOnlinePublicFolder") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $exchangePublicFolderConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangePublicFolderConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $exchangePublicFolderConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangePublicFolderConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $exchangePublicFolderConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $exchangePublicFolderConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1             
        }

    }
    elseif ($endpointType -eq "Office365Groups") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $url = (Read-Host -prompt "Please enter the Office 365 group URL").trim()
            }while ($url -eq "")
        
            $msg = "INFO: Office 365 group URL is '$url'."
            Write-Host $msg
            Log-Write -Message $msg 
        
        
            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $office365GroupsConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.SharePointConfiguration" -Property @{
                "Url"                          = $url;
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $office365GroupsConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.SharePointConfiguration" -Property @{
                "Url"                          = $endpointConfiguration.Url;
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $office365GroupsConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 
                
                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $office365GroupsConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1           
        }
    }
    elseif ($endpointType -eq "OneDrivePro") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $oneDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $oneDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $oneDriveConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $oneDriveConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1                   
        }
    }
    elseif ($endpointType -eq "OneDriveProAPI") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            do {
                $confirm = (Read-Host -prompt "Do you want to use your own Azure Storage account?  [Y]es or [N]o")
                if ($confirm.ToLower() -eq "y") {
                    $microsoftStorage = $false
                }
                else {
                    $microsoftStorage = $true
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

            if (!$microsoftStorage) {
                do {
                    $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
                }while ($azureAccountName -eq "")
        
                $msg = "INFO: Azure storage account name is '$azureAccountName'."
                Write-Host $msg
                Log-Write -Message $msg 

                do {
                    $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key").trim()
                }while ($secretKey -eq "")
        
                $msg = "INFO: Azure storage account access key is '$secretKey'."
                Write-Host $msg
                Log-Write -Message $msg 
            }
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            if ($microsoftStorage) {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $adminUsername;
                    "AdministrativePassword"             = $adminPassword;
                    "UseSharePointOnlineProvidedStorage" = $true
                }
            }
            else {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials" = $true;
                    "AdministrativeUsername"       = $adminUsername;
                    "AdministrativePassword"       = $adminPassword;
                    "AzureStorageAccountName"      = $azureAccountName;
                    "AzureAccountKey"              = $secretKey
                }
            }
        }
        else {
            if ($endpointConfiguration.UseSharePointOnlineProvidedStorage) {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "UseSharePointOnlineProvidedStorage" = $true
                }
            }
            else {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials" = $true;
                    "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"       = $endpointConfiguration.administrativePassword;
                    "AzureStorageAccountName"      = $endpointConfiguration.AzureStorageAccountName;
                    "AzureAccountKey"              = $azureAccountKey
                }
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $oneDriveProAPIConfiguration -ErrorAction Stop
                 
                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $oneDriveProAPIConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1                 
        }
    }
    elseif ($endpointType -eq "SharePoint") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{   
                "Url"                          = $Url;           
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{  
                "Url"                          = $endpointConfiguration.Url;             
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $spoConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 
                
                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $spoConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1                    
        }
    }
    elseif ($endpointType -eq "SharePointOnlineAPI") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            <#
            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")
        
            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key").trim()
            }while ($secretKey -eq "")
        
            $msg = "INFO: Azure storage account access key is '$secretKey'."
            Write-Host $msg
            Log-Write -Message $msg 
            #>
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{   
                "Url"                                = $Url;               
                "UseAdministrativeCredentials"       = $true;
                "AdministrativeUsername"             = $adminUsername;
                "AdministrativePassword"             = $adminPassword;
                #"AzureStorageAccountName" = $azureAccountName;
                #"AzureAccountKey" = $secretKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }
        else {
            if ($endpointConfiguration.UseSharePointOnlineProvidedStorage) {
                $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{   
                    "Url"                                = $endpointConfiguration.Url;              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "UseSharePointOnlineProvidedStorage" = $true 
                }
            }
            else {
                $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{   
                    "Url"                                = $endpointConfiguration.Url;              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "AzureStorageAccountName"            = $endpointConfiguration.AzureStorageAccountName;
                    "AzureAccountKey"                    = $endpointConfiguration.azureAccountKey
                    "UseSharePointOnlineProvidedStorage" = $false 
                }            
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $spoConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $spoConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1                     
        }
 
    }
    elseif ($endpointType -eq "SharePointBeta") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $url = (Read-Host -prompt "Please enter the SharePoint Online Root URL").trim()
            }while ($url -eq "")
        
            $msg = "INFO: SharePoint Online Root URL is '$url'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $script:BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            
            do {
                $confirm = (Read-Host -prompt "Do you want to use Microsoft provided Azure Storage?  [Y]es or [N]o")
                if ($confirm.ToLower() -eq "n") {
    
                    do {
                        $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
                    } while ($AzureStorageAccountName -eq "")

                    $msg = "INFO: Azure storage account name is '$azureAccountName'."
                    Write-Host $msg
                    Log-Write -Message $msg 
        
                    do {
                        $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key").trim()
                    } while ($AzureAccountKey -eq "")

                    $msg = "INFO: Azure storage account access key is '$secretKey'."
                    Write-Host $msg
                    Log-Write -Message $msg 

                    $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointBetaConfiguration' -Property @{   
                        "Url"                                = $Url;               
                        "UseAdministrativeCredentials"       = $true;
                        "AdministrativeUsername"             = $adminUsername;
                        "AdministrativePassword"             = $adminPassword;
                        "AzureStorageAccountName"            = $azureAccountName;
                        "AzureAccountKey"                    = $secretKey
                        "UseSharePointOnlineProvidedStorage" = $false 
                    }
                }
                else {

                    $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointBetaConfiguration' -Property @{   
                        "Url"                                = $Url;               
                        "UseAdministrativeCredentials"       = $true;
                        "AdministrativeUsername"             = $adminUsername;
                        "AdministrativePassword"             = $adminPassword;
                        #"AzureStorageAccountName"            = $AzureStorageAccountName;
                        #"AzureAccountKey"                    = $AzureAccountKey
                        "UseSharePointOnlineProvidedStorage" = $true 
                    }
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
        }
        else {
            if ($endpointConfiguration.UseSharePointOnlineProvidedStorage) {
                $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointBetaConfiguration' -Property @{   
                    "Url"                                = $endpointConfiguration.Url;              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "UseSharePointOnlineProvidedStorage" = $true 
                }
            }
            else {
                $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointBetaConfiguration' -Property @{   
                    "Url"                                = $endpointConfiguration.Url;              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "AzureStorageAccountName"            = $endpointConfiguration.AzureStorageAccountName;
                    "AzureAccountKey"                    = $endpointConfiguration.azureAccountKey
                    "UseSharePointOnlineProvidedStorage" = $false 
                }            
            }
        }

        #####################################################################################################################
        # Create endpoint. 
        #####################################################################################################################

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $spoConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $spoConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1                     
        }
 
    }
    elseif ($endpointType -eq "MicrosoftTeamsSourceParallel") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            <#
            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")
        
            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key").trim()
            }while ($secretKey -eq "")
        
            $msg = "INFO: Azure storage account access key is '$secretKey'."
            Write-Host $msg
            Log-Write -Message $msg 
            #>
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsSourceConfiguration' -Property @{          
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword;
            }
        }
        else {
            $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsSourceConfiguration' -Property @{             
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword;
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1                     
        }
 
    }
    elseif ($endpointType -eq "MicrosoftTeamsDestinationParallel") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            <#
            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")
        
            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key").trim()
            }while ($secretKey -eq "")
        
            $msg = "INFO: Azure storage account access key is '$secretKey'."
            Write-Host $msg
            Log-Write -Message $msg 
            #>
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsDestinationConfiguration' -Property @{          
                "UseAdministrativeCredentials"       = $true;
                "AdministrativeUsername"             = $adminUsername;
                "AdministrativePassword"             = $adminPassword;
                #"AzureStorageAccountName" = $azureAccountName;
                #"AzureAccountKey" = $secretKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }
        else {
            if ($endpointConfiguration.UseSharePointOnlineProvidedStorage) {
                $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsDestinationConfiguration' -Property @{             
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "UseSharePointOnlineProvidedStorage" = $true 
                }
            }
            else {
                $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsDestinationConfiguration' -Property @{             
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "AzureStorageAccountName"            = $endpointConfiguration.AzureStorageAccountName;
                    "AzureAccountKey"                    = $azureAccountKey
                    "UseSharePointOnlineProvidedStorage" = $false 
                }       
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1                    
        }
 
    }
    elseif ($endpointType -eq "Pst") {

        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")

            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key ").trim()
            }while ($secretKey -eq "")


            do {
                $containerName = (Read-Host -prompt "Please enter the container name").trim()
            }while ($containerName -eq "")

            $msg = "INFO: Azure subscription ID is '$containerName'."
            Write-Host $msg
            Log-Write -Message $msg 

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $azureAccountName;     
                "AccessKey"                    = $secretKey;  
                "ContainerName"                = $containerName;       
            }
        }
        else {
            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;  
                "AccessKey"                    = $endpointConfiguration.AccessKey;    
                "ContainerName"                = $endpointConfiguration.ContainerName 
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1       
        }  
    }
    elseif ($endpointType -eq "IMAP") {

        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $hostName = (Read-Host -prompt "Please enter the server name").trim()
            }while ($hostName -eq "")

            $msg = "INFO: Server name is '$hostName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $portNumber = (Read-Host -prompt "Please enter server port").trim()
            }while ($portNumber -eq "" -and (isNumeric($portNumber)))

            do {
                $confirm = (Read-Host -prompt "Is SSL enabled?  [Y]es or [N]o")
                if ($confirm.ToLower() -eq "y") {
                    $UseSsl = $true
                }
                else {
                    $UseSsl = $false
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $imapConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.HostConfiguration' -Property @{  
                "Host"                         = $hostName;     
                "Port"                         = $portNumber; 
                "UseSsl"                       = $UseSsl; 
                "UseAdministrativeCredentials" = $false;       
            }
        }
        else {
            $imapConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.HostConfiguration' -Property @{         
                "Host"                         = $endpointConfiguration.hostName;     
                "Port"                         = $endpointConfiguration.portNumber;  
                "UseSsl"                       = .$endpointConfiguration.UseSsl; 
                "UseAdministrativeCredentials" = $false;   
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $imapConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $imapConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1       
        }  
    }
    elseif ($endpointType -eq "Lotus") {

        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $extractorName = (Read-Host -prompt "Please enter the Lotus Extractor name (bt- identified)").trim()
            }while ($extractorName -eq "")

            $msg = "INFO: Lotus Extractor name is '$extractorName'."
            Write-Host $msg
            Log-Write -Message $msg 

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $imapConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.ExtractorConfiguration' -Property @{  
                "ExtractorName"                = $extractorName;     
                "UseAdministrativeCredentials" = $true;       
            }
        }
        else {
            $imapConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.ExtractorConfiguration' -Property @{         
                "ExtractorName"                = $endpointConfiguration.extractorName;     
                "UseAdministrativeCredentials" = $true;   
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $imapConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $imapConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1       
        }  
    }  
    <#
        elseif(endpointType -eq "WorkMail"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name WorkMailRegion -Value $endpoint.WorkMailRegion

             
        }
        elseif(endpointType -eq "Zimbra"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif(endpointType -eq "ExchangeOnlinePublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            
        }
        elseif(endpointType -eq "ExchangeOnlineUsGovernment"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            
        }
        elseif(endpointType -eq "ExchangeOnlineUsGovernmentPublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            
        }
        elseif(endpointType -eq "ExchangeServer"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            
        }
        elseif(endpointType -eq "ExchangeServerPublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            
        }  
        elseif(endpointType -eq "GoogleDrive"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdminEmailAddress -Value $endpoint.AdminEmailAddress
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            return $endpointCredentials   
        }
        elseif(endpointType -eq "GoogleVault"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdminEmailAddress -Value $endpoint.AdminEmailAddress
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            return $endpointCredentials   
        }
        elseif(endpointType -eq "GroupWise"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name TrustedAppKey -Value "ChangeMe"

            return $endpointCredentials  
        }
        elseif(endpointType -eq "PstInternalStorage") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
   
        }
        elseif(endpointType -eq "OX"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url

        }#>
}

# Function to get endpoint data
Function Get-MSPC_EndpointData {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId,
        [parameter(Mandatory = $true)] [guid]$endpointId
    )

    $script:CustomerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId

    try {
        $endpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Id $endpointId -IsDeleted False -IsArchived False | Select-Object -Property Name, Type -ExpandProperty Configuration   
        
        $msg = "SUCCESS: Endpoint '$($endpoint.Name)' retrieved." 
        write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg  

        if ($endpoint.Type -eq "AzureFileSystem") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername

            return $endpointCredentials        
        }
        elseif ($endpoint.Type -eq "AzureSubscription") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name SubscriptionID -Value $endpoint.SubscriptionID

            return $endpointCredentials
        
        } 
        elseif ($endpoint.Type -eq "BoxStorage") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessToken -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name RefreshToken -Value $administrativePassword
            return $endpointCredentials
        }
        elseif ($endpoint.Type -eq "DropBox") {
            $endpointCredentials = New-Object PSObject
            return $endpointCredentials
        }
        elseif ($endpoint.Type -eq "ExchangeOnline2") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "ExchangeOnlinePublicFolder") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "ExchangeOnlineUsGovernment") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "ExchangeOnlineUsGovernmentPublicFolder") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "ExchangeServer") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "ExchangeServerPublicFolder") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "Gmail") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            if ($script:userMailboxesWithResourceMailboxes -or $script:resourceMailboxes) {
                Export-GoogleResources $endpoint.UseAdministrativeCredentials
            }
            
            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "GSuite") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name CredentialsFileName -Value $endpoint.CredentialsFileName

            if ($script:userMailboxesWithResourceMailboxes -or $script:resourceMailboxes) {
                Export-GoogleResources $endpoint.UseAdministrativeCredentials
            }
            
            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "GoogleSharedDrive") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name CredentialsFileName -Value $endpoint.CredentialsFileName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name VersionsCountToMigrate -Value $endpoint.VersionsCountToMigrate
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name MigrateSharedDriveMembership -Value $endpoint.MigrateSharedDriveMembership
            
            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "GoogleDrive") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdminEmailAddress -Value $endpoint.AdminEmailAddress
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "GoogleVault") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdminEmailAddress -Value $endpoint.AdminEmailAddress
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "GroupWise") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name TrustedAppKey -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "IMAP") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Host -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Port -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseSsl -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "Lotus") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ExtractorName -Value $endpoint.UseAdministrativeCredentials

            $msg = "INFO: Extractor Name '$($endpoint.ExtractorName)'." 
            write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg  

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "Office365Groups") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "OneDrivePro") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "OneDriveProAPI") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $azureAccountKey

            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "OX") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "Pst") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials

            return $endpointCredentials        
        }
        elseif ($endpoint.Type -eq "PstInternalStorage") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials

            return $endpointCredentials        
        }
        elseif ($endpoint.Type -eq "SharePoint") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "SharePointOnlineAPI") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $azureAccountKey

            return $endpointCredentials     
        }        
        elseif ($endpoint.Type -eq "SharePointBeta") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $azureAccountKey

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "MicrosoftTeamsSource") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "MicrosoftTeamsDestination") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $azureAccountKey

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "MicrosoftTeamsSourceParallel") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "MicrosoftTeamsDestinationParallel") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $azureAccountKey

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "WorkMail") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name WorkMailRegion -Value $endpoint.WorkMailRegion

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "Zimbra") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }

    }
    catch {
        $msg = "ERROR: Failed to retrieve endpoint '$($endpoint.Name)' credentials."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
    }
}

# Function to display the workgroups created by the user
Function Select-MSPC_Workgroup {

    #######################################
    # Display all mailbox workgroups
    #######################################

    $workgroupPageSize = 100
    $workgroupOffSet = 0
    $workgroups = @()

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC workgroups..."

    do {
        try {
            #default workgroup in the 1st position
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffset -PageSize 1 -IsDeleted false -CreatedBySystemUserId $script:ticket.SystemUserId )
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC workgroups."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }

        if ($workgroupsPage) {
            $workgroups += @($workgroupsPage)
        }

        $workgroupOffset += 1
    } while ($workgroupsPage)

    $workgroupOffSet = 0

    do { 
        try {
            #add all the workgroups including the default workgroup, so there will be 2 default workgroups
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffSet -PageSize $workgroupPageSize -IsDeleted false | Where-Object { $_.CreatedBySystemUserId -ne $script:ticket.SystemUserId })   
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC workgroups."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
        if ($workgroupsPage) {
            $workgroups += @($workgroupsPage)
            foreach ($Workgroup in $workgroupsPage) {
                Write-Progress -Activity ("Retrieving workgroups (" + $($workgroups.Length - 1) + ")") -Status $Workgroup.Id
            }

            $workgroupOffset += $workgroupPageSize
        }
    } while ($workgroupsPage)

    if ($workgroups -ne $null -and $workgroups.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $($workgroups.Length - 1).ToString() + " Workgroup(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No workgroups found." 
        Exit
    }

    #######################################
    # Prompt for the mailbox Workgroup
    #######################################
    if ($workgroups -ne $null) {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a Workgroup:" 
        Write-Host -Object "INFO: A default workgroup has no name, only Id. Your default workgroup is the number 0 in yellow." 

        for ($i = 0; $i -lt $workgroups.Length; $i++) {
            
            $Workgroup = $workgroups[$i]

            if ([string]::IsNullOrEmpty($Workgroup.Name)) {
                if ($i -eq 0) {
                    $defaultWorkgroupId = $Workgroup.Id.Guid
                    Write-Host -ForegroundColor Yellow -Object $i, "-", $defaultWorkgroupId
                }
                else {
                    if ($Workgroup.Id -ne $defaultWorkgroupId) {
                        Write-Host -Object $i, "-", $Workgroup.Id
                    }
                }
            }
            else {
                Write-Host -Object $i, "-", $Workgroup.Name
            }
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do {
            if ($workgroups.count -eq 1) {
                $msg = "INFO: There is only one workgroup. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $Workgroup = $workgroups[0]
                $global:btWorkgroupOrganizationId = $Workgroup.WorkgroupOrganizationId
                Return $Workgroup.Id
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($workgroups.Length - 1) + ", or x")
            }
            
            if ($result -eq "x") {
                Exit
            }
            if (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $workgroups.Length)) {
                $Workgroup = $workgroups[$result]
                $global:btWorkgroupOrganizationId = $Workgroup.WorkgroupOrganizationId
                Return $Workgroup.Id
            }
        }
        while ($true)

    }

}

# Function to display all customers
Function Select-MSPC_Customer {

    param 
    (      
        [parameter(Mandatory = $true)] [String]$WorkgroupId
    )

    #######################################
    # Display all mailbox customers
    #######################################

    $customerPageSize = 100
    $customerOffSet = 0
    $customers = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC customers..."

    do {   
        try { 
            $customersPage = @(Get-BT_Customer -WorkgroupId $global:btWorkgroupId -IsDeleted False -IsArchived False -PageOffset $customerOffSet -PageSize $customerPageSize)
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC customers."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
    
        if ($customersPage) {
            $customers += @($customersPage)
            foreach ($customer in $customersPage) {
                Write-Progress -Activity ("Retrieving customers (" + $customers.Length + ")") -Status $customer.CompanyName
            }
            
            $customerOffset += $customerPageSize
        }

    } while ($customersPage)

    

    if ($customers -ne $null -and $customers.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $customers.Length.ToString() + " customer(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No customers found." 
        Return "-1"
    }

    #######################################
    # {Prompt for the mailbox customer
    #######################################
    if ($customers -ne $null) {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a customer:" 

        for ($i = 0; $i -lt $customers.Length; $i++) {
            $customer = $customers[$i]
            Write-Host -Object $i, "-", $customer.CompanyName
        }
        Write-Host -Object "b - Go back to workgroup selection menu"
        Write-Host -Object "x - Exit"
        Write-Host

        do {
            if ($customers.count -eq 1) {
                $msg = "INFO: There is only one customer. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $customer = $customers[0]

                try {
                    if ($script:confirmImpersonation) {
                        $script:CustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ImpersonateId $global:btMspcSystemUserId -ErrorAction Stop
                    }
                    else {
                        $script:CustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ErrorAction Stop
                    }
                }
                Catch {
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket under Select-MSPC_Customer()." 
                }

                $global:btcustomerName = $Customer.CompanyName

                Return $customer
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($customers.Length - 1) + ", b or x")
            }

            if ($result -eq "b") {
                Return "-1"
            }
            if ($result -eq "x") {
                Exit
            }
            if (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $customers.Length)) {
                $customer = $customers[$result]
    
                try {
                    if ($script:confirmImpersonation) {
                        $script:CustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ImpersonateId $global:btMspcSystemUserId -ErrorAction Stop
                    }
                    else { 
                        $script:CustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ErrorAction Stop
                    }
                }
                Catch {
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket under Select-MSPC_Customer()." 
                }

                $global:btcustomerName = $Customer.CompanyName

                Return $Customer
            }
        }
        while ($true)

    }

}

# Function to display all endpoints under a customer
Function Select-MSPC_Endpoint {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId,
        [parameter(Mandatory = $false)] [String]$endpointType,
        [parameter(Mandatory = $false)] [String]$endpointName,
        [parameter(Mandatory = $false)] [object]$endpointConfiguration,
        [parameter(Mandatory = $false)] [String]$exportOrImport,
        [parameter(Mandatory = $false)] [String]$projectType,
        [parameter(Mandatory = $false)] [boolean]$deleteEndpointType

    )

    #####################################################################################################################
    # Display all MSPC endpoints. If $endpointType is provided, only endpoints of that type
    #####################################################################################################################

    $endpointPageSize = 100
    $endpointOffSet = 0
    $endpoints = $null

    $sourceMailboxEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "Gmail", "IMAP", "GroupWise", "zimbra", "OX", "WorkMail", "Lotus", "Office365Groups")
    $destinationeMailboxEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "Gmail", "IMAP", "OX", "WorkMail", "Office365Groups", "Pst")
    $sourceStorageEndpointList = @("OneDrivePro", "OneDriveProAPI", "SharePoint", "SharePointOnlineAPI", "GoogleDrive", "AzureFileSystem", "BoxStorage"."DropBox", "Office365Groups")
    $destinationStorageEndpointList = @("OneDrivePro", "OneDriveProAPI", "SharePoint", "SharePointOnlineAPI", "GoogleDrive", "BoxStorage"."DropBox", "Office365Groups")
    $sourceArchiveEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "GoogleVault", "PstInternalStorage", "Pst")
    $destinationArchiveEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "Gmail", "IMAP", "OX", "WorkMail", "Office365Groups", "Pst")
    $sourcePublicFolderEndpointList = @("ExchangeServerPublicFolder", "ExchangeOnlinePublicFolder", "ExchangeOnlineUsGovernmentPublicFolder")
    $destinationPublicFolderEndpointList = @("ExchangeServerPublicFolder", "ExchangeOnlinePublicFolder", "ExchangeOnlineUsGovernmentPublicFolder", "ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment")
    $sourceTeamworkEndpointList = @("MicrosoftTeamsSourceParallel")
    $destinationTeamworkEndpointList = @("MicrosoftTeamsDestinationParallel")

    Write-Host
    if ($endpointType -ne "") {
        Write-Host -Object  "INFO: Retrieving MSPC $exportOrImport $endpointType endpoints..."
    }
    else {
        Write-Host -Object  "INFO: Retrieving MSPC $exportOrImport endpoints..."

        if ($projectType -ne "") {
            switch ($projectType) {
                "Mailbox" {
                    if ($exportOrImport -eq "Source") {
                        $availableEndpoints = $sourceMailboxEndpointList 
                    }
                    else {
                        $availableEndpoints = $destinationeMailboxEndpointList
                    }
                }

                "Storage" {
                    if ($exportOrImport -eq "Source") { 
                        $availableEndpoints = $sourceStorageEndpointList
                    }
                    else {
                        $availableEndpoints = $destinationStorageEndpointList
                    }
                }

                "Archive" {
                    if ($exportOrImport -eq "Source") {
                        $availableEndpoints = $sourceArchiveEndpointList 
                    }
                    else {
                        $availableEndpoints = $destinationArchiveEndpointList
                    }
                }

                "PublicFolder" {
                    if ($exportOrImport -eq "Source") { 
                        $availableEndpoints = $publicfolderEndpointList
                    }
                    else {
                        $availableEndpoints = $publicfolderEndpointList
                    }
                } 

                "Teamwork" {
                    if ($exportOrImport -eq "Source") { 
                        $availableEndpoints = $sourceTeamworkEndpointList
                    }
                    else {
                        $availableEndpoints = $destinationTeamworkEndpointList
                    }
                } 
            }          
        }
    }

    $script:CustomerTicket = Get-BT_Ticket -OrganizationId $global:btCustomerOrganizationId

    do {
        try {
            if ($endpointType -ne "") {
                $endpointsPage = @(Get-BT_Endpoint -Ticket $script:CustomerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize -type $endpointType )
            }
            else {
                $endpointsPage = @(Get-BT_Endpoint -Ticket $script:CustomerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize | Sort-Object -Property Type)
            }
        }

        catch {
            $msg = "ERROR: Failed to retrieve MSPC endpoints."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
            Exit
        }

        if ($endpointsPage) {
            
            $endpoints += @($endpointsPage)

            foreach ($endpoint in $endpointsPage) {
                Write-Progress -Activity ("Retrieving endpoint (" + $endpoints.Length + ")") -Status $endpoint.Name
            }
            
            $endpointOffset += $endpointPageSize
        }
    } while ($endpointsPage)

    Write-Progress -Activity " " -Completed

    if ($endpoints -ne $null -and $endpoints.Length -ge 1) {
        Write-Host -ForegroundColor Green "SUCCESS: $($endpoints.Length) endpoint(s) found."
    }
    else {
        Write-Host -ForegroundColor Red "INFO: No endpoints found." 
    }

    #####################################################################################################################
    # Prompt for the endpoint. If no endpoints found and endpointType provided, ask for endpoint creation
    #####################################################################################################################
    if ($endpoints -ne $null) {


        if ($endpointType -ne "") {
            
            Write-Host -ForegroundColor Yellow -Object "ACTION: Select the $exportOrImport $endpointType endpoint:" 

            for ($i = 0; $i -lt $endpoints.Length; $i++) {
                $endpoint = $endpoints[$i]
                Write-Host -Object $i, "-", $endpoint.Name
            }
        }
        elseif ($endpointType -eq "" -and $projectType -ne "") {

            Write-Host -ForegroundColor Yellow -Object "ACTION: Select the $exportOrImport $projectType endpoint:" 

            for ($i = 0; $i -lt $endpoints.Length; $i++) {
                $endpoint = $endpoints[$i]
                if ($endpoint.Type -in $availableEndpoints) {
                    
                    Write-Host $i, "- Type: " -NoNewline 
                    Write-Host -ForegroundColor White $endpoint.Type -NoNewline                      
                    Write-Host "- Name: " -NoNewline                    
                    Write-Host -ForegroundColor White $endpoint.Name   
                }
            }
        }


        Write-Host -ForegroundColor Yellow -Object "c - Create a new $endpointType endpoint"
        Write-Host -Object "x - Exit"
        Write-Host

        do {
            if ($endpoints.count -eq 1) {
                $result = Read-Host -Prompt ("Select 0, c or x")
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($endpoints.Length - 1) + ", c or x")
            }
            
            if ($result -eq "c") {
                if ($endpointName -eq "") {
                
                    if ($endpointConfiguration -eq $null) {
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType                     
                    }
                    else {
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration          
                    }        
                }
                else {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
                }
                Return $endpointId
            }
            elseif ($result -eq "x") {
                Exit
            }
            elseif (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $endpoints.Length)) {
                $endpoint = $endpoints[$result]
                Return $endpoint.Id
            }
        }
        while ($true)

    } 
    elseif ($endpoints -eq $null -and $endpointType -ne "") {

        do {
            $confirm = (Read-Host -prompt "Do you want to create a $endpointType endpoint ?  [Y]es or [N]o")
        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        if ($confirm.ToLower() -eq "y") {
            if ($endpointName -eq "") {
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration 
            }
            else {
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
            }
            Return $endpointId
        }
    }
}

# Function to
Function Get-CustomerUrlId {
    param 
    (      
        [parameter(Mandatory = $true)] [String]$customerOrganizationId
    )

    $customerUrlId = (Get-BT_Customer -OrganizationId $customerOrganizationId).Id

    Return $customerUrlId

}

# Function to delete all endpoints under a customer
Function Remove-MSPC_Endpoints {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId,
        [parameter(Mandatory = $false)] [String]$endpointType,
        [parameter(Mandatory = $false)] [String]$endpointName
    )

    $endpointPageSize = 100
    $endpointOffSet = 0
    $endpoints = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC $endpointType endpoints matching '$endpointName' endpoint name..."

    do {
        
        $endpointsPage = @(Get-BT_Endpoint -Ticket $script:CustomerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize -type $endpointType)

        if ($endpointsPage) {
            
            $endpoints += @($endpointsPage)

            foreach ($endpoint in $endpointsPage) {
                Write-Progress -Activity ("Retrieving endpoint (" + $endpoints.Length + ")") -Status $endpoint.Name
            }
            
            $endpointOffset += $endpointPageSize
        }
    } while ($endpointsPage)

    

    if ($endpoints -ne $null -and $endpoints.Length -ge 1) {
        Write-Host -ForegroundColor Green "SUCCESS: $($endpoints.Length) endpoint(s) found."
    }
    else {
        Write-Host -ForegroundColor Red "INFO: No endpoints found." 
    }

    $deletedEndpointsCount = 0

    if ($endpoints -ne $null) {
        Write-Host -ForegroundColor Yellow -Object "INFO: Deleting $endpointType endpoints:" 

        for ($i = 0; $i -lt $endpoints.Length; $i++) {
            $endpoint = $endpoints[$i]

            Try {
                if (($endpoint.Name -match "SRC-OD4B-" -and $endpointName -match "SRC-OD4B-") -or `
                    ($endpoint.Name -match "DST-OD4B-" -and $endpointName -match "DST-OD4B-") -or `
                    ($endpoint.Name -match "SRC-SPO-" -and $endpointName -match "SRC-SPO-") -or `
                    ($endpoint.Name -match "DST-SPO-" -and $endpointName -match "DST-SPO-") -or `
                    ($endpoint.Name -match "SRC-PF-" -and $endpointName -match "SRC-PF-") -or `
                    ($endpoint.Name -match "DST-PF-" -and $endpointName -match "DST-PF-") -or `
                    ($endpoint.Name -match "SRC-Teams-" -and $endpointName -match "SRC-Teams-") -or `
                    ($endpoint.Name -match "DST-Teams-" -and $endpointName -match "DST-Teams-") -or `
                    ($endpoint.Name -match "SRC-O365G-" -and $endpointName -match "SRC-O365G-") -or `
                    ($endpoint.Name -match "DST-O365G-" -and $endpointName -match "DST-O365G-")) {

                    remove-BT_Endpoint -Ticket $script:CustomerTicket -Id $endpoint.Id -force
             
                    Write-Host -ForegroundColor Green "SUCCESS: $($endpoint.Name) endpoint deleted." 
                    $deletedEndpointsCount += 1
                }

            }
            catch {
                $msg = "ERROR: Failed to delete endpoint $($endpoint.Name)."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message   
            }            
        }

        if ($deletedEndpointsCount -ge 1 ) {
            Write-Host
            Write-Host -ForegroundColor Green "SUCCESS: $deletedEndpointsCount $endpointType endpoint(s) deleted." 
        }
        elseif ($deletedEndpointsCount -eq 0) {
            Write-Host -ForegroundColor Blue "INFO: No $endpointType endpoint was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1" 
        }
    }
}

# Function to delete all mailbox connectors created by scripts
Function Remove-MW_Connectors {

    param 
    (      
        [parameter(Mandatory = $true)] [guid]$CustomerOrganizationId,
        [parameter(Mandatory = $false)] [String]$ProjectType,
        [parameter(Mandatory = $false)] [String]$exportType,
        [parameter(Mandatory = $false)] [String]$importType,
        [parameter(Mandatory = $false)] [String]$ProjectName
    )
   
    $connectorPageSize = 100
    $connectorOffSet = 0
    $connectors = $null

    if (-not [string]::IsNullOrEmpty($ProjectName)) {
        Write-Host
        Write-Host -Object  "INFO: Retrieving $projectType connectors matching '$ProjectName' project name..."

        do {   

            if ($projectType -eq "Storage") {
                $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:MwTicket -OrganizationId $customerOrganizationId -ProjectType "Storage" -PageOffset $connectorOffSet -PageSize $connectorPageSize | Sort-Object Name | Where-Object { $_.Name -match $ProjectName })
            }
    
            if ($connectorsPage) {
                $connectors += @($connectorsPage)
                foreach ($connector in $connectorsPage) {
                    Write-Progress -Activity ("Retrieving connectors (" + $connectors.Length + ")") -Status $connector.Name
                }
    
                $connectorOffset += $connectorPageSize
            }
    
        } while ($connectorsPage)
    }
    elseif (-not [string]::IsNullOrEmpty($exportType) -and -not [string]::IsNullOrEmpty($importType)) {
        Write-Host
        Write-Host -Object  "INFO: Retrieving $projectType connectors matching '$exportType,$importType' migration scenario..."

        do {   

            if ($projectType -eq "Storage") {
                $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:MwTicket -OrganizationId $customerOrganizationId -ProjectType "Storage" -ExportType $exportType -ImportType $importType -PageOffset $connectorOffSet -PageSize $connectorPageSize | Sort-Object Name)
            }

    
            if ($connectorsPage) {
                $connectors += @($connectorsPage)
                foreach ($connector in $connectorsPage) {
                    Write-Progress -Activity ("Retrieving connectors (" + $connectors.Length + ")") -Status $connector.Name
                }
    
                $connectorOffset += $connectorPageSize
            }
    
        } while ($connectorsPage)
    }
    


    if ($connectors -ne $null -and $connectors.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $connectors.Length.ToString() + " $projectType connector(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No $projectType connectors found." 
        Return
    }

    Write-Host
    do {
        $confirm = (Read-Host -prompt "Are you sure you want to delete all '$exportType,$importType' MigrationWiz projects ?  [Y]es or [N]o")
        if ($confirm.ToLower() -eq "n") {
            Return
        }
    } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    $deletedMailboxConnectorsCount = 0
    $deletedDocumentConnectorsCount = 0
    if ($connectors -ne $null) {
        
        Write-Host -ForegroundColor Yellow -Object "INFO: Deleting $projectType connectors:" 

        for ($i = 0; $i -lt $connectors.Length; $i++) {
            $connector = $connectors[$i]

            Try {
                if ($projectType -eq "Storage") {
                    if ($ProjectName -match "FS-OD4B-" -and $connector.Name -match "FS-OD4B-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif ($ProjectName -match "FS-GoogleDrive-" -and $connector.Name -match "FS-GoogleDrive-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                }       
            }
            catch {
                $msg = "ERROR: Failed to delete $projectType connector $($connector.Name)."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 
            } 
        }

        
        if (($deletedDocumentConnectorsCount -ge 1 -and $projectType -eq "Storage") -or ($deletedMailboxConnectorsCount -ge 1 -and $projectType -eq "Mailbox")) {
            Write-Host
            Write-Host -ForegroundColor Green "SUCCESS: $deletedDocumentConnectorsCount $projectType connector(s) deleted." 
        }
        elseif (($deletedDocumentConnectorsCount -eq 0 -and $projectType -eq "Storage") -or ($deletedMailboxConnectorsCount -eq 0 -and $projectType -eq "Mailbox")) {
            if ($projectName -match "FS-OD4B-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Migrate-MW_AzureBlobContainerToOD4B.ps1."    
            }
            elseif ($projectName -match "FS-DropBox-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-MW_AzureBlobContainerToDropBox.ps1."    
            }    
            elseif ($projectName -match "O365Group-Document-" -or $projectName -match "ClassicSPOSite-Document-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1."   
            } 
            elseif ($projectName -match "OneDrive-Document-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1."   
            }     
            elseif ($projectName -match "O365-Archive-User Mailboxes-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1."   
            }
            elseif ($projectName -match "O365-Mailbox-User Mailboxes-" -or $projectName -match "O365-RecoverableItems-User Mailboxes-" -or $projectName -match "O365Group-Mailbox-All conversations" ) {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1."   
            }
            elseif ($projectName -match "Teams-Collaboration-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1."   
            }  
        }
    }
}

function DownloadGitHubRepository { 
    param( 
        [Parameter(Mandatory = $false)] 
        [string] $Name = "BitTitanMigrationWizFileServerPSTScripting", 
         
        [Parameter(Mandatory = $false)] 
        [string] $Author = "pablogalan1981", 
         
        [Parameter(Mandatory = $False)] 
        [string] $Branch = "main", 
         
        [Parameter(Mandatory = $False)] 
        [string] $Location = "c:\Scripts"
    ) 
    
    Write-Host
    Write-Host "Downloading 4 PowerShell Scripts from GitHub repository:
- Concatenate-Create-Start-FileServerToOD4BHCLVersion.ps1
- Create-FileServerToOD4BMWMigrationProjectsHCLVersion.ps1
- Create-PstToMailboxMWMigrationProjectsHCLVersion.ps1
- Start-MWMigrationsFromCSVFileHCLVersion.ps1"


    Write-Host 'Starting downloading the GitHub Repository'
    try {
        # Force to create a zip file 
        $ZipFile = "$location\$Name.zip"
        New-Item $ZipFile -ItemType File -Force
 
        $RepositoryZipUrl = "https://github.com/$Author/$Name/archive/$Branch.zip" 
        # download the zip     
        Invoke-RestMethod -Uri $RepositoryZipUrl -OutFile $ZipFile
        Write-Host 'Download finished'
 
        #Extract Zip File
        Write-Host 'Starting unzipping the GitHub Repository locally'
        Expand-Archive -Path $ZipFile -DestinationPath $location -Force
        Write-Host 'Unzip finished'
    }
    catch {
        Write-Host
        Write-Host -ForegroundColor Red 'ERROR: executing Invoke-RestMethod.'

        Write-Host
        Write-Host -ForegroundColor Yellow "ACTION: Download all files in https://github.com/pablogalan1981/BitTitanMigrationWizFileServerPSTScripting to $Location"

        Write-Host
        WaitForKeyPress
        Write-Host
    }
     
    # remove the zip file
    Remove-Item -Path $ZipFile -Force
}

Function WaitForKeyPress {
    $msg = "ACTION: If you have downloaded and saved all .ps1 files then press any key to continue. Press 'Ctrl + C' to exit." 
    Write-Host $msg
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}


#######################################################################################################################
#                                   MENU
#######################################################################################################################
# Function to display the main menu
Function Menu {

    #Main menu
    do {
        write-host 
        $msg = "#######################################################################################################################`
                       ACTION SELECTION                 `
#######################################################################################################################"
        Write-Host $msg
        
        $confirm = (Read-Host -prompt "
1. Prepare your BitTitan FileServer Home Directory to OneDrive For Business migration
2. Delete created MigrationWiz endpoints and projects
-----------------------------------------------------------------------------------------------------------------------
3. Exit

Select 1-3")

        if ($confirm -eq 1) {
            $script:createMigrationWizProjects = $true
        }
        elseif ($confirm -eq 2) {
            $script:createMigrationWizProjects = $false
        }
        elseif ($confirm -eq 3) {
            write-Host
            Exit
        }

    } while (!(isNumeric($confirm)) -or $confirm -notmatch '[1-3]')
        
    Return 1
}

$script:createMigrationWizProjects = $true

#######################################################################################################################
#                   MAIN PROGRAM
#######################################################################################################################

Import-MigrationWizPowerShellModule

#######################################################################################################################
#                   CUSTOMIZABLE VARIABLES  
#######################################################################################################################

$updateEndpoint = $true
$updateConnector = $true

###################################################################################################################
$script:srcGermanyCloud = $false
$script:srcUsGovernment = $False

$script:dstGermanyCloud = $False
$script:dstUsGovernment = $false
                        
$ZoneRequirement1 = "NorthAmerica"   #North America (Virginia). For Azure: Both AZNAE and AZNAW.
$ZoneRequirement2 = "WesternEurope"  #Western Europe (Amsterdam for Azure, Ireland for AWS). For Azure: AZEUW.
$ZoneRequirement3 = "AsiaPacific"    #Asia Pacific (Singapore). For Azure: AZSEA
$ZoneRequirement4 = "Australia"      #Australia (Asia Pacific Sydney). For Azure: AZAUE - NSW.
$ZoneRequirement5 = "Japan"          #Japan (Asia Pacific Tokyo). For Azure: AZJPE - Saltiama.
$ZoneRequirement6 = "SouthAmerica"   #South America (Sao Paolo). For Azure: AZSAB.
$ZoneRequirement7 = "Canada"         #Canada. For Azure: AZCAD.
$ZoneRequirement8 = "NorthernEurope" #Northern Europe (Dublin). For Azure: AZEUN.
$ZoneRequirement9 = "China"          #China.
$ZoneRequirement10 = "France"         #France.
$ZoneRequirement11 = "SouthAfrica"    #South Africa.

if ([string]::IsNullOrEmpty($global:btZoneRequirement)) {
    $global:btZoneRequirement = $ZoneRequirement1
}
#######################################################################################################################
#                       SELECT WORKING DIRECTORY  
#######################################################################################################################

Write-Host
Write-Host
Write-Host -ForegroundColor Yellow "      BitTitan File Server to OneDrive For Business migration project creation tool."
Write-Host

write-host 
$msg = "#######################################################################################################################`
                       SELECT WORKING DIRECTORY                  `
#######################################################################################################################"
Write-Host $msg
write-host 

#Working Directorys
$global:btWorkingDir = "C:\scripts"

if (!$global:btCheckWorkingDirectory) {
    do {
        $confirm = (Read-Host -prompt "The default working directory is '$global:btWorkingDir'. Do you want to change it?  [Y]es or [N]o")
        if ($confirm.ToLower() -eq "y") {
            #Working Directory
            $global:btWorkingDir = [environment]::getfolderpath("desktop")
            Get-Directory $global:btWorkingDir            
        }

        $global:btCheckWorkingDirectory = $true

    } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
}

#Logs directory
$logDirName = "LOGS"
$logDir = "$global:btWorkingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format "yyyyMMddTHHmmss")_Create-O365T2TMigrationWizProjects.log"
$script:logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $global:btWorkingDir -logDir $logDir

Write-Host 
Write-Host -ForegroundColor Yellow "WARNING: Minimal output will appear on the screen." 
Write-Host -ForegroundColor Yellow "         Please look at the log file '$($script:logFile)'."
Write-Host -ForegroundColor Yellow "         Generated CSV file will be in folder '$($global:btWorkingDir)'."
Write-Host 
Start-Sleep -Seconds 1

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg 

#######################################################################################################################
#         CONNECTION TO YOUR BITTITAN ACCOUNT 
#######################################################################################################################

write-host 
$msg = "#######################################################################################################################`
                       CONNECTION TO YOUR BITTITAN ACCOUNT                  `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "CONNECTION TO YOUR BITTITAN ACCOUNT" 
write-host 

Connect-BitTitan

#######################################################################################################################
#         INFINITE LOOP FOR MENU
#######################################################################################################################

# keep looping until specified to exit

#Select Action
$action = Menu
if ($action -ne $null) {

    if ($script:createMigrationWizProjects) {
        write-host 
        $msg = "#######################################################################################################################`
                           AZURE CLOUD SELECTION                 `
#######################################################################################################################"
        Write-Host $msg
        Write-Host

        if ($script:srcGermanyCloud) {
            Write-Host -ForegroundColor Magenta "WARNING: Connecting to (source) Azure Germany Cloud." 

            Write-Host
            do {
                $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
                if ($confirm.ToLower() -eq "y") {
                    $script:srcGermanyCloud = $false
                }  
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
            Write-Host 
        }
        elseif ($script:srcUsGovernment ) {
            Write-Host -ForegroundColor Magenta "WARNING: Connecting to (source) Azure Goverment Cloud." 

            Write-Host
            do {
                $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
                if ($confirm.ToLower() -eq "y") {
                    $script:srcUsGovernment = $false
                }  
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
            Write-Host 
        }

        if ($script:dstGermanyCloud) {
            Write-Host -ForegroundColor Magenta "WARNING: Connecting to (destination) Azure Germany Cloud." 

            Write-Host
            do {
                $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
                if ($confirm.ToLower() -eq "y") {
                    $script:dstGermanyCloud = $false
                }  
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
            Write-Host 
        }
        elseif ($script:dstUsGovernment) {
            Write-Host -ForegroundColor Magenta "WARNING: Connecting to (destination) Azure Goverment Cloud." 

            Write-Host
            do {
                $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
                if ($confirm.ToLower() -eq "y") {
                    $script:dstUsGovernment = $false
                }  
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
            Write-Host 
        }

        Write-Host -ForegroundColor Green "INFO: Using Azure $global:btZoneRequirement Datacenter." 

        if ([string]::IsNullOrEmpty($BitTitanAzureDatacenter)) {
            if (!$global:btCheckAzureDatacenter) {
                Write-Host
                do {
                    $confirm = (Read-Host -prompt "Do you want to switch the Azure Datacenter to another region?  [Y]es or [N]o")  
                    if ($confirm.ToLower() -eq "y") {
                        do {
                            $ZoneRequirementNumber = (Read-Host -prompt "`
        1. NorthAmerica   #North America (Virginia). For Azure: Both AZNAE and AZNAW.
        2. WesternEurope  #Western Europe (Amsterdam for Azure, Ireland for AWS). For Azure: AZEUW.
        3. AsiaPacific    #Asia Pacific (Singapore). For Azure: AZSEA
        4. Australia      #Australia (Asia Pacific Sydney). For Azure: AZAUE - NSW.
        5. Japan          #Japan (Asia Pacific Tokyo). For Azure: AZJPE - Saltiama.
        6. SouthAmerica   #South America (Sao Paolo). For Azure: AZSAB.
        7. Canada         #Canada. For Azure: AZCAD.
        8. NorthernEurope #Northern Europe (Dublin). For Azure: AZEUN.
        9. China          #China.
        10. France        #France.
        11. SouthAfrica   #South Africa.

        Select 0-11")
                            switch ($ZoneRequirementNumber) {
                                1 { $ZoneRequirement = 'NorthAmerica' }
                                2 { $ZoneRequirement = 'WesternEurope' }
                                3 { $ZoneRequirement = 'AsiaPacific' }
                                4 { $ZoneRequirement = 'Australia' }
                                5 { $ZoneRequirement = 'Japan' }
                                6 { $ZoneRequirement = 'SouthAmerica' }
                                7 { $ZoneRequirement = 'Canada' }
                                8 { $ZoneRequirement = 'NorthernEurope' }
                                9 { $ZoneRequirement = 'China' }
                                10 { $ZoneRequirement = 'France' }
                                11 { $ZoneRequirement = 'SouthAfrica' }
                            }
                        } while (!(isNumeric($ZoneRequirementNumber)) -or !($ZoneRequirementNumber -in 1..11))

                        $global:btZoneRequirement = $ZoneRequirement
                
                        Write-Host 
                        Write-Host -ForegroundColor Yellow "WARNING: Now using Azure $global:btZoneRequirement Datacenter." 

                        $global:btCheckAzureDatacenter = $true
                    }  
                    if ($confirm.ToLower() -eq "n") {
                        $global:btCheckAzureDatacenter = $true
                    }
                } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
            }
            else {
                Write-Host
                $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different Azure datacenter."
                Write-Host -ForegroundColor Yellow $msg
            }
        }
        else {
            $global:btZoneRequirement = $BitTitanAzureDatacenter
        }

    }

    write-host 
    $msg = "#######################################################################################################################`
                       WORKGROUP, CUSTOMER AND ENDPOINTS SELECTION              `
#######################################################################################################################"
    Write-Host $msg
    Log-Write -Message "WORKGROUP, CUSTOMER AND ENDPOINTS SELECTION" 

    if (!$global:btCheckCustomerSelection) {
        do {
            #Select workgroup
            $global:btWorkgroupId = Select-MSPC_WorkGroup

            #Select customer
            $customer = Select-MSPC_Customer -Workgroup $global:btWorkgroupId
        }
        while ($customer -eq "-1")

        $global:btCustomerOrganizationId = $customer.OrganizationId.Guid
        $global:btCustomerId = $customer.Id.Guid

        $global:btCheckCustomerSelection = $true  
    }
    else {
        Write-Host
        $msg = "INFO: Already selected workgroup '$global:btWorkgroupId' and customer '$global:btcustomerName'."
        Write-Host -ForegroundColor Green $msg

        Write-Host
        $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different workgroups/customers."
        Write-Host -ForegroundColor Yellow $msg
    }

    #$script:customerTicket = Get-BT_Ticket -Ticket $script:ticket -OrganizationId $global:btCustomerOrganizationId 
    #$script:workgroupTicket = Get-BT_Ticket -Ticket $script:ticket -OrganizationId $global:btWorkgroupOrganizationId 

    if (!$script:createMigrationWizProjects) {

        #######################################################################################################################
        #         MIGRATIONWIZ ACCOUNT CLEAN-UP
        #######################################################################################################################
        write-host 
        $msg = "#######################################################################################################################`
                       MIGRATIONWIZ ACCOUNT CLEAN-UP                  `
#######################################################################################################################"
        Write-Host $msg
        Log-Write -Message "MIGRATIONWIZ ACCOUNT CLEAN-UP" 
        Write-Host
    
        Write-Host
        $msg = "INFO: Deleting MigrationWiz projects."
        Write-Host $msg
        Log-Write -Message $msg 
        Write-Host 
        #delete projects

        Remove-MW_Connectors -CustomerOrganizationId $global:btCustomerOrganizationId -ProjectType "Storage" -ExportType "AzureFileSystem" -ImportType "OneDriveProAPI" -ProjectName "FS-OD4B-"

    }
    else { 
        write-host 
        $msg = "#######################################################################################################################`
                       ENDPOINT SELECTION                  `
#######################################################################################################################"
        Write-Host $msg
        Log-Write -Message "ENDPOINT SELECTION" 

        Write-Host
        $msg = "INFO: Getting the connection information for the FileServer Home Directories and Azure Storage Account."
        Write-Host $msg
        Log-Write -Message $msg  

        if (!$global:btExportEndpointId) {
            #Select source endpoint
            $global:btExportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType 'AzureFileSystem'
        }
        else {
            Write-Host
            $msg = "INFO: Already selected 'AzureFileSystem' endpont '$global:btExportEndpointId'."
            Write-Host -ForegroundColor Green $msg

            Write-Host
            $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to use a different 'AzureFileSystem' endpont."
            Write-Host -ForegroundColor Yellow $msg
        }
        if (!$global:btRootPath) {
            #Get folder path to the FileServer root
            Write-host 
            do {    
                do {
                    Write-host  "Enter the folder path to the FileServer root: "  -NoNewline
                    $fileServerPath = Read-Host
                    $global:btRootPath = $fileServerPath
                    #$rootPath = "`'$fileServerPath`'"
    
                } while ($global:btRootPath -eq "")
            
                Write-host -ForegroundColor Yellow  "ACTION: If $global:btRootPath is correct press [C] to continue. If not, press any key to re-enter: " -NoNewline
                $confirm = Read-Host 
    
            } while ($confirm -ne "C")

        }
        else {
            Write-Host
            $msg = "INFO: Already selected folder path to the FileServer root '$global:btRootPath'."
            Write-Host -ForegroundColor Green $msg

            Write-Host
            $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to use a different folder path to the FileServer root."
            Write-Host -ForegroundColor Yellow $msg
        }

        Write-Host
        $msg = "INFO: Getting the connection information for the PST files and Azure Storage Account."
        Write-Host $msg
        Log-Write -Message $msg  

        if (!$global:btExportPstEndpointId) {
            #Select source endpoint
            $global:btExportPstEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType 'Pst'
        }
        else {
            Write-Host
            $msg = "INFO: Already selected 'Pst' endpont '$global:btExportPstEndpointId'."
            Write-Host -ForegroundColor Green $msg

            Write-Host
            $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to use a different 'Pst' endpont."
            Write-Host -ForegroundColor Yellow $msg
        }

        if (!$global:btAzureCredentials) {
            
            if (!$global:btAzureCredentials) {
                #Select source endpoint
                $azureSubscriptionEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointType "AzureSubscription"
                if ($azureSubscriptionEndpointId.count -gt 1) { $azureSubscriptionEndpointId = $azureSubscriptionEndpointId[1] }
        
                if ($azureSubscriptionEndpointId -eq "-1") {    
                    do {
                        $confirm = (Read-Host -prompt "Do you want to skip the Azure Check ?  [Y]es or [N]o")
                        if ($confirm.ToLower() -eq "n") {
                            $skipAzureCheck = $false    
                
                            Write-Host
                            $msg = "ACTION: Provide the following credentials that cannot be retrieved from endpoints:"
                            Write-Host -ForegroundColor Yellow $msg
                            Log-Write -Message $msg 
                
                            Write-Host
                            do {
                                $administrativeUsername = (Read-Host -prompt "Please enter the Azure account email address")
                            }while ($administrativeUsername -eq "")
                        }
                        if ($confirm.ToLower() -eq "y") {
                            $skipAzureCheck = $true
                        }    
                    } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))    
                }
                else {
                    $skipAzureCheck = $false
                }
            }
            else {
                Write-Host
                $msg = "INFO: Already selected 'AzureSubscription' endpoint '$azureSubscriptionEndpointId'."
                Write-Host -ForegroundColor Green $msg
        
                Write-Host
                $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to 'AzureSubscription'."
                Write-Host -ForegroundColor Yellow $msg
            }
        
            if (!$global:btAzureCredentials) {
                #Get source endpoint credentials
                [PSObject]$azureSubscriptionEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $azureSubscriptionEndpointId 
        
                #Create a PSCredential object to connect to Azure Active Directory tenant
                $administrativeUsername = $azureSubscriptionEndpointData.AdministrativeUsername
                    
                if (!$script:AzureSubscriptionPassword) {
                    Write-Host
                    $msg = "ACTION: Provide the following credentials that cannot be retrieved from endpoints:"
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg 
        
                    do {
                        $AzureAccountPassword = (Read-Host -prompt "Please enter the Azure Account Password" -AsSecureString)
                    }while ($AzureAccountPassword -eq "")
                }
                else {
                    $AzureAccountPassword = $script:AzureSubscriptionPassword
                }
        
                $global:btAzureCredentials = New-Object System.Management.Automation.PSCredential ($administrativeUsername, $AzureAccountPassword)
            }
              
            if (!$global:btSecretKey) {
                #Get Azure Storage Primary Access Key source endpoint
                do {
                    $script:secretKeySecureString = (Read-Host -prompt "Please enter the Azure Storage Account Primary Access Key" -AsSecureString)
        
                    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($script:secretKeySecureString)
                    $global:btSecretKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        
                }while ($global:btSecretKey -eq "")
            }
            else {
                Write-Host
                $msg = "INFO: Already selected Azure Storage Account Primary Access Key."
                Write-Host -ForegroundColor Green $msg
    
                Write-Host
                $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to use a different Azure Storage Account Primary Access Key."
                Write-Host -ForegroundColor Yellow $msg
            }

            if (!$global:btAzureSubscriptionID) {
                #Get Azure Storage Primary Access Key source endpoint
                do {
                    $global:btAzureSubscriptionID = (Read-Host -prompt "Please enter the Azure subscription ID").trim()
                }while ($global:btAzureSubscriptionID -eq "")
            }
            else {
                Write-Host
                $msg = "INFO: Already selected Azure subscription ID '$global:btAzureSubscriptionID'."
                Write-Host -ForegroundColor Green $msg

                Write-Host
                $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to use a different Azure Storage Account Primary Access Key."
                Write-Host -ForegroundColor Yellow $msg
            }
        }

        Write-Host
        $msg = "INFO: Getting the connection information for OneDrive For Business."
        Write-Host $msg
        Log-Write -Message $msg  

        if (!$global:btImportEndpointId) {
            #Select destination endpoint
            $global:btImportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType 'OneDriveProAPI'

        }
        else {
            Write-Host
            $msg = "INFO: Already selected 'OneDriveProAPI' endpont '$global:btImportEndpointId'."
            Write-Host -ForegroundColor Green $msg

            Write-Host
            $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to use a different 'OneDriveProAPI' endpont."
            Write-Host -ForegroundColor Yellow $msg
        }

        Write-Host
        $msg = "INFO: Getting the connection information for Exchange Online."
        Write-Host $msg
        Log-Write -Message $msg  

        if (!$global:btImportPstEndpointId) {
            #Select destination PST endpoint

            $global:btImportPstEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType 'ExchangeOnline2'


        }
        else {
            Write-Host
            $msg = "INFO: Already selected 'ExchangeOnline2' endpont '$global:btImportPstEndpointId'."
            Write-Host -ForegroundColor Green $msg

            Write-Host
            $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to use a different 'ExchangeOnline2' endpont."
            Write-Host -ForegroundColor Yellow $msg
        }

        if (!$global:btMigrateToArchive ) {
            Write-Host
            do { 
                $confirm = (Read-Host -prompt "Do you want to migrate to mailbox or to archive?  [M]ailbox or [A]rchive")    
                if ($confirm.ToLower() -eq "a") {
                    $target = "Archive"
                    $global:btMigrateToArchive = $true
                }
                elseif ($confirm.ToLower() -eq "m") {
                    $target = "Mailbox"
                    $global:btMigrateToArchive = $false
                }
            } while (($confirm.ToLower() -ne "m") -and ($confirm.ToLower() -ne "a"))
        }
        else {
            Write-Host
            if ($global:btMigrateToArchive) {
                $msg = "INFO: PST files will be migrated to archive mailbox."
            }
            else {
                $msg = "INFO: PST files will be migrated to mailbox."
            }
            Write-Host -ForegroundColor Green $msg

            Write-Host
            $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to use a different target."
            Write-Host -ForegroundColor Yellow $msg
        }

        if (!$global:btApplyCustomFolderMapping ) {
            Write-Host
            do { 
                $confirm = (Read-Host -prompt "Do you want to migrate the PST files under a folder with the original PST file name?  [Y]es or [N]o")    
                if ($confirm.ToLower() -eq "y") {
                    $global:btApplyCustomFolderMapping = $true
                }
                elseif ($confirm.ToLower() -eq "n") {
                    $global:btApplyCustomFolderMapping = $false
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
        }
        else {
            Write-Host
            $msg = "INFO: PST files under a folder with the original PST file name."        
            Write-Host -ForegroundColor Green $msg

            Write-Host
            $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to change the PST folder mapping."
            Write-Host -ForegroundColor Yellow $msg
        }
    
        write-host 
        $msg = "#######################################################################################################################`
                       DOWNLOADING POWERSHELL SCRIPTS FROM GITHUB                `
#######################################################################################################################"
        Write-Host $msg
        Log-Write -Message "DOWNLOADING POWERSHELL SCRIPTS FROM GITHUB " 

        DownloadGitHubRepository -Location $global:btWorkingDir

        write-host 
        $msg = "#######################################################################################################################`
                        CHANGING DIRECTORY               `
#######################################################################################################################"
        Write-Host $msg
        Log-Write -Message "CHANGING DIRECTORY" 

        write-host
        write-host "PowerShell scripts have been downloaded to '$global:btWorkingDir\BitTitanMigrationWizFileServerPSTScripting-main':"
        Set-Location "$global:btWorkingDir\BitTitanMigrationWizFileServerPSTScripting-main"

        write-host 
        $msg = "#######################################################################################################################`
                        COMMANDLINE TO EXECUTE FILE SERVER ANALYSIS                `
#######################################################################################################################"
        Write-Host $msg
        Log-Write -Message "COMMANDLINE TO EXECUTE FILE SERVER ANALYSIS" 

        write-host
        write-host "Copy this commandline in yellow and execute it from the directory you are now:"

        write-host
        write-host  -ForegroundColor Yellow ".\Analyze-FileServer.ps1"

        write-host 
        $msg = "#######################################################################################################################`
                        COMMANDLINE TO EXECUTE MIGRATIONWIZ AUTOMATION                `
#######################################################################################################################"
        Write-Host $msg
        Log-Write -Message "COMMANDLINE TO EXECUTE MIGRATIONWIZ AUTOMATION" 

        write-host
        write-host "Copy this commandline in yellow and execute it from the directory you are now:"

        write-host
        write-host "For verify credentials but not migration:"
        write-host

        write-host  -ForegroundColor Yellow ".\Concatenate-Create-Start-FileServerToOD4BHCLVersion.ps1 -BitTitanWorkGroupId $global:btWorkgroupId ``
    -BitTitanCustomerId $global:btCustomerId ``
    -WorkingDirectory 'C:\Scripts' ``
    -downloadLatestVersion `$false ``
    -BitTitanAzureDatacenter $global:btZoneRequirement ``
    -AzureStorageAccessKey $global:btSecretKey ``
    -FileServerRootFolderPath '$global:btRootPath' ``
    -BitTitanSourceEndpointId  $global:btExportEndpointId ``
    -BitTitanSourcePstEndpointId  $global:btExportPstEndpointId ``
    -BitTitanDestinationEndpointId $global:btImportEndpointId ``
    -BitTitanDestinationPstEndpointId $global:btImportPstEndpointId ``
    -CheckFileServer `$true ``
    -CheckOneDriveAccounts `$false ``
    -MigrationWizFolderMapping 'MigratedHomeDir' ``
    -OwnAzureStorageAccount `$false ``
    -ApplyUserMigrationBundle `$true ``
    -BitTitanMigrationScope All ``
    -BitTitanMigrationType Verify"  

        write-host
        write-host "For migration:"
        write-host

        write-host  -ForegroundColor Yellow ".\Concatenate-Create-Start-FileServerToOD4BHCLVersion.ps1 -BitTitanWorkGroupId $global:btWorkgroupId ``
    -BitTitanCustomerId $global:btCustomerId ``
    -WorkingDirectory 'C:\Scripts' ``
    -downloadLatestVersion `$false ``
    -BitTitanAzureDatacenter $global:btZoneRequirement ``
    -AzureStorageAccessKey $global:btSecretKey ``
    -AzureSubscriptionID $global:btAzureSubscriptionID ``
    -FileServerRootFolderPath '$global:btRootPath' ``
    -BitTitanSourceEndpointId  $global:btExportEndpointId ``
    -BitTitanSourcePstEndpointId  $global:btExportPstEndpointId ``
    -BitTitanDestinationEndpointId $global:btImportEndpointId ``
    -BitTitanDestinationPstEndpointId $global:btImportPstEndpointId ``
    -CheckFileServer `$true ``
    -CheckOneDriveAccounts `$false ``
    -MigrationWizFolderMapping 'MigratedHomeDir' ``
    -OwnAzureStorageAccount `$false ``
    -ApplyUserMigrationBundle `$true ``
    -MigrateToArchive `$$global:btMigrateToArchive  ``
    -ApplyCustomFolderMapping `$$global:btApplyCustomFolderMapping  ``
    -BitTitanMigrationScope All ``
    -BitTitanMigrationType Full"  

    }
}
#End if($action -ne $null)
else {
    ##END SCRIPT 
    Write-Host

    $msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
    Log-Write -Message $msg

    if ($script:sourceO365Session) {
        try {
            Write-Host "INFO: Opening directory $global:btWorkingDir where you will find all the generated CSV files."
            Invoke-Item $global:btWorkingDir
            Write-Host
        }
        catch {
            $msg = "ERROR: Failed to open directory '$global:btWorkingDir'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
        }

        Remove-PSSession $script:sourceO365Session
        if ($script:destinationO365Session) {
            Remove-PSSession $script:destinationO365Session
        }
    }

    Exit
}



$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg 

##END SCRIPT

