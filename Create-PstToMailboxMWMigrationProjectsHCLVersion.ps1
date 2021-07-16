<#
Copyright 2020 BitTitan, Inc.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
#>

<#
.SYNOPSIS
     .SYNOPSIS
    This script will create a MigrationWiz project to migrate FileServer Home Directories to OneDrive For Business accounts.
    It will generate a CSV file with the MigrationWiz project and all the migrations that will be used by the script 
    Start-MWMigrationsFromCSVFile.ps1 to submit all the migrations.
    Another output 

.DESCRIPTION
    This script will download the UploaderWiz exe file and execute it to create Azure blob containers per each home directory 
    found in the File Server and upload each home directory to the corresponding blob container. After that the script will 
    create the MigrationWiz projects to migrate from Azure blob containers to the OneDrive For Business accounts.
    The output of this script will be 4 CSV files:
    1. PstToMailboxProject-<date>.csv with the project name of the home directory processed that will be passed to Start-MWMigrationsFromCSVFile.ps1
                                          to start automatically all created MigrationWiz projects. 
    2. ProccessedHomeDirectories-<date>.csv with the project names of the home directories processed the same day that will be passed to Get-MWMigrationProjectStatistics.ps1
                                          to generate the migration statistics and error reports of MigrationWiz projects processed that day. 
    3. AllFailedHomeDirectories.csv with ALL the project names of the home directories processed to date.
    4. AllProccessedHomeDirectories.csv with ALL the project names of the home directories failed to process to date.

.PARAMETER WorkingDirectory
    This parameter defines the folder path where the CSV files generated during the script execution will be placed.
    This parameter is optional. If you don't specify an output folder path, the script will output the CSV files to 'C:\scripts'.  
.PARAMETER BitTitanAzureDatacenter
    This parameter defines the Azure data center the MigrationWiz project will be using. These are the accepted Azure data centers:
    'NorthAmerica','WesternEurope','AsiaPacific','Australia','Japan','SouthAmerica','Canada','NorthernEurope','China','France','SouthAfrica'
    This parameter is optional. If you don't specify an Azure data center, the script will use 'NorthAmerica' by default.  
.PARAMETER BitTitanWorkgroupId
    This parameter defines the BitTitan Workgroup Id.
    This parameter is optional. If you don't specify a BitTitan Workgroup Id, the script will display a menu for you to manually select the workgroup.  
.PARAMETER DownloadLatestVersion
    This parameter defines if the script must download the latest version of UploaderWiz before starting the UploaderWiz execution for file server upload. 
    This parameter is optional. If you don't specify the parameter with true value, the script will check if UploaderWIz was downloaded and if it was, 
    it will skip the new download. If UploaderWiz wasn´t previously downloaded it will download it for the first time. 
.PARAMETER BitTitanWorkgroupId
    This parameter defines the BitTitan Workgroup Id.
    This parameter is optional. If you don't specify a BitTitan Workgroup Id, the script will display a menu for you to manually select the workgroup.  
.PARAMETER BitTitanCustomerId
    This parameter defines the BitTitan Customer Id.
    This parameter is optional. If you don't specify a BitTitan Customer Id, the script will display a menu for you to manually select the customer.  
.PARAMETER BitTitanSourceEndpointId
    This parameter defines the BitTitan source endpoint Id.
    This parameter is optional. If you don't specify a source endpoint Id, the script will display a menu for you to manually select the source endpoint.  
    The selected source endpoint Id will be save in a global variable for next executions.
 .PARAMETER AzureStorageAccessKey
    This parameter defines the Azure storage primary access key.
    This parameter is mandatory in unattended execution. If you don't specify the Azure storage primary access key, the script will prompt for it in every execution
    making the unattended execution an interactive execution.  
.PARAMETER BitTitanDestinationEndpointId
    This parameter defines the BitTitan destination endpoint Id.
    This parameter is optional. If you don't specify a destination endpoint Id, the script will display a menu for you to manually select the destination endpoint.  
    The selected destination endpoint Id will be save in a global variable for next executions.
 .PARAMETER FileServerRootFolderPath
    This parameter defines the folder path to the file server root folder holding all end user home directories. 
    This parameter is mandatory in unattended execution. If you don't specify the folder path to the file server root folder, the script will prompt for it in every execution
    making the unattended execution an interactive execution.  
.PARAMETER HomeDirectorySearchPattern
    This parameter defines which projects you want to process, based on the project name search term. There is no limit on the number of characters you define on the search term.
    This parameter is optional. If you don't specify a project search term, all projects in the customer will be processed.
    Example: to process all projects starting with "Batch" you enter '-ProjectSearchTerm Batch'  
.PARAMETER ApplyUserMigrationBundle
    This parameter defines if the migration added to the MigrationWiz project must be licensed with an existing User Migration Bundle.
    This parameter is optional. If you don't specify the parameter with true value, the migration won't be automatically licensed. 

.NOTES
    Author          Pablo Galan Sabugo <pablog@bittitan.com> from the BitTitan Technical Sales Specialist Team
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>

Param
(
    [Parameter(Mandatory = $false)] [String]$WorkingDirectory,
    [Parameter(Mandatory = $false)] [Boolean]$DownloadLatestVersion,
    [Parameter(Mandatory = $false)] [ValidateSet('NorthAmerica', 'WesternEurope', 'AsiaPacific', 'Australia', 'Japan', 'SouthAmerica', 'Canada', 'NorthernEurope', 'China', 'France', 'SouthAfrica')] [String]$BitTitanAzureDatacenter,
    [Parameter(Mandatory = $false)] [String]$BitTitanWorkgroupId,
    [Parameter(Mandatory = $false)] [String]$BitTitanCustomerId,
    [Parameter(Mandatory = $false)] [String]$BitTitanSourceEndpointId,
    [Parameter(Mandatory = $false)] [String]$AzureStorageAccessKey,
    [Parameter(Mandatory = $false)] [String]$AzureSubscriptionID,
    [Parameter(Mandatory = $false)] [String]$BitTitanDestinationEndpointId,
    [Parameter(Mandatory = $false)] [String]$FileServerRootFolderPath,
    [Parameter(Mandatory = $false)] [String]$HomeDirectorySearchPattern,
    [Parameter(Mandatory = $false)] [Object]$HomeDirToUserPrincipalNameMapping,
    [Parameter(Mandatory = $false)] [Boolean]$ApplyUserMigrationBundle,
    [Parameter(Mandatory = $false)] [Boolean]$ApplyCustomFolderMapping
)


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

    if (!(((Get-Module -Name "AzureAD") -ne $null) -or ((Get-InstalledModule -Name "AzureAD" -ErrorAction SilentlyContinue) -ne $null))) {
        Write-Host
        $msg = "INFO: AzureAD PowerShell module not installed."
        Write-Host $msg     
        $msg = "INFO: Installing AzureAD PowerShell module."
        Write-Host $msg

        Sleep 5
    
        try {
            Install-Module -Name AzureAD -force -ErrorAction Stop
        }
        catch {
            $msg = "ERROR: Failed to install AzureAD module. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Write-Host
            $msg = "ACTION: Run this script 'As administrator' to intall the AzureAD module."
            Write-Host -ForegroundColor Yellow $msg
            Exit
        }
        Import-Module AzureAD
    }
}

function Import-MigrationWizPowerShellModule {
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

Function Get-CsvFile {
    Write-Host
    Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import the user email addresses."
    Get-FileName $script:workingDir

    # Import CSV and validate if headers are according the requirements
    try {
        $lines = @(Import-Csv $script:inputFile)
    }
    catch {
        $msg = "ERROR: Failed to import '$inputFile' CSV file. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message   
        Exit   
    }

    # Validate if CSV file is empty
    if ( $lines.count -eq 0 ) {
        $msg = "ERROR: '$inputFile' CSV file exist but it is empty. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Exit
    }

    # Validate CSV Headers
    $CSVHeaders = "UserEmailAddress,FirstName"
    foreach ($header in $CSVHeaders) {
        if ($lines.$header -eq "" ) {
            $msg = "ERROR: '$inputFile' CSV file does not have all the required columns. Required columns are: '$($CSVHeaders -join "', '")'. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg   
            Exit
        }
    }

    Return $lines
 
}

# Function to get a CSV file name or to create a new CSV file
Function Get-FileName {
    param 
    (      
        [parameter(Mandatory = $true)] [String]$initialDirectory

    )

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $script:inputFile = $OpenFileDialog.filename

    if ($OpenFileDialog.filename -ne "") {		    
        $msg = "SUCCESS: CSV file '$($OpenFileDialog.filename)' selected."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg 
    }
    else {
        $msg = "ERROR: CSV file has not been selected. Script aborted"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
        Exit
    }
}
 
# Function to write information to the Log File
Function Log-Write {
    param
    (
        [Parameter(Mandatory = $true)]    [string]$Message
    )

    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
    Add-Content -Path $logFile -Value $lineItem
}
 
# Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory = $true)] [string]$workingDir,
        [parameter(Mandatory = $true)] [string]$logDir,
        [parameter(Mandatory = $false)] [string]$metadataDir
    )
    if ( !(Test-Path -Path $workingDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $workingDir -Force -ErrorAction Stop
            $msg = "SUCCESS: Folder '$($workingDir)' for CSV files has been created."
            Write-Host -ForegroundColor Green $msg
        }
        catch {
            $msg = "ERROR: Failed to create '$workingDir'. Script will abort."
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
    if ( $metadataDir -and !(Test-Path -Path $metadataDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $metadataDir -Force -ErrorAction Stop      

            $msg = "SUCCESS: Folder '$($metadataDir)' for PST metadata files has been created."
            Write-Host -ForegroundColor Green $msg 
        }
        catch {
            $msg = "ERROR: Failed to create log directory '$($metadataDir)'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
        } 
    }
}

# Function to wait for the user to press any key to continue
Function WaitForKeyPress {
    param 
    (      
        [parameter(Mandatory = $true)] [string]$message
    )
    
    Write-Host $message
    Log-Write -Message $message 
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}

# Function to download a file from a URL
Function Download-File {
    param 
    (      
        [parameter(Mandatory = $true)] [String]$url,
        [parameter(Mandatory = $true)] [String]$outFile
    )

    $fileName = $($url.split("/")[-1])
    $folderName = $fileName.split(".")[0]

    $msg = "INFO: Downloading the latest version of '$fileName' agent (~12MB) from BitTitan..."
    Write-Host $msg
    Log-Write -Message $msg 

    #Download the latest version of UploaderWiz from BitTitan server
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    try {
        $result = Invoke-WebRequest -Uri $url -OutFile $outFile
        $msg = "SUCCESS: '$fileName' file downloaded into '$PSScriptRoot'."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg 
    }
    catch {
        $msg = "ERROR: Failed to download '$fileName'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg    
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Unzip-File $outFile 

    #Open the zip file 
    try {
    
        Start-Process -FilePath "$PSScriptRoot\$folderName"


    }
    catch {
        $msg = "ERROR: Failed to open '$PSScriptRoot' folder."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        Exit
    }

    
    # else {
    #     $msg = 
    #     "ERROR: Failed to download  UploaderWiz agent from BitTitan."
    #     Write-Host -ForegroundColor Red  $msg
    #     Log-Write -Message $msg 
    # }
 
}

# Function to unzip a file
Function Unzip-File {
    param 
    (      
        [parameter(Mandatory = $true)] [String]$zipfile
    )

    $folderName = (Get-Item $zipfile).Basename
    $fileName = $($zipfile.split("\")[-1])

    $result = New-Item -ItemType directory -Path $folderName -Force 

    try {
        $result = Expand-Archive $zipfile -DestinationPath $folderName -Force

        $msg = "SUCCESS: '$fileName' file unzipped into '$PSScriptRoot\$folderName'."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg 
    }
    catch {
        $msg = "ERROR: Failed to unzip '$fileName' file."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message
        Exit
    }
}

Function isNumeric($x) {
    $x2 = 0
    $isNum = [System.Int32]::TryParse($x, [ref]$x2)
    return $isNum
}

######################################################################################################################################
#                                    CONNECTION TO BITTITAN
######################################################################################################################################

# Function to authenticate to BitTitan SDK
Function Connect-BitTitan {
    #[CmdletBinding()]

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
    $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'
    
    if (!$script:creds) {
        $credentials = (Get-Credential -Message "Enter BitTitan credentials")
        if (!$credentials) {
            $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Exit
        }
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

# Function to create an endpoint under a customer
# Configuration Table in https://www.bittitan.com/doc/powershell.html#PagePowerShellmspcmd%20
Function Create-MSPC_Endpoint {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$CustomerOrganizationId,
        [parameter(Mandatory = $false)] [String]$endpointType,
        [parameter(Mandatory = $false)] [String]$endpointName,
        [parameter(Mandatory = $false)] [object]$endpointConfiguration,
        [parameter(Mandatory = $false)] [String]$exportOrImport
    )

    $customerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId
    
    if ($endpointType -eq "AzureFileSystem") {
        
        ########################################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        ########################################################################################################################################
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

            if ([string]::IsNullOrEmpty($script:secretKey)) {
                do {
                    $script:secretKey = (Read-Host -prompt "Please enter the Azure storage account access key ").trim()
                }while ($script:secretKey -eq "")
            }

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $azureFileSystemConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $azureAccountName; #Azure Storage Account Name        
                "AccessKey"                    = $script:secretKey; #Azure Storage Account SecretKey         
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
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $customerTicket -Name $endpointName -Type $endpointType -Configuration $azureFileSystemConfiguration 

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
        }    
    }
    elseif ($endpointType -eq "AzureSubscription") {
           
        ########################################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        ########################################################################################################################################
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
                $script:AzureSubscriptionPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($script:AzureSubscriptionPassword -eq "")

            do {
                $global:btAzureSubscriptionID = (Read-Host -prompt "Please enter the Azure subscription ID").trim()
            }while ($global:btAzureSubscriptionID -eq "")

            $msg = "INFO: Azure subscription ID is '$global:btAzureSubscriptionID'."
            Write-Host $msg
            Log-Write -Message $msg 

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureSubscriptionConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $adminUsername;     
                "AdministrativePassword"       = $script:AzureSubscriptionPassword;         
                "SubscriptionID"               = $global:btAzureSubscriptionID
            }
        }
        else {
            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureSubscriptionConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;  
                "AdministrativePassword"       = $endpointConfiguration.AdministrativePassword;    
                "SubscriptionID"               = $endpointConfiguration.SubscriptionID 
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $customerTicket -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration 

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
        }   
    }
    elseif ($endpointType -eq "BoxStorage") {
        ########################################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        ########################################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $boxStorageConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.BoxStorageConfiguration' -Property @{  
            }
        }
        else {
            $boxStorageConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.BoxStorageConfiguration' -Property @{  
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $customerTicket -Name $endpointName -Type $endpointType -Configuration $boxStorageConfiguration 

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
        }  
    }
    elseif ($endpointType -eq "DropBox") {

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

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

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $dropBoxConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }      
    elseif ($endpointType -eq "Gmail") {

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
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
                         
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $GoogleMailboxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleMailboxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "Domains"                      = $Domains;
            }
        }
        else {
            $GoogleMailboxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleMailboxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "Domains"                      = $endpointConfiguration.Domains;
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleMailboxConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif ($endpointType -eq "GoogleDrive") {

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
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
                         
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

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

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleDriveConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif ($endpointType -eq "GoogleDriveCustomerTenant") {

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
                $workingDir = "C:\scripts"
                Get-FileName $workingDir

                #Read CSV file
                try {
                    $jsonFileContent = get-content $script:inputFile 
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'."
                    Write-Host -ForegroundColor Red  $msg
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message
                    Return -1    
                } 
            }while ($jsonFileContent -eq "")
        
            $msg = "INFO: The file path to the JSON file is '$script:inputFile'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################
          
            $GoogleDriveCustomerTenantConfiguration = New-BT_GSuiteConfiguration -AdministrativeUsername $adminUsername `
                -CredentialsFileName $script:inputFile `
                -Credentials $jsonFileContent.ToString()   

        }
        else {
            $adminUsername = $endpointConfiguration.AdministrativeUsername
            do {
                Write-host -NoNewline "Please enter the file path to the Google service account credentials using JSON file:"
                $workingDir = "C:\scripts"
                Get-FileName $workingDir

                #Read CSV file
                try {
                    $jsonFileContent = get-content $script:inputFile 
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'."
                    Write-Host -ForegroundColor Red  $msg
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message
                    Return -1 
                } 
            }while ($jsonFileContent -eq "")
        
            $msg = "INFO: The file path to the JSON file is '$script:inputFile'."
            Write-Host $msg
            Log-Write -Message $msg 
            $GoogleDriveCustomerTenantConfiguration = New-BT_GoogleDriveCustomerTenantConfiguration -AdministrativeUsername $adminUsername `
                -CredentialsFileName $script:inputFile `
                -Credentials $jsonFileContent.ToString()   
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleDriveCustomerTenantConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

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
        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
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
                $script:o365AdminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $script:o365AdminPassword
            }
        }
        else {
            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.AdministrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $exchangeOnline2Configuration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }

    }
    elseif ($endpointType -eq "Office365Groups") {

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

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
                "AdministrativePassword"       = $endpointConfiguration.AdministrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $office365GroupsConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif ($endpointType -eq "OneDrivePro") {

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
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
                $script:adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($script:adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$script:adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $oneDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $script:adminPassword
            }
        }
        else {
            $oneDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.AdministrativePassword
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $oneDriveConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif ($endpointType -eq "OneDriveProAPI") {

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
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
                $administrativePasswordSecureString = (Read-Host -prompt "Please enter the Office 365 admin password" -AsSecureString)
            }while ($administrativePasswordSecureString -eq "") 
    
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($administrativePasswordSecureString)
            $script:adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            do {
                $confirm = (Read-Host -prompt "Do you want to use your own Azure Storage account?  [Y]es or [N]o")
                if ($confirm.ToLower() -eq "y") {
                    $script:microsoftStorage = $false
                }
                else {
                    $script:microsoftStorage = $true
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
            
            if (!$script:microsoftStorage) {
                do {
                    $script:dstAzureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
                }while ($script:dstAzureAccountName -eq "")
            
                $msg = "INFO: Azure storage account name is '$script:dstAzureAccountName'."
                Write-Host $msg
                Log-Write -Message $msg 

                do {
                    $secretKeySecureString = (Read-Host -prompt "Please enter the Azure storage account access key" -AsSecureString)
                }while ($secretKeySecureString -eq "")

                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secretKeySecureString)
                $script:secretKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            }
    
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            if (!$script:microsoftStorage) {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $adminUsername;
                    "AdministrativePassword"             = $script:adminPassword;
                    "AzureStorageAccountName"            = $script:dstAzureAccountName;
                    "AzureAccountKey"                    = $script:secretKey
                    "UseSharePointOnlineProvidedStorage" = $true
                }
            }
            else {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $adminUsername;
                    "AdministrativePassword"             = $script:adminPassword;
                    "UseSharePointOnlineProvidedStorage" = $true
                }
            }            
        }
        else {
            $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                "UseAdministrativeCredentials"       = $true;
                "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"             = $endpointConfiguration.AdministrativePassword;
                #"AzureStorageAccountName" = $endpointConfiguration.AzureStorageAccountName;
                #"AzureAccountKey" = $endpointConfiguration.AzureAccountKey
                "UseSharePointOnlineProvidedStorage" = $true
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $oneDriveProAPIConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif ($endpointType -eq "SharePoint") {
        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

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
                "AdministrativePassword"       = $endpointConfiguration.AdministrativePassword
            }
        }

        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $spoConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif ($endpointType -eq "SharePointOnlineAPI") {

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")
        
            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $script:secretKey = (Read-Host -prompt "Please enter the Azure storage account access key").trim()
            }while ($script:secretKey -eq "")
    
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{   
                "Url"                                = $Url;               
                "UseAdministrativeCredentials"       = $true;
                "AdministrativeUsername"             = $adminUsername;
                "AdministrativePassword"             = $adminPassword;
                #"AzureStorageAccountName" = $azureAccountName;
                #"AzureAccountKey" = $script:secretKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }
        else {
            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{   
                "Url"                                = $endpointConfiguration.Url;              
                "UseAdministrativeCredentials"       = $true;
                "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"             = $endpointConfiguration.AdministrativePassword;
                #"AzureStorageAccountName" = $endpointConfiguration.AzureStorageAccountName;
                #"AzureAccountKey" = $endpointConfiguration.AzureAccountKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $spoConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
 
    }
    elseif ($endpointType -eq "MicrosoftTeamsSource") {

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")
        
            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key").trim()
            }while ($secretKey -eq "")
    
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

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
                "AdministrativePassword"       = $endpointConfiguration.AdministrativePassword;
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
 
    }
    elseif ($endpointType -eq "MicrosoftTeamsDestination") {

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 

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
    
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

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
            $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsDestinationConfiguration' -Property @{             
                "UseAdministrativeCredentials"       = $true;
                "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"             = $endpointConfiguration.AdministrativePassword;
                #"AzureStorageAccountName" = $endpointConfiguration.AzureStorageAccountName;
                #"AzureAccountKey" = $endpointConfiguration.AzureAccountKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
 
    }
    elseif ($endpointType -eq "Pst") {

        ########################################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        ########################################################################################################################################
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

            if ([string]::IsNullOrEmpty($script:secretKey)) {
                do {
                    $script:secretKey = (Read-Host -prompt "Please enter the Azure storage account access key ").trim()
                }while ($script:secretKey -eq "")
            }


            do {
                $containerName = (Read-Host -prompt "Please enter the container name").trim()
            }while ($containerName -eq "")

            $msg = "INFO: Azure container name is '$containerName'."
            Write-Host $msg
            Log-Write -Message $msg 

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $azureAccountName;     
                "AccessKey"                    = $script:secretKey;  
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
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $customerTicket -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration 

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
        }  
    }

    <#
        elseif(endpointType -eq "WorkMail"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name WorkMailRegion -Value $endpoint.WorkMailRegion

             
        }
        elseif(endpointType -eq "Zimbra"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            return $endpointCredentials  
        }
        elseif(endpointType -eq "ExchangeOnlinePublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            
        }
        elseif(endpointType -eq "ExchangeOnlineUsGovernment"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            
        }
        elseif(endpointType -eq "ExchangeOnlineUsGovernmentPublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            
        }
        elseif(endpointType -eq "ExchangeServer"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            
        }
        elseif(endpointType -eq "ExchangeServerPublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            
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
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name TrustedAppKey -Value $endpoint.AdministrativePassword

            return $endpointCredentials  
        }
        elseif(endpointType -eq "IMAP"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Host -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Port -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseSsl -Value $endpoint.AdministrativePassword

            return $endpointCredentials  
        }
        elseif(endpointType -eq "Lotus"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ExtractorName -Value $endpoint.UseAdministrativeCredentials

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
                -OrganizationId $global:btCustomerOrganizationId `
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
                    -OrganizationId $global:btCustomerOrganizationId `
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

# Function to get the ULR to the MSPComplete Customerdashboard
Function Get-CustomerUrlId {
    param 
    (      
        [parameter(Mandatory = $true)] [String]$customerOrganizationId
    )

    $customerUrlId = (Get-BT_Customer -OrganizationId $customerOrganizationId).Id

    Return $customerUrlId

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
                        $script:btCustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ImpersonateId $global:btMspcSystemUserId -ErrorAction Stop
                    }
                    else {
                        $script:btCustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ErrorAction Stop
                    }
                }
                Catch {
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket under Select-MSPC_Customer()." 
                }

                $global:btCustomerName = $Customer.CompanyName

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
                        $script:btCustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ImpersonateId $global:btMspcSystemUserId -ErrorAction Stop
                    }
                    else { 
                        $script:btCustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ErrorAction Stop
                    }
                }
                Catch {
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket under Select-MSPC_Customer()." 
                }

                $global:btCustomerName = $Customer.CompanyName

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

    ########################################################################################################################################
    # Display all MSPC endpoints. If $endpointType is provided, only endpoints of that type
    ########################################################################################################################################

    $endpointPageSize = 100
    $endpointOffSet = 0
    $endpoints = $null

    $sourceMailboxEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "Gmail", "IMAP", "GroupWise", "zimbra", "OX", "WorkMail", "Lotus", "Office365Groups")
    $destinationeMailboxEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "Gmail", "IMAP", "OX", "WorkMail", "Office365Groups", "Pst")
    $sourceStorageEndpointList = @("OneDrivePro", "OneDriveProAPI", "SharePoint", "SharePointOnlineAPI", "GoogleDrive", "GoogleDriveCustomerTenant", "AzureFileSystem", "BoxStorage"."DropBox", "Office365Groups")
    $destinationStorageEndpointList = @("OneDrivePro", "OneDriveProAPI", "SharePoint", "SharePointOnlineAPI", "GoogleDrive", "GoogleDriveCustomerTenant", "BoxStorage"."DropBox", "Office365Groups")
    $sourceArchiveEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "GoogleVault", "PstInternalStorage", "Pst")
    $destinationArchiveEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "Gmail", "IMAP", "OX", "WorkMail", "Office365Groups", "Pst")
    $sourcePublicFolderEndpointList = @("ExchangeServerPublicFolder", "ExchangeOnlinePublicFolder", "ExchangeOnlineUsGovernmentPublicFolder")
    $destinationPublicFolderEndpointList = @("ExchangeServerPublicFolder", "ExchangeOnlinePublicFolder", "ExchangeOnlineUsGovernmentPublicFolder", "ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment")

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
            }          
        }
    }

    $customerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId

    do {
        try {
            if ($endpointType -ne "") {
                $endpointsPage = @(Get-BT_Endpoint -Ticket $customerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize -type $endpointType )
            }
            else {
                $endpointsPage = @(Get-BT_Endpoint -Ticket $customerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize | Sort-Object -Property Type)
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

    ########################################################################################################################################
    # Prompt for the endpoint. If no endpoints found and endpointType provided, ask for endpoint creation
    ########################################################################################################################################
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


        Write-Host -Object "c - Create a new $endpointType endpoint"
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
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType                     
                    }
                    else {
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration          
                    }        
                }
                else {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
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
            if ($confirm.ToLower() -eq "y") {
                if ($endpointName -eq "") {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration 
                }
                else {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
                }
                Return $endpointId
            }
            elseif ($confirm.ToLower() -eq "n") {
                Return -1
            }
        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    }
}

# Function to get endpoint data
Function Get-MSPC_EndpointData {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId,
        [parameter(Mandatory = $true)] [guid]$endpointId
    )

    $script:btCustomerTicket = Get-BT_Ticket -Ticket $script:ticket -OrganizationId $customerOrganizationId

    try {
        $endpoint = Get-BT_Endpoint -Ticket $script:btCustomerTicket -Id $endpointId -IsDeleted False -IsArchived False | Select-Object -Property Name, Type -ExpandProperty Configuration   
        
        if (!$endpoint) {
            $msg = "ERROR: Endpoint '$endpointId' not found." 
            write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg 

            Return -1
        }

        $msg = "SUCCESS: Endpoint '$($endpoint.Name)' '$endpointId' retrieved." 
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
        elseif ($endpoint.Type -eq "GoogleDrive") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdminEmailAddress -Value $endpoint.AdminEmailAddress
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "GoogleDriveCustomerTenant") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name CredentialsFileName -Value $endpoint.CredentialsFileName

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
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseSharePointOnlineProvidedStorage -Value $endpoint.UseSharePointOnlineProvidedStorage

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
        else {
            Return -1
        }

    }
    catch {
        $msg = "ERROR: Failed to retrieve endpoint '$($endpoint.Name)' credentials."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
    }
}


######################################################################################################################################
#                                    CONNECTION TO AZURE
######################################################################################################################################

# Function to connect to Azure
Function Connect-Azure {
    param(
        [Parameter(Mandatory = $true)] [PSObject]$azureCredentials,
        [Parameter(Mandatory = $false)] [String]$subscriptionID
    )

    $msg = "INFO: Connecting to Azure to create a blob container."
    Write-Host $msg
    Log-Write -Message $msg   

    #Load-Module ("Az")

    Try {
        if ($subscriptionID -eq $null) {
            $result = Connect-AzAccount  -Environment "AzureCloud" -ErrorAction Stop -Credential $azureCredentials
        }
        else {
            $result = Connect-AzAccount -Environment "AzureCloud" -SubscriptionId $subscriptionID -ErrorAction Stop -Credential $azureCredentials
        }
    }
    catch {
        $msg = "ERROR: Failed to connect to Azure. You must use multi-factor authentication to access Azure subscription '$subscriptionID'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        
        Try {
            if ($subscriptionID -eq $null) {
                $result = Connect-AzAccount  -Environment "AzureCloud" -ErrorAction Stop 
            }
            else {
                $result = Connect-AzAccount -Environment "AzureCloud" -SubscriptionId $subscriptionID -ErrorAction Stop 
            }
        }
        catch {
            $msg = "ERROR: Failed to connect to Azure. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg   
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message   
            Exit
        }
    }

    try {

        #If there are multiple Azure Subscriptions in the tenant ensure the current context is set correctly.
        $result = Get-AzSubscription -SubscriptionID $subscriptionID | Set-AzContext

        $azureAccount = (Get-AzContext).Account.Id
        $subscriptionName = (Get-AzSubscription -SubscriptionID $subscriptionID).Name
        $msg = "SUCCESS: Connection to Azure: Account: $azureAccount Subscription: '$subscriptionName'."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg   

    }
    catch {
        $msg = "ERROR: Failed to get the Azure subscription. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message   
        Exit
    }
}

# Function to check if AzModule is installed
Function Check-AzModule {
    Try {
        $result = Get-InstalledModule -Name 'Az' -AllVersions -ErrorAction Stop
        if ($result) {
            $msg = "INFO: Ready to execute Azure PowerShell module $($result.moduletype), $($result.version), $($result.name)"
            Write-Host $msg
            Log-Write -Message $msg 
        }
        Else {
            $msg = "INFO: Az module is not installed."
            Write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg 

            Install-Module -Name Az -Scope CurrentUser -Repository PSGallery -AllowClobber -Force
            Import-Module Az

            Try {
                
                $result = Get-InstalledModule -Name 'Az' -AllVersions -ErrorAction Stop
                
                If ($result) {
                    write-information "INFO: Ready to execute PowerShell module $($result.moduletype), $($result.version), $($result.name)"
                }
                Else {
                    $msg = "ERROR: Failed to install and import the Az module. Script aborted."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg    
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $_.Exception.Message 
                    Return
                }
            }
            Catch {
                $msg = "ERROR: Failed to check if the Az module is installed. Script aborted."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg    
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message 
                Return
            }
        }

    }
    Catch {
        $msg = "ERROR: Failed to check if the AzModule module is installed. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg    
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        Return
    } 
}

# Function to check if a blob container exists
Function Check-BlobContainer {
    param 
    (      
        [parameter(Mandatory = $true)] [String]$blobContainerName,
        [parameter(Mandatory = $true)] [PSObject]$storageAccount
    )   

    try {
        $result = Get-AzStorageContainer -Name $blobContainerName -Context $storageAccount.Context -ErrorAction SilentlyContinue

        if ($result) {
            $msg = "SUCCESS: Blob container '$($blobContainerName)' found under the Storage account '$($storageAccount.StorageAccountName)'."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg   
            Return $true
        }
        else {
            Return $false
        }
    }
    catch {
        $msg = "ERROR: Failed to get the blob container '$($blobContainerName)' under the Storage account '$($storageAccount.StorageAccountName)'. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg    
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        Exit

    }
}

# Function to check if a StorageAccount exists
Function Check-StorageAccount {
    param 
    (      
        [parameter(Mandatory = $true)] [String]$storageAccountName,
        [parameter(Mandatory = $true)] [String]$userPrincipalName
    )   

    try {
        $storageAccount = Get-AzStorageAccount -ErrorAction Stop | ? { $_.StorageAccountName -eq $storageAccountName }
        $resourceGroupName = $storageAccount.ResourceGroupName

        $msg = "SUCCESS: Azure storage account '$storageAccountName' found."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg   

        $msg = "SUCCESS: Resource Group Name '$resourceGroupName' found."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg   

    }
    catch {
        $msg = "ERROR: Failed to find the Azure storage account '$storageAccountName'. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        Exit

    }

    if ($storageAccount) {
        Return $storageAccount
    }
    else {
        Return $false
    }
}

# Function to create a Blob Container
Function Create-BlobContainer {
    param 
    (      
        [parameter(Mandatory = $true)] [String]$blobContainerName,
        [parameter(Mandatory = $true)] [PSObject]$storageAccount
    )   

    try {
        $result = New-AzureStorageContainer -Name $blobContainerName -Context $storageAccount.Context -ErrorAction Stop

        $msg = "SUCCESS: Blob container '$($blobContainerName)' created under the Storage account '$($storageAccount.StorageAccountName)'."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg    
    }
    catch {
        $msg = "ERROR: Failed to create blob container '$($blobContainerName)' under the Storage account '$($storageAccount.StorageAccountName)'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg    
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        
        $msg = "ACTION: Create blob container '$($blobContainerName)' under the Storage account '$($storageAccount.StorageAccountName)'. Script aborted."
        Write-Host -ForegroundColor Yellow  $msg

        Exit

    }

}

# Function to create a SAS Token
Function Create-SASToken {
    param 
    (      
        [parameter(Mandatory = $true)] [String]$blobContainerName,
        [parameter(Mandatory = $true)] [String]$BlobName,
        [parameter(Mandatory = $true)] [PSObject]$storageAccount
    )   

    $StartTime = Get-Date
    $EndTime = $startTime.AddHours(8760.0) #1 year

    # Read access - https://docs.microsoft.com/en-us/powershell/module/azure.storage/new-azurestoragecontainersastoken
    $SasToken = New-AzureStorageContainerSASToken -Name $blobContainerName `
        -Context $storageAccount.Context -Permission rl -StartTime $StartTime -ExpiryTime $EndTime
    $SasToken | clip

    # Construnct the URL & Test
    $url = "$($storageAccount.Context.BlobEndPoint)$($blobContainerName)/$($BlobName)$($SasToken)"
    $url | clip

    Return $url
}

# Function to get the licenses of each of the Office 365 users
Function Get-OD4BAccounts {
    param 
    (      
        [parameter(Mandatory = $true)] [Object]$Credentials

    )

    try {
        #Prompt for destination Office 365 global admin Credentials
        $msg = "INFO: Connecting to Azure Active Directory."
        Write-Host $msg
        Log-Write -Message $msg 

        Connect-MsolService -Credential $Credentials -ErrorAction Stop
	    
        $msg = "SUCCESS: Connection to Azure Active Directory."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg 
    }
    catch {
        $msg = "ERROR: Failed to connect to Azure Active Directory."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg 
        Start-Sleep -Seconds 5

        try {
            #Prompt for destination Office 365 global admin Credentials
            $msg = "INFO: Connecting to Azure Active Directory."
            Write-Host $msg
            Log-Write -Message $msg 

            Connect-MsolService -ErrorAction Stop
	    
            $msg = "SUCCESS: Connection to Azure Active Directory."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg 
        }
        catch {
            $msg = "ERROR: Failed to connect to Azure Active Directory."
            Write-Host $msg -ForegroundColor Red
            Log-Write -Message $msg 
            Start-Sleep -Seconds 5
            Exit
        }
    }

    $msg = "INFO: Exporting all users with OneDrive For Business."
    write-Host $msg
    Log-Write -Message $msg 

    $od4bArray = @()

    $allUsers = Get-MsolUser -All | Select-Object UserPrincipalName, primarySmtpAddress -ExpandProperty Licenses 
    ForEach ($user in $allUsers) {
        $userUpn = $user.UserPrincipalName
        $accountSkuId = $user.AccountSkuId
        $services = $user.ServiceStatus

        ForEach ($service in $services) {
            $serviceType = $service.ServicePlan.ServiceType
            $provisioningStatus = $service.ProvisioningStatus

            if ($serviceType -eq "SharePoint" -and $provisioningStatus -ne "disabled") {
                $properties = @{UserPrincipalName = $userUpn
                    Office365Plan                 = $accountSkuId
                    Office365Service              = $serviceType
                    ServiceStatus                 = $provisioningStatus
                    SourceFolder                  = $userUpn.split("@")[0]
                }

                $obj1 = New-Object -TypeName PSObject -Property $properties 

                $od4bArray += $obj1 
                Break
            }            
        }
    }

    $od4bArray = $od4bArray | sort-object UserPrincipalName -Unique

    Return $od4bArray 
}

Function Load-Module ($m) {

    # If module is imported say that and do nothing
    if (Get-Module | Where-Object { $_.Name -eq $m }) {
        write-host "INFO: Module $m is already imported."
    }
    else {

        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object { $_.Name -eq $m }) {
            Import-Module $m -Verbose
        }
        else {

            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object { $_.Name -eq $m }) {
                Install-Module -Name $m -Force -Verbose -Scope CurrentUser
                Import-Module $m -Verbose
            }
            else {

                # If module is not imported, not available and not in online gallery then abort
                write-host -ForegroundColor Red "ERROR: Module $m not imported, not available and not in online gallery, exiting."
                EXIT 1
            }
        }
    }
}

function Invoke-Command() {
    param ( [string]$program = $(throw "Please specify a program" ),
        [string]$argumentString = "",
        [switch]$waitForExit )

    $psi = new-object "Diagnostics.ProcessStartInfo"
    $psi.FileName = $program 
    $psi.Arguments = $argumentString
    $proc = [Diagnostics.Process]::Start($psi)
    if ( $waitForExit ) {
        $proc.WaitForExit();
    }
}

#######################################################################################################################
#                                               MAIN PROGRAM
#######################################################################################################################
Import-PowerShellModules
Import-MigrationWizPowerShellModule

#######################################################################################################################
#                   CUSTOMIZABLE VARIABLES  
#######################################################################################################################
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

write-host 
$msg = "#######################################################################################################################"
Write-Host -ForegroundColor Magenta $msg
write-host 

Write-Host
Write-Host -ForegroundColor Magenta "             BitTitan PST file to mailbox migration project creation tool."
Write-Host

write-host 
$msg = "#######################################################################################################################"
Write-Host -ForegroundColor Magenta $msg
write-host 

write-host 
$msg = "#######################################################################################################################`
                       SELECT WORKING DIRECTORY  '$($HomeDirectorySearchPattern.ToUpper())'                   `
#######################################################################################################################"
Write-Host -ForegroundColor Yellow $msg
write-host 

#Working Directorys
$script:workingDir = "C:\scripts"

if ([string]::IsNullOrEmpty($WorkingDirectory)) {
    if (!$global:btCheckWorkingDirectory) {
        do {
            $confirm = (Read-Host -prompt "The default working directory is '$script:workingDir'. Do you want to change it?  [Y]es or [N]o")
            if ($confirm.ToLower() -eq "y") {
                #Working Directory
                $script:workingDir = [environment]::getfolderpath("desktop")
                Get-Directory $script:workingDir            
            }

            $global:btCheckWorkingDirectory = $true

        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    }
}
else {
    $script:workingDir = $WorkingDirectory
}

#Logs directory
$logDirName = "LOGS"
$logDir = "$script:workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format "yyyyMMddTHHmmss")_Create-FileServerToOD4BMWMigrationProjects.log"
$script:logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $script:workingDir -logDir $logDir

Write-Host 
Write-Host -ForegroundColor Yellow "WARNING: Minimal output will appear on the screen." 
Write-Host -ForegroundColor Yellow "         Please look at the log file '$($script:logFile)'."
Write-Host -ForegroundColor Yellow "         Generated CSV file will be in folder '$($script:workingDir)'."
Write-Host 
Start-Sleep -Seconds 1

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg 

#######################################################################################################################
#         CONNECTION TO YOUR BITTITAN ACCOUNT 
#######################################################################################################################

write-host 
$msg = "#######################################################################################################################`
                       CONNECTION TO YOUR BITTITAN ACCOUNT  '$($HomeDirectorySearchPattern.ToUpper())'                   `
#######################################################################################################################"
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message "CONNECTION TO YOUR BITTITAN ACCOUNT" 
write-host 

Connect-BitTitan

write-host 
$msg = "#######################################################################################################################`
                       AZURE CLOUD SELECTION  '$($HomeDirectorySearchPattern.ToUpper())'                     `
#######################################################################################################################"
Write-Host -ForegroundColor Yellow $msg
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

write-host 
$msg = "#######################################################################################################################`
                       WORKGROUP AND CUSTOMER SELECTION  '$($HomeDirectorySearchPattern.ToUpper())'               `
#######################################################################################################################"
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message "WORKGROUP AND CUSTOMER SELECTION"   

if (-not [string]::IsNullOrEmpty($BitTitanCustomerId)) {
    $global:btCustomerOrganizationId = (Get-BT_Customer | where { $_.id -eq $BitTitanCustomerId }).OrganizationId

    $global:btWorkgroupId = (Get-BT_Customer -OrganizationId $global:btCustomerOrganizationId).WorkgroupId
    $global:btWorkgroupOrganizationId = (Get-BT_WorkGroup -Id $global:btWorkgroupId).WorkgroupOrganizationId
    
    Write-Host
    $msg = "INFO: Selected workgroup '$global:btWorkgroupId' and customer '$global:btCustomerOrganizationId'."
    Write-Host -ForegroundColor Green $msg
}
else {
    if (-not [string]::IsNullOrEmpty($BitTitanWorkgroupId)) {
        $global:btWorkgroupId = $BitTitanWorkgroupId

        #Select customer
        $customer = Select-MSPC_Customer -WorkgroupId $global:btWorkgroupId
        
        $global:btCustomerOrganizationId = $customer.OrganizationId.Guid
                
        Write-Host
        $msg = "INFO: Selected customer '$($customer.CompanyName)'."
        Write-Host -ForegroundColor Green $msg
                
        Write-Progress -Activity " " -Completed
    }
    #[string]::IsNullOrEmpty($BitTitanWorkgroupId) -and [string]::IsNullOrEmpty($BitTitanCustomerId)
    else {
        if (!$global:btCheckCustomerSelection -or !$global:btWorkgroupId -or !$global:btCustomerOrganizationId) {
            do {
                #Select workgroup
                $global:btWorkgroupId = Select-MSPC_WorkGroup
    
                Write-Host
                $msg = "INFO: Selected workgroup '$global:btWorkgroupId'."
                Write-Host -ForegroundColor Green $msg
    
                Write-Progress -Activity " " -Completed
    
                #Select customer
                $customer = Select-MSPC_Customer -WorkgroupId $global:btWorkgroupId
    
                $global:btCustomerOrganizationId = $customer.OrganizationId.Guid
    
                Write-Host
                $msg = "INFO: Selected customer '$($customer.CompanyName)'."
                Write-Host -ForegroundColor Green $msg
    
                Write-Progress -Activity " " -Completed
            }
            while ($customer -eq "-1")
            
            $global:btCheckCustomerSelection = $true  
        }
        else {
            Write-Host
            $msg = "INFO: Already selected workgroup '$global:btWorkgroupId' and customer '$global:btCustomerOrganizationId'."
            Write-Host -ForegroundColor Green $msg
    
            Write-Host
            $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different workgroups/customers."
            Write-Host -ForegroundColor Yellow $msg
        }
    }
}

#Create a ticket for project sharing
try {
    $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -WorkgroupId $global:btWorkgroupId -IncludeSharedProjects
}
catch {
    $msg = "ERROR: Failed to create MigrationWiz ticket for project sharing. Script aborted."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg 
    Exit
}
#Create a ticket for workgroup ticket to apply UMB
try {
    $global:btWorkgroupTicket = Get-BT_Ticket -Ticket $script:ticket -OrganizationId $global:btWorkgroupOrganizationId
}
catch {
    $msg = "ERROR: Failed to create a ticket scoped to workgroup to apply UMB. Script aborted."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg 
    Exit
}

write-host 
$msg = "#######################################################################################################################`
                       UPLOADERWIZ DOWNLOAD AND UNZIPPING  '$($HomeDirectorySearchPattern.ToUpper())'                `
#######################################################################################################################"
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message "AZURE AND FILE SYSTEM ENDPOINT SELECTION"   
Write-Host

$url = "https://api.bittitan.com/secure/downloads/UploaderWiz.zip"   
$outFile = "$PSScriptRoot\UploaderWiz.zip" 
$path = "$PSScriptRoot\UploaderWiz"

$downloadUploaderWiz = $false

if ([string]::IsNullOrEmpty($downloadLatestVersion)) {
    if (!$global:btCheckPath) {
        if (Test-Path $outFile) {
            if (Test-Path $path) {
                $lastWriteTime = (get-Item -Path $path -ErrorAction Stop).LastWriteTime

                do {
                    $confirm = (Read-Host -prompt "UploaderWiz was downloaded on $lastWriteTime. Do you want to download it again?  [Y]es or [N]o")
    
                    if ($confirm.ToLower() -eq "y") {
                        $downloadUploaderWiz = $true
                    }
                    elseif ($confirm.ToLower() -eq "n") {
                        $global:btCheckPath = $true
                    }
    
                } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
            }
            else {
                $downloadUploaderWiz = $true
            }  
        }
        else {
            $downloadUploaderWiz = $true
        }
    }
    else {
        if (Test-Path $outFile) {
            if (Test-Path $path) {
                $lastWriteTime = (get-Item -Path $path -ErrorAction Stop).LastWriteTime

                $msg = "INFO: UploaderWiz was downloaded on $lastWriteTime."
                Write-Host -ForegroundColor Green $msg
    
                Write-Host
                $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to download it again."
                Write-Host -ForegroundColor Yellow $msg
            }
            else {
                $downloadUploaderWiz = $true
            }  
        }
        else {
            $downloadUploaderWiz = $true   
        }
    }
}
else {    
    if (Test-Path $outFile) {
        if (Test-Path $path) {
            $lastWriteTime = (get-Item -Path $path -ErrorAction Stop).LastWriteTime

            $msg = "INFO: UploaderWiz was downloaded on $lastWriteTime."
            Write-Host -ForegroundColor Green $msg

            $downloadUploaderWiz = $downloadLatestVersion
        }
        else {
            $downloadUploaderWiz = $true
        }          
    }
    else {
        $downloadUploaderWiz = $true   
    }    
}

if ($downloadUploaderWiz) {    
    Download-File -Url $url -OutFile $outFile
    $global:btCheckPath = $true
}

write-host 
$msg = "#######################################################################################################################`
                       AZURE AND PST ENDPOINT SELECTION  '$($HomeDirectorySearchPattern.ToUpper())'               `
#######################################################################################################################"
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message "AZURE AND PST ENDPOINT SELECTION"   
Write-Host

$skipAzureCheck = $false

if ([string]::IsNullOrEmpty($AzureStorageAccessKey) -or !$global:btAzureCredentials) {
    $msg = "INFO: Getting the connection information to the Azure Storage Account."
    Write-Host $msg
    Log-Write -Message $msg   
    
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

    if (!$skipAzureCheck) {
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

        if (!$AzureSubscriptionID) {
            if (!$global:btAzureSubscriptionID) {
                do {
                    $global:btAzureSubscriptionID = (Read-Host -prompt "Please enter the Azure Subscription ID").trim()
                }while ($global:btAzureSubscriptionID -eq "")
            }
        }
        else {
            $global:btAzureSubscriptionID = $AzureSubscriptionID
        }

        if ($AzureStorageAccessKey) {
            $script:secretKey = $AzureStorageAccessKey
        }
        else {
            if (!$script:secretKey) {
                Write-Host
                do {
                    $script:secretKeySecureString = (Read-Host -prompt "Please enter the Azure Storage Account Primary Access Key" -AsSecureString)
        
                    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($script:secretKeySecureString)
                    $script:secretKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        
                }while ($script:secretKey -eq "")
            }
        }
    }
}
else {
    $script:secretKey = $AzureStorageAccessKey
}

if (!$global:btExportPstEndpointId) {
    if ([string]::IsNullOrEmpty($BitTitanSourceEndpointId)) {
        #Select source endpoint
        $global:btExportPstEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType "Pst"
        if ($global:btExportPstEndpointId.count -gt 1) { $global:btExportPstEndpointId = $global:btExportPstEndpointId[1] }
    } 
    else {
        $global:btExportPstEndpointId = $BitTitanSourceEndpointId
    }   
}
else {
    Write-Host
    $msg = "INFO: Already selected 'PST' endpoint '$global:btExportPstEndpointId'."
    Write-Host -ForegroundColor Green $msg

    Write-Host
    $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different 'Pst'."
    Write-Host -ForegroundColor Yellow $msg
    Write-Host 
}

#Get source endpoint credentials
[PSObject]$exportEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $global:btExportPstEndpointId 

#if ([string]::IsNullOrEmpty($AzureStorageAccessKey)) {
#if (!$skipAzureCheck) {
write-host 
$msg = "#######################################################################################################################`
                    CONNECTION TO YOUR AZURE ACCOUNT  '$($HomeDirectorySearchPattern.ToUpper())'                   `
#######################################################################################################################"
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message "CONNECTION TO YOUR AZURE ACCOUNT" 
write-host 

if (!$global:btAzureStorageAccountChecked) {
    $msg = "INFO: Checking the Azure Storage Account."
    Write-Host $msg
    Log-Write -Message $msg 
    Write-Host

    # AzModule module installation
    Check-AzModule
    # Azure log in
    if ($azureSubscriptionEndpointData.SubscriptionID) {
        Connect-Azure -AzureCredentials $global:btAzureCredentials -SubscriptionID $azureSubscriptionEndpointData.SubscriptionID
    }
    elseif ($global:btAzureSubscriptionID) {
        Connect-Azure -AzureCredentials $global:btAzureCredentials -SubscriptionID $global:btAzureSubscriptionID
    }
    else {
        $msg = "ERROR: Wrong Azure Subscription ID provided."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg     
        Exit
    }
    #Azure storage account
    $storageAccount = Check-StorageAccount -StorageAccountName $exportEndpointData.AdministrativeUsername -UserPrincipalName $global:btAzureCredentials.UserName

    if ($storageAccount) {
        $global:btAzureStorageAccountChecked = $true  
    }
}
else {
    $msg = "INFO: Already validated Azure Storage account with subscription ID: '$global:btAzureSubscriptionID'."
    Write-Host -ForegroundColor Green $msg
    
    Write-Host
    $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different Azure Storage account."
    Write-Host -ForegroundColor Yellow $msg      
}
#}
#}

write-host 
$msg = "#######################################################################################################################`
                       SELECTING MAILBOXES  '$($HomeDirectorySearchPattern.ToUpper())'                  `
#######################################################################################################################"
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message "SELECTING MAILBOXES" 
write-host 

$msg = "INFO: Creating or selecting existing ExchangeOnline2 endpoint."
Write-Host $msg
Log-Write -Message $msg 

if (!$global:btImportPstEndpointId) {
    if ([string]::IsNullOrEmpty($BitTitanDestinationEndpointId)) {
        #Select destination endpoint
        $global:btImportPstEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "ExchangeOnline2"
        if ($global:btImportPstEndpointId.count -gt 1) { $global:btImportPstEndpointId = $global:btImportPstEndpointId[1] }
    } 
    else {
        $global:btImportPstEndpointId = $BitTitanDestinationEndpointId
    } 
}
else {
    [PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $global:btImportPstEndpointId 
    if ($importEndpointData -eq -1) {
        $global:btImportPstEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "ExchangeOnline2"
        if ($global:btImportPstEndpointId.count -gt 1) { $global:btImportPstEndpointId = $global:btImportPstEndpointId[1] }
    }
}

if (!$global:btO365Credentials) {   
    #Get source endpoint credentials
    if (!$importEndpointData) { [PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $global:btImportPstEndpointId }
    if ($importEndpointData -eq -1) {
        $global:btImportPstEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "ExchangeOnline2"
        if ($global:btImportPstEndpointId.count -gt 1) { $global:btImportPstEndpointId = $global:btImportPstEndpointId[1] }
    }

    write-host
    Connect-Office365

    <#
    if(!$script:adminPassword) {
        do {
            $administrativePasswordSecureString = (Read-Host -prompt "Please enter the Office 365 admin password" -AsSecureString)
        }while ($administrativePasswordSecureString -eq "") 

        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($administrativePasswordSecureString)
        $importAdministrativePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    }
    else{
        $importAdministrativePassword = $script:adminPassword
    }

    #Create a PSCredential object to connect to Azure Active Directory tenant
    $importAdministrativeUsername = $importEndpointData.AdministrativeUsername
    $importAdministrativePassword = ConvertTo-SecureString -String $($importAdministrativePassword) -AsPlainText -Force
    $global:btO365Credentials = New-Object System.Management.Automation.PSCredential ($importAdministrativeUsername, $importAdministrativePassword)
    #>
}

$importAdministrativeUsername = $global:btO365Credentials.Username
$SecureString = $global:btO365Credentials.Password
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
$importAdministrativePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

if (!$CheckOneDriveAccounts -and !$HomeDirToUserPrincipalNameMapping) {
    write-host
    $filterHomeDirs = $false
    do {
        $confirm = (Read-Host -prompt "Do you want to retrieve the OneDrive For Business accounts from Office 365 tenant?  [Y]es or [N]o")

        if ($confirm.ToLower() -eq "y") {
            $filterHomeDirs = $true  

            Write-Host
            $msg = "INFO: Retrieving OneDrive For Business accounts from destination Office 365 tenant."
            Write-Host $msg
            Log-Write -Message $msg 

            $od4bArray = Get-OD4BAccounts -Credentials $global:btO365Credentials  
        }

    } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

}
else {
    if (!$HomeDirToUserPrincipalNameMapping) {
        $filterHomeDirs = $true  
        $od4bArray = Get-OD4BAccounts -Credentials $global:btO365Credentials   
    } 
    else {
        $od4bArray = $HomeDirToUserPrincipalNameMapping

    }
}

$date = (Get-Date -Format yyyyMMddHHmm)
$date2 = (Get-Date -Format yyyyMMdd)

if ($od4bArray) {
    do {
        try {
            if ($filterHomeDirs) {
                #Export users with OD4B to CSV file filtering by the folder names in the root path
                $od4bArray | Select-Object SourceFolder, UserPrincipalName | sort { $_.UserPrincipalName } | where { $filteredHomeDirectories.Name -match $_.SourceFolder } | Export-Csv -Path $script:workingDir\OD4BAccounts-$date.csv -NoTypeInformation -Force
            }
            else {
                $od4bArray | Select-Object SourceFolder, UserPrincipalName | sort { $_.UserPrincipalName } | Export-Csv -Path $script:workingDir\OD4BAccounts-$date.csv -NoTypeInformation -Force
            }
            Break
        }
        catch {
            $msg = "WARNING: Close CSV file '$script:workingDir\OD4BAccounts-$date.csv' open."
            Write-Host -ForegroundColor Yellow $msg

            Start-Sleep 5
        }
    } while ($true)
}
else {

    Write-Host
    $msg = "INFO: Retrieving already processed home directories from File Server."
    Write-Host $msg
    Log-Write -Message $msg 

    foreach ($folderName in $filteredHomeDirectories.Name) {
        [array]$od4bArray += New-Object PSObject -Property @{UserPrincipalName = ''; SourceFolder = $folderName; }
    }

    $od4bArray | Export-Csv -Path $script:workingDir\OD4BAccounts-$date.csv -NoTypeInformation -Force

    $msg = "ACTION: Provide the UserPrincipalName in the CSV file for each home directory processed."
    Write-Host -ForegroundColor Yellow $msg
    Log-Write -Message $msg   
}

if (!$CheckOneDriveAccounts -and !$HomeDirToUserPrincipalNameMapping) {
    #Open the CSV file
    try {
        
        Start-Process -FilePath $script:workingDir\OD4BAccounts-$date.csv

        $msg = "SUCCESS: CSV file '$script:workingDir\OD4BAccounts-$date.csv' processed, exported and open."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg   
    }
    catch {
        $msg = "ERROR: Failed to open '$script:workingDir\OD4BAccounts-$date.csv' CSV file."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message   
        Exit
    }

    WaitForKeyPress -Message "ACTION: If you have edited and saved the CSV file then press any key to continue." 
}

#Re-import the edited CSV file
Try {
    $users = @(Import-CSV "$script:workingDir\OD4BAccounts-$date.csv" | where-Object { $_.PSObject.Properties.Value -ne "" } | sort { $_.UserPrincipalName.length } )
    $totalLines = $users.Count

    if ($totalLines -eq 0) {
        $msg = "INFO: No Office 365 users found with OneDrive For Business matching the home directory names. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        
        if (-not [string]::IsNullOrEmpty($HomeDirectorySearchPattern)) {
            $FailedHomeDirectory = New-Object -TypeName PSObject -Property @{
                "FailedHomeDirectory" = $HomeDirectorySearchPattern
            }
            $FailedHomeDirectory | Export-Csv -Path $script:workingDir\AllFailedHomeDirectories.csv -NoTypeInformation -Append
        }

        $output = "$script:workingDir\PstToMailboxProject-$date.csv"
        
        write-host $output

        Return $output 
    } 
}
Catch [Exception] {
    $msg = "ERROR: Failed to import the CSV file '$script:workingDir\OD4BAccounts-$date.csv'."
    Write-Host -ForegroundColor Red  $msg
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $msg   
    Log-Write -Message $_.Exception.Message   
    Exit
}

write-host 
$msg = "#######################################################################################################################`
                       MIGRATIONWIZ PST->MAILBOX PROJECT CREATION  '$($HomeDirectorySearchPattern.ToUpper())'                   `
#######################################################################################################################"
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message "MIGRATIONWIZ PST->MAILBOX PROJECT CREATION" 
write-host 

#Create Pst-OneDriveProAPI Document project
Write-Host
$msg = "INFO: Creating MigrationWiz PST to Mailbox projects."
Write-Host $msg
Log-Write -Message $msg   
Write-Host

if (-not [string]::IsNullOrEmpty($ApplyUserMigrationBundle) -and $ApplyUserMigrationBundle) {
    $msg = "INFO: Checking User Migration Bundle licenses available in the BitTItan account:"
    Write-Host $msg
    Log-Write -Message $msg   

    #Get the product ID
    #$productId = Get-BT_ProductSkuId -Ticket $script:ticket -ProductName MspcEndUserYearlySubscription
    $productId = '39854d8c-b41d-11e6-a82f-e4b31882dc3b'
    <#If ($productid) {
        $msg = "SUCCESS: Product ID obtained..."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
    Else {
        $msg = "ERRO: Invalid Product ID"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Break
    }#>

    # Validate if the account have the required subscriptions
    # On this query, all the expired Subscriptions licenses will be excluded
    $curDate = Get-Date
    $licensesPacks = @(Get-MW_LicensePack -Ticket $script:MwTicket -WorkgroupOrganizationId $global:btWorkgroupOrganizationId -ProductSkuId $productId  -RetrieveAll | Where-Object { $_.ExpireDate -gt $curDate } | Where-Object { $_.ExpireDate -gt $curDate } | Where-Object { ($_.Purchased + $_.Granted) -gt ($_.Revoked + $_.Used) })

    if (!($licensesPack)) {
        $msg = "ERROR: No valid license pack found under this BitTitan Workgroup / Account"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
    }
    else {
        $licensesPacksAvailable = $licensesPack.Count 
        $msg = "INFO: $licensesPacksAvailable User Migration Bundle license packages found under this MSPC Workgroup / Account"
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg

        $remainingLicenses = 0
        foreach ($licensesPack in $licensesPacks) {
            $remainingLicenses += (($licensesPack.Purchased + $licensesPack.Granted) - ($licensesPack.Revoked + $licensesPack.Used))
        }

        $msg = "INFO: $remainingLicenses total User Migration Bundle licenses found under this MSPC Workgroup / Account"
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }           
}

Write-Host

$processedLines = 0
$existingMigrationList = @()
$PstToMailboxProject = @()

foreach ($user in $users) {    

    $containerName = $user.SourceFolder
    $containerName = $containerName.ToLower()

    Create-Working-Directory -workingDir $workingDir -logDir $logDir -metadataDir "$script:workingDir\PSTMetadata-$containerName\"

    write-host 
    $msg = "#######################################################################################################################`
                    ANALYZING PST FILES IN HOME DIRECTORY  '$($HomeDirectorySearchPattern.ToUpper())'                   `
#######################################################################################################################"
    Write-Host -ForegroundColor Yellow $msg
    Log-Write -Message "GENERATING PST ASSESSMENT REPORT " 
    write-host 

    # AzModule module installation
    Check-AzModule
                    
    $StorageContext = New-AzStorageContext -StorageAccountName $exportEndpointData.AdministrativeUsername -StorageAccountKey $script:secretKey
        
    if (!$storageAccount) {
        #Azure storage account
        Write-Host
        $storageAccount = Check-StorageAccount -StorageAccountName $exportEndpointData.AdministrativeUsername -UserPrincipalName $exportEndpointData.AdministrativeUsername
    }
        
    $result = Check-BlobContainer -BlobContainerName $containerName -StorageAccount $storageAccount
    if ($result) {
        Write-Host
        $msg = "INFO: Downloading metadata files from the Azure Blob Containers 'migrationwizpst'."
        Write-Host $msg
        Log-Write -Message $msg 
        Write-Host

        Write-Progress -Activity " " -Completed
        
        $result = Get-AzStorageBlob -Context $StorageContext -Container $containerName | Where-Object { $_.Name -like "*metadata" } | Get-AzStorageBlobContent -Destination "$script:workingDir\PSTMetadata-$containerName" -force
    }

    Write-Progress -Activity " " -Completed

    $split_variables = ","
    # This word will be used to split the PST Files when more than one exists.
    # The metadata file is not very well structured
    $split_entries = "UpdatedInTicks"
    $output = @()
    
    $fileList = Get-ChildItem "$script:workingDir\PSTMetadata-$containerName" -Filter "*-file.metadata"
    
    
    $msg = "SUCCESS: $($fileList.count) metadata file(s) found in Azure Blob Container and downloaded to '$script:workingDir\PSTMetadata-$containerName\'."
    Write-Host $msg -ForegroundColor Green
    Log-Write -Message $msg 
    
    ForEach ( $file in $fileList ) {   
        
        #Count the number of PST Files will be based on the occurences of the 'UpdatedInTicks' field
        $workXML = Get-Content "$script:workingDir\PSTMetadata-$containerName\$($file.name)"
        
        if ( $workXML -ne "[]" ) {
            $numberOfPSTs = (Select-String -InputObject $workXML -Pattern $split_entries -AllMatches).Matches.Count
    
            $msg = "INFO: $numberOfPSTs file(s) found in '$($file.Name)'."
            Write-Host $msg 
            Log-Write -Message $msg 
            
            #Split PSTs entries and remove all the double quotes because they are not needed
            $PST_entries = $workXML -Split $split_entries -replace """"
            for ( $Count = 0; $Count -lt ($PST_entries).Count - 1; $Count++ ) {
                #using .split and -split give different results
                $Variables = ($PST_entries[$Count]).split($split_variables)
                foreach ( $value in $Variables ) {
                    $key = $value -split (":")
                    switch ($key[0]) {
                        "SourcePath" {
                            $FullPath = $value -replace "SourcePath:"
                            $FullPath = $FullPath.replace("\\", "\") 
                            $FullPathElements = $value -split ("\\")  
                            $OriginalPSTName = $FullPathElements[$FullPathElements.count - 1]
                        }
                        "DestinationPath" {
                            $PSTName = $key[1]
                        }
                        "Owner" {
                            $owner = $key[1]
                            $owner = $owner.replace("\\", "\") 
                        }
                        "OwnerUpn" {
                            $ownerUPN = $key[1]
                        }
                        "SizeInBytes" {
                            $Size_MB = ($key[1]) / 1024 / 1024
                        }
                    }
                    if ($Size_MB -ne 0 -and $FullPath -match ".pst"  ) {
                        $lineItem = @{    
                            OriginalPSTFolderPath = $FullPath              
                            OriginalPSTName       = $OriginalPSTName    
                            Size_MB               = "{0:N2}" -f $Size_MB
                            Owner                 = $owner
                            OwnerUpn              = $ownerUPN                
                            PSTName               = $PSTName                    
                        }
                    }
                }
                if ($Size_MB -ne 0 -and $FullPath -match ".pst") {
                    $output += New-Object psobject -Property $lineItem
                }
            }
    
            if ($output.Count -ge 1) {
                $msg = "SUCCESS: $($output.Count) PST file(s) found."
                Write-Host $msg -ForegroundColor Green
                Log-Write -Message $msg 
            }
        }

    }

    if ($output) {
    
        Write-Host
        $ProjectName = "PST-Mailbox-$containerName" #-$(Get-Date -Format yyyyMMddHHmm)
        $ProjectType = "Archive"   
        $exportType = "Pst" 
        $importType = "ExchangeOnline2"
    
        

        $exportTypeName = "MigrationProxy.WebApi.AzureConfiguration"
        $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
            "AdministrativeUsername"       = $exportEndpointData.AdministrativeUsername;
            "AccessKey"                    = $script:SecretKey;
            "ContainerName"                = $containerName;
            "UseAdministrativeCredentials" = $true
        }

        $importTypeName = "MigrationProxy.WebApi.ExchangeConfiguration"
        $importConfiguration = New-Object -TypeName $importTypeName -Property @{
            "AdministrativeUsername"       = $importAdministrativeUsername;
            "AdministrativePassword"       = $importAdministrativePassword;
            "UseAdministrativeCredentials" = $true;           
        }

   
        #$advancedOptions = "InitializationTimeout=8 RenameConflictingFiles=1 ShrinkFoldersMaxLength=200"
        $advancedOptions = "UseEwsImportImpersonation=1"
    
        $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
            -ProjectName $ProjectName `        -ProjectType $ProjectType `
            -importType $importType `
            -exportType $exportType `
            -exportEndpointId $global:btExportPstEndpointId `
            -importEndpointId $global:btImportPstEndpointId `
            -exportConfiguration $exportConfiguration `
            -importConfiguration $importConfiguration `
            -advancedOptions $advancedOptions `
            -maximumSimultaneousMigrations $totalLines `
            -ZoneRequirement $global:btZoneRequirement

        $msg = "INFO: Adding advanced options '$advancedOptions' to the project."
        Write-Host $msg
        Log-Write -Message $msg   
    
        $msg = "INFO: Adding PST migrations to the project:"
        Write-Host $msg
        Log-Write -Message $msg    

        $SourceFolder = $user.SourceFolder
        $importEmailAddress = $user.UserPrincipalName 

        foreach ($pstFile in $output) {

            $pstFileName = $pstFile.OriginalPSTName
            $pstFilePath = $pstFile.OriginalPSTFolderPath 
            $pstFilePath = $pstFilePath.replace( "$FileServerRootFolderPath\", "")
            $pstFilePath = $pstFilePath.TrimStart("\")
            $pstFilePath = $pstFilePath -ireplace( "$containerName", "")
            $pstFilePath = $pstFilePath.TrimStart("\")
            $pstFilePath = $pstFilePath -replace "\\", "/"
        
            $folderMapping = ""
            #Double Quotation Marks
            [string]$CH34 = [CHAR]34
            if ($applyCustomFolderMapping) {

                $folderMapping = "FolderMapping=" + $CH34 + "^->$($pstFileName -replace('.pst',''))/" + $destinationFolder + $CH34
            }

            $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -PublicFolderPath $pstFilePath -ImportEmailAddress $importEmailAddress -ErrorAction SilentlyContinue
            if (!$result) {
                try {
                    $suppressOutput = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -PublicFolderPath $pstFilePath -ImportEmailAddress $importEmailAddress -advancedOptions $folderMapping

                    $tab = [char]9
                    $msg = "SUCCESS: PST migration '$pstFilePath->$importEmailAddress' of HomeDir '$SourceFolder' added to the project $ProjectName."
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg   

                    $ProcessedLines += 1

                    [array]$PstToMailboxProject += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = 'Archive'; ConnectorId = $connectorId; MailboxId = $result.id; PstFilePath = $pstFilePath; EmailAddress = $importEmailAddress; CreateDate = $(Get-Date -Format yyyyMMddHHmm) } 
                }
                catch {
                    $msg = "ERROR: Failed to add PST file and destination primary SMTP address." 
                    write-Host -ForegroundColor Red $msg
                    Log-Write -Message $msg     
                    Exit
                }
            }
            else {
                $msg = "WARNING: PST file to mailbox migration '$pstFilePath->$importEmailAddress' already exists in connector."  
                write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg  

                $existingMigrationList += "'$pstFilePath->$importEmailAddress'`n"
                $existingMigrationCount += 1

                [array]$PstToMailboxProject += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = 'Archive'; ConnectorId = $connectorId; MailboxId = $result.id; PstFilePath = $pstFilePath; EmailAddress = $importEmailAddress; CreateDate = $(Get-Date -Format yyyyMMddHHmm) } 
            }
        }

        if ($SourceFolder -ne "" -and $importEmailAddress -ne "") {
             

            ########################################################################    
            #                            APPLY UMB
            ######################################################################## 
            if (-not [string]::IsNullOrEmpty($ApplyUserMigrationBundle) -and $ApplyUserMigrationBundle) {

                $msg = "INFO: Checking and applying User Migration Bundle."
                Write-Host $msg
                Log-Write -Message $msg    

                ########################################################################    
                # If mailbox is added and was previously licensed    
                ######################################################################## 
                do {
                    $CustomerEndUserId = $result.CustomerEndUserId
                    if (!$CustomerEndUserId) {
                        $CustomerEndUserId = (Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportEmailAddress $importEmailAddress -ErrorAction SilentlyContinue).CustomerEndUserId
                    }
                    $mspcUser = (Get-BT_CustomerEndUser -Ticket $script:ticket -OrganizationId $global:btCustomerOrganizationId -id $CustomerEndUserId -Environment "BT" -IsDeleted $false -ErrorAction SilentlyContinue) 

                    if (!$mspcUser) {
                        Write-host -ForegroundColor Yellow "WARNING: User '$importEmailAddress' not found in MSPComplete yet. Waiting for replication."
                        Start-Sleep -Seconds 10 
                    }
                    else {
                        Break
                    }
                }
                while (!$mspcUser)    

                if (!$mspcUser) {
                    Write-host -ForegroundColor Red "ERROR: User '$importEmailAddress' not found in MSPComplete."
                }

                if ($mspcUser) {
                    $subscriptionEndDate = (Get-BT_Subscription -Ticket $script:ticket -Id $mspcuser.SubscriptionId.guid).SubscriptionEndDate

                    if ( $mspcuser.ActiveSubscriptionId -eq "00000000-0000-0000-0000-000000000000" ) {
                        Write-host -ForegroundColor Yellow "WARNING: User '$($mspcuser.PrimaryEmailAddress)' does not have a subscription applied."

                        $isUmbApplied = $false
                    }
                    else {
                        Write-host -ForegroundColor Green "SUCCESS: User '$($mspcuser.PrimaryEmailAddress)' has a subscription applied that will expire in '$subscriptionEndDate'. "

                        $isUmbApplied = $true
                    } 

                    if (!$isUmbApplied) {

                        Try {
                            $subscription = Add-BT_Subscription -ticket $global:btWorkgroupTicket -ReferenceEntityType CustomerEndUser -ReferenceEntityId $mspcuser.Id -ProductSkuId $productId -WorkgroupOrganizationId $global:btWorkgroupOrganizationId -ErrorAction Stop
                        
                            $msg = "SUCCESS: User Migration Bundle subscription assigned to MSPC User '$($mspcUser.PrimaryEmailAddress)' and migration '$SourceFolder->$importEmailAddress'."
                            Write-Host -ForegroundColor Blue  $msg
                            Log-Write -Message $msg 

                            $changeCount += 1 
                        }
                        Catch {
                            $msg = "ERROR: Failed to assign User Migration License subscription to MSPC User '$($mspcUser.PrimaryEmailAddress)'."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg
                            Write-Host -ForegroundColor Red $($_.Exception.Message)
                            Log-Write -Message $($_.Exception.Message) 
                        }
                    }    
                }
            }

        }
        else {
            if ($SourceFolder -eq "") {
                $msg = "ERROR: Missing source folder in the CSV file. Skipping '$importEmailAddress' user processing."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg   
                Continue    
            } 
            if ($importEmailAddress -eq "") {
                $msg = "ERROR: Missing destination OneDrive For Business email address in the CSV file. Skipping '$SourceFolder' source folder processing."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg   
                Continue    
            }         
        }              
    }
    else {     
        $msg = "ERROR: No PST file information found on file $($file.Name). Script will skip."
        Write-Host -ForegroundColor Red $msg
    }
}

if ($ProcessedLines -gt 0) {
    write-Host
    $msg = "SUCCESS: $ProcessedLines out of $totalLines migrations have been processed." 
    write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg 
}
if (-not ([string]::IsNullOrEmpty($existingMigrationList)) -and $existingMigrationCount -ne 0 ) {
    write-Host
    $msg = "WARNING: $existingMigrationCount out of $totalLines migrations have not been added because they already exist: `n$existingMigrationList" 
    write-Host -ForegroundColor Yellow $msg
    Log-Write -Message $msg 
}

if ([string]::IsNullOrEmpty($ApplyUserMigrationBundle)) {
    $customerUrlId = Get-CustomerUrlId -CustomerOrganizationId $global:btCustomerOrganizationId

    $url = "https://manage.bittitan.com/customers/$customerUrlId/users?qp_currentWorkgroupId=$workgroupId"

    Write-Host
    $msg = "ACTION: Apply User Migration Bundle licenses to the OneDrive For Business email addresses in MSPComplete."
    Write-Host -ForegroundColor Yellow $msg
    Log-Write -Message $msg   
    $msg = "INFO: Opening '$url' in your default web browser."
    Write-Host $msg
    Log-Write -Message $msg   

    $result = Start-Process $url
    Start-Sleep 5
    WaitForKeyPress -Message "ACTION: If you have applied the User Migration Bundle to the users, press any key to continue"
    Write-Host

    $url = "https://migrationwiz.bittitan.com/app/projects" + "?qp_currentWorkgroupId=$workgroupId"  

    Write-Host
    $msg = "INFO: MigrationWiz projects created."
    Write-Host $msg
    Log-Write -Message $msg   
    $msg = "INFO: Opening '$url' in your default web browser."
    Write-Host $msg
    Log-Write -Message $msg  

    $result = Start-Process $url
    Start-Sleep 5

    Write-Host
    $msg = "INFO: Opening the CSV file that will be used by 'Start-MWMigrationsFromCSVFile.ps1' script."
    Write-Host $msg
    Log-Write -Message $msg  
} 

do {    
    try {
        #export the project info to CSV file
        $PstToMailboxProject | Select-Object ProjectName, ProjectType, ConnectorId, MailboxId, SourceFolder, EmailAddress | sort { $_.UserPrincipalName } | Export-Csv -Path $script:workingDir\PstToMailboxProject-$date.csv -NoTypeInformation -force
        $PstToMailboxProject | Select-Object ProjectName, ProjectType, ConnectorId, MailboxId, SourceFolder, EmailAddress | sort { $_.UserPrincipalName } | Export-Csv -Path $script:workingDir\AllProccessedHomeDirectories.csv -NoTypeInformation -Append
        $PstToMailboxProject | Select-Object ProjectName, ProjectType, ConnectorId, MailboxId, SourceFolder, EmailAddress | sort { $_.UserPrincipalName } | Export-Csv -Path $script:workingDir\ProccessedHomeDirectories-$date2.csv -NoTypeInformation -Append 
        
        if ([string]::IsNullOrEmpty($ApplyUserMigrationBundle)) {
            #Open the CSV file
            Start-Process -FilePath $script:workingDir\PstToMailboxProject-$date.csv
        }

        $msg = "SUCCESS: CSV file CSV file with the script output '$script:workingDir\PstToMailboxProject-$date.csv' opened."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg   
        $msg = "INFO: This CSV file will be used by Start-MWMigrationsFromCSVFile.ps1 script to automatically submit all home directories for migration."
        Write-Host $msg
        Log-Write -Message $msg   
        Write-Host

        Break
    }
    catch {
        $msg = "WARNING: Close CSV file '$script:workingDir\PstToMailboxProject-$date.csv' open."
        Write-Host -ForegroundColor Yellow $msg

        Start-Sleep 5
    }
} while ($true)

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg   

$output = "$script:workingDir\PstToMailboxProject-$date.csv"

Return $output 


##END SCRIPT
