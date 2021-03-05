<#
#### READ FIRST
  - **** DO NOT JUST RUN THIS WHOLE SCRIPT
  - Best method for use is to open the PowerShell ISE and run it in sections as you need them
    - Highlight the section of code you need and hit F8 or hit the little icon for run selection
  - This is broken down into sections separated by headers you can't miss so go section by section
#>

#/---------------------------------------------------------------------------
# SET VARIABLES
#/---------------------------------------------------------------------------

# Set the path to the root of the extracted Exchange CU files or mounted ISO
$EXCUBitsPath = "J:\"

if (-not (Test-Path $env:ExchangeInstallPath)) {
  <#
  ### ExchangeInstallPath ENVIRONMENT VARIABLE
  Aight, either you're not on an Exchange Server, someone removed it or it was never created.
    - You can go ahead and set the variable for this session using the code below
    - Either way, you need to go the system advanced properties and set the ExchangeInstallPath environment variable
    - Exchange uses this internally in some cases and every PowerShell script you see on the web will also utilize it
    - Adding it here will only set it for your session for convenience. Go ahead and do it manually.
    - If you don't set it in the session you'll have to close and relaunch PowerShell for the variable to propagate.
  #>

  Write-Host "[$(Get-Date)] The `$env:ExchangeInstallPath environment variable is not populated. What's the path?"
  Write-Host "[$(Get-Date)] i.e., C:\Program Files\Microsoft\Exchange Server\V15\ ... including the trailing slash"
  $env:ExchangeInstallPath = "C:\Program Files\Microsoft\Exchange Server\V15\"
  $env:ExchangeInstallPath = Read-Host -Prompt "[$env:ExchangeInstallPath]"

  if (-not (Test-Path $env:ExchangeInstallPath)) {
    Write-Host "[$(Get-Date)] Wrong - The path @ $env:ExchangeInstallPath does not exist." -ForegroundColor Red
  }
  elseif (-not $env:ExchangeInstallPath.EndsWith("\")) {
    Write-Host "[$(Get-Date)] Wrong - Read the instructions, you need that trailing `"\`"" -ForegroundColor Red
  }
}


#/---------------------------------------------------------------------------
# CHECK SCHEMA VERSION
#/---------------------------------------------------------------------------

# Enabled the Schema Management MMC Snap-in for schema updates
regsvr32 schmmgmt.dll

<# 
##### Check the FSMO roles for the Schema Master
  - Make sure the Schema Master role holder is in the same site
  - If not:
    - Run the schema update commands with the Exchange CU bits directly on the schema master
    - Or, update Exchange in the site where the schema master exists first
  - You can run the schema updates during the day before the downtime window to save time if desired
  - Schema updates are backwards compatible and generally non-disruptive
#>
netdom query fsmo

# Check the AD schema version before running the schema updates to confirm they actually change
Import-Module ActiveDirectory

$SchemaVersions = @()
$SchemaPartition = (Get-ADRootDSE).NamingContexts | Where-Object { $_.SubString(0, 9).ToLower() -eq "cn=schema" }
$SchemaVersionAD = (Get-ADObject $SchemaPartition -Property objectVersion).objectVersion
$SchemaVersions += @{ "Active Directory" = $SchemaVersionAD }
$SchemaPathExchange = "CN=ms-Exch-Schema-Version-Pt,$SchemaPartition"

if (Test-Path "AD:$SchemaPathExchange") {
  $SchemaVersionExchange = (Get-ADObject $SchemaPathExchange -Property rangeUpper).rangeUpper
}
else {
  $SchemaVersionExchange = 0
}

$SchemaVersions += @{ "Exchange" = $SchemaVersionExchange }

Write-Output $SchemaVersions 

<#
### YOU CAN RUN THE ABOVE AGAIN TO ENSURE THE VERSIONS CHANGE

# Example Output
Name                           Value                                                                                                                                                                                                                             
----                           -----                                                                                                                                                                                                                             
Active Directory               88                                                                                                                                                                                                                                
Exchange                       15332 

#>


#/---------------------------------------------------------------------------
# RUN THE CUMULATIVE UPDATE
#/---------------------------------------------------------------------------

Write-Host "[$(Get-Date)] Change the directory to $EXCUBitsPath"
Set-Location $EXCUBitsPath 

# This is here to stop fools from simply running the entire script without reading above
Write-Host "[$(Get-Date)] Ummm.... YOU NEED TO READ THE INSTRUCTIONS BEFORE USING THIS. Exiting..."
Start-Sleep -Seconds 10
exit

<#
# Before running the updates you need to make sure - 
  - You have Enterprise and Schema Admins for the account you're using
  - The Schema Master is in the current site you're in
  - You can run the schema updates directly from the schema master using the Exchange CU bits and commands below
  - Also, you can run the updates a preliminary step before installing Exchange updates during the downtime window
  - The updates are backwards compatible and you can, generally, run them any time
#>

Write-Host "[$(Get-Date)] Run AD Schema Updates"

# Upgrade the AD schema
.\setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
.\setup.exe /PrepareAD /IAcceptExchangeServerLicenseTerms
.\setup.exe /PrepareDomain /IAcceptExchangeServerLicenseTerms

Write-Host "[$(Get-Date)] Run the CU Installation"

# Install the Exchange CU
.\setup.exe /m:upgrade /IAcceptExchangeServerLicenseTerms


#/---------------------------------------------------------------------------
# FIXES
#/---------------------------------------------------------------------------

<#
### BUSTED ECP BECAUSE MS THROUGH SOME LAZY CODE IN THERE FOR  US
*** THIS MAY ONLY BE REQUIRED WITH Exchange 2013
If ECP throws some FileNotFound mambo-jahambo in the browser, you may need to do this
  - Manually changing this setting in IIS is simple if you prefer
    - Open IIS and go to the "Exchange Back End" site
    - Click on ECP
    - Go to "Application Settings" under the ASP.NET section in the center pane
    - Double click the "BinSearchFolders" element and you'll likely see %ExchangeInstallDir% (which is a non-existent environment variable)
    - Replace %ExchangeInstallDir% with the path to the Exchange installation - i.e., C:\Program Files\Microsoft\Exchange Server\V15\
    - Do not forget the trailing slash
    - Or, 
#>

$ECPWebConfigFile = "$env:ExchangeInstallPath`ClientAccess\ecp\Web.config"

if (Test-Path $ECPWebConfigFile) {
  $BackupFile = "$ECPWebConfigFile.$(Get-Date -Format "yyyyMMddHHmm").bak"
  Write-Host "[$(Get-Date)] Creating a backup of the Web.config @ $BackupFile" -ForegroundColor Red
  try {
    Copy-Item -Path $ECPWebConfigFile -Destination $BackupFile  -Force | Out-Null
    $ECPWebConfig = [xml](Get-Content $ECPWebConfigFile)
    $BinSearchFolders = $ECPWebConfig.configuration.appSettings.GetElementsByTagName("add") | Where-Object { $_.Key -eq "BinSearchFolders" }
    if ($BinSearchFolders.Value -match "%ExchangeInstallDir%") {
      $NewValue = $BinSearchFolders.Value.Replace("%ExchangeInstallDir%",$env:ExchangeInstallPath)
      $BinSearchFolders.SetAttribute("value", $NewValue)
      $ECPWebConfig.Save("/Users/bthompson/Documents/Scripts/Web.config")

      $Restart = Read-Host -Prompt "- Safe to restart IIS?"
      switch -Wildcard ($Restart) {
        "*Y*" { 
          Write-Host "[$(Get-Date)] Restarting IIS"
          iisreset | Out-Null
        }
        default { Write-Host "[$(Get-Date)] Skipping IIS Restart" -ForegroundColor Yellow }
      }
    }
  }
  catch {
    Write-Host "[$(Get-Date)] Creating a backup of the Web.config failed. Just go through the manual process above." -ForegroundColor Red
  }
}
else {
  Write-Host "[$(Get-Date)] ECP Web.config file not found @ $ECPWebConfigFile. Make sure the path is correct." -ForegroundColor Red
}


<#
### ECP AND OWA ARE BUSTED WITH FILE NOT FOUND OR ASSEMBLY ERRORS
# Run the UpdateCas.ps1 script
#>

."$($env:ExchangeInstallPath)bin\UpdateCas.ps1"


#/---------------------------------------------------------------------------
# CHECK FOR COMPROMISE
#/---------------------------------------------------------------------------

<#
### COOTIE CHECK
This will gather details about the infected files if they're found and give you a list to work with
#>

$KnownFiles = (Invoke-RestMethod -Uri "https://raw.githubusercontent.com/brittonthompson/exremediation/master/KnownFiles.txt" -Method GET).Split("`n")
$KnownPaths = @(
  "C:\inetpub\wwwroot\aspnet_client",
  "C:\inetpub\wwwroot\aspnet_client\system_web\",
  "$env:ExchangeInstallPath\FrontEnd\HttpProxy\owa\auth"
)

$Remediate = @()
foreach ($P in $KnownPaths) {
  Write-Host "[$(Get-Date)] Check Path: $P"
  $KnownFiles | ForEach-Object { 
    Write-Host " - Check File: $_"
    if (Test-Path "$P\$_") {
      $Remediate += Get-Item "$P\$_"
    }
  }
}

if ($Remediate) {
  Write-Host "[$(Get-Date)] Discovered Files to Remediate:"
  $File | Format-Table CreationTime, FullName
}