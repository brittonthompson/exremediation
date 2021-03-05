# Set the path to the root of the extracted Exchange CU files or mounted ISO
$EXCUBitsPath = "J:\"

# Enabled the Schema Management MMC Snap-in for schema updates
regsvr32 schmmgmt.dll

# Check the FSMO roles for the Schema Master
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

# Example Output
Name                           Value                                                                                                                                                                                                                             
----                           -----                                                                                                                                                                                                                             
Active Directory               88                                                                                                                                                                                                                                
Exchange                       15332 

#>

Set-Location $EXCUBitsPath 

<#
# Before running the updates you need to make sure - 
  - You have Enterprise and Schema Admins for the account you're using
  - The Schema Master is in the current site you're in
  - You can run the schema updates directly from the schema master using the Exchange CU bits and commands below
  - Also, you can run the updates a preliminary step before installing Exchange updates during the downtime window
  - The updates are backwards compatible and you can, generally, run them any time

#>
# Upgrade the AD schema
.\setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
.\setup.exe /PrepareAD /IAcceptExchangeServerLicenseTerms
.\setup.exe /PrepareDomain /IAcceptExchangeServerLicenseTerms

# Install the Exchange CU
.\setup.exe /m:upgrade /IAcceptExchangeServerLicenseTerms

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


<#
### COOTIE CHECK
This will gather details about the infected files if they're found and give you a list to work with
#>


