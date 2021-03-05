function Get-1PExchangeCompromisedFiles {
  class CootiePath {
    [string]$Path
    [bool]$ExtensionsCheck
    [bool]$Recurse
    [string]$PathRegex
    [string[]]$AllButThese

    CootiePath() { $this }
  }

  $Extensions = @(".aspx")
  $KnownPaths = @(
    [CootiePath]@{
      Path        = "$($env:ExchangeInstallPath)\FrontEnd\HttpProxy\ecp\auth"
      AllButThese = @("TimeoutLogout.aspx")
    },
    [CootiePath]@{
      Path        = "$($env:ExchangeInstallPath)FrontEnd\HttpProxy\owa\auth"
      AllButThese = @(
        "errorFE.aspx",
        "ExpiredPassword.aspx",
        "frowny.aspx",
        "getidtoken.htm",
        "logoff.aspx",
        "logon.aspx",
        "OutlookCN.aspx",
        "RedirSuiteServiceProxy.aspx",
        "signout.aspx"
      )
    },
    [CootiePath]@{
      Path            = "$($env:ExchangeInstallPath)FrontEnd\HttpProxy\owa\auth\Current"
      ExtensionsCheck = $true
      Recurse         = $true
    },
    [CootiePath]@{
      Path            = "$($env:ExchangeInstallPath)FrontEnd\HttpProxy\owa\auth"
      ExtensionsCheck = $true
      Recurse         = $true
      PathRegex       = "[0-9]{2}\.[0-9]\.[0-9]{4}"
    },
    [CootiePath]@{
      Path            = "C:\inetpub\wwwroot\aspnet_client"
      ExtensionsCheck = $true
      Recurse         = $true
    }
  )

  $Remediate = @()
  foreach ($P in $KnownPaths) {
    $Path = $P.Path
    $Recurse = $P.Recurse

    if (Test-Path $Path) {

      if ($P.PathRegex) { 
        $Path = (Get-ChildItem -Path $Path | Where-Object { 
            $_.PSIsContainer -and
            $_.Name -match $P.PathRegex 
          }).FullName
      }

      if ($Path -and (Test-Path $Path)) {
        Write-Host "[$(Get-Date)] Check Path: $Path"

        if ($P.ExtensionsCheck) {
          Write-Host " - Check Extensions"
          $Remediate += Get-ChildItem -Path $Path -Recurse:$Recurse  | Where-Object { -not $_.PSIsContainer -and $_.Extension -in $Extensions }
        }
      
        if ($P.AllButThese) {
          Write-Host " - Check All But These"
          $Remediate += Get-ChildItem -Path $Path -Recurse:$Recurse  | Where-Object { -not $_.PSIsContainer -and $_.Name -notin $P.AllButThese }
        }  
      }
      else {
        Write-Host "[$(Get-Date)] Path Not Found with Regex $($P.PathRegex): $($P.Path)"
      }
    }
    else {
      Write-Host "[$(Get-Date)] Path Not Found: $Path"
    }
  }

  if ($Remediate) {
    Write-Host "[$(Get-Date)] Discovered Files to Remediate:"
    $Remediate
  }
}

$Files = Get-1PExchangeCompromisedFiles
$Files | Format-Table CreationTime, FullName


function Get-1PExchangeAttackLogs {
  param(
    [switch]$GetIISLogs,
    [switch]$GetExchangeLogs
  )

  $Return = @()
  if ($GetExchangeLogs) { 
    $Logs = (Get-ChildItem -Path "$($env:ExchangeInstallPath)Logging\HttpProxy" -Recurse -Filter "*.log").FullName
    $Cooties = Import-Csv -Path $Logs | Where-Object { -not $_.AuthenticatedUser -and $_.AnchorMailbox -like "ServerInfo~*/*" } 
    $Return += [PSCustomObject]@{
      Type  = "ExchangeLogs"
      Value = $Cooties
    }
  }

  if ($GetIISLogs) {
    $BadGuys = @(
      "103.77.192.219",
      "104.140.114.110",
      "104.250.191.110",
      "108.61.246.56",
      "149.28.14.163",
      "157.230.221.198",
      "167.99.168.251",
      "185.250.151.72",
      "192.81.208.169",
      "203.160.69.66",
      "211.56.98.146",
      "5.254.43.18",
      "80.92.205.81"
    )
  
    $IPRegex = "\b(?!10\.|169\.254|192\.168|172\.(2[0-9]|1[6-9]|3[0-2]))[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}"
    $IISLogs = (Get-ChildItem -Path "C:\inetpub\logs\LogFiles" -Recurse -Filter "*.log").FullName
    $X = 0
    $Logs = $IISLogs | ForEach-Object { 
      Write-Host "[$(Get-Date)] Checking Log File ($($X+1) of $($IISLogs.Count)): $_"
      Get-Content $_ | Where-Object {
        $_ -match "POST \/owa\/auth\/Current\/" -or 
        $_ -match "POST \/ecp\/default\.flt" -or 
        $_ -match "POST \/ecp\/main\.css" -or 
        $_ -match "POST \/ecp/[a-zA-Z0-9]\.js" -or 
        $(if ($_ -match $IPRegex) {
            $BadGuys -contains $Matches[0].Value
          })
      } | ForEach-Object { [PSCustomObject]@{ LogFile = $IISLogs[$X]; String = $_ }
      }
      $X++
    }

    $Return += [PSCustomObject]@{
      Type  = "IISLogs"
      Value = $Logs
    }
  }

  return $Return
}

$Logs = Get-1PExchangeAttackLogs -GetExchangeLogs -GetIISLogs
$Logs | Select-Object Type, @{ E = { if (-not $_.Value.Count) { 1 } else { $_.Value.Count } }; N = "CootieCount" }