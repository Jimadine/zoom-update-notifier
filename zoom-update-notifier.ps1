#Requires -Version 5.1

<#
  .SYNOPSIS
  A Powershell script to check the installed version of Zoom and the latest available version of Zoom, and pop a dialogue or send an email if there is a mismatch.
  Note that if there is an error, this will be displayed in a window even if you have chosen to send an email.

  .PARAMETER IsTest
  Unconditionally pops the 'Update available' window or sends an email (depending on whether the -EnableEmail switch was passed)

  .PARAMETER EnableEmail
  Whether to send an email instead of popping a window [switch]

  .PARAMETER MailFromAddress
  For email notifications, specify the From: address. Only applicable when -EnableEmail is supplied. [string]

  .PARAMETER MailServer
  For email notifications, specify the FQDN of an SMTP server. Only applicable when -EnableEmail is supplied. [string]

  .PARAMETER MailSubject
  For email notifications, specify the subject line of the email. Only applicable when -EnableEmail is supplied. The default subject is 'Generic SFTP Upload notification' [string]

  .PARAMETER MailToAddress
  For email notifications, specify the To: address. Only applicable when -EnableEmail is supplied. [string]

  .PARAMETER ZoomExePath
  The full file path to zoom.exe [String]

#>
[CmdletBinding(DefaultParameterSetName = 'window')]
param(
  [Parameter(Mandatory, ParameterSetName = 'email')]
  [switch]$EnableEmail,

  [Parameter(ParameterSetName = 'window')]
  [Parameter(ParameterSetName = 'email')]
  [switch]$IsTest,

  [Parameter(Mandatory, ParameterSetName = 'email')]
  [string]$MailFromAddress,

  [Parameter(Mandatory, ParameterSetName = 'email')]
  [string]$MailServer,

  [Parameter(ParameterSetName = 'email')]
  [string]$MailSubject = 'Zoom Update Notifier: Update Available',

  [Parameter(Mandatory, ParameterSetName = 'email')]
  [string]$MailToAddress,

  [Parameter(ParameterSetName = 'window')]
  [Parameter(ParameterSetName = 'email')]
  [string]$ZoomExePath = 'C:\Program Files\Zoom\bin\zoom.exe'
)

<#
    .PARAMETER Path
    The full file path to zoom.exe [String]

    #>
function Get-ZoomExeVersion {
  If (!(Test-Path -Path $Path)) {
    Add-Type -AssemblyName System.Windows.Forms
    $message = "zoom.exe does not exist at '${Path}'"
    $null = [System.Windows.Forms.MessageBox]::Show($message, 'Zoom Update Notifier: Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    Exit 1
  }
  $installedVersion = (Get-Item -Path $Path).VersionInfo.FileVersionRaw.ToString()
  If ([string]::IsNullOrEmpty($installedVersion)) {
    Add-Type -AssemblyName System.Windows.Forms
    $message = "Failed to retrieve the version number from '${Path}'"
    $null = [System.Windows.Forms.MessageBox]::Show($message, 'Zoom Update Notifier: Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    Exit 1
  }
  Write-Output $installedVersion
}

function Get-ZoomLatestVersion {
  [string]$uas = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36'
  If ([Environment]::Is64BitOperatingSystem) {
    [uri]$url = 'https://zoom.us/client/latest/ZoomInstallerFull.msi?archType=x64'
  }
  Else {
    [uri]$url = 'https://zoom.us/client/latest/ZoomInstallerFull.msi'
  }

  # Code here is to handle dIfferences between Powershell 5.1 & 7.4 - https://github.com/PowerShell/PowerShell/issues/4534
  Try {
    $response = Invoke-WebRequest -Uri "$url" -UserAgent "$uas" -Method Head -MaximumRedirection 0 -ErrorAction Ignore
  }
  Catch [Microsoft.PowerShell.Commands.HttpResponseException] {
    [uri]$location = $_.Exception.Response.Headers.Location
  }
  Finally {
    If ($null -eq $location -or [string]::IsNullOrEmpty($location.AbsoluteUri)) {
      [uri]$location = $response.Headers.Location
    }
  }

  $latestVersion = $location.Segments[2].TrimEnd('/') | Where-Object { $_ -match '\d\.\d\.\d\.\d{1,5}' } | Select-Object -First 1
  If ([string]::IsNullOrEmpty($latestVersion)) {
    Add-Type -AssemblyName System.Windows.Forms
    $null = [System.Windows.Forms.MessageBox]::Show('Failed to retrieve the latest version number from the web.', 'Zoom Update Notifier: Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    Exit 1
  }
  Write-Output $latestVersion
}

<#
    .PARAMETER Title
    The title of the shown window [String]

    .PARAMETER Message
    The message of the shown window [String]

    #>
function Show-Window {
  param(
    [string]$Title,
    [string]$Message
  )
  Add-Type -AssemblyName System.Windows.Forms
  $null = [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
  Exit 2
}

<#

    .PARAMETER MailFromAddress
    The mail from address for the email [String]

    .PARAMETER MailToAddress
    The mail to address for the email [String]

    .PARAMETER MailServer
    The message of the email [String]

    .PARAMETER MailSubject
    The subject line of the email [String]

    .PARAMETER Message
    The body of the email [String]

    #>
function Send-Email {
  param(
    [string]$MailFromAddress,
    [string]$MailToAddress,
    [string]$MailServer,
    [string]$MailSubject,
    [string]$Message
  )
  $emailBody = @"
${Message}

This email was generated by $(Split-Path -Path $PSCommandPath -Leaf) at $((Get-Date).ToString('dd/MM/yyyy HH:mm:ss')) on $(([System.Net.Dns]::GetHostByName(($env:computerName))).Hostname).
"@
  Send-MailMessage -SmtpServer $MailServer -From $MailFromAddress -To $MailToAddress -Subject $MailSubject -Body $emailBody
}

<#
    .PARAMETER InstalledVersion
    The version string of the currently installed version [String]

    .PARAMETER LatestVersion
    The version string of the latest available version [String]

    #>
function Initialize-Notification {
  param(
    [string]$InstalledVersion,
    [string]$LatestVersion
  )
  If ($LatestVersion -ne $InstalledVersion -or $Script:IsTest) {
    If ($Script:EnableEmail) {
      Send-Email `
        -MailFromAddress $Script:MailFromAddress `
        -MailToAddress $Script:MailToAddress `
        -MailServer $Script:MailServer `
        -MailSubject $Script:MailSubject `
        -Message "Installed version is ${InstalledVersion}`nLatest version is ${LatestVersion}`nUpdate available!"
    }
    Else {
      Show-Window -Title 'Zoom Update Notifier: Update Available' -Message "Installed version is ${InstalledVersion}`nLatest version is ${LatestVersion}`nUpdate available!"
    }
  }
}

$installedVersion = Get-ZoomExeVersion -Path $Script:ZoomExePath
$latestVersion = Get-ZoomLatestVersion
Initialize-Notification -InstalledVersion $installedVersion -LatestVersion $latestVersion
