#requires -version 5

<#
  This Sample Code/POC is provided for the purpose of illustration only and is not intended to be used in a production environment.

  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.

  We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that You agree:
  (i) to not use Our name, logo, or trademarks to market Your software product in which the
    Sample Code is embedded;
  (ii) to include a valid copyright notice on Your software product in which the Sample Code
    is embedded; and
  (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any
    claims or lawsuits, including attorneys’ fees, that arise or result from the use or
    distribution of the Sample Code.




This script should be used to download the export once the export job has been run successfully.  With the caseId and exportId obtained from the graph APIs, run the following script with those parameters to download it to the path provided. Application Id is the one you obtained in the previous step.
	 
To run, .\DownloadExport.ps1 -appId "00bdc236-263d-4ea7-9882-50838b5b834c” -caseId "93563cc6-1f1e-497b-80cf-6a7558f5a658" -exportId "d6a0a11e50234776890dd5951475137c" -path "c:\temp"
	 
You can check in the download location for the downloaded export files - summary.csv as well as the. zips after the command are completed.




#>


[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$appId,
    [Parameter(Mandatory = $true)]
    [string]$caseId,
    [Parameter(Mandatory = $true)]
    [string]$exportId,
    [Parameter(Mandatory = $true)]
    [string]$path="D:\Temp"
    
)

if($PSVersionTable.PSVersion.Major -gt 5){
    Write-Warning "This script only tested on PowerShell 5"
}

if(-not(Get-Module -Name Microsoft.Graph -ListAvailable)){
    Write-Host "Installing Microsoft.Graph module"
    Install-Module Microsoft.Graph -Scope CurrentUser
}

if(-not(Get-Module -Name MSAL.PS -ListAvailable)){
    Write-Host "Installing MSAL.PS module"
    Install-Module MSAL.PS -Scope CurrentUser
}

if(-not(Get-MgContext)){
    Write-Host "Connect with credentials of a ediscovery admin (token for graph)"
    Connect-MgGraph -ClientId $appId  -Scopes "eDiscovery.ReadWrite.All"
}

Write-Host "Connect with credentials of a ediscovery admin (token for export)"
$exportToken=Get-MsalToken -ClientId $appId -Scopes "b26e684c-5068-4120-a679-64a5d2c909d9/.default" -RedirectUri "http://localhost" -Interactive


#
$exportToken 
#


$export=Invoke-MgGraphRequest -Uri "/beta/security/cases/ediscoveryCases/$($caseId)/operations/$($exportId)";
$export.exportFileMetadata | %{
    Write-Host "Downloading $($_.fileName)"
    Invoke-WebRequest -Uri $_.downloadUrl -OutFile "$($path)\$($_.fileName)" -Headers @{"Authorization"="Bearer $($exportToken.AccessToken)";"X-AllowWithAADToken"="true"}
    $exportToken
    $_.downloadUrl
}