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
#>


<#

This script helps you register the application, if already created in the tenant for Microsoft graph. It will provision the application with user delegated read access to the Microsoft Purview application. Please note, it is necessary to be signed in as tenant admin to system for providing the provisioning the access.
Note: If the application name provided is not already created, it will create an application for you.
	 
This registration of application will help in automating the sign-in process without having to include interactive sign-in page and avoid the manual intervention.
	 
To run the script

.\RegisterDownloadExportApp.ps1 -appName "MyeDApp" 
	 
This will do the above mentioned for the application name “MyeDApp". It will return the complete application information. Please note the application Id obtained in the information. You can also check the application id from Microsoft Azure Portal.


#>




[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$appName
)

if(-not(Get-Module -Name Microsoft.Graph -ListAvailable)){
    Write-Host "Installing Microsoft.Graph module"
    Install-Module Microsoft.Graph -Scope CurrentUser
}

if(-not(Get-MgContext)){
    Write-Host "Connect with credentials of a tenant admin"
    Connect-MgGraph  -Scopes "Application.ReadWrite.All"
}

$app = Get-MgApplication -Search "DisplayName:$appName" -ConsistencyLevel "Eventual" | Select-Object -First 1;
if(-not($app)){
    Write-Host "Creating application"
    $app=New-MgApplication -DisplayName $appName  -PublicClient @{ RedirectUris = "http://localhost" };
}

if($app){
    $mgSpn=Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
    $edSpn = Get-MgServicePrincipal -Filter "AppId eq 'b26e684c-5068-4120-a679-64a5d2c909d9'"  | Select-Object -First 1;
    if(-not($edSpn)){
        Write-Host "Creating eDiscovery app";
        $spId=@{"AppId" = "b26e684c-5068-4120-a679-64a5d2c909d9"}
        New-MgServicePrincipal -BodyParameter $spId;
        $edSpn = Get-MgServicePrincipal -Filter "AppId eq 'b26e684c-5068-4120-a679-64a5d2c909d9'"  | Select-Object -First 1;
        $rt=0;
        if(-not($edSpn) -and $rt -lt 3){
            Write-Host "Waiting for SPN";
            Start-Sleep 30;
            $rt=$rt+1;
        }
    }

    if(-not($edSpn)){
        Write-Warning "The eDiscovery app is not available in your organization";
    }
    else{
        Write-Host "Adding permissions";
        $perms=@();
        $mgSpn.Oauth2PermissionScopes | Where-Object{$_.Value -like 'ediscovery.readwrite*'} | ForEach-Object{
            $perms+=@{ResourceAppId=$mgSpn.AppId;ResourceAccess = @(@{Id =$_.Id;Type = "Scope"})}
        };
        $edSpn.Oauth2PermissionScopes | Where-Object{$_.Value -like 'ediscovery.*'} | ForEach-Object{
            $perms+=@{ResourceAppId=$edSpn.AppId;ResourceAccess = @(@{Id =$_.Id;Type = "Scope"})}
        };
        Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $perms;
        Write-Host "App regsistered, here are the app details";
        $app | Format-List;
    }
}
else{
    Write-Warning "The  app is not available/could not be created";
}