


##CONFIGURE THE FOLLOWING=================================

$newCustodian = "admin@M365x661152.onmicrosoft.com"

$YammerSite =  "https://m365x661152.sharepoint.com/sites/askhr"

$YammerEmail = "askhr@M365x661152.onmicrosoft.com"

##========================================================




$ranum = get-random -Maximum 1000
$NewCaseName = "TestCase_$($ranum)"
$exportName = "Export_$($ranum)"


<#
1. Create a case
2. Create Custodian

3. Add a User Source (the Yammer Group Email)
4. Add Custodial SiteSource to case
5. Apply Hold
6. Create a Collection looking for all of Yammer
7. Estimate Search
7a. Get Status of Estimate
##!!!WAIT FOR ESTIMATE TO FINISH!!!!#
8. Create Review Set
9. Add Search to Review Set
9a. Check status of transfer to review set.

#!!! Wait for data to be sent to review set!!!!#
10. Create ReviewSet Query
11. Export the Review Set based on Query
12. Get the Export ID by the Export Name + Loop to Check the status of the export until its 100%. 
13. Get the Export URLS




#>
#region 1. Create a case
    $params = @{
	    DisplayName = $NewCaseName
	    Description = "Testing Graph/Yammer" + (get-date).DateTime 
	    ExternalId = "$RandomNumber"
    }

    $CaseGUID =  (New-MgSecurityCaseEdiscoveryCase -BodyParameter $params).Id
#endregion


#region 2. Create Custodian

    $params = @{
	    email = $NewCustodian
        holdstatus = "true"
    }

    $CustID = (New-MgSecurityCaseEdiscoveryCaseCustodian -EdiscoveryCaseId $caseGUID -BodyParameter $params).id

#endregion

#region 3. Add a User Source (the Yammer Group Email)

$params = @{
	
		
       "Email" =  $NewCustodian

	}

$userSourceID = (New-MgSecurityCaseEdiscoveryCaseCustodianUserSource -EdiscoveryCaseId $CaseGUID -EdiscoveryCustodianId $CustID -BodyParameter $params).ID


$params = @{
	
		
       "Email" = $YammerEmail

	}

$YammerSourceID = (New-MgSecurityCaseEdiscoveryCaseCustodianUserSource -EdiscoveryCaseId $CaseGUID -EdiscoveryCustodianId $CustID -BodyParameter $params).ID

#endregion



#region 4. Add Custodial SiteSource to case

    $params = @{
           site = @{
                  webUrl = "$YammerSite"
           }
           holdstatus = "true"
    }

    $SiteSource = New-MgSecurityCaseEdiscoveryCaseCustodianSiteSource -EdiscoveryCaseId $caseGUID -EdiscoveryCustodianId $CustID -BodyParameter $params 

#endregion


#region 5. Apply Hold

Add-MgSecurityCaseEdiscoveryCaseCustodianHold -EdiscoveryCaseId $caseGUID -EdiscoveryCustodianId $CustID


#endregion


#region 6. Create a Collection looking for all of Yammer

    #get the UserSource

$UserSource = Get-MgSecurityCaseEdiscoveryCaseCustodianUserSource -EdiscoveryCaseId $caseid -EdiscoveryCustodianId $Custid

$caseid = $CaseGUID


$params = @{
	DisplayName = "My Graph Search-yammer " + (get-random -Minimum 0 -Maximum 50) #<---------Search Name
	Description = "Testing Search" #<---------Search Description
	ContentQuery = "(ItemClass=IPM.Yammer.message)(ItemClass=IPM.Yammer.poll)(ItemClass=IPM.Yammer.praise)(ItemClass=IPM.Yammer.question)" #<----------Specify the Query
	"CustodianSources@odata.bind" = @(
		"https://graph.microsoft.com/beta/security/cases/ediscoveryCases/$($caseid)/custodians/$($Custid)/userSources/$($UserSource.Id)"

	)
	#"NoncustodialSources@odata.bind" = @("https://graph.microsoft.com/v1.0/security/cases/ediscoveryCases/$($caseid)/noncustodialdatasources/35393639323133394345384344303043"
	#)
}

$Search = New-MgSecurityCaseEdiscoveryCaseSearch -EdiscoveryCaseId $caseid -BodyParameter $params
#endregion

#region 7. Estimate Search
Invoke-MgEstimateSecurityCaseEdiscoveryCaseSearchStatistics -EdiscoveryCaseId $CaseGUID -EdiscoverySearchId $Search.Id
#endregion

##!!!WAIT FOR ESTIMATE TO FINISH!!!!#


#region 7a. Get Status of Estimate
$EstimateStatus = ""
Do {
$EstimateStatus = Get-MgBetaComplianceEdiscoveryCaseSourceCollectionLastEstimateStatisticsOperation -CaseId $CaseGUID -SourceCollectionId $Search.Id
Clear-Host
write-host "Estimating Search... $($EstimateStatus.PercentProgress) Complete." -ForegroundColor Yellow
Start-Sleep -Seconds 3
} Until
($EstimateStatus.PercentProgress -eq "100")



#endregion



#region 8. Create Review Set
$params = @{
	DisplayName = "RS_$($Search.DisplayName_)$(get-random -Minimum 0 -Maximum 500)"
}

$RS = New-MgSecurityCaseEdiscoveryCaseReviewSet -EdiscoveryCaseId $CaseGUID -BodyParameter $params
#endregion



#region 9. Add Search to Review Set
$params = @{
	Search = @{
		Id = "$($Search.id)"
	}
	AdditionalDataOptions = "linkedFiles"
}

Add-MgSecurityCaseEdiscoveryCaseReviewSetToReviewSet -EdiscoveryCaseId $CaseGUID -EdiscoveryReviewSetId $RS.Id -BodyParameter $params
#endregion

#!!! Wait for data to be sent to review set!!!!#
"adding to review set"
#region 9a. Check status of transfer to review set.
$RS_OP = Get-MgSecurityCaseEdiscoveryCaseSearchAddToReviewSetOperation -EdiscoveryCaseId $CaseGUID -EdiscoverySearchId $Search.Id


$Estimate_RS_OP = ""
Do {
$Estimate_RS_OP = Get-MgSecurityCaseEdiscoveryCaseSearchAddToReviewSetOperation -EdiscoveryCaseId $CaseGUID -EdiscoverySearchId $Search.Id
Clear-Host
write-host "Transfering to Review Set... $($Estimate_RS_OP.PercentProgress) Complete." -ForegroundColor Yellow
If (($Estimate_RS_OP.PercentProgress -eq "100") -and ($Estimate_RS_OP.Status -eq "succeeded")) {continue}

Start-Sleep -Seconds 30

} Until
(($Estimate_RS_OP.PercentProgress -eq "100") -and ($Estimate_RS_OP.Status -eq "succeeded")) 





#endregion


#region 10. Create ReviewSet Query
$params = @{
	displayName = "Q-5"
	contentquery = @"
Keywords:"storylinepost2"
"@
}

$RSQuery = New-MgSecurityCaseEdiscoveryCaseReviewSetQuery -EdiscoveryCaseId $CaseGUID -EdiscoveryReviewSetId $RS.Id -BodyParameter $params
#endregion

#region 11. Export the Review Set based on Query
$params = @{
	OutputName = $ExportName
	Description = "Testing"
	ExportOptions = "originalFiles,fileInfo,tags"
	ExportStructure = "directory"
}
$Exportinfo = Export-MgSecurityCaseEdiscoveryCaseReviewSetQuery -EdiscoveryCaseId $CaseGUID -EdiscoveryReviewSetId $rs.Id -EdiscoveryReviewSetQueryId $RSQuery.Id -BodyParameter $params

#endregion








#region 12. Get the Export ID by the Export Name 
$ContentExports = Get-MgSecurityCaseEdiscoveryCaseOperation -EdiscoveryCaseId $CaseGUID | where {$_.Action -eq 'ContentExport'}
 
$export_info = $null
$exports = @()
 
foreach ($item in $ContentExports.id)
{
   $export_info = Get-MgSecurityCaseEdiscoveryCaseOperation -EdiscoveryCaseId $CaseGUID  -CaseOperationId $item 
   $exports += $export_info | Where {($_ | select -ExpandProperty AdditionalProperties).outputName -eq $exportname}
 
}
 
$ExportID = $exports.id

$ExportStatus = ""
Do {
$ExportStatus= Get-MgSecurityCaseEdiscoveryCaseOperation -EdiscoveryCaseId $CaseGUID -CaseOperationId $ExportID
Clear-Host
write-host "Export Status... $($ExportStatus.PercentProgress) Complete." -ForegroundColor Yellow
Start-Sleep -Seconds 3
} Until
($ExportStatus.PercentProgress -eq "100")
#endregion



#region 13. Get the Export URLS


#After choosing an operation (the export), get the information for that operation.
$export_info = Get-MgSecurityCaseEdiscoveryCaseOperation -EdiscoveryCaseId $CaseGUID -CaseOperationId $ExportID

#Get the Export Information
$Export_Properties = ($export_info | select -ExpandProperty AdditionalProperties)

#Get the Export File Meta Data - This will store in an array. So the Filename, Size, DownloadURL are in each element.
$ExportFileMeta = ($Export_Properties.exportFileMetadata)

#Grab the URLS - One for The summary.csv and for the Content.
$ExportURL0 = $ExportFileMeta[0].downloadUrl
$ExportURL1 = $ExportFileMeta[1].downloadUrl

#Display the URLS
$ExportURL0

$ExportURL1 

#endregion

