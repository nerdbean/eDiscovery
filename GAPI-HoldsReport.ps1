#requires -version 5

<#
This is a sample script to build a csv with all holds and sources from Premium eDiscovery using Graph API.
You will need to have already logged in and obtained an oauth token to use this script.

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

$cases = Get-MgComplianceEdiscoveryCase
$holds = @()
$Array = @()
$MyObject2Array = @()
$HoldsANDcases = @()


Foreach ($case in $cases){



$tmp_holds = Get-MgComplianceEdiscoveryCaseLegalHold -CaseId $case.id

    Foreach ($Hold in $tmp_holds){
        


            $myObject = [PSCustomObject]@{
                HoldId     = $($hold.id)
                CaseID = $($case.id) 
                HoldName =  $($hold.Displayname)
                CaseName = $($case.Displayname)
            }

        $HoldsANDcases += $myObject


    }

}


Foreach($Combo in $HoldsANDcases){


    $UserSources = Get-MgComplianceEdiscoveryCaseLegalHoldUserSource -CaseId $Combo.caseid -LegalHoldId $Combo.HoldID 
    $SiteSources = Get-MgComplianceEdiscoveryCaseLegalHoldSiteSource -CaseId $Combo.caseid -LegalHoldId $Combo.HoldID  
    $sites = (Get-MgComplianceEdiscoveryCaseLegalHoldSiteSource -CaseId $Combo.caseid -LegalHoldId $Combo.HoldID).site


            Foreach ($usersource in $usersources){

                    #write-host ("UserEmail: " + $usersource.email)

                            $myObject2 = [PSCustomObject]@{
                                        HoldId     = $($Combo.HoldId)
                                        CaseID = $($Combo.CaseID)
                                        CaseName = $($Combo.casename)
                                        UserSourceID = $($usersource.ID)
                                        UserDisplayname = $($usersource.Displayname)
                                        Email = $($usersource.Email)
                                        SiteSourceID = $null
                                        SiteDisplayname = $null
                                        Site = $null
                                        CreatedBy = $($usersource.CreatedBY.User.DisplayName)
                                        Created = $($usersource.Createddatetime)
                                        Type = "Usersource"
                                        }
                            

                    $MyObject2Array += $myObject2

                             }



 $Combo.HoldId
            Foreach ($SiteSource in $sitesources){
             
              $SiteSource.site

              #$SiteSource.Site

                            $myObject3 = [PSCustomObject]@{
                                        HoldId     = $($Combo.HoldId)
                                        CaseID = $($Combo.CaseID)
                                        CaseName = $($Combo.casename)
                                        UserSourceID = $null
                                        UserDisplayname = $null
                                        Email = $null
                                        SiteSourceID = $($Sitesource.ID)
                                        SiteDisplayname = $($Sitesource.Displayname)
                                        Site = $($Sitesource.Site).id
                                        CreatedBy = $($SiteSource.CreatedBy.User.DisplayName)
                                        Created = $($sitesource.CreatedDateTime)
                                        Type = "Sitesource"
                                        }              

                                          

                      $MyObject2Array += $myObject3

                 }

               

}




$MyObject2Array | export-csv C:\temp\graph_holds_report.csv -NoTypeInformation

write-host ("Done!") -BackgroundColor Black -ForegroundColor Yellow


