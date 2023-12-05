#requires -version 5

<#
  This Sample Code is provided for the purpose of illustration only and is not intended to
  be used in a production environment.

  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED WITHOUT WARRANTY OF ANY
  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
  MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.

  We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to
  reproduce and distribute the object code form of the Sample Code, provided that You agree:
  (i) to not use Our name, logo, or trademarks to market Your software product in which the
    Sample Code is embedded;
  (ii) to include a valid copyright notice on Your software product in which the Sample Code
    is embedded; and
  (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any
    claims or lawsuits, including attorneys’ fees, that arise or result from the use or
    distribution of the Sample Code.
#>

#Configure this.............................................................................
$Filepath = "C:\TEMP\$((get-date).Month;(get-date).Day;(get-date).Year)_TeamsMembers.csv"
$Filepath = ($Filepath -replace " ","_")
#...........................................................................................


#Gather a list of teams.
$Teams = Get-MgTeam


$result = $null

#Loop thru each team and create an Object
foreach ($Team in $teams)
{

$members = $Null

$members = $((Get-MgTeamMember -TeamID $team.id -All) | Select -ExpandProperty AdditionalProperties)
$member_name = $((Get-MgTeamMember -TeamID $team.id -All) | select -Property DisplayName)

$properties = @{
   'TeamName' = ($Team).DisplayName
   'TeamID'   = ($team).Id
   'Member_Id' = $members.userId
   'Member_Email' = $members.email
   'Member_Names' = $member_name.DisplayName
   
    
}

[array]$result += New-Object –TypeName PSObject -Property $properties   


}





remove-item $filepath -ErrorAction SilentlyContinue -Force 


$result2 = $null
###Foreach team add the member info and the Team info to a csv
Foreach($item in $result){



    $strTeamID = $item.TeamID
    $strTeamName = $item.TeamName
    $TeamMembers = $item.Member_Id
    
    
    
    Foreach ($itemA in $TeamMembers){
    
        $hash = @{
            'TeamName' = $strTeamName
            'TeamID'   = $strTeamID
            'Member_Id' = $itemA
            'Member_UPN' = (get-mguser -UserId $itemA).UserPrincipalName
            }


        [Array]$result2 += New-Object –TypeName PSObject -Property $hash


    }


}


$result2 | Export-csv $filepath -NoTypeInformation




