﻿##############################################
 # InPuts for the Scripts
 #Author: Mohammed, Jawed
 #Date: 2019-08-01
 #Discription: Inputs for the script to run queries in TFS and send email as reminder to asisgnedto person.
 ##########################################################################

param (
 [string]$AzureDevOpsUri,
 [string]$SendEmailFrom,
 [string]$SmtpServerName,
 [string]$Recipients,
 [string]$SearchBase ,
 [boolean]$SendEmail,
 [boolean]$testemail,
 [string]$TeamProject,
 [string]$queryId,
 [string]$Operation,
 [string]$subject,
 [string]$Preline,
 [string]$title
 )

$currentDate=(Get-Date).ToString("yyyy-MM-dd") 
  
$subject="[$($currentDate)] " + $subject

[string]$Postline="<br /><b> For any assitance contact-:ALOZIE, Anthony.`r`n </b><br /><br /> - Auto-generated by an utility created by, Mohammed Jawed (a.k.a MD)"

 $uri_Queries="$($AzureDevOpsUri)/$($TeamProject)/_apis/wit/wiql/$($queryId)?api-version=5.0"

 $PathToSaveHtmlOutPut=Join-Path (Join-Path $PSScriptRoot "WorkItemResults") $currentDate
 
 #Create Path to Save Html File Output
try
{
   "Create Path to Save Html File Output"
   New-Item -path "$PathToSaveHtmlOutPut" -Type "directory" -Force

}
catch
{
  $ErrorMessage=$_.Exception.Message
  Write-Host "`n ERROR: " $ErrorMessage  
}

$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

     $tableName = "QueriesResultTable"
        #Create Table object
        $table = New-Object system.Data.DataTable “$tableName”

        #Define Columns
        $ID = New-Object system.Data.DataColumn ID,([string])        
        $Type = New-Object system.Data.DataColumn Type,([string])  
        $State = New-Object system.Data.DataColumn State,([string])        
        $Title = New-Object system.Data.DataColumn Title,([string]) 
        $AssignedTo = New-Object system.Data.DataColumn AssignedTo,([string])        
        $RequestedBy = New-Object system.Data.DataColumn RequestedBy,([string])  
        $NeedDate = New-Object system.Data.DataColumn NeedDate,([string])        
  
      
        #Add the Columns
        $table.columns.add($ID)
        $table.columns.add($Type)
        $table.columns.add($State)
        $table.columns.add($Title)
        $table.columns.add($AssignedTo)
        $table.columns.add($RequestedBy)
        $table.columns.add($NeedDate)
 
        ###################################
             
       
        #Get Overall pointer for WorkItem Queries
        $response_queries=Invoke-RestMethod -Method Get -Uri $uri_Queries -UseDefaultCredentials -ContentType application/json
        

        if($response_queries.workItems.Count -gt 0)
        {
            foreach($workItemUrl in $response_queries.workItems.url )
            {
               "==============================="
                 $workItemresult=Invoke-RestMethod -Method Get -Uri "$workItemUrl" -UseDefaultCredentials -ContentType application/json
                
                $WorkItemIDLink="$($AzureDevOpsUri)/$($TeamProject)/_workitems/edit/$($workItemresult.id)"
                
                $WorkItemIDWithLink= "`<a href`=`"$WorkItemIDLink`"`>$($workItemresult.id)`<`/a`>"

                "WorkItem Id: $($workItemresult.id)"
                 "WorkItem Type : $($workItemresult.fields.'System.WorkItemType')"
                  
                 "WorkItem State: $($workItemresult.fields.'System.State')"
                 "WorkItem Title: $($workItemresult.fields.'System.Title')"
                 "WorkItem AssignedTo: $($workItemresult.fields.'System.AssignedTo'.displayName)"
                 "WorkItem RequestedBy: $($workItemresult.fields.'Microsoft.VSTS.Common.RequestedBy.displayName')"
                 "WorkItem NeedDate: $($workItemresult.fields.'Microsoft.VSTS.Scheduling.NeedDate')"

                "============================="
                 #Create a row
                $row = $table.NewRow()
                #Enter data in the row
                $row.ID = $CountReleases

                 #Add the Columns
               $row.ID=$WorkItemIDWithLink #$workItemresult.id
                 $row.Type=$workItemresult.fields.'System.WorkItemType'
                 $row.State=$workItemresult.fields.'System.State'
                 $row.Title= $workItemresult.fields.'System.Title'
                 $row.AssignedTo=$workItemresult.fields.'System.AssignedTo'.displayName
                 $row.RequestedBy=$workItemresult.fields.'Microsoft.VSTS.Common.RequestedBy'.displayName
                 $row.NeedDate=$workItemresult.fields.'Microsoft.VSTS.Scheduling.NeedDate'
                       
                #Add the row to the table
                $table.Rows.Add($row)
            }
            
            #Convert to HTML and write to file and save
        $table | ConvertTo-Html -Property ID,Type,State,Title,AssignedTo,RequestedBy,NeedDate -Title "$title" -Body $currentDate -Pre "<P>$Preline</P>" -Post "$Postline" -Head $Header| Out-File  "$PathToSaveHtmlOutPut\$Operation-$($currentDate).htm"
        
       # Get-Content "$PSScriptRoot\WorkItemResult-$($currentDate).htm"
        #Get AsisgnedTo person List

        $Unique_AssignedTo=$table|Select-Object -Unique AssignedTo
        foreach($AssignedToName in $Unique_AssignedTo)
        {
            
            $assignedToPerson=$AssignedToName.AssignedTo
           $Results=$table|Where-Object {$_.AssignedTo -eq $assignedToPerson};
            "-------------------------- $assignedToPerson ------------ "
            "Result based on Assigned To: $($Results)"
            
            $Results | ConvertTo-Html -Property ID,Type,State,Title,AssignedTo,RequestedBy,NeedDate -Title "$title" -Body $currentDate -Pre "<P>$Preline</P>" -Post "<P>$Postline</P>" -Head $Header| Out-File  "$PathToSaveHtmlOutPut\$Operation-$assignedToPerson.htm" #$($AssignedToName.AssignedTo).htm"
            
            $messageBody =Get-Content "$PathToSaveHtmlOutPut\$Operation-$assignedToPerson.htm" |Out-String
            
            $messagebody= $messageBody.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"')
            $messagebody|Out-File  "$PathToSaveHtmlOutPut\$Operation-$assignedToPerson.htm"

            try
            {
             if($SendEmail)
               {
                   $UserEmail= Get-ADUser -Filter {DisplayName -eq $assignedToPerson} -SearchBase $SearchBase -Properties mail |Select-object mail
                   $Recipient=$UserEmail.mail

                   Write-Host "Email ID: $Recipient"
               }

            }
            catch
            {
                $ErrorMessage=$_.Exception.Message
                Write-Host "`n ERROR: " $ErrorMessage  

            }
            try
            {
               if($SendEmail)
               {
                if($testemail)
                {
                  $Recipient= [regex]::split($Recipients, ";")
                }
                 Send-MailMessage -From $SendEmailFrom -To $Recipient -SmtpServer $SmtpServerName -Body $messageBody.ToString() -Subject "$Subject" -BodyAsHtml

                "Done sending an email"
              }
            
            }
            catch
            {
              $ErrorMessage=$_.Exception.Message
              Write-Host "`n ERROR: " $ErrorMessage  

            }
            "-------------------------------"

        }


      }
      else
      {
          Write-Host "Nothing for Today"
      }
Write-Host "END"
