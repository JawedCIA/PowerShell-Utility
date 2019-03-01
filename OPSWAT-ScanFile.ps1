#Discription: Script to Scan files over OPSWat using Metdefender API and display results of scan
#Auther: MD

function OPSWATScanFile 
{
 param(
        [string]$site,
        [string]$FolderPathForScan
     )

Try 
    {     
        if (Test-path $FolderPathForScan) 
        {

            $FilesPathInFolder=Get-ChildItem -Recurse "$FolderPathForScan" | Where { ! $_.PSIsContainer } 
            foreach($FileInfo in $FilesPathInFolder)
                {             
                                   
                    Write-Host $('-' * 50)
                    write-host "Uploading File: $FileInfo over Metadefender using API for Scan"
                    $FilePath=$FileInfo.FullName
                    $FileName=$FileInfo.Name
                    $uri = "$site/file"
                    $ProgressPreference = 'SilentlyContinue'
                    $response =  Invoke-RestMethod -Uri $uri -Method Post -InFile $FilePath -UseDefaultCredentials
                    $response = $response -Replace "@{data_id=","" -Replace "}",""
                    $resultURI = "$uri/$response"
                    $resultResponse = Invoke-RestMethod -Uri $resultURI -Method Get -UseDefaultCredentials | ConvertTo-Json | Format-Json
                    if ($resultResponse -clike '*Processing*') 
                        {
                            write-host "Processing..."
                    DO 
                        {
                            Start-Sleep -s 3
                            $resultResponse = Invoke-RestMethod -Uri $resultURI -Method Get -UseDefaultCredentials | ConvertTo-Json | Format-Json
    
                        } while ($resultResponse -clike '*Processing*')
                        }
                    if($FullDetails)
                        {
                            write-host $resultResponse
                        } 
                    Else 
                        {
                            $resultResponse = $resultResponse | ConvertFrom-Json
                            if (!$resultResponse.process_info.blocked_reason)
                                {
                                    write-host " File:" $FileName `n " Data ID:" $resultResponse.data_id `n "Profile:" $resultResponse.process_info.profile `n "Result:" $resultResponse.process_info.result `n "Process time:" $resultResponse.process_info.processing_time -ForegroundColor Green
                                    write-host `n
                                    continue
                                }
                            Else 
                                {
                                    write-host " File:" $FileName `n "Data ID:" $resultResponse.data_id `n "Profile:" $resultResponse.process_info.profile `n "Result:" $resultResponse.process_info.result `n "Process time:" $resultResponse.process_info.processing_time -ForegroundColor Red
                                    write-host `n "Blocked Reason:" $resultResponse.process_info.blocked_reason -ForegroundColor Red
                                }
                        }
                      
                }
        }
        Else
            {
                Write-Host "Provide Folder Path DoesNot Exist: $FolderPathForScan" -ForegroundColor Red
            }
    }
    catch
        {
           $ErrorMessage = $_.Exception.Message
           Write-Host "ERROR: $ErrorMessage" -ForegroundColor Red
           Break
        }
}

OPSWATScanFile -site "https://MDCore.site.com" -FolderPathForScan "C:\FileToScan"
