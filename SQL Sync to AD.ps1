<#
.SYNOPSIS
Nightly sync for  SQL DB to Active Directory.
.DESCRIPTION
Pulls the tables from SQL and matches users in the project to the corresponding groups in AD.
.NOTES
Author: David Findley & Kory Dumas
Date: 08/27/2018
Version: v 3.0
#>

Import-Module ActiveDirectory -ErrorAction Stop

$server = "sqlsv.domain.local\instance"
$database = "DB Name"
$query = "SELECT * FROM dbo.ProjectMembers ORDER BY Name;"
$activeGroupName = ''


$extractFile = @"
C:\Users\Projects.csv
"@

$connectionTemplate = "Data Source={0};Integrated Security=SSPI;Initial Catalog={1};"
$connectionString = [string]::Format($connectionTemplate, $server, $database)
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString

$command = New-Object System.Data.SqlClient.SqlCommand
$command.CommandText = $query
$command.Connection = $connection

$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $command
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$connection.Close()

# dump the data to a csv
$DataSet.Tables[0] | Export-Csv $extractFile 


$GroupLine = Import-Csv -Path C:\Users\Projects.csv

Get-Date | Out-File -FilePath C:\Users\Log.txt -Append
"Member Additions `n" | Out-File -FilePath C:\Users\Log.txt -Append

Foreach ($GroupLine in $GroupLine) {
    TRY {
    Get-ADGroup "$($GroupLine.Name)"
    }

    CATCH {
    New-ADGroup -Name "$($GroupLine.Name)" -SamAccountName "$($GroupLine.Name)" -GroupCategory Distribution -GroupScope Universal -DisplayName "$($GroupLine.Name)" -Path "OU=OU NAME,DC=SOMECOMPANY,DC=local"
    }

    Add-ADGroupMember -Identity "$($GroupLine.Name)" -Members $($GroupLine.ObjectGuid) -ErrorAction SilentlyContinue
    $GroupLine.Name| Out-File -FilePath C:\Users\Log.txt -Append
    "Member Added: $($GroupLine.ObjectGuid) `n" | Out-File -FilePath C:\Users\Log.txt -Append
}


##Populate initial variable and declare deletion section in log
$GroupLine = Import-Csv -Path C:\Users\Projects.csv
"Member Deletions `n" | Out-File -FilePath C:\Users\Log.txt -Append
$MemberList = Get-ADGroupMember -Identity $GroupName | Select-Object -ExpandProperty objectGUID
[System.Collections.ArrayList]$membersToRemove = $MemberList

Foreach ($GroupLine in $GroupLine) {
    $GroupName = $GroupLine.Name

    if($activeGroupName -ne $GroupLine.Name)    
        {

        if ($membersToRemove) 
            {
            
            Remove-ADGroupMember -Identity "$($activeGroupName)" -Members $membersToRemove -Confirm:$False 
            
            ##Log Member Deletions
            $GroupName | Out-File -FilePath C:\Users\Log.txt -Append            
            "AD Group Members $($MemberList)" | Out-File -FilePath C:\Users\Log.txt -Append
            "Members to Remove $($membersToRemove)" | Out-File -FilePath C:\Users\Log.txt -Append

            ##Repopulate Variables
            $MemberList = Get-ADGroupMember -Identity $GroupName | Select-Object -ExpandProperty objectGUID  
            [System.Collections.ArrayList]$membersToRemove = @($MemberList)
            
            $activeGroupName = $GroupLine.Name
        
            }
        else
            {
            
            ##Repopulate Variables
            $MemberList = Get-ADGroupMember -Identity $GroupName | Select-Object -ExpandProperty objectGUID
            [System.Collections.ArrayList]$membersToRemove = @($MemberList)

            ##Remove Project Team Member from list of members to remove
            if($membersToRemove)
                {
                $membersToRemove.Remove($GroupLine.ObjectGuid)
                }
            else
                {
                }

            $activeGroupName = $GroupLine.Name

            }
        }
    else
        {

        ##Remove Project Team Member from list of members to remove
        if($membersToRemove)
            {
            $membersToRemove.Remove($GroupLine.ObjectGuid)
            }
        else
            {
            }

        }
    
}

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchsvr/powershell 
Import-PSSession $Session -AllowClobber

$GroupLine | ForEach-Object {
Enable-DistributionGroup -Identity $_.Name -ErrorAction SilentlyContinue
}

Remove-PSSession $Session
