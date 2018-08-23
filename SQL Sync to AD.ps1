<#
.SYNOPSIS
Nightly sync for SQL DB to Active Directory.
.DESCRIPTION
Pulls the tables from SQL and matches users in the project to the corresponding groups in AD.
.NOTES
Author: David Findley
Date: 07/10/2018
Version: v1.0
Requires PS v4.0 or above and the AD Powershell tools. 
#>

Import-Module ActiveDirectory -ErrorAction Stop

$server = "SQL-SVR\NAME"
$database = "DBNAME"
$query = "SELECT * FROM dbo.ProjectMembers;"


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

#Exports to a CSV file. 
$DataSet.Tables[0] | Export-Csv $extractFile 

$DestinationOU = "OU=SomeOUName,DC=Server,DC=local"
$GroupLines = Import-Csv -Path C:\Users\Projects.csv

#Runs through the script line by line and checks for the group in AD. If doesn't exist, it creates the groups and adds users. If exists, it just adds users.
Foreach ($GroupLine in $GroupLines) {
    try {
        Get-ADGroup $GroupLine.Name
    }
    catch {
        New-ADGroup -Name $GroupLine.Name -SamAccountName $GroupLine.Name -GroupCategory Distribution -GroupScope Universal -DisplayName $GroupLine.Name -Path $DestinationOU
    }
    Add-ADGroupMember -Identity $GroupLine.Name -Members $GroupLine.ObjectGUID
}

#Same principal as above, but it removes the users if they are in a specific project but not in the exported CSV file. 
$GroupName = "Horizon $($GroupLine.Name)"
$objectGUID = Get-ADGroupMember -Identity $GroupName | Select-Object -ExpandProperty objectGUID

Foreach ($GroupLine in $GroupLine) {
    if ($($GroupLine.ObjectGuid) -notcontains $objectGUID) {
        Remove-ADGroupMember -Identity $GroupName -Members $objectGUID -Confirm:$False
    }
    else {
    }
}