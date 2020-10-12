<#
.SYNOPSIS
Collects free disk space information on VMware volumes, Windows drives and mount points, and stores them in a
database.

.DESCRIPTION
This script will scan all VMware datastores (it will ignore local datastores when the server is attached to NFS),
Windows drives and Windows mount points (Exchange), then populate a database with the info it finds.
You can then use Report-FreeDiskSpace.ps1 or History-FreeDiskSpace.ps1 to obtain a report with the latest info
or historical trends, respectively.

This script should be run as an account that has administrative permissions to those targets.

This script requires SQLPS (SQL) and PowerCLI (VMware) Powershell modules

SQL is setup with a primary key on HOST, DRIVE and DATE, that way there will only be one data point for each drive per day.
This script will display an error message when trying to populate SQL if an entry already exists for that day.

#>

#######################################
# Make sure to review and change the variables below as appropriate for your environment
#######################################

$serverlist = "c:\tasks\serverList.txt"

$dbServer = "name of db server including instance"
$dbName   = "name of database"
$vCenters = @("array of vCenter servers")

$date = (get-date).ToString('yyyy/MM/dd')

import-module sqlps -WarningAction SilentlyContinue

#VMware integration
& "C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1"

Function writeDbInfo {
    param($hostname,$drive,$label,$frSpace,$totSpace)

    $totSpace=[math]::Round(($totSpace),0)
    $frSpace=[Math]::Round(($frSpace),0)

    #Write SQL info
    $query = "INSERT INTO freeSpace values ('$hostname','$drive','$label','$date',$totSpace,$frSpace)"
    invoke-sqlcmd -HostName $dbServer -Database $dbName -query $query
}

Function sendEmail {
    param($from,$to,$subject,$smtphost,$htmlFileName)

    $body = Get-Content $htmlFileName
    $smtp= New-Object System.Net.Mail.SmtpClient $smtphost
    $msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
    $msg.isBodyhtml = $true
    $smtp.send($msg)
}


#VM Hosts (more than one vCenter)
$vCenters | % {
    Connect-VIServer -Server $_ -Protocol https
    $vmDatacenters = Get-Datacenter
    $vmDatacenters
    foreach ($datacenter in $vmDatacenters) {

        $vmStorage = Get-Datastore -Location $datacenter | Select-Object -Property Name,
            Type,
            State,
            @{Name="Capacity";Expression={"{0:N0}" -f ($_.CapacityGB)}},
            @{Name="FreeSpace";Expression={"{0:N0}" -f ($_.FreeSpaceGB)}} |`
                where {($datacenter.Name -eq "Remote Offices") -or ($_.Type -eq "NFS") -and ($_.State -eq "Available")} |`
                Sort-Object -Property Name

        foreach ($store in $vmStorage) {
            writeDbInfo $datacenter $store.Name $store.type $store.FreeSpace $store.Capacity
            $store.name
        }
    }
}

#Windows Servers
foreach ($server in Get-Content $serverlist) {
    $server = $server.ToUpper()

    $server
    Clear-Variable dp -ErrorAction SilentlyContinue
    $dp = Get-WmiObject win32_logicaldisk -ComputerName $server |  Where-Object {$_.drivetype -eq 3}

    foreach ($item in $dp) {
        #Write-Host  $item.DeviceID  $item.VolumeName $item.FreeSpace $item.Size
        writeDbInfo $server $item.DeviceID $item.VolumeName ($item.FreeSpace/1GB) ($item.Size/1GB)
    }

    #Exchange (mount points)
    if ($server -like "*dag*") {
        $dv = Get-WmiObject win32_volume -ComputerName $server | `
            where {($_.driveletter -eq $null) -and ($_.label -ne "System Reserved")} | `
            select -Property name, label, capacity, freespace | sort -Property name
        $dv | % {
            writeDbInfo $server $_.name $_.label ($_.FreeSpace/1GB) ($_.capacity/1GB)
        }
    }
}
