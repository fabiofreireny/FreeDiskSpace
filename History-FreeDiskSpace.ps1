# This script should be run as an account that has db_reader to the database
# This script requires SQLPS (SQL) and PowerCLI (VMware) Powershell modules

import-module sqlps -WarningAction SilentlyContinue

# Free space percentage thresholds
$warning  = 25
$critical = 10

$freeSpaceFileName = "c:\temp\FreeSpace.html"
$serverList        = get-content "c:\tasks\serverList.txt"

$dbServer = "databaseServerName"
$dbname   = "databseName"

$today     = (get-date).toString('yyyy-MM-dd')
$yesterday = (get-date).adddays(-1).toString('yyyy-MM-dd')
$lastWeek  = (get-date).adddays(-7).toString('yyyy-MM-dd')
$lastMonth = (get-date).adddays(-30).toString('yyyy-MM-dd')

$exceptionExclusions = @(
    "New York",
    "exclusion2",
    "exclusion3"
)

New-Item -ItemType file $freeSpaceFileName -Force

Function writeHtmlHeader {
    param($fileName)

    Add-Content $fileName "<html>"
    Add-Content $fileName "<head>"
    Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
    add-content $fileName '<STYLE TYPE="text/css">'
    add-content $fileName  "<!--"
    add-content $fileName  "td {"
    add-content $fileName  "font-family: Tahoma;"
    add-content $fileName  "font-size: 11px;"
    add-content $fileName  "border-top: 1px solid #999999;"
    add-content $fileName  "border-right: 1px solid #999999;"
    add-content $fileName  "border-bottom: 1px solid #999999;"
    add-content $fileName  "border-left: 1px solid #999999;"
    add-content $fileName  "padding-top: 0px;"
    add-content $fileName  "padding-right: 0px;"
    add-content $fileName  "padding-bottom: 0px;"
    add-content $fileName  "padding-left: 0px;"
    add-content $fileName  "}"
    add-content $fileName  "body {"
    add-content $fileName  "margin-left: 5px;"
    add-content $fileName  "margin-top: 5px;"
    add-content $fileName  "margin-right: 0px;"
    add-content $fileName  "margin-bottom: 10px;"
    add-content $fileName  ""
    add-content $fileName  "table {"
    add-content $fileName  "border: thin solid #000000;"
    add-content $fileName  "}"
    add-content $fileName  "-->"
    add-content $fileName  "</style>"
    Add-Content $fileName "</head>"
    Add-Content $fileName "<body>"
}

Function writeTableHeader {
    param($fileName)

    Add-Content $fileName "<table width='100%'>"
    Add-Content $fileName "<tr bgcolor=#CCCCCC>"
    Add-Content $fileName "<td width='26%' align='left'>Drive</td>"
    Add-Content $fileName "<td width='26%' align='left'>Drive Label</td>"
    Add-Content $fileName "<td width='8%' align='left'>Total Capacity (GB)</td>"
    Add-Content $fileName "<td style='border-left:solid 2px' width='8%' align='left'>Free Last Month (GB)</td>"
    Add-Content $fileName "<td width='8%' align='left'>Free Last Week (GB)</td>"
    Add-Content $fileName "<td width='8%' align='left'>Free Yesterday (GB)</td>"
    Add-Content $fileName "<td width='8%' align='left'>Free Space (GB)</td>"
    Add-Content $fileName "<td width='8%' align='left'>Free Space %</td>"
    Add-Content $fileName "</tr>"
}

Function writeHostname {
    param($fileName,$hostname)

    Add-Content $fileName "<tr bgcolor='#CCCCCC'>"
    Add-Content $fileName "<td width='100%' align='left' colSpan=8><font face='tahoma' color='#003399' size='2'><strong> $hostname </strong></font></td>"
    Add-Content $fileName "</tr>"
}

Function writeHtmlFooter {
    param($fileName)

    Add-Content $fileName "</body>"
    Add-Content $fileName "</html>"
}

Function colorPick {
    param([int]$prior,[int]$current)

    $colorscale = @("#ff0000","#b30000","#660000","BLACK","#006600","#00b300","#00ff00")

    if ($prior -ge $current) {
        if ($current -eq 0) { return $colorScale[3] }
        $percent = (($prior - $current)/$current)
        }
    else {
        if ($prior -eq 0) { return $colorScale[3] }
        $percent = (($current - $prior)/$prior * -1)
        }

    #change color at 5%, 20% and 40% (plus or minus)
    switch ($percent) {
        {$percent -le -0.4                         } {$index = 0}
        {$percent -le -0.2  -and $percent -gt  -0.4} {$index = 1}
        {$percent -le -0.05 -and $percent -gt  -0.2} {$index = 2}
        {$percent -gt -0.05 -and $percent -lt  0.05} {$index = 3}
        {$percent -ge  0.05 -and $percent -lt   0.2} {$index = 4}
        {$percent -ge  0.2  -and $percent -lt   0.4} {$index = 5}
        {$percent -ge  0.4                         } {$index = 6}
    }

    return $colorScale[$index]
}


Function writeDiskInfo {
    param($filename,$drive,$label,$frSpace,$totSpace,$yesterday,$lastWeek,$lastMonth)
    #$usedSpace = [int]$totSpace - $frspace
    $freePercent = ($frspace/$totSpace)*100

    $usedSpace=[Math]::Round($totSpace - $frspace ,0)
    $freePercent = [Math]::Round($freePercent,0)

    Add-Content $fileName "<tr>"
    Add-Content $fileName "<td>$drive</td>"
    Add-Content $fileName "<td>$label</td>"
    Add-Content $fileName "<td align=right>$totSpace</td>"
    #Add-Content $fileName "<td align=right>$usedSpace</td>"
    Add-Content $fileName "<td  style='border-left:solid 2px' align=right>$lastmonth</td>"

    #Change color based on trend (relative to prior period)
    $lastWeekColor = colorPick $lastWeek $lastMonth

    $yesterdayColor = colorPick $yesterday $lastWeek

    $todayColor = colorPick $frSpace $yesterday

    Add-Content $fileName "<td align=right><font color=$lastWeekColor>$lastweek</td>"
    Add-Content $fileName "<td align=right><font color=$yesterdayColor>$yesterday</td>"
    Add-Content $fileName "<td align=right><font color=$todayColor>$frSpace</td>"

    #Change color based on free space percentage (absolute; warning at $warning, critical at $critical)
    switch ($freePercent) {
        {$_ -gt $warning}  { $freeColor = "bgcolor='PALEGREEN'" }
        {$_ -le $critical} { $freeColor = "bgcolor='RED'><font color='WHITE'" }
        default            { $freeColor = "bgcolor='ORANGE'" }
    }

    Add-Content $fileName "<td align=center $freeColor>$freePercent</td>"
    Add-Content $fileName "</tr>"
}

Function sendEmail {
    param($from,$to,$subject,$smtphost,$htmlFileName)

    $body = Get-Content $htmlFileName
    $smtp= New-Object System.Net.Mail.SmtpClient $smtphost
    $msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body
    $msg.isBodyhtml = $true
    $smtp.send($msg)
}

writeHtmlHeader $freeSpaceFileName
writeTableHeader $freeSpaceFileName

#Get host list
$query      = "select distinct host from freeSpace where date like '$today'"
$allServers = invoke-sqlcmd -HostName $dbServer -Database $dbName -Query $query

#Put VMware in front and sort the rest
$allServers = $exceptionExclusions + (Compare-Object $allservers.host  $exceptionExclusions -PassThru | sort )

#for each host
$allservers | % {
    $hostname = $_
    writeHostname $freeSpaceFileName $hostname

    #get drives for each host
    $query     = "select distinct drive from freeSpace where date like '$today' and host like '$hostname'"
    $allDrives = invoke-sqlcmd -HostName $dbServer -Database $dbName -Query $query

    #for each drive on each host
    $allDrives | % {
        $drive       = $_.drive
        $query       = "select * from freeSpace where host like '$hostname' and date >= '$lastMonth' and drive like '$drive'"
        $freeHistory = invoke-sqlcmd -HostName $dbServer -Database $dbName -Query $query

        #get free space values for today, -1, -7 and -30 days
        $freeToday     = ($freeHistory | where {$_.date -eq $today}     | select -Property freeSpace).freeSpace
        $freeYesterday = ($freeHistory | where {$_.date -eq $yesterday} | select -Property freeSpace).freeSpace
        $freeLastWeek  = ($freeHistory | where {$_.date -eq $lastWeek}  | select -Property freeSpace).freeSpace
        $freeLastMonth = ($freeHistory | where {$_.date -eq $lastMonth} | select -Property freeSpace).freeSpace

        $label         = ($freeHistory | where {$_.date -eq $today}     | select -Property label    ).label
        $total         = ($freeHistory | where {$_.date -eq $today}     | select -Property totalsize).totalsize

        writeDiskInfo $freeSpaceFileName $_.drive $label $freeToday $total $freeYesterday $freeLastWeek $freeLastMonth
    }
}

#Missing hosts
$diff = Compare-Object ($serverList + $exceptionExclusions) $allServers | Select @{Name="Missing Hosts";Expression={$_.InputObject}} | ConvertTo-Html -Fragment -Property "Missing Hosts"
Add-Content $freeSpaceFileName "</table><br>$diff"

writeHtmlFooter $freeSpaceFileName

sendEmail noreply@fabio.nyc fabio@fabio.nyc "Disk Space Trend Report - $today" relayHost $freeSpaceFileName
