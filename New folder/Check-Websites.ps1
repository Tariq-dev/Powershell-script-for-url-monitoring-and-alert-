#Script generates email report for Website checks in html format
#powershell v3 or higher is required!
#Created by Jatin Patel

$ErrorActionPreference = "SilentlyContinue";
$scriptpath = $MyInvocation.MyCommand.Definition 
$dir = Split-Path $scriptpath
if(!$dir){$dir=$pwd}

$AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols

$alert=$null;$preOutput=@();$alertOutput=@();$output=@();$allSites = gc .\Websites.txt

foreach ($site in $allSites) {
    For ($i=0; $i -lt 4; $i++) {
        $request = $null;$reqtime = (Measure-Command {$request = Invoke-WebRequest $site -UseBasicParsing}).TotalMilliseconds
        if ($request) {$output += [PSCustomObject] @{"Time"=Get-Date;"Site"=$site;"ResponseTime"=[math]::Round($reqtime/1000,1);"StatusCode"=[int]$request.StatusCode;"Description"=[string]$request.StatusDescription}; break}
        if (!($request) -AND ($i -eq 3)) {
            try {$request = $null;$reqtime = (Measure-Command {$request = Invoke-WebRequest $site -UseBasicParsing}).TotalMilliseconds}
            catch {$exception = $null;$exception = $_.Exception;$request = $_.Exception.Response;$reqtime = $null}
            If ($request) {$output += [PSCustomObject] @{"Time"=Get-Date;"Site"=$site;"ResponseTime"=[math]::Round($reqtime/1000,1);"StatusCode"=[int]$request.StatusCode;"Description"=[string]$request.StatusDescription}}
            Else {$output += [PSCustomObject] @{"Time"=Get-Date;"Site"=$site;"ResponseTime"=[math]::Round($reqtime/1000,1);"StatusCode"=402;"Description"=[string]$exception.Status}}
        }#if
    }#for
}#foreach

[array]$preOutput = Import-Csv $dir\output.csv
$output | Export-Csv $pwd\output.csv -Force
$Changes = Compare-Object $output $preOutput -Property Site,StatusCode,Description | group site | select -exp Name

If ($Changes) {
$alert = $true
    foreach ($change in $changes) {
        $changeOutput = $output | ?{$_.site -eq $change}
        $changePreOutput = $preOutput | ?{$_.site -eq $change}
        $alertOutput += [PSCustomObject] @{"Site"=$change;"LastCheckTime"=$changePreOutput.Time;"CheckTime"=$changeOutput.Time;"LastStatusCode"=$changePreOutput.StatusCode;"ChangedStatusCode"=[int]$changeOutput.StatusCode;"LastDescription"=[string]$changePreOutput.Description;"ChangedDescription"=[string]$changeOutput.Description}
    }
}

$output = $output | sort statuscode -desc
$subjectTime = Get-Date -Format dddd-hhmmtt

#EMAIL VARIABLES
$smtpServer = "smtp.1and1.com"
$credentials = new-object Management.Automation.PSCredential “tariq.shafiq@tenpearls.com”, (“10Pearls123%” | ConvertTo-SecureString -AsPlainText -Force)
$ReportSender = "Operations Alerts <tariq.shafiq@tenpearls.com>"
If ($alert) {
    $to = "tariq.shafiq@tenpearls.com","tariq.shafiq@tenpearls.com"
    $cc = "tariq.shafiq@tenpearls.com"
    $MailSubject = "Alert: Websites Availibility Report for $subjectTime"
} else {
   # $to = "tariq.shafiq@tenpearls.com"
   # $cc = "tariq.shafiq@tenpearls.com"
   # $MailSubject = "Websites Availibility Report for $subjectTime"
}

$reportPath = "$pwd\Logs\"
$reportName = "SiteCheckRpt_$(get-date -uformat %H%M).html";
$websiteReport = $reportPath + $reportName
If(Test-Path $websiteReport) {Remove-Item $websiteReport}
$redColor = "#FF6666"
$blueColor = "#4DB8FF"
$orangeColor = "#FFB84D"
$whiteColor = "#FFFFFF"
$greenColor = "#66FF66"
$i=$null
$arrayGreen = @($output | ?{$_.StatusCode -lt 401} | select -exp Statuscode); [int]$numGreen = $arrayGreen.Count #informational/redirect codes
$arrayBlue = @($output | ?{$_.StatusCode -eq 401} | select -exp Statuscode); [int]$numBlue = $arrayBlue.Count #auth required code
$arrayOrange = @($output | ?{$_.StatusCode -eq 402 -or $_.StatusCode -eq 403} | select -exp Statuscode); [int]$numOrange = $arrayOrange.Count #SSL/Forbidden client error codes
$arrayRed = @($output | ?{$_.StatusCode -ge 404} | select -exp Statuscode); [int]$numRed = $arrayRed.Count #404+ client errors and 5xx server error codes

$header = "
		<html>
		<head>
		<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
		<title>Site Check Report</title>
		<STYLE TYPE='text/css'>
		<!--
        table {
            border-collapse: collapse;
            border: thin solid #999999;
        	}
		td {
			font-family: 'Tahoma','Arial';
			font-size: 11px;
            border: thin solid #999999;
		    }
		body {
            font-family: 'Tahoma','Arial';
			margin-left: 5px;
			margin-top: 5px;
			margin-right: 5px;
			margin-bottom: 10px;
		    }
        .noBorder {
            font-family: 'Tahoma','Arial';
            border: none !important;
            text-align: center;
            font-weight: bold;
            font-size: 14px;
            }
        #blink_text {
			animation: blink 1s infinite;	
			}
		@keyframes blink {  
			0% { opacity: 1.0; }
			50% { opacity: 0.25; }
			100% { opacity: 1.0; }
			}
		-->
		</style>
		</head>
		<body>
"
 Add-Content $websiteReport $header

 If($alert) {
    $AlertSummary = "
        <table width='100%'>
        <tr bgcolor=`'$redColor'`>
	    <td colspan='7' height='25' align='center'>
	    <font id='blink_text' color='#FFFF00' face='Arial' size='3'><strong>Alert: Please check, site status has changed.</strong></font>
	    </td>
	    </tr>
	    </table>
        "
        Add-Content $websiteReport $AlertSummary

    $AlertTableHeader = "
        <table width='100%'>
            <tr bgcolor='#EEEEEE'>
	        <td colspan='7' height='25' align='center'>
	        <font face='Arial' color='#990000' size='4'><strong>Status Changed Website Report</strong></font>
	        </td>
	        </tr>
	        </table>
         <table width='100%'><tbody>
	        <tr bgcolor=#EEEEEE>
            <td align='center'><strong>No.</strong></td>
            <td align='center'><strong>Site</strong></td>
	        <td align='center'><strong>Last Check Time</strong></td>
	        <td align='center'><strong>Check Time</strong></td>
	        <td align='center'><strong>Last Status Code</strong></td>
	        <td align='center'><strong>Changed Status Code</strong></td>
	        <td align='center'><strong>Last Description</strong></td>
	        <td align='center'><strong>Changed Description</strong></td>
	        </tr>
        "
        Add-Content $websiteReport $AlertTableHeader

    foreach ($a in $alertOutput) {
        if ($arrayGreen -contains $a.LastStatusCode) {$aprecolor = $greenColor} elseif ($arrayBlue -contains $a.LastStatusCode) {$aprecolor = $blueColor} elseif ($arrayOrange -contains $a.LastStatusCode) {$aprecolor = $orangeColor} else {$aprecolor = $redColor}
        if ($arrayGreen -contains $a.ChangedStatusCode) {$acolor = $greenColor} elseif ($arrayBlue -contains $a.ChangedStatusCode) {$acolor = $blueColor} elseif ($arrayOrange -contains $a.ChangedStatusCode) {$acolor = $orangeColor} else {$acolor = $redColor}
        $dataRow = "
		    <tr>
		    <td><strong>$([array]::IndexOf($alertOutput,$a)+1)</strong></td>
            <td>$($a.Site)</td>
		    <td>$($a.LastCheckTime)</td>
		    <td>$($a.CheckTime)</td>
		    <td>$($a.LastStatusCode)</td>
		    <td>$($a.ChangedStatusCode)</td>
		    <td bgcolor=`'$aprecolor`'>$($a.LastDescription)</td>
		    <td bgcolor=`'$acolor`'>$($a.ChangedDescription)</td>
		    </tr>
        "
        Add-Content $websiteReport $dataRow
    }
    $AlertFooter = "</table><br>"
    Add-Content $websiteReport $AlertFooter
 }

 $tableSummary = "
    <table class='noBorder' width='100%' bgcolor='#505050' cellpadding='5'><tbody><tr>
    <td class='noBorder'><font color=`'$greenColor'`>Accessible: $numGreen</td>
    <td class='noBorder'><font color=`'$blueColor'`>Unauthorized: $numBlue</td>
    <td class='noBorder'><font color=`'$orangeColor'`>InvalidSSL: $numOrange</td>
    <td class='noBorder'><font color=`'$redColor'`>Inaccessible: $numRed</td>
    </tr></table>
 "
 Add-Content $websiteReport $tableSummary

 $tableHeader = "
<table width='100%'>
    <tr bgcolor='#EEEEEE'>
	<td colspan='7' height='25' align='center'>
	<font face='Arial' color='#557799' size='4'><strong>Website Availibity Report</strong></font>
	</td>
	</tr>
	</table>
 <table width='100%'><tbody>
	<tr bgcolor=#EEEEEE>
	<td align='center'><strong>No.</strong></td>
	<td align='center'><strong>Site</strong></td>
    <td align='center'><strong>Check Initiated</strong></td>
	<td align='center'><strong>Response Time (Sec)</strong></td>
	<td align='center'><strong>Status Code</strong></td>
	<td align='center'><strong>Description</strong></td>
	</tr>
"
Add-Content $websiteReport $tableHeader

foreach ($i in $output) {
    if ($arrayGreen -contains $i.StatusCode) {$color = $greenColor} elseif ($arrayBlue -contains $i.StatusCode) {$color = $blueColor} elseif ($arrayOrange -contains $i.StatusCode) {$color = $orangeColor} else {$color = $redColor}
    $dataRow = "
		<tr>
		<td><strong>$([array]::IndexOf($output,$i)+1)</strong></td>
		<td>$($i.Site)</td>
        <td>$($i.Time)</td>
		<td>$($i.ResponseTime)</td>
		<td>$($i.StatusCode)</td>
		<td bgcolor=`'$color`'>$($i.Description)</td>
		</tr>
"
Add-Content $websiteReport $dataRow;
}

 $tableDescription = "
 </table><br><table width='50%' bgcolor='white'>
	<tr>
    <td bgcolor=`'$greenColor'`>200</td>
    <td>The site is accessible</td>
    </tr><tr>
	<td bgcolor=`'$blueColor'`>401</td>
    <td>Access is denied because credential for admin account needed</td>
    </tr><tr>
    <td bgcolor=`'$orangeColor'`>402</td>
    <td>Placeholder code for SSL errors, Timeouts and Connect failures</td>
    </tr><tr>
    <td bgcolor=`'$orangeColor'`>403</td>
    <td>Internet usage to this site is monitored and logged</td>
    </tr><tr>
    <td bgcolor=`'$redColor'`>404</td>
    <td>The requested resource could not be found</td>
    </tr><tr>
    <td bgcolor=`'$redColor'`>503</td>
    <td>The site is either unavailable or invalid url</td>
	</tr></table>
"
Add-Content $websiteReport $tableDescription
Add-Content $websiteReport "</body></html>"

### BEGIN Send Email ###
$body = [System.IO.File]::ReadAllText($websiteReport)
If($cc) {Send-MailMessage -SmtpServer $smtpServer -Credential $credentials -Subject $MailSubject -Body $body -to $to -Cc $cc -From $ReportSender -BodyAsHtml -UseSsl}
else {Send-MailMessage -SmtpServer $smtpServer -Credential $credentials -Subject $MailSubject -Body $body -to $to -From $ReportSender -BodyAsHtml -UseSsl}
### END Send Email ###>