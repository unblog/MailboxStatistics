<#
  MailboxStatistics-html-Report.ps1 generate mailbox statistics reporting to out-file
  Version 1.0.3 (27.03.2017) by DonMatteo
  Mail: think@unblog.ch
  Blog: think.unblog.ch
#>

$a = "<style>"
$a = $a + "BODY{background-color:GhostWhite;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:Gold;}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:Azure;}"
$a = $a + "</style>"

Add-PSSnapin Microsoft.Exchange.Management.Powershell.Admin -erroraction silentlyContinue

$email = "your@email.com"
$date = Get-Date -format F
$exch = [system.environment]::MachineName
$build = (Get-ExchangeServer -Identity $env:exch | ft name,AdminDisplayVersion -HideTableHeaders -AutoSize)
$path = [Environment]::GetFolderPath('ApplicationData')
$file = "$path\report.html"
	if (test-Path $path\report.html) { remove-Item $path\report.html;
	Write-Host -ForegroundColor white -BackgroundColor Red    "Old file removed"
	}
ConvertTo-Html -Head $a  -Title "Exchange Mailbox statistics and health status for $exch" -Body "<h1> Computer Name : $exch </h1>" >  "$path\report.html"
$mailboxdata = (Get-MailboxStatistics -Server $exch | Sort-Object TotalItemSize -Descending | select DisplayName,@{label="TotalItemSize(MB)";expression={$_.TotalItemSize.Value.ToMB()}},ItemCount,LastLogonTime)
$HealthReport = (Get-HealthReport -Server $exch | Where-Object { $_.alertvalue -ne "Healthy" } | Select Server,State,HealthSet,HealthGroup,AlertValue,LastTransitionTime,MonitorCount,HaImpactingMonitorCount)
$disk = (Get-WMIObject Win32_Logicaldisk -ComputerName $exch | Select PSComputername,DeviceID,@{Name="SizeGB";Expression={$_.Size/1GB -as [int]}},@{Name="FreeGB";Expression={[math]::Round($_.Freespace/1GB,2)}})
$date | Out-file $file -append
$build | Out-file $file -append
$mailboxdata | ConvertTo-html -Body "<H2> Mailbox Statistics </H2>" >> $file
$HealthReport | ConvertTo-html -Body "<H2> Exchange Health Report </H2>" >> $file
$disk | ConvertTo-html -Body "<H2> Disk Space usage </H2>" >> $file
$smtpServer = "127.0.0.1"
$att = new-object System.Net.Mail.Attachment($file)
$msg = new-object System.Net.Mail.MailMessage
$smtp = new-object System.Net.Mail.SmtpClient($smtpServer)
$msg.From = "no_reply@$exch"
$msg.To.Add($email)
$msg.Subject = "Notification report from $exch"
$msg.Body = "$exch MailboxStatistics generated $date reported to attachment $file"
$msg.Attachments.Add($att)
$smtp.Send($msg)
$att.Disposen
