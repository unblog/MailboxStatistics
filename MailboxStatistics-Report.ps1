<#
  MailboxStatistics-Report.ps1 generate mailbox statistics reporting to out-file
  Version 1.0.3 (27.03.2017) by DonMatteo
  Mail: think@unblog.ch
  Blog: think.unblog.ch
#>
Add-PSSnapin Microsoft.Exchange.Management.Powershell.Admin -erroraction silentlyContinue
$date = Get-Date -format F
$exch = [system.environment]::MachineName
$build = (Get-ExchangeServer -Identity $env:exch | ft name,AdminDisplayVersion -HideTableHeaders -AutoSize)
$path = [Environment]::GetFolderPath('ApplicationData')
$file = "$path\report.txt"
$mailboxdata = (Get-MailboxStatistics -Server $exch | Sort-Object TotalItemSize -Descending | ft DisplayName,@{label="TotalItemSize(MB)";expression={$_.TotalItemSize.Value.ToMB()}},ItemCount,LastLogonTime -auto)
$HealthReport = (Get-HealthReport -Server $exch | Where-Object { $_.alertvalue -ne "Healthy" })
$disk = (Get-WMIObject Win32_Logicaldisk -ComputerName $exch | Select PSComputername,DeviceID,@{Name="SizeGB";Expression={$_.Size/1GB -as [int]}},@{Name="FreeGB";Expression={[math]::Round($_.Freespace/1GB,2)}})
$date | Out-File $file
$build | Out-File $file -append
$mailboxdata | Out-File $file -append
$HealthReport | Out-File $file -append
$disk | Out-File $file -append
$smtpServer = "127.0.0.1"
$att = new-object System.Net.Mail.Attachment($file)
$msg = new-object System.Net.Mail.MailMessage
$smtp = new-object System.Net.Mail.SmtpClient($smtpServer)
$msg.From = "no_reply@$exch"
$msg.To.Add("your@email.com")
$msg.Subject = "Notification from $exch"
$msg.Body = "$exch MailboxStatistics generated $date reported to attachment $file"
$msg.Attachments.Add($att)
$smtp.Send($msg)
$att.Disposen
