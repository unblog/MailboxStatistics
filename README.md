# MailboxStatistics
A powershell script to generate ms exchange mailbox statistics and health report will sending per email.
 
# Usage:
Ran on exchange host as a task scheduled job at your desired time and action prog *powershell.exe*<br>
use ``example`` param:
```sh
-psconsolefile "C:\Program Files\Microsoft\Exchange Server\V15\Bin\exshell.psc1" -file "C:\windows\system32\MailboxStatistics-Report.ps1"
```
# References
* [Source](http://think.unblog.ch/winstat-user-status/)
---
Screenshot:
<img src="report-screen.png">
