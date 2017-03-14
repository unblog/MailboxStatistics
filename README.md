# MailboxStatistics
A powershell script to generate MS Exchange mailbox statistics and health report will sending per email.

[![Travis](https://img.shields.io/travis/rust-lang/rust.svg)](https://github.com/donkey/MailboxStatistics)
[![GitHub issues](https://img.shields.io/github/issues/donkey/MailboxStatistics.svg)](https://github.com/donkey/MailboxStatistics/issues)
[![Github Releases](https://img.shields.io/github/downloads/atom/atom/latest/total.svg)](https://github.com/donkey/MailboxStatistics)
[![GitHub forks](https://img.shields.io/github/forks/donkey/MailboxStatistics.svg)](https://github.com/donkey/MailboxStatistics/network)
[![GitHub stars](https://img.shields.io/github/stars/donkey/MailboxStatistics.svg)](https://github.com/donkey/MailboxStatistics/stargazers)
[![GitHub license](https://img.shields.io/badge/license-AGPL-blue.svg)](https://raw.githubusercontent.com/donkey/MailboxStatistics/master/LICENSE)
[![WordPress rating](https://img.shields.io/wordpress/plugin/r/akismet.svg)](https://github.com/donkey/MailboxStatistics)

## Installation
It requires that the script is executed in the Exchange management shell ``exshell.psc1``.

## Usage:
Ran on exchange host as a task scheduled job at your desired time and action prog *powershell.exe*<br>
use ``example`` param:
```sh
-psconsolefile "C:\Program Files\Microsoft\Exchange Server\V15\Bin\exshell.psc1" -file "C:\windows\system32\MailboxStatistics-Report.ps1"
```
## References
* [Source](http://think.unblog.ch/winstat-user-status/)
---
Screenshot:
<img src="report-screen.png">
