# o2010s
PowerShell script to download the free 
[Microsoft Office Starter 2010][msoffice-starter-2010] as a locally-cached copy 
for reuse or archival.

Created after watching 
[CyberCPU Tech: How To Get Legit Microsoft Office For Free][cybercpu] :tv:

[![CyberCPU Tech: How To Get Legit Microsoft Office For Free](https://img.youtube.com/vi/ud0WTQcTgSE/0.jpg)][cybercpu]

Inspired and rewritten from 
[Office Starter 2010 downloader AutoIt script by Jemboy][downloader].

Written for studying [AutoIt][autoit] and practising [PowerShell][pwsh] 
scripting.

# How it works

In [Microsoft's own words][msoffice-starter-2010],

> Microsoft Office Starter 2010 is a simplified, ad-funded version of Microsoft 
> Office 2010 that comes pre-loaded and ready to use on your computer. Office 
> Starter includes the spreadsheet program Microsoft Excel Starter 2010 and the 
> word processing program Microsoft Word Starter 2010.

Microsoft Office Starter 2010 is packaged using the [Microsoft App-V][ms-app-v] 
virtualization technology for downloading and running on-demand without prior 
installation.

`o2010s` is a PowerShell script that downloads the appropriate required files 
for Microsoft Office Starter 2010 to install and run successfully. To do so, it
downloads and parses the App-V's manifest and catalog files in XML format. 

The downloaded files can be archived and preserved as a local copy for future 
reuse.

# Caution

A word of caution against using Microsoft Office 2010 though.

> [!CAUTION]
>
> There is NO REASON to use Microsoft Office 2010 today.
>
> It is a *seriously outdated and unmaintained* software, with no security 
> updates and perhaps plenty of security vulnerabilities ever since 
> [support ended on October 13, 2020][support-end].
> 
> There is a safer yet free alternative available (see recommendations).

# Use Case

If use of Microsoft Office 2010 is discouraged today, what is it useful for 
then?

> [!TIP]
> If it is for doing *offline* document processing on a PC/laptop disconnected 
> from the Internet, and we cannot utilize the cloud-hosted SaaS solutions like 
> [Office 365][office-web-apps] and [Google Docs][gdocs], perhaps it is a good
> enough freeware solution.

# Recommendations

- Use [Microsoft Office Web Apps][office-web-apps] with a 
  [Microsoft account][ms-acct]. They are available for free.
- If Microsoft Office 2010 is used for whatever reason, use it with the 
  following security measures:
  1. Run from within a virtual machine
  2. Without Internet connectivity
  3. With up-to-date anti-virus and anti-malware protection
  4. Open or edit only trusted documents from known sources

# Usage

*o2010s* only requires *PowerShell 3.0* and above.

Windows 10 and 11 come with *PowerShell 5.1* out-of-the-box, so no further 
action is required.

From `PowerShell`

```pwsh
.\o2010s-downloader.ps1
```

From `Command Prompt`

```bat
powershell.exe o2010s-downloader.ps1
```

# Settings and Customizations

There are some variables and constants in the PowerShell script that allow users
to specify custom settings before executing the script. They include:

- `$DEBUG`: Runs in debug mode which simulates and skips downloading files from 
  servers while reusing existing files already on storage. Used only for 
  testing during development only. (default: `false`)
- `$lang`: Downloads the "English-US" version by default. (default: `en-us`)
  Languages still available for download include:

| ID    | Language   | ID    | Language   | ID    | Language   | ID    | Language   |
| ----- | ---------- | ----- | ---------- | ----- | ---------- | ----- | ---------- |
| ar-sa | Arabic     | fr-fr | French     | ko-kr | Korean     | sv-se | Swedish    |
| da-dk | Danish     | he-il | Hebrew     | nl-nl | Dutch      | th-th | Thai       |
| de-de | German     | hi-in | Hindi      | pl-pl | Polish     | zh-cn | Chinese S. |
| en-us | English    | it-it | Italian    | pt-br | Portuguese | zh-tw | Chinese T. |
| es-es | Spanish    | ja-jp | Japanese   | ru-ru | Russian    |       |            |

[msoffice-starter-2010]: https://support.microsoft.com/en-gb/office/getting-started-with-office-starter-379fba5a-6d82-4e19-aa2e-d41627f5ea5e
[downloader]: https://www.autoitscript.com/forum/topic/205471-office-2010-starter-downloader/
[autoit]: https://www.autoitscript.com/site/
[pwsh]: https://learn.microsoft.com/en-us/powershell/
[cybercpu]: https://www.youtube.com/watch?v=ud0WTQcTgSE
[ms-app-v]: https://en.wikipedia.org/wiki/Microsoft_App-V
[support-end]: https://support.microsoft.com/en-us/office/end-of-support-for-office-2010-3a3e45de-51ac-4944-b2ba-c2e415432789
[office-web-apps]: https://www.office.com/?ms.officeurl=webapps
[gdocs]: https://docs.google.com
[ms-acct]: https://account.microsoft.com/account/Account
