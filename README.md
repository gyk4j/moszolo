# o2010s
PowerShell script to download the free Microsoft Office 2010 Starter as a locally-cached copy for reuse or archival.

Inspired and rewritten from 
[Office 2010 Starter downloader AutoIt script by Jemboy][downloader].

Written for learning and practising [AutoIt][autoit] and [PowerShell][pwsh] 
scripting.

Created after watching :tv:

[![CyberCPU Tech: How To Get Legit Microsoft Office For Free](https://img.youtube.com/vi/ud0WTQcTgSE/0.jpg)](https://www.youtube.com/watch?v=ud0WTQcTgSE)

# How it works

Office 2010 Starter Edition is packaged using the [Microsoft App-V][ms-app-v] 
virtualization technology to enable it to be downloaded and run on-demand 
without prior installation.

`o2010s` is a PowerShell script that downloads and parses the App-V's manifest 
and catalog files in XML format to identify and download the appropriate 
required files for Office 2010 Starter Edition to install and run successfully.

The downloaded files can be archived and preserved as a local copy for future 
reuse.

# Caution Against Using Microsoft Office 2010

> [!CAUTION]
>
> There is NO REASON to use Microsoft Office 2010 today.
>
> It is a *seriously outdated and unmaintained* software, with no security 
> updates and perhaps plenty of security vulnerabilities ever since support 
> ended on **October 13, 2020**.

# Recommendations

- Use [Office 365][office-365] with a [free Microsoft account][ms-acct] if you 
  can.
- If you insist on using Microsoft Office 2010 for whatever reason, use it only
  from within a virtual machine and/or without Internet connectivity.

# Usage

*o2010s* only requires *PowerShell 3.0* or above. Higher versions of PowerShell
are pre-installed with modern versions of Windows in use. Windows 10 and 11 
come with *PowerShell 5.1* out-of-the-box, so no further action is required.

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
  - ar-sa (Arabic)
  - da-dk (Danish)
  - de-de (German)
  - en-us (English)
  - es-es (Spanish)
  - fr-fr (French)
  - he-il (Hebrew)
  - hi-in (Hindi)
  - it-it (Italian)
  - ja-jp (Japanese)
  - ko-kr (Korean)
  - nl-nl (Dutch)
  - pl-pl (Polish)
  - pt-br (Portuguese)
  - ru-ru (Russian)
  - sv-se (Swedish)
  - th-th (Thai)
  - zh-cn (Chinese Simplified)
  - zh-tw (Chinese Traditional)

[downloader]: https://www.autoitscript.com/forum/topic/205471-office-2010-starter-downloader/
[autoit]: https://www.autoitscript.com/site/
[pwsh]: https://learn.microsoft.com/en-us/powershell/
[ms-app-v]: https://en.wikipedia.org/wiki/Microsoft_App-V
[office-365]: https://www.office.com/
[ms-acct]: https://account.microsoft.com/account/Account
