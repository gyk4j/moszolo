[bool]$DEBUG = $false

[string]$lang = 'ja-jp'
[string]$version = '14.0.4763.1000'

function Download-Urls( [Collections.Generic.List[string]]$Urls ){
    [string[]]$skipBigFile = @()
    # ".sft", ".dsft", ".msi", ".exe", ".dll", ".cab"

    $urls | ForEach-Object {
        if(!$DEBUG){
            [string]$url = $_;
            [string]$filename = Split-Path $_ -Leaf
            

            [string]$extension = [System.IO.Path]::GetExtension($filename)

            
            if($skipBigFile.Contains($extension.ToLower())){
                $filename
                New-Item -ItemType File -Name $filename -Force | Out-Null
            }
            else{
                $url
                Invoke-WebRequest -Uri $url -OutFile $filename
            }
        }
        else{
            [string]$path = $_.Replace("http://", "..\..\").Replace("/", "\")
            [string]$filename = Split-Path $_ -Leaf

            if(Test-Path -Path $path){
                $path
                Copy-Item -Path $path -Destination $filename
            }
            else{
                $filename
                New-Item -ItemType File -Name $filename -Force | Out-Null
            }
        }
    }
    $urls.Clear()
}

function Parse-Xml( [string]$XmlFile, [string]$XPath, [string]$Attribute, [Collections.Generic.List[string]]$DownloadQueue ){
    if(Test-Path -Path $XmlFile){
        Select-Xml -Path $XmlFile -XPath $XPath | ForEach-Object { 
            $DownloadQueue.Add($_.Node.Attributes[$Attribute].Value)
        }
    }
    else{
        [string]::Format("Error: {0} is missing", $XmlFile)
    }
}

function Extract-Cabinet( [string]$Cab, [string]$Directory ){
    if((Test-Path -Path $Cab) -and (Get-Item $Cab).Length -gt 0){
        "Extracting to " + $Directory
        Start-Process -Wait -NoNewWindow -FilePath "C:\Windows\System32\extrac32.exe" -ArgumentList "/Y", "/A", $Cab, "/E", "/L", $Directory
    }
    else{
        "Skipped " + $Cab + ". File does not exist or it is empty."
    }
}

function Write-Readme(){
    [string]$content = @'
To install Office 2010 Starter, go to into the 'bin' folder and to the right language directory.
E.g. '.../bin/de-de' is the folder for the German Office 2010 Starter installation files.
 
1. Run SetupConsumerC2ROLW.exe with right mouse click 'Run as Administrator'
2. After the installation, double click click2run-x-none.msp (KB2598285) within de update folder and restart
   your computer. https://www.catalog.update.microsoft.com/Search.aspx?q=2598285
 
After the above reboot, Office 2010 Starter will work on Windows 10 20H2.
 
To copy the shortcuts to your desktop:
a. Go to search and type 'Word'
b. At 'Microsoft Word Starter 2010', click on 'Open file location'
c. Now copy the shortcuts (Word and Excel) to your desktop.
 
Office 2010 Starter works with a virtualization technology, which creates a 
system drive Q:. If you are already using drive Q: then Office 2010 Starter won't start.
 
Office 2010 starter files have been unavailable for some time, but Microsoft quietly 
activated the downloads again a few years ago.

'@

    $content | Out-file -FilePath "README.TXT"
    "README.TXT"
}

[string]$scriptDir = $PSCommandPath | Split-Path -Parent
[string]$downloadBaseDir = [string]::Format("{0}\bin\{1}", $scriptDir, $lang)

New-Item -Path $downloadBaseDir -ItemType Directory -Force | Out-Null
Set-Location $downloadBaseDir

[string]$remoteBaseUrl = [string]::Format("http://c2r.microsoft.com/consumerC2R/{0}/{1}/", $lang, $version)
[string]$localBasePath = [string]::Format(".\c2r.microsoft.com\ConsumerC2R\{0}\{1}\", $lang, $version)
[string]$windowsUpdateCatalogBaseUrl = "http://download.windowsupdate.com/d/msdownload/update/software/crup/"

[Collections.Generic.List[string]]$downloadQueue = New-Object Collections.Generic.List[string]

"--- Phase 1: Downloading base files..."
$downloadQueue.Add($remoteBaseUrl + "setupconsumerc2rolw.exe")
$downloadQueue.Add($remoteBaseUrl + "PackageProperties.xml")

Download-Urls -Urls $downloadQueue

"--- Phase 2: Downloading package..."

Parse-Xml -XmlFile "PackageProperties.xml" `
    -XPath '/PackageProperties/Files/File' `
    -Attribute "ServerLocation" `
    -DownloadQueue $downloadQueue

Download-Urls -Urls $downloadQueue

"--- Phase 3: Downloading differential SFT files..."

Parse-Xml -XmlFile "descriptor.xml" `
    -XPath '/ToothpickDescriptor/DiffSftList/File' `
    -Attribute "Url" `
    -DownloadQueue $downloadQueue

Download-Urls -Urls $downloadQueue

"--- Phase 4: Downloading Office updates..."

# KB2598285 (Sep 18, 2013)
$downloadQueue.Add($windowsUpdateCatalogBaseUrl + "2013/09/click2run-x-none_f74703316deaa94b7b7e72bfcf7bd718910e26a4.cab")
                    
# KB2837578 (Apr 14, 2015)
$downloadQueue.Add($windowsUpdateCatalogBaseUrl + "2015/04/click2run-x-none_235e9328493ac7313f6862e502a8b4b0a14e8066.cab")

# KB2986257 (May 12, 2015)
$downloadQueue.Add($windowsUpdateCatalogBaseUrl + "2015/05/click2run-x-none_943e2c5aea7957af886fd3700268c72750513797.cab")

Download-Urls -Urls $downloadQueue

"--- Phase 5: Extracting update cabinet files"
Extract-Cabinet -Cab "click2run-x-none_f74703316deaa94b7b7e72bfcf7bd718910e26a4.cab" -Directory "KB2598285_20130918"
Extract-Cabinet -Cab "click2run-x-none_235e9328493ac7313f6862e502a8b4b0a14e8066.cab" -Directory "KB2837578_20150414"
Extract-Cabinet -Cab "click2run-x-none_943e2c5aea7957af886fd3700268c72750513797.cab" -Directory "KB2986257_20150512"

"--- Phase 6: Writing information files"
Write-Readme

"--- Done"
