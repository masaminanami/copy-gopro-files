#
# Copy-GoPro --- copy GoPro files to local folder and OneDrive (for business)
#
using module .\lib\Config.psm1
using module .\GoPro.psm1
using module .\OneDrive.psm1
using module .\Copy-Gopro.resources.ja-jp.psm1
using module .\MsalWrapper.psm1
using module .\FileInfo.psm1

param(
[string]$Name,
[Alias('Dest')]
[string]$LocalDestination,
[Alias('ODDest','OneDrive','OD')]
[string]$OneDriveDestination,
[string]$Remark = "特別な1日",
[string]$Date = $null,
[string]$Cache,
[int]$BufferSize = 100,
[switch]$NoRename,
[string]$LogFile = "log.txt",
[string]$AppConfig = "AppConfig.json",
[string]$LibPath = ".\lib",
[switch]$Help
)

$ErrorActionPreference = "stop"

$script:GoProDevice # initialized by AppConfig.json
$script:GoProDrive # initialized by AppConfig.json
$script:AppClientId # initialized by AppConfig.json
$script:TenantId # initialized by AppConfig.json
$script:AppRedirectUri # initialized by AppConfig.json

$script:ScriptRoot = $MyInvocation.PSScriptRoot ? $MyInvocation.PSScriptRoot : ($MyInvocation.InvocationName -match '\.ps1$' ? (Split-Path -Parent $MyInvocation.InvocationName) : (Get-Location).Path)

Import-Module (Join-Path $script:ScriptRoot $script:LibPath "stdps.psm1")

<#
# Logfile locator
#>
function getLogFilePath() {
    if ($fp = $script:LogFile) {
        if (-not ([System.IO.Path]::IsPathRooted($fp))) {
            $fp = Join-Path $script:ScriptRoot $fp
        }
    }
    $fp
}

<#--- covert byte to MB, GB, TB ---#>
function toByteCountString([double]$s) {
    switch ($s) {
        {$_ -le 1GB} { return "$([Math]::Round($s/1MB, 2)) MB" }
        {$_ -le 1TB} { return "$([Math]::Round($s/1GB, 2)) GB" }
    }
    "$([Math]::Round($s/1TB, 2)) TB"
}

Set-StrictMode -Version latest
[Config]::Init((Join-Path $ScriptRoot $script:AppConfig))
RunApp ([Main]::New()) (getLogfilePath) 4
return;

class Main {
    $goproFiles;
    $msal;

    Main() {
        $this.goproFiles = @()
    }

    Run() {
        if (-not $this.checkEnv()) { return }

        #--- getting files from gopro devices
        $this.getGoProFiles($scrit:Name, $script:GoProDevice, $script:GoProDrive)
        if (-not $this.goproFiles) { return }

        #--- sign in to OD/ODB if specified
        if ($this.msal -and -not $this.OneDriveSignIn()) { return }

        if ($script:LocalDestination) {
            foreach ($dest in $script:LocalDestination -split(',')) {
                $dest = [IO.Path]::GetFullPath($dest)
                log ([CopyGoProMessages]::CopyingTo -f $dest)
                if (-not $this.copyToLocal($dest)) { return }
            }
        }

        if ($this.msal) {
            $dest = $script:OneDriveDestination
            log ([CopyGoProMessages]::CopyingTo -f $dest)
            $this.CopyToOneDrive($dest, $script:Cache)
        }

        log ([CopyGoProMessages]::AllDone)
    }

    #--- envrionmental check
    [bool] checkEnv() {
        #--- Check LocalDestination
        if ($script:Cache) {
            #--- current cache implementtion is No-LocalDDestination
            if (-not (Test-Path $script:Cache)) {
                log ([CopyGoProMessages]::NoDestinationFolder -f $script:Cache)
                return $false
            }
        }

        if ($script:LocalDestination) {
            foreach ($dir in $script:LocalDestination -split(',')) {
                if (-not (Test-Path $dir -PathType Container)) {
                    log ([CopyGoProMessages]::NoDestinationFolder -f $dir)
                    return $false
                }
            }
        }

        if (-not $script:OneDriveDestination) {
            log ([CopyGoProMessages]::OneDriveDestinationNotSpecified)
        } else {
            $this.msal = [MSALWrapper]::New()
        }
        return $true

        if (-not $script:LocalDestination -and -not $script:OneDriveDestination) {
            log ([CopyGoProMessages]::LocalDestinationNotSpecified)
            return $false
        }
    }

    getGoProFiles($devName, $devSpec, $drvSpec) {
        $drives = [GoProDrive]::Discover($devName, $devSpec, $drvSpec)
        $files = @()
        #--- find files from folders
        $dateLimit = $null
        if ($script:Date) {
            $dateLimit = Get-Date $script:Date
            log ([CopyGoProMessages]::DateLimitSet -f "$($datelimit.tostring('yyyy/MM/dd HH:mm'))")
        } else {
            $datelimit = Get-Date "1900/1/1"
        }
        foreach ($drv in $drives) {
            if ($r = $drv.GetFiles()) {
                $r = @($r |? { $_.date -ge $datelimit })
                if ($r) {
                    log ([CopyGoProMessages]::MTPFileFound -f $r.Count)
                    $files += $r
                }
            }
        }
        if ($drives.count -gt 1) {
            log ([CopyGoProMessages]::MTPFileFoundTotal -f $files)
        } elseif (-not $files) {
            log ([CopyGoProMessages]::NoFileToCopy)
        }
        $this.goproFiles = $files
        return
    }

    [bool] copyToLocal($destDir) {
        foreach ($findex in 0..($this.goproFiles.count - 1)) {
            $file = $this.goproFiles[$findex]

            $yyyy = $file.yyyy
            $mm = $file.mm
            $dd = $file.dd
            $destdirpath = Join-Path $destDir $yyyy $mm
            if (-not (Test-Path $destdirpath)) {
                New-Item -ItemType Directory $destdirpath -Verbose:1
            }
            #--- here, parent folder exists (even blank) as $destdirpath
            #--- try to find subfolder yyyyMMdd or yyyyMMdd<some-remarks>
            $subfolder = Get-ChildItem $destdirpath -Directory -Filter "$yyyy$mm$dd*" |Select -Last 1
            if (-not $subfolder) {
                $subfolderName = "$yyyy$mm$dd"
                if ($script:Remark) {
                    $subfoldername += " $($script:Remark)"
                }
                $destdirpath = Join-Path $destdirpath $subfoldername
                New-Item -ItemType Directory $destdirpath -Verbose:1
            } else {
                $destdirpath = $subfolder.FullName
            }

            $destFiles = Get-ChildItem -File -Path $destdirpath
            $fileExists = $false
            $fp = $null
            foreach ($al in $file.altNames) {
                if (-not $destFiles) { <# no file in the folder #> break }
                if ($al -in $destFiles.Name) {
                    $fp = Join-Path $destdirpath $al
                    if ($al -eq $file.newName) {
                        log ([CopyGoProMessages]::FileExistsSkip -f $al)
                    } else {
                        if ($file.newName) {
                            log ([CopyGoProMessages]::FileExistsRename -f $al)
                            Rename-Item -Path $fp -NewName $file.newName -verbose:1
                        }
                    }
                    $fileExists = $true
                }
            }
            if (-not $fileExists) {
                #--- here it's new!
                if ($file.getFileSize() -ge ($diskfree = $this.getFree($destdirpath))) {
                    logv "file too large: $($file.getFileFize()) <-> diskfree: $diskfree"
                    log ([CopyGoProMessages]::NoDiskFree -f "$([math]::Round($diskfree/1GB, 2)) GB")
                    return $false
                }

                log ([CopyGoProMessages]::Copying -f $file.newName)
                $file.CopyTo($destdirpath)
                $fp = Join-Path $destdirpath $file.Name
                Rename-Item -Path $fp -NewName $file.newName -verbose:1
            }
            $copiedFile = Get-Item (Join-Path $destdirpath $file.newName)
            $newFileObj = [DirectAccessFile]::New($copiedFile)
            $newFileObj.altNames = $this.goproFiles[$findex].altNames
            $newFileObj.newName = $this.goproFiles[$findex].newName
            $this.goproFiles[$findex] = $newFileObj
        }
        return $true
    }

    [bool] OneDriveSignIn() {
        log ([CopyGoProMessages]::OneDriveSignIn)
        return $this.msal.SignIn($script:AppClientId, $script:AppRedirectUri, 'Files.ReadWrite.All', $script:TenantId)
    }

    copyToOneDrive($rootFolder, $cachedir) {
        # check if source file is not MTP
        $files = $this.goproFiles |?{ -not $_.hasDirectAccess }
        if ($files) {
            log ([CopyGoProMessages]::CannotUploadODFromMTP)
            return
        }

        $od = [OneDrive]::New($script:BufferSize)
        if (-not $od.SetOrCreateLocation($this.msal, $rootFolder, $false)) {
            log ([CopyGoProMessages]::NoDestinationFolder -f $rootFolder)
            return
        }

        foreach ($e in $this.goproFiles) {
            $yyyy = $e.yyyy
            $mm = $e.mm
            $dd = $e.dd
            log ([CopyGoProMessages]::CopyOD -f $e.Name, $e.GetFullPath())
            $destdirpath = Join-Path $rootFolder $yyyy $mm
            $parent = $od.SetOrCreateLocation($this.msal, $destdirpath, $true)

            <#--- Determine the full directory structure <root>/<yyyy>/<mm>/<yyyymmdd-remark> #>
            $res = $od.GetChildItem($this.msal, $parent)
            $dir = $res.value |? { $_.Name -match "$yyyy$mm$dd" } |Select -Last 1
            if ($dir) {
                logv "Subfolder found: $($dir.Name) for $yyyy/$mm/$dd"
                $dirName = $dir.Name
            } else {
                $dirName = "$yyyy$mm$dd"
                if ($script:Remark) {
                    $dirName += " $($script:Remark)"
                }
                log ([CopyGoProMessages]::SubFolderCreate -f $dirName,$parent.Name)
            }
            $dir = $od.SetOrCreateLocation($this.msal, $dirName, $parent, $true)
            $destdirpath = Join-Path $destdirpath $dirName

            if ($this.fileExists($od, $dir, $e)) { continue }

            <#--- Now try uploading!! #>
            if ($cacheDir) {
                $folderObj = (New-Object -ComObject Shell.Application).NameSpace($cachedir)
                $e.CopyTo($folderObj)
            }
            log ([CopyGoProMessages]::Uploading -f $e.name,$e.filesize,$destdirpath)
            $res = $od.UploadFile($this.msal, $destdirpath, $e.GetFullpath(), $e.filesize)

            # rename when upload OK _and_ newName is set (e.g., directly uploaded from GoPro)
            if ($res -and $e.newName) {
                $od.RenameItem($this.msal, $res, $e.newName)
            }
            if ($cachedir) {
                Remove-Item $e.localfp -Verbose:1
            }
        }
    }

    [bool] fileExists($od, $dir, $file) {
        <#--- Fiding file, perhaps exists as alternate name?? ---#>
        $files = ($od.GetChildItem($this.msal, $dir)).value
        if (-not $files) { return $false }

        $fileFound = $false
        foreach ($al in $file.altNames) {
            if ($al -in $files.Name) {
                if ($al -eq $file.newName) {
                    log ([CopyGoProMessages]::FileExistsSkip -f $al)
                } else {
                    log ([CopyGoProMessages]::FileExistsRename -f $al)
                    $fileObj = $files |? { $al -eq $_.Name } |Select -First 1
                    $od.RenameItem($this.msal, $fileObj, $file.newName)
                }
                $fileFound = $true
            }
        }
        return $fileFound
    }

    [long] getFree($path) {
        $txt = & cmd /c dir $path |Select -Last 1
        if ($txt -match '([\d,]+)\s*\w*$') {
            return [long]($matches[1])
        } else {
            throw "Cannot get free space: $path"
        }
    }
}
