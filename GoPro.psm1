#
# GoPro MTP driver
#
using module .\GoPro.resources.ja-jp.psm1
using module .\FileInfo.psm1

class GoProDrive {
    $name;
    $dcimFolder;

    GoProDrive() {}
    GoProDrive($n, $dcim) {
        $this.name = $n
        $this.dcimFolder = $dcim
    }

    static [GoProDrive[]] Discover($devName, $devSpec, $drvSpec) {
        $devices = @()
        $sh = New-Object -ComObject Shell.Application
        $items = $sh.NameSpace(0x11).Items()
        if ($devName) {
            if ($r = [GoProDevice]::DiscoverByName($items, $devName)) {
                $devices += $r
            }
        } else {
            log ([GoProMessages]::Discovery -f "$($devSpec.DeviceType),$($drvSpec.DeviceType)")
            if ($r = [GoProDevice]::Discover($items, $devSpec)) {
                $devices += $r
            }
            if ($r = [GoProMemory]::Discover($items, $drvSpec)) {
                $devices += $r
            }
        }
        if (-not $devices) {
            throw ([GoProMessages]::DeviceNotFound -f ($devName ? " : $devName" : ""))
        }
        return $devices
    }
}

<#
 # Media Transfer Protocol class
 #>
class GoProDevice : GoProDrive {
    [int]$indexName;
    [int]$indexSize;
    [int]$indexDate;
    [bool]$isInitDone;

    GoProDevice() { $this.initVars() }
    GoProDevice($n, $dcim) : base($n, $dcim) { $this.initVars() }
    initVars() {
        $this.indexName = $this.indexSize = $this.indexDate = -1
        $this.isInitDone = $false;
    }

    static [GoProDevice[]] Discover($items, $spec) {
        log ([GoProMessages]::Discovery -f $spec.DeviceType)
        $devices = $items |? { $_.Type -match $spec.DeviceType }
        logv "Devices matching type: $($devices.Name -join(', '))"
        $devFound = @()

        foreach ($dev in $devices) {
            $disk = $null
            try {
                $disk = $dev.GetFolder.Items() |? { $_.Name -match $devSpec.GoProDiskNameKeyword } |Select -First 1
            } catch {
                log ([GoProMessages]::NotSupported -f $dev.Name)
                logv "Device found but failed to access subfolder $($dev.Name)"
            }
            if ($disk) {
                $dcim = $disk.GetFolder.Items() |? { $_.Name -eq "DCIM" }
                if ($dcim) {
                    log ([GoProMessages]::DeviceFound -f $dev.Name)
                    $devFound += [GoProDevice]::New($dev.Name, $dcim)

                } else {
                    log ([GoProMessages]::NotSupported -f $dev.Name)
                    logv "Device found but no DCIM folder: $($dev.name)"
                }
            } else {
                log ([GoProMessages]::NotSupported -f $dev.Name)
                logv "Device found but not GoPro or No MTP disk: $($dev.name)"
            }
        }
        return $devFound
    }

    [FileInfo[]] GetFiles() {
        logv "GoPro.GetFiles started"
        if (-not $this.isInitDone) { $this.Init() }
        $subdirs = $this.dcimFolder.GetFolder.Items() |? { $_.Name -match "...GOPRO" }
        logv "GoPro.GetFiles: folders=$($subdirs.Name -join(', '))"

        $files = @()
        foreach ($subd in $subdirs) {
            if ($r = $this.GetFiles($subd)) {
                $files += $r
            }
        }
        return $files
    }

    [FileInfo[]] GetFiles($dirObj) {
        $files = $dirObj.GetFolder.Items() |?{ $_.Name -match "\.(mp4|jpg)$" }
        $FileFound = @()
        foreach ($file in $files) {
            $name = $file.Name
            $date = Get-Date ($dirObj.GetFolder.GetDetailsOf($file, $this.indexDate))
            $fi = [GoProFile]::New($file, $name, $date)
            $fi.sizeStr = $dirObj.GetFolder.GetDetailsOf($file, $this.indexSize)
            $FileFound += $fi
        }
        return $fileFound
    }


    #--- initialize internal info
    # get index numbers for item's attribute
    Init() {
        foreach ($idx in 0..49) {
            $key = $this.dcimFolder.GetFolder.GetDetailsOf($null, $idx)
            switch -regex ($key) {
                {$_ -eq "更新日時" } {
                    if ($true <# $this.indexWriteTime -eq -1 #>) { $this.indexDate = $idx }
                    break;
                }
                {$_ -eq "名前" } {
                    if ($this.indexName -eq -1) { $this.indexName = $idx }
                    break;
                }
                {$_ -eq "サイズ" } {
                    if ($this.indexSize -eq -1) { $this.indexSize = $idx }
                    break;
                }
                default {
                    #logv "DEBUG> unmatched $idx|$key|"
                }
            }
        }
        if ($this.indexName -eq -1 -or $this.indexSize -eq -1 -or $this.indexDAte -eq -1) {
            throw  ([GoProMessages]::IndexNotFound -f "Name=$($this.indexName), Size=$($this.indexSize) WT=$($this.indexDate)")
        }
        $this.isInitDone = $true
    }


} # end of GoProDevice

class GoProFile : FileInfo {

    GoProFile() : base() {}
    GoProFile($s, $n, $d) :  base($s, $n, $d) {
        $this.filesize = -1
        $this.renameRequired = $true
        $this.hasDirectAccess = $false
    }

    [long] GetFileSize() {
        if ($this.sizeStr -match '([\d\.,]+)\s*MB') { return $matches[1] * 1MB }
        elseif ($this.sizeStr -match '([\d\.,]+)\s*GB') { return $matches[1] * 1GB }
        elseif ($this.sizeStr -match '[\d\.,]+') { return $matches }
        return 1
    }

    [bool] CopyTo($folder) {
        log ([GoProMessages]::Copying -f $this.Name)
        logv "copyto $($this.name) -> $folder"
        $folderObj = (New-Object -ComObject Shell.Application).NameSpace($folder)
        $folderObj.CopyHere($this.source)
        return $true
    }
}



<#
 #--- GOProMemory
 # handles files on USB memory directly attached to the PC
 #>
class GoProMemory : GoProDrive {
    GoProMemory() {}
    GoProMemory($name, $dcim) : base($name, $dcim) {}

    static [GoProMemory[]] Discover($items, $spec) {
        log ([GoProMessages]::Discovery -f $spec.DeviceType)
        $devices = $items |? { $_.Type -match $spec.DeviceType }
        logv "Devices matching type: $(($devices |%{ $_.Name }) -join(', '))"
        $devFound = @()

        foreach ($dev in $devices) {
            $dcim = Get-ChildItem $dev.Path -Filter "DCIM"
            if ($dcim) {
                $devFound += [GoProMemory]::New($dev.Name, $dcim)
                log ([GoProMessages]::DeviceFound -f $dev.Name)
            } else {
                log ([GoProMessages]::NotSupported -f $dev.Name)
                logv "Device found but no DCIM folder: $($dev.name)"
            }
        }
        return $devFound
    }

    [FileInfo[]] GetFiles() {
        $files = @()
        $subdirs = Get-ChildItem -Path $this.dcimFolder.Fullname -Filter "*GOPRO" -Directory
        foreach ($subd in $subdirs) {
            if ($r = $this.GetFiles($subd)) {
                $files += $r
            }
        }
        return $files
    }

    [FileInfo[]] GetFiles($dir) {
        $fileFound = @()
        $files = (Get-ChildItem -Path $dir.Fullname -Filter "*.mp4") +
            (Get-ChildItem -Path $dir.Fullname -Filter "*.jpg")
        foreach ($file in $files) {
            $fileFound += [DirectAccessFile]::New($file)
        }
        return $fileFound
    }
}
