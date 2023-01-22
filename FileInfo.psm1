#
# FileInfo
#

class FileInfo {
    [Object]$source;
    [string]$name;
    [DateTime]$date;
    [string]$yyyy;
    [string]$mm;
    [string]$dd;
    [string[]]$altNames;
    [string]$newName;

    [string]$sizeStr;
    [long]$filesize;
    [bool]$renameRequired;
    [bool]$hasDirectAccess;

    FileInfo($s, $n, $d) {
        $this.source = $s
        $this.name = $n
        $this.date = $d
        $this.set_dateinfo()
        $this.set_nameinfo()
    }

    set_dateinfo() {
        $this.yyyy = $this.date.ToString('yyyy')
        $this.mm = $this.date.ToString('MM')
        $this.dd = $this.date.ToString('dd')
    }

    set_nameinfo() {
        $this.altNames = @($this.name)
        $this.newName = $null
        if ($this.name -match 'GX(\d\d)(\d\d\d\d)(\..*)') {
            $datestr = $this.date.ToString('yyyyMMdd-HHmm')
            $basename = "GX$($matches[2])-$($matches[1])$($matches[3])"
            $this.altNames += $basename
            $this.newName = "$($datestr).$basename"
            $this.altNames += $this.newName
        } else {
            $this.newName = $this.name
        }
    }
}

class DirectAccessFile : FileInfo {
    DirectAccessFile($item) : base($item, $item.Name, $item.LastWriteTime) {
        $this.filesize = $item.Length
        $this.sizestr = toByteCountString $item.Length
        $this.renameRequired = $false
        $this.hasDirectAccess = $true
    }

    [bool] CopyTo($destdir) {
        $destfp = Join-Path $destdir $this.newName
        Copy-Item -LiteralPath $this.source.Fullname -Destination $destfp -verbose:1
        return $true
    }

    [string] GetFullPath() { return $this.source.Fullname }

    [long] GetFileSize() { return $this.filesize }
}
