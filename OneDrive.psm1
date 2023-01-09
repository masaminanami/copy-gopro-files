#
# OneDrive API
#
using module .\OneDrive.resources.ja-jp.psm1
using module .\MsalWrapper.psm1

class OneDrive {
    $baseUri;
    $bufferSize;
    $itemCache;
    $elapsed_format
    static $MaxRetryCount = 3

    OneDrive() { $this.init() }
    OneDrive($buffUnit) {
        $this.init()
        $this.bufferSize = 320 * 1024 * $buffUnit
        log ([OneDriveMessages]::BufferSizeChanged -f $buffUnit)
    }

    Init() {
        $this.baseUri = 'https://graph.microsoft.com/v1.0'
        $this.bufferSize = 320 * 1024 * 100 # do not change "320*1024" as it is SPO/ODB requirement
        $this.itemCache = @{}
        $this.elapsed_format = 'h\:mm\:ss\.fff'
    }

    <#
    # Upload file (large)
    #>
    [Object] UploadFile($msal, $folderpath, $file) {
        $fd = Get-Item -path $file
        return $this.UploadFile($msal, $folderpath, $file, $fd.Length)
    }

    [Object] UploadFile($msal, $folderpath, $file, $size) {
        $folderpath = $folderpath -replace '^\\','' -replace '\\$',''
        $basename = Split-Path -Leaf $file
        $uri = "/me/drive/root:/$folderpath/$($basename):/createUploadSession"
        $body = @{ "@microsoft.graph.conflictBehavior"="fail"; name=$basename; filesize=$size}
        $session = $this.apipost($msal, $uri, $body)
        logv "uploadFile: url=$($session.uploadUrl) nextExpectedRanges=$($session.nextExpectedRanges)"

        $this.bufferSize
        $fileStream = $null
        $binReader = $null
        $res = $null
        $totalStart = Get-Date
        $fileStream = [System.IO.FileStream]::New([System.IO.Path]::GetFullPath($file), [System.IO.FileMode]::Open)
        $binReader = [System.IO.BinaryReader]::New($fileStream)
        $pos = 0
        $retryCount = 0

        while ($buf = $binReader.ReadBytes($this.bufferSize)) {
            log ([OneDriveMessages]::ReadData -f (toByteCountString $buf.Length),(toByteCountString $pos))
            $hdr = @{}
            $hdr.Authorization = $msal.CreateHeader()
            $range = "bytes $pos-$($pos+$buf.Length-1)/$size"
            $hdr.'Content-Range' = $range;
            logv "OneDrive.Upload: Range=$range"

            $tStart = Get-Date
            try {
                $res = Invoke-WebRequest -Method Put -Uri $session.uploadUrl -Headers $hdr -Body $buf -SkipHeaderValidation
            } catch {
                log "ERROR! $_"
                if ($res) {
                    log "Result: $res"
                    $msg = $res.Content |ConvertFrom-Json -AsHashtable
                    ++$retryCount
                    if ($retryCount -le [OneDrive]::MaxRetryCount -and $msg.Contains('nextExpectedRanges')) {
                        $nextRange = $msg.nextExpectedRanges |Sort |Select -First 1
                        if ($nextRange -match '^(\d+)-\d+') {
                            $pos = [long]$matches[1]
                            [void]$fileStream.Seek($pos, [IO.SeekOrigin]::Begin)
                            log ([OneDriveMessages]::Retrying -f $pos,(toByteCountString $pos))
                            continue
                        }
                    }
                }
                #--- error is not recoverable
                if ($session) {
                    $res = Invoke-RestMethod -Method Delete -Uri $session.uploadUrl -Headers @{Authoriaztion=$msal.CreateHeader() }
                    log ([OneDriveMessages]::Aborted)
                }
                $res = $null
                break
            }

            #--- Chunk upload was successfull (Invoke-WebRequest Put)
            $tEnd = Get-Date
            $elapsed = $tEnd - $tStart
            $pos += $buf.Length
            $retryCount = 0

            log ([OneDriveMessages]::PutSuccess -f $elapsed.ToString($this.elapsed_format),(toByteCountString $buf.Length/$elapsed.TotalSeconds),(toByteCountString $size - $pos))
            logv "Response=$res"

            $msal.RefreshToken()
        }

        #--- uploading loop ends
        logv "Upload final response: $($res)"

        if ($res) {
            $totalEnd = Get-Date
            $elapsed = $totalEnd - $totalStart
            log ([OneDriveMessages]::Completed -f $elapsed.ToString($this.elapsed_format),(toByteCountString $size/$elapsed.TotalSeconds))
        }

        #--- close the streams anyway
        if ($binReader) {
            $binReader.Close()
            $binReader.Dispose()
        }
        if ($fileStream) {
            $fileStream.Dispose()
        }
        return $res
    }

    <#
     # Rename Item
     #>
    [object] RenameItem($msal, $file, $newName) {
        $uri = $this.baseUri + '/me/drive/items/' + $file.id
        $body = @{ name = $newName }
        logv "OneDrive.Rename: $uri Body=$($body |Convertto-json)"
        $res = Invoke-WebRequest -Method Patch -Uri $uri -Headers @{ Authorization=$msal.CreateHeader(); "Content-Type"="application/json" } -Body ($body |ConvertTo-Json -Depth 10 -Compress)

        logv $res
        return $res.StatusCode -eq 200
    }

    <#
     # GetChildItem
     #>
    [object] GetChildItem($msal, $parent) {
        logv "GetChildItem: $($parent.Name) type=$($parent.Gettype()) parent=$parent"
        return $this.apiget($msal, "/me/drive/items/$($parent.Id)/children")
    }

    <#
    #--- ChDir
    #>
    [Object] SetOrCreateLocation($msal, $path, [bool]$fCreate) {
        $root = $this.apiget($msal, '/me/drive/root')
        logv "ROOT=$root"
        return $this.SetOrCreateLocation($msal, $path, $root, $fCreate)
    }

    [Object] SetOrCreateLocation($msal, $path, $parent, [bool]$fCreate) {
        $path = $path -replace '^\\',''
        foreach ($subf in $path -split('\\')) {
            $children = $this.GetChildItem($msal, $parent)

            #log "DEBUG> $(($children.value|%{ $_.Name }) -join(', '))"

            $dir = $children.value |? { $_.Name -match "^$subf`$" } |select -First 1
            logv "Checking subfolder: $subf"
            if ($dir) {
                logv "Subfolder found: $($dir.Name)"
                $parent = $dir
            } elseif ($fCreate) {
                $res = $this.apipost($msal, "/me/drive/items/$($parent.Id)/children", @{ name=$subf; folder=@{}; '@microsoft.graph.conflictBehavior'="fail"; })
                if ($res) {
                    log ([OneDriveMessages]::FolderCreated -f $res.Name)
                    $parent = $res
                }
            } else {
                log ([OneDriveMessages]::FolderNotFound -f $subf)
                $parent = $null
            }
        }
        return $parent
    }

    <#
     # API Get
     #>
    [Object] apiget($msal, $uri) {
        if ($this.itemCache.Contains($uri)) { return $this.itemCache.$uri }
        logv "OneDrive.apiget: $uri"
        $res =  Invoke-RestMethod -Method Get -Uri ($this.baseUri + $uri) -Headers @{ Authorization=$msal.CreateHeader() }
        $this.itemCache.$uri = $res
        return $res
    }

    <#
     # API Post
     #>
    [Object] apipost($msal, $uri, $body) {
        $bstr = $body |ConvertTo-Json -Depth 10 -Compress
        logv "apipost: body=$bstr"
        return $this.apipost($msal, $uri, $bstr, $null)
    }

    [Object] apipost($msal, $uri, $bodyRaw, $optHeaders) {
        $rUrl = $uri -match '^http' ? $uri : $this.baseUri + $uri
        logv "OneDrive.apipost: $rUrl body.size=$($bodyRaw.Length)"
        $hdr = @{ Authorization = $msal.CreateHeader() }
        if ($optHeaders) {
            $optHeaders.Keys |%{ $hdr.$_ = $optHeaders.$_; log "apipost: Adding $k = $($optHeaders.$k)" }
        }
        $res = Invoke-RestMethod -Method Post -Uri $rUrl -Headers $hdr  -Body $bodyRaw -ContentType "application/json; charset=utf-8"
        return $res
    }
}
