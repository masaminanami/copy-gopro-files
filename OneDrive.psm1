#
# OneDrive API
#
using module .\OneDrive.resources.ja-jp.psm1
using module .\MsalWrapper.psm1


class OneDriveBase {
    $msal;
    $baseUri;
    $bufferSize;
    $itemCache;
    $elapsed_format
    static $MaxRetryCount = 3

    OneDriveBase() { $this.init() }
    OneDriveBase($buffUnit) {
        $this.init()
        $this.bufferSize = 320 * 1024 * $buffUnit
        log ([OneDriveMessages]::BufferSizeChanged -f $buffUnit)
    }

    Init() {
        $this.bufferSize = 320 * 1024 * 100 # do not change "320*1024" as it is SPO/ODB requirement
        $this.elapsed_format = 'h\:mm\:ss\.fff'
        $this.msal = [MSALWrapper]::New()
        $this.ClearCache()
    }

    ClearCache() { $this.itemCache = @{} }

    [bool] SignIn($clientId, $redirectUri, $scope, $tenantId) {
        return $this.msal.SignIn($clientId, $redirectUri, $scope, $tenantId)
    }

    <#
    # Upload file (large)
    #>
    [Object] getUploadSession($folderpath, $file, $size) { throw "ERROR! subclass must implement getUploadUrl()"}

    [Object] UploadFile($folderpath, $file) {
        $fd = Get-Item -path $file
        return $this.UploadFile($folderpath, $file, $fd.Length)
    }

    [Object] UploadFile($folderpath, $file, $size) {
        $folderpath = $folderpath -replace '^\\','' -replace '\\$',''
        $session = $this.getUploadSession($folderpath, $file, $size)

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
            $hdr.Authorization = $this.msal.CreateHeader()
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
                    $res = Invoke-RestMethod -Method Delete -Uri $session.uploadUrl -Headers @{Authoriaztion=$this.msal.CreateHeader() }
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

            log ([OneDriveMessages]::PutSuccess -f $elapsed.ToString($this.elapsed_format),(toByteCountString ($buf.Length/$elapsed.TotalSeconds)),(toByteCountString ($size - $pos)))
            logv "Response=$res"

            $this.msal.RefreshToken()
        }

        #--- uploading loop ends
        logv "Upload final response: $($res)"

        if ($res) {
            $totalEnd = Get-Date
            $elapsed = $totalEnd - $totalStart
            log ([OneDriveMessages]::Completed -f $elapsed.ToString($this.elapsed_format),(toByteCountString ($size/$elapsed.TotalSeconds)))
            $res = $res |ConvertFrom-Json -AsHashtable
            $this.ClearCache()
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
    [object] RenameItem($file, $newName) {
        $uri = $this.baseUri + '/drive/items/' + $file.id
        $body = @{ name = $newName }
        logv "OneDrive.Rename: $uri Body=$($body |Convertto-json)"
        $res = Invoke-WebRequest -Method Patch -Uri $uri -Headers @{ Authorization=$this.msal.CreateHeader(); "Content-Type"="application/json" } -Body ($body |ConvertTo-Json -Depth 10 -Compress)

        logv $res
        $this.ClearCache()
        return $res.StatusCode -eq 200
    }

    <#
     # GetChildItem
     #>
    [object] GetChildItem($parent) {
        logv "GetChildItem: $($parent.Name) type=$($parent.Gettype()) parent=$parent"
        return $this.apiget("/drive/items/$($parent.Id)/children")
    }

    <#
    #--- ChDir
    #>
    [Object] SetOrCreateLocation($path, [bool]$fCreate) {
        $root = $this.apiget('/drive/root')
        logv "ROOT=$root"
        return $this.SetOrCreateLocation($path, $root, $fCreate)
    }

    [Object] SetOrCreateLocation($path, $parent, [bool]$fCreate) {
        $path = $path -replace '^\\',''
        foreach ($subf in $path -split('\\')) {
            $children = $this.GetChildItem($parent)

            #log "DEBUG> $(($children.value|%{ $_.Name }) -join(', '))"

            $dir = $children.value |? { $_.Name -match "^$subf`$" } |select -First 1
            logv "Checking subfolder: $subf"
            if ($dir) {
                logv "Subfolder found: $($dir.Name)"
                $parent = $dir
            } elseif ($fCreate) {
                $res = $this.apipost("/drive/items/$($parent.Id)/children", @{ name=$subf; folder=@{}; '@microsoft.graph.conflictBehavior'="fail"; })
                if ($res) {
                    $this.ClearCache()
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
    [Object] apiget($uri) {
        if ($this.itemCache.Contains($uri)) { return $this.itemCache.$uri }
        logv "OneDrive.apiget: $uri"
        $res =  Invoke-RestMethod -Method Get -Uri ($this.baseUri + $uri) -Headers @{ Authorization=$this.msal.CreateHeader() }
        $this.itemCache.$uri = $res
        return $res
    }

    <#
     # API Post
     #>
    [Object] apipost($uri, $body) {
        $bstr = $body |ConvertTo-Json -Depth 10 -Compress
        logv "apipost: body=$bstr"
        return $this.apipost($uri, $bstr, $null)
    }

    [Object] apipost($uri, $bodyRaw, $optHeaders) {
        $rUrl = $uri -match '^http' ? $uri : $this.baseUri + $uri
        logv "OneDrive.apipost: $rUrl body=$($bodyRaw)"
        $hdr = @{ Authorization = $this.msal.CreateHeader() }
        if ($optHeaders) {
            $optHeaders.Keys |%{ $hdr.$_ = $optHeaders.$_; logv "apipost: Adding $k = $($optHeaders.$k)" }
        }
        $res = Invoke-RestMethod -Method Post -Uri $rUrl -Headers $hdr  -Body $bodyRaw -ContentType "application/json; charset=utf-8"
        return $res
    }
}

<# for Personal #>
class OneDrive : OneDriveBase {
    OneDrive() : Base() {}
    OneDrive($buffUnit) : Base($buffUnit) {
        $this.baseUri = 'https://api.onedrive.com/v1.0'
    }

    [Object] getUploadSession($folderpath, $file, $size) {
        $folderObj = $this.SetOrCreateLocation($folderpath, $false)
        $basename = Split-Path -Leaf $file
        $uri = "/drive/items/$($folderObj.id):/$($basename):/createUploadSession"
        $body = @{ "@microsoft.graph.conflictBehavior"="fail" }
        $session = $this.apipost($uri, $body)
        logv "uploadFile: url=$($session.uploadUrl) nextExpectedRanges=$($session.nextExpectedRanges)"
        return $session
    }
}

<# for Business #>
class OneDriveBusiness : OneDriveBase {
    OneDriveBusiness() : Base() {}
    OneDriveBusiness($buffUnit) : Base($buffUnit) {
        $this.baseUri = 'https://graph.microsoft.com/v1.0/me'
    }

    [Object] getUploadSession($folderpath, $file, $size) {
        $basename = Split-Path -Leaf $file

        $uri = "/drive/root:/$folderpath/$($basename):/createUploadSession"
        $body = @{ "@microsoft.graph.conflictBehavior"="fail"; name=$basename; filesize=$size}
        $session = $this.apipost($uri, $body)
        logv "uploadFile: url=$($session.uploadUrl) nextExpectedRanges=$($session.nextExpectedRanges)"
        return $session
    }

}
