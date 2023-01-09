#
# Msal Wrapper
#
Import-Module MSAL.PS

class MSALWrapper {
    $token;
    $_clientId;
    $_redirectUri;
    $_scope;
    $_tenantId;

    MSALWrapper() {
        $this.token = $null;
    }

    [bool] SignIn($clientId, $redirectUri, $scope, $tenantId) {
        if ($this.getToken($clientId, $redirectUri, $scope, $tenantId)) {
            $this._clientId = $clientId
            $this._redirectUri = $redirectUri
            $this._scope = $scope
            $this._tenantId = $tenantId
            $this.logToken()
            return $true
        } else {
            return $false
        }
    }
    
    RefreshToken() {
        [void]$this.getToken($this._clientId, $this._redirectUri, $this._scope, $this._tenantId)
        $this.logToken()
    }

    [bool] getToken($clientId, $redirectUri, $scope, $tenantId) {
        try {
            if ($tenantId) {
                $this.token = Get-MsalToken -ClientId $clientId -RedirectUri $redirectUri -Scopes $scope -TenantId $tenantId
            } else {
                $this.token = Get-MsalToken -ClientId $clientId -RedirectUri $redirectUri -Scopes $scope
            }
            $this._clientId = $clientId
            $this._redirectUri = $redirectUri
            $this._scope = $scope
            $this._tenantId = $tenantId
        } catch {
            logerror $_
            $this.token = $null
            return $false
        }
        return $true
    }

    [string] CreateHeader() { return $this.token.CreateAuthorizationHeader() }

    logToken() {
        logv "MSAL Token=$($this.token.AccessToken.substring(0,10))...$($this.token.AccessToken.substring($this.token.AccessToken.Length-10, 10))"
    }
}