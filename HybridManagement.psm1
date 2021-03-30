$ErrorActionPreference = 'Stop';

Function GetAuthToken
{
    param
        (
        [Parameter(Mandatory=$false)]
        $credentials
        )

    Import-Module Azure
    if (!$credentials) {
        $credential = Get-Credential
    }
    else {
     $credential = $credentials;
    }
    
    $userName = $credential.UserName;
    $tenantName = $userName.split('@')[1];

    if (!$tenantName) {
      Write-Error("Unable to determine tenant from user:" + $credential.UserName);
      return;
    }

    $clientId = "1950a258-227b-4e31-a9cf-717495945fc2" 
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $resourceAppIdURI = "https://graph.microsoft.com"
    $authority = "https://login.microsoftonline.com/$TenantName"
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority

    $AADCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential" -ArgumentList $credential.UserName,$credential.Password
    $authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId,$AADCredential)
    return $authResult
}

Function GetAuthHeader
{
    param
        (
        [Parameter(Mandatory=$true)]
        $token)

    #------Building Rest Api header with authorization token------#
    $authHeader = @{
        'Content-Type'='application\json'
        'Authorization'=$token.CreateAuthorizationHeader()
        }
    return $authHeader;
}

Function Get-HybridApplication
    {
    [CmdletBinding()]
    param
        (
        [Parameter(Mandatory=$true)]
        $appId,
        [Parameter(Mandatory=$true)]
        $credential
        )
    #------Get the authorization token------#
    $token = GetAuthToken -credential $credential;
  
    #------Build the Rest Api header with authorization token------#
    $authHeader = GetAuthHeader -Token $token;

    #------Initial URI Construction------#
    $uri = "https://graph.microsoft.com/edu/$tenant/applications/$appId/onPremisesPublishing"
    $applications = Invoke-RestMethod -Uri $uri –Headers $authHeader –Method Get
    $applications;
    return $foundApplications
}

Function Update-HybridApplication
{
    [CmdletBinding()]
    param
        (
        [Parameter(Mandatory=$true)]
        $appId,
        [Parameter(Mandatory=$true)]
        $targetUri,
        [Parameter(Mandatory=$true)]
        $credential
        )
    #------Get the authorization token------#
    $token = GetAuthToken -credential $credential;
  
    #------Build the Rest Api header with authorization token------#
    $authHeader = GetAuthHeader -Token $token;

    $application = @{
        "onPremisesPublishing" = @{
                                    "applicationServerTimeout" = "Default"
                                    "applicationType" = "microsoftapp"
                                    "externalAuthenticationType" = "passthru"
                                    "externalUrl" = "https://$appId.resource.mailboxmigration.his.msappproxy.net:443/"
                                    "internalUrl" = $targetUri
                                    "isOnPremPublishingEnabled" = $true
                                    "isTranslateHostHeaderEnabled" = $false
                                    "isTranslateLinksInBodyEnabled" = $false
                                    "singleSignOnSettings" = $null
                                    "verifiedCustomDomainCertificatesMetadata" = $null
                                    "verifiedCustomDomainKeyCredential" = $null
                                    "verifiedCustomDomainPasswordCredential" = $null
                                    }
                    }

    $uri = "https://graph.microsoft.com/edu/$tenant/applications/$appId"
    $json = $application | ConvertTo-Json
    Invoke-RestMethod -Uri $uri –Headers $authHeader –Method Patch -Body $json -ContentType 'application/json'

    #now display results
    Get-HybridApplication -appId $appId -credential $credential
}

Function New-HybridApplication
{
    [CmdletBinding()]
    param
        (
        [Parameter(Mandatory=$false)]
        $appId,
        [Parameter(Mandatory=$true)]
        $targetUri,
        [Parameter(Mandatory=$true)]
        $credential
        )

    if (!$appId) {
      $appId = New-Guid;
    }

    Update-HybridApplication -appId $appId -credential $credential -targetUri $targetUri;
}

Function Remove-HybridApplication
{
    [CmdletBinding()]
    param
        (
        [Parameter(Mandatory=$true)]
        $appId,
        [Parameter(Mandatory=$true)]
        $credential
        )
    #------Get the authorization token------#
    $token = GetAuthToken -credential $credential;
  
    #------Build the Rest Api header with authorization token------#
    $authHeader = GetAuthHeader -Token $token;

    $application = @{
        "onPremisesPublishing" = $null
                    }
    $json = $application | ConvertTo-Json

    $uri = "https://graph.microsoft.com/edu/$tenant/applications/$appId"
    Invoke-RestMethod -Uri $uri –Headers $authHeader –Method Patch -Body $json -ContentType 'application/json'
}

Function Get-HybridAgent
    {
    [CmdletBinding()]
    param
        (
        [Parameter(Mandatory=$true)]
        $credential
        )
    #------Get the authorization token------#
    $token = GetAuthToken -credential $credential;
    #------Build the Rest Api header with authorization token------#
    $authHeader = GetAuthHeader -Token $token;

    #------Initial URI Construction------#
    $uri = "https://graph.microsoft.com/edu/connectorGroups?`$expand=members"
    $result = Invoke-RestMethod -Uri $uri –Headers $authHeader –Method Get    
    return $result.value.members;
}

Function TestoneEndpoint
{
    param
        (
        [Parameter(Mandatory=$true)]
        $endpoint,
        $testGet)
    Write-Host("Testing connection to " + $endpoint.uri + " on port " + $endpoint.port)
    test-netconnection $endpoint.uri -Port $endpoint.port

    if ($testGet) {
        $uri = 'https://' + $endpoint.uri + ':' + $endpoint.port;
        Write-Host("Performing GET on " + $uri)
        $result = Invoke-WebRequest -Method Get -Uri $uri
        if ($result.StatusCode -ne 200)
        {
            Write-Host("Failed to get content with status " + $result.StatusCode);
        }
    }
}

Function TestEndpoints
{
    param
        (
        [Parameter(Mandatory=$true)]
        $endpoints,
        $testGet)

    foreach ($endpoint in $endpoints) {
      $result = TestoneEndpoint $endpoint $testGet
      if ($result.TcpTestSucceeded -ne $true) {
        Write-Host("Failed to connecto to " + $endpoint.uri + ":" + $endpoint.port); 
      }
    }
}

Function TestProxySettings
{
    $proxy = [System.Net.WebProxy]::GetDefaultProxy() | select address
    if ($proxy.Address -ne $null)
    {
        Write-Host("WebRequest is configured for proxy:" + $proxy.Address);
        Write-Host("Ensure connector is configured for proxy or whitelisted to bypass proxy.");
    }

    $browserProxies = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer
    if ($browserProxies -ne $null)
    {
        Write-Host("Browser is configured for proxy:" + $browserProxies.Address);
        Write-Host("Ensure connector is configured for proxy or whitelisted to bypass proxy.");
    }
}

Function TestTLSSettings
{
  $regKey = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2"
  $tlsKeyExists = Test-Path $regKey
  if ($tlsKeyExists -eq $false) {
    Write-Warning("Registry Key does not exist:" + $regKey);
    Write-Warning("TLS 1.2 is not explicitly enabled. Please enable it.");
    return;
  }

  $regKey = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client";
  $tlsClientKeyExists = Test-Path $regKey
  if ($tlsClientKeyExists -eq $false) {
    Write-Warning("Registry Key does not exist:" + $regKey);
    Write-Warning("TLS 1.2 is not explicitly enabled. Please enable it.");
    return;
  }
  
  $clientEnabled = (Get-ItemProperty -Path $regKey).Enabled
  $clientDisabledByDefault = (Get-ItemProperty -Path $regKey).DisabledByDefault
  if (($clientDisabledByDefault) -eq $true -or ($clientEnabled -eq $false)) {
    Write-Warning("TLS 1.2 is not explicitly enabled. Please enable it.");
    return;
  }

  $regKey = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server";
  $tlsServerKeyExists = Test-Path $regKey
  if ($tlsServerKeyExists -eq $false) {
    Write-Warning("TLS is either not enabled or disabled by default:" + $regKey);
    Write-Warning("TLS 1.2 is not explicitly enabled. Please enable it.");
    return;
  }
  
  $serverEnabled = (Get-ItemProperty -Path $regKey).Enabled
  $serverDisabledByDefault = (Get-ItemProperty -Path $regKey).DisabledByDefault
  if (($serverDisabledByDefault) -eq $true -or ($serverEnabled -eq $false)) {
    Write-Warning("TLS is either not enabled or disabled by default:" + $regKey);
    Write-Warning("TLS 1.2 is not explicitly enabled. Please enable it.");
    return;
  }
}

Function Test-HybridConnectivity
{
  [CmdletBinding()]
  param
        (
        [Parameter(Mandatory=$false)]
        [Switch]
        $testO365Endpoints)

    $httpEndpoints = @(
        [pscustomobject]@{uri = 'mscrl.microsoft.com'; port = '80' },
        [pscustomobject]@{uri = 'crl.microsoft.com'; port = '80' },
        [pscustomobject]@{uri = 'ocsp.msocsp.com'; port = '80' },
        [pscustomobject]@{uri = 'www.microsoft.com'; port= '80' })

    $httpsEndpoints = @(
        [pscustomobject]@{uri = 'login.windows.net'; port = '443' },
        [pscustomobject]@{uri = 'login.microsoftonline.com'; port= '443' })

    $o365Endpoints = @(
        [pscustomobject]@{uri = 'outlook.office.com'; port = '443' },
        [pscustomobject]@{uri = 'outlook.office365.com'; port = '443' },
        [pscustomobject]@{uri = 'nexus.microsoftonline-p.com'; port = '443' },
        [pscustomobject]@{uri = 'login.microsoftonline.com'; port= '443' })

    # CRL Endpoints
    TestEndpoints $httpEndpoints

    # Logon Endpoints
    TestEndpoints $httpsEndpoints
    
    if ($testO365Endpoints) {
      # O365 Endpoints
      TestEndpoints $o365Endpoints $true
    }
  
    # Test TLS Configuration
    TestTLSSettings

    # Test Proxy Settings
    TestProxySettings
}

