$ErrorActionPreference = "Stop"
$proxyDirectory = $PSScriptRoot
$secretPath = Join-Path $proxyDirectory "hostinger-api-token.dpapi"

Write-Host "Cole o token NOVO da API Hostinger." -ForegroundColor Cyan
Write-Host "O valor nao sera exibido e sera criptografado para o usuario atual do Windows." -ForegroundColor DarkGray
$secureToken = Read-Host "Token" -AsSecureString

$credential = New-Object System.Management.Automation.PSCredential("hostinger", $secureToken)
$plainToken = $credential.GetNetworkCredential().Password
if ([string]::IsNullOrWhiteSpace($plainToken) -or $plainToken.Length -lt 20) {
    throw "Token vazio ou curto demais. Nada foi salvo."
}

$secureToken | ConvertFrom-SecureString | Set-Content -LiteralPath $secretPath -Encoding ASCII
$acl = Get-Acl -LiteralPath $secretPath
$acl.SetAccessRuleProtection($true, $false)
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    [System.Security.Principal.WindowsIdentity]::GetCurrent().Name,
    "Read,Write",
    "Allow"
)
$acl.SetAccessRule($rule)
Set-Acl -LiteralPath $secretPath -AclObject $acl

$plainToken = $null
Write-Host "Token salvo com protecao DPAPI em: $secretPath" -ForegroundColor Green
Write-Host "Agora execute C:\Volume\proxy\start_reverse_proxy.ps1 para recarregar o Caddy." -ForegroundColor Yellow
