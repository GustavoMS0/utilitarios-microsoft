<#
.SYNOPSIS
  Cria o app registration no Entra ID para rodar Backup-Caixa.ps1 / Restaurar-Backup.ps1
  sem login interativo (ex: Agendador de Tarefas).

.DESCRIPTION
  - Gera um certificado autoassinado no repositório do Windows (a chave privada não sai da máquina)
  - Cria o aplicativo com as permissões de aplicativo do Graph necessárias
  - Concede o consentimento de administrador (exige Administrador Global ou
    Administrador de Função Privilegiada); use -SemConsentimento para conceder depois no portal

  ATENÇÃO: permissões de aplicativo valem para TODAS as caixas do tenant.
  Quem tiver o certificado consegue ler (e, com -PermitirRestauracao, gravar) qualquer caixa.

.EXAMPLE
  .\Criar-AppBackup.ps1
  .\Criar-AppBackup.ps1 -PermitirRestauracao -Repositorio LocalMachine
#>
param(
    [string]$NomeApp = 'Backup Exchange - utilitarios-microsoft',

    # Inclui as permissões de escrita usadas por Restaurar-Backup.ps1
    [switch]$PermitirRestauracao,

    # LocalMachine permite que a tarefa agendada rode com outra conta (exige PowerShell como administrador)
    [ValidateSet('CurrentUser', 'LocalMachine')]
    [string]$Repositorio = 'CurrentUser',

    [int]$ValidadeAnos = 2,

    [switch]$SemConsentimento
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\GraphMailbox.ps1')

$graphAppId = '00000003-0000-0000-c000-000000000000'
$permissoes = @('User.Read.All', 'MailboxFolder.Read.All', 'MailboxItem.Read.All', 'MailboxItem.Export.All')
if ($PermitirRestauracao) { $permissoes += 'MailboxFolder.ReadWrite.All', 'MailboxItem.ImportExport.All' }

$escopos = @('Application.ReadWrite.All')
if (-not $SemConsentimento) { $escopos += 'AppRoleAssignment.ReadWrite.All' }
Connect-GraphBackup -Escopos $escopos
$tenantId = (Get-MgContext).TenantId


# ==============================================================================
# PASSO 1: IDs DAS PERMISSÕES NO GRAPH
# ==============================================================================
$graphSp = Invoke-GraphRetry -Uri "v1.0/servicePrincipals(appId='$graphAppId')?`$select=id,appRoles"
$roles = foreach ($p in $permissoes) {
    $r = $graphSp.appRoles | Where-Object { $_.value -eq $p -and $_.allowedMemberTypes -contains 'Application' } | Select-Object -First 1
    if (-not $r) { throw "Permissão de aplicativo '$p' não encontrada no Microsoft Graph deste tenant." }
    [pscustomobject]@{ Nome = $p; Id = $r.id }
}


# ==============================================================================
# PASSO 2: CERTIFICADO
# ==============================================================================
Write-Host "`nGerando certificado em Cert:\$Repositorio\My ..." -ForegroundColor Cyan
$cert = New-SelfSignedCertificate -Subject "CN=$NomeApp" -CertStoreLocation "Cert:\$Repositorio\My" `
    -KeyExportPolicy Exportable -KeySpec Signature -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears($ValidadeAnos)
Write-Host "Thumbprint: $($cert.Thumbprint) (válido até $($cert.NotAfter.ToString('dd/MM/yyyy')))" -ForegroundColor Green


# ==============================================================================
# PASSO 3: APLICATIVO E SERVICE PRINCIPAL
# ==============================================================================
Write-Host "`nCriando aplicativo '$NomeApp'..." -ForegroundColor Cyan
$app = Invoke-GraphRetry -Method POST -Uri 'v1.0/applications' -Body @{
    displayName            = $NomeApp
    signInAudience         = 'AzureADMyOrg'
    requiredResourceAccess = @(@{
        resourceAppId  = $graphAppId
        resourceAccess = @($roles | ForEach-Object { @{ id = $_.Id; type = 'Role' } })
    })
    keyCredentials         = @(@{
        type        = 'AsymmetricX509Cert'
        usage       = 'Verify'
        key         = [Convert]::ToBase64String($cert.RawData)
        displayName = "CN=$NomeApp"
    })
}

# O aplicativo recém-criado pode levar alguns segundos para ser reconhecido
$sp = $null
for ($t = 1; -not $sp; $t++) {
    try {
        $sp = Invoke-GraphRetry -Method POST -Uri 'v1.0/servicePrincipals' -Body @{ appId = $app.appId }
    } catch {
        if ($t -ge 6) { throw }
        Start-Sleep -Seconds 5
    }
}
Write-Host "Aplicativo criado. ClientId: $($app.appId)" -ForegroundColor Green


# ==============================================================================
# PASSO 4: CONSENTIMENTO DO ADMINISTRADOR
# ==============================================================================
if ($SemConsentimento) {
    Write-Warning "Consentimento NÃO concedido. Conceda em:"
    Write-Warning "  https://entra.microsoft.com > Aplicativos > Registros de aplicativo > $NomeApp > Permissões de API > Conceder consentimento"
} else {
    foreach ($r in $roles) {
        Invoke-GraphRetry -Method POST -Uri "v1.0/servicePrincipals/$($sp.id)/appRoleAssignments" -Body @{
            principalId = $sp.id
            resourceId  = $graphSp.id
            appRoleId   = $r.Id
        } | Out-Null
        Write-Host "  ✔ $($r.Nome)" -ForegroundColor Green
    }
}


# ==============================================================================
# RESUMO
# ==============================================================================
$exe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
Write-Host "`n--- DADOS PARA OS SCRIPTS ---" -ForegroundColor Yellow
Write-Host "TenantId:              $tenantId"
Write-Host "ClientId:              $($app.appId)"
Write-Host "CertificateThumbprint: $($cert.Thumbprint)"
Write-Host "`nTeste:" -ForegroundColor Yellow
Write-Host "  .\Backup-Caixa.ps1 -Caixas <caixa@empresa.com.br> -TenantId $tenantId -ClientId $($app.appId) -CertificateThumbprint $($cert.Thumbprint)"
Write-Host "`nAgendar backup diário às 22h:" -ForegroundColor Yellow
Write-Host "  `$acao = New-ScheduledTaskAction -Execute '$exe' -Argument '-NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $PSScriptRoot 'Backup-Caixa.ps1')`" -Caixas <caixa@empresa.com.br> -TenantId $tenantId -ClientId $($app.appId) -CertificateThumbprint $($cert.Thumbprint)'"
Write-Host "  Register-ScheduledTask -TaskName 'Backup Exchange' -Action `$acao -Trigger (New-ScheduledTaskTrigger -Daily -At 22:00)"
Write-Host "`nA permissão pode levar alguns minutos para começar a valer." -ForegroundColor DarkGray
