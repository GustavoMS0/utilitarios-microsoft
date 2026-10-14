<#
.SYNOPSIS
  Cria o app registration no Entra ID para rodar Backup-Caixa.ps1 / Restaurar-Backup.ps1
  sem login interativo (ex: Agendador de Tarefas).

.DESCRIPTION
  - Gera um certificado autoassinado no repositório do Windows com chave privada NÃO
    exportável (a chave não sai da máquina; para outro servidor, rode o script de novo lá)
  - Cria o aplicativo com as permissões de aplicativo do Graph necessárias
  - Concede o consentimento de administrador (exige Administrador Global ou
    Administrador de Função Privilegiada); use -SemConsentimento para conceder depois no portal

  ATENÇÃO: permissões de aplicativo consentidas no Entra ID valem para TODAS as caixas do tenant.
  Quem tiver o certificado consegue ler (e, com -PermitirRestauracao, gravar) qualquer caixa.

  RECOMENDADO: restringir o app às caixas do backup com o RBAC para Aplicativos do Exchange Online
  (New-ManagementScope + New-ManagementRoleAssignment para o service principal do app). Com
  -GrupoCaixas, o script NÃO consente as permissões de caixa no Entra ID (só User.Read.All) e
  mostra os comandos do Exchange que concedem MailboxFolder/MailboxItem apenas aos membros
  diretos do grupo informado. Os comandos são só exibidos: rode-os conectado ao Exchange Online
  como Organization Management. As permissões do Entra ID e do RBAC se somam: se as de caixa
  forem consentidas no portal depois, a restrição deixa de valer.
  https://learn.microsoft.com/exchange/permissions-exo/application-rbac

.EXAMPLE
  .\Criar-AppBackup.ps1
  .\Criar-AppBackup.ps1 -PermitirRestauracao -Repositorio LocalMachine

.EXAMPLE
  # App restrito às caixas que são membros diretos do grupo "Backup-Caixas"
  .\Criar-AppBackup.ps1 -GrupoCaixas Backup-Caixas@empresa.com.br
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$NomeApp = 'Backup Exchange - utilitarios-microsoft',

    # Inclui as permissões de escrita usadas por Restaurar-Backup.ps1
    [switch]$PermitirRestauracao,

    # LocalMachine permite que a tarefa agendada rode com outra conta (exige PowerShell como administrador)
    [ValidateSet('CurrentUser', 'LocalMachine')]
    [string]$Repositorio = 'CurrentUser',

    [int]$ValidadeAnos = 2,

    [switch]$SemConsentimento,

    # Grupo (e-mail ou nome) cujos membros diretos o app poderá acessar, via RBAC para Aplicativos do
    # Exchange Online. As permissões de caixa não são consentidas no Entra ID; o script mostra os
    # comandos do Exchange para concedê-las só a esse grupo.
    [string]$GrupoCaixas
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\GraphMailbox.ps1')

$graphAppId = '00000003-0000-0000-c000-000000000000'
$permissoesCaixa = @('MailboxFolder.Read.All', 'MailboxItem.Read.All', 'MailboxItem.Export.All')
if ($PermitirRestauracao) { $permissoesCaixa += 'MailboxFolder.ReadWrite.All', 'MailboxItem.ImportExport.All' }
# Com -GrupoCaixas as permissões de caixa vêm do RBAC do Exchange (com escopo), não do Entra ID
$permissoes = @('User.Read.All') + $(if (-not $GrupoCaixas) { $permissoesCaixa })

$escopos = @('Application.ReadWrite.All')
if (-not $SemConsentimento) { $escopos += 'AppRoleAssignment.ReadWrite.All' }

if (-not $PSCmdlet.ShouldProcess("Entra ID / certificado em Cert:\$Repositorio\My", "Criar o aplicativo '$NomeApp' com as permissões $($permissoes -join ', ')")) { return }

Connect-GraphBackup -Escopos $escopos
$tenantId = (Get-MgContext).TenantId


try {
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
        -KeyExportPolicy NonExportable -KeySpec Signature -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
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
    # PASSO 5: RESTRIÇÃO ÀS CAIXAS DO GRUPO (RBAC PARA APLICATIVOS DO EXCHANGE)
    # ==============================================================================
    if ($GrupoCaixas) {
        $nomeSql  = $NomeApp -replace "'", "''"
        $escopo   = "$nomeSql - caixas"
        $comandos = @(
            "New-ServicePrincipal -AppId $($app.appId) -ObjectId $($sp.id) -DisplayName '$nomeSql'"
            "`$dn = (Get-Recipient -Identity '$($GrupoCaixas -replace "'", "''")').DistinguishedName"
            "New-ManagementScope -Name '$escopo' -RecipientRestrictionFilter `"MemberOfGroup -eq '`$dn'`""
        )
        foreach ($p in $permissoesCaixa) {
            $comandos += "New-ManagementRoleAssignment -App $($sp.id) -Role 'Application $($p -replace '\.All$')' -CustomResourceScope '$escopo'"
        }
        $comandos += "Test-ServicePrincipalAuthorization -Identity $($sp.id) -Resource <caixa@empresa.com.br>"

        Write-Host "`n--- RESTRINGIR O APP ÀS CAIXAS DE '$GrupoCaixas' ---" -ForegroundColor Yellow
        Write-Host "As permissões de caixa NÃO foram concedidas no Entra ID. Para concedê-las só às caixas do grupo,"
        Write-Host "rode no Exchange Online (Connect-ExchangeOnline, com Organization Management):"
        foreach ($c in $comandos) { Write-Host "  $c" }
        Write-Host "Só membros DIRETOS do grupo entram no escopo. A permissão pode levar até 2 horas para valer." -ForegroundColor DarkGray
        Write-Host "Não conceda as permissões de caixa no portal do Entra ID: elas se somam ao RBAC e anulam a restrição." -ForegroundColor DarkGray
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
} finally {
    Disconnect-GraphSeConectado
}
