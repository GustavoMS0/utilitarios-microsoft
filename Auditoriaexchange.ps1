<#
.SYNOPSIS
  Habilita/ajusta a auditoria de caixas de correio no Exchange Online.

.DESCRIPTION
  Restaura o conjunto padrão de ações auditadas pela Microsoft para acessos
  delegados e acrescenta "Move" (não incluído por padrão). Usar -AuditDelegate
  com lista fixa SUBSTITUI o padrão e desliga ações importantes como
  UpdateInboxRules (regras de encaminhamento), por isso aqui usamos @{Add=...}.

.EXAMPLE
  .\Auditoriaexchange.ps1 -Caixas financeiro@empresa.com.br, boletos@empresa.com.br
#>
param(
    [Parameter(Mandatory = $true, HelpMessage = "E-mails das caixas a configurar (ex: financeiro@empresa.com.br)")]
    [string[]]$Caixas,

    # Ações acrescentadas ao padrão da Microsoft
    [string[]]$AcoesDelegateExtras = @('Move'),
    [string[]]$AcoesOwnerExtras    = @('Move')
)

# ==============================================================================
# PASSO 0: CONEXÃO
# ==============================================================================
if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
    Write-Host "Instalando módulo: ExchangeOnlineManagement ..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -Repository PSGallery
}
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected')) {
    Connect-ExchangeOnline -ShowBanner:$false
}

# Se a auditoria estiver desligada na organização, nada do que for feito por caixa vale
if ((Get-OrganizationConfig).AuditDisabled) {
    Write-Warning "A auditoria de caixas está DESATIVADA na organização (AuditDisabled = True)."
    Write-Warning "Para ativar: Set-OrganizationConfig -AuditDisabled `$false"
}


# ==============================================================================
# PASSO 1: CONFIGURAÇÃO POR CAIXA
# ==============================================================================
foreach ($caixa in $Caixas) {
    Write-Host "`nConfigurando auditoria para: $caixa" -ForegroundColor Cyan

    try {
        $null = Get-Mailbox -Identity $caixa -ErrorAction Stop

        # Restaura o padrão da Microsoft (desfaz customizações antigas que substituíram a lista)
        Set-Mailbox -Identity $caixa -AuditEnabled $true -DefaultAuditSet Delegate, Owner -ErrorAction Stop

        # Acrescenta ações extras sem remover as do padrão
        Set-Mailbox -Identity $caixa -AuditDelegate @{ Add = $AcoesDelegateExtras } -AuditOwner @{ Add = $AcoesOwnerExtras } -ErrorAction Stop

        # Contas com bypass não geram log nenhum nesta caixa
        $bypass = Get-MailboxAuditBypassAssociation -Identity $caixa -ErrorAction SilentlyContinue
        if ($bypass.AuditBypassEnabled) {
            Write-Warning "A caixa $caixa está com AuditBypassEnabled = True. Acessos a ela NÃO são auditados."
        }

        $mbx = Get-Mailbox -Identity $caixa -ErrorAction Stop
        Write-Host "  AuditEnabled:  $($mbx.AuditEnabled)"
        Write-Host "  AuditDelegate: $($mbx.AuditDelegate -join ', ')"
        Write-Host "  AuditOwner:    $($mbx.AuditOwner -join ', ')"
        Write-Host "Auditoria configurada com sucesso para $caixa" -ForegroundColor Green
    } catch {
        Write-Warning "Falha ao configurar ${caixa}: $($_.Exception.Message)"
    }
}
