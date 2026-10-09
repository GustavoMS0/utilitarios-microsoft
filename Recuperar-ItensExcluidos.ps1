<#
.SYNOPSIS
  Recuperação granular de itens excluídos no Exchange Online (sem precisar de backup).

.DESCRIPTION
  Lista os itens da pasta Itens Excluídos e de Itens Recuperáveis com
  Get-RecoverableItems, deixa escolher quais restaurar e os devolve com
  Restore-RecoverableItems. Por padrão cada item volta para a pasta onde estava;
  com -PastaDestino eles vão para uma pasta específica.

  Só alcança o que ainda está dentro do prazo de retenção da caixa
  (RetainDeletedItemsFor: 14 dias por padrão, máximo 30) ou preservado por Hold
  ou política de retenção.

  Permissão: função "Mailbox Import Export". Por padrão ela não é atribuída a ninguém:
    New-ManagementRoleAssignment -Role "Mailbox Import Export" -User admin@empresa.com.br
  (pode levar alguns minutos para valer; reconecte ao Exchange depois)

.EXAMPLE
  # Lista e escolhe o que restaurar
  .\Recuperar-ItensExcluidos.ps1 -Caixa financeiro@empresa.com.br -Assunto boleto

.EXAMPLE
  # E-mails apagados nos últimos 3 dias, todos de volta numa pasta separada
  .\Recuperar-ItensExcluidos.ps1 -Caixa financeiro@empresa.com.br -Tipo IPM.Note -De (Get-Date).AddDays(-3) -Todos -PastaDestino "Recuperados"
#>
param(
    [Parameter(Mandatory = $true, HelpMessage = "E-mail da caixa (ex: financeiro@empresa.com.br)")]
    [ValidateNotNullOrEmpty()]
    [string]$Caixa,

    # Trecho do assunto
    [string]$Assunto,

    # Classe do item: IPM.Note (e-mail), IPM.Appointment, IPM.Contact, IPM.Task...
    [string]$Tipo,

    # Período pela data da última modificação (que, para itens apagados, é a data da exclusão)
    [datetime]$De,
    [datetime]$Ate,

    [ValidateSet('DeletedItems', 'RecoverableItems', 'PurgedItems', 'DiscoveryHoldsItems')]
    [string]$Origem,

    # Pasta (a partir da raiz da caixa) para onde restaurar; padrão: pasta original de cada item
    [string]$PastaDestino,

    [int]$Max = 1000,

    # Restaura todos os itens encontrados sem perguntar quais
    [switch]$Todos,

    # Seleção pela janela interativa (Out-GridView) em vez da lista no console
    [switch]$GridView,

    # Não pede confirmação
    [switch]$Force,

    [string]$Relatorio = "C:\Temp\Recuperacao-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

# ==============================================================================
# PASSO 0: CONEXÃO E PERMISSÃO
# ==============================================================================
if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
    Write-Host "Instalando módulo: ExchangeOnlineManagement ..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -Repository PSGallery
}
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected')) {
    Connect-ExchangeOnline -ShowBanner:$false
}

# Sem a função Mailbox Import Export os cmdlets nem aparecem na sessão
if (-not (Get-Command Get-RecoverableItems -ErrorAction SilentlyContinue)) {
    Write-Warning "Get-RecoverableItems não está disponível: sua conta não tem a função 'Mailbox Import Export'."
    Write-Warning "Peça a um administrador do Exchange para rodar:"
    Write-Warning "  New-ManagementRoleAssignment -Role 'Mailbox Import Export' -User <sua conta>"
    Write-Warning "Depois desconecte (Disconnect-ExchangeOnline) e rode este script de novo."
    exit 1
}


# ==============================================================================
# PASSO 1: BUSCA
# ==============================================================================
$filtros = @{ Identity = $Caixa; ResultSize = $Max }
if ($Assunto) { $filtros.SubjectContains = $Assunto }
if ($Tipo)    { $filtros.FilterItemType  = $Tipo }
if ($Origem)  { $filtros.SourceFolder    = $Origem }
if ($PSBoundParameters.ContainsKey('De'))  { $filtros.FilterStartTime = $De }
if ($PSBoundParameters.ContainsKey('Ate')) { $filtros.FilterEndTime   = $Ate.Date.AddDays(1).AddTicks(-1) }

Write-Host "Buscando itens excluídos em $Caixa..." -ForegroundColor Cyan
$itens = @(Get-RecoverableItems @filtros | Sort-Object LastModifiedTime -Descending)

if ($itens.Count -eq 0) {
    Write-Host "Nenhum item excluído encontrado com esses filtros." -ForegroundColor Yellow
    return
}
if ($itens.Count -ge $Max) { Write-Warning "Atingido o limite de $Max itens. Refine os filtros ou aumente -Max." }

$origens = @{
    DeletedItems        = 'Itens Excluídos'
    RecoverableItems    = 'Recuperáveis'
    PurgedItems         = 'Expurgados'
    DiscoveryHoldsItems = 'Retidos (Hold)'
}


# ==============================================================================
# PASSO 2: SELEÇÃO
# ==============================================================================
if ($Todos) {
    $selecionados = $itens
} elseif ($GridView -and (Get-Command Out-GridView -ErrorAction SilentlyContinue)) {
    $selecionados = @($itens | Out-GridView -Title "Selecione os itens para restaurar (Ctrl+clique para vários)" -PassThru)
} else {
    for ($i = 0; $i -lt $itens.Count; $i++) {
        $it = $itens[$i]
        $local = if ($origens["$($it.SourceFolder)"]) { $origens["$($it.SourceFolder)"] } else { "$($it.SourceFolder)" }
        Write-Host ("  [{0,3}] {1:dd/MM/yyyy HH:mm}  {2,-16} {3,-15} {4}" -f $i, $it.LastModifiedTime, $it.ItemClass, $local, $it.Subject)
        if ($it.LastParentPath) { Write-Host ("        pasta original: {0}" -f $it.LastParentPath) -ForegroundColor DarkGray }
    }
    $resp = Read-Host "`nNúmeros (ex: 0,3,5-8), [T] para todos ou [Enter] para cancelar"
    $selecionados = @()
    if ($resp.Trim().ToUpper() -eq 'T') {
        $selecionados = $itens
    } elseif ($resp) {
        $indices = foreach ($parte in $resp -split ',') {
            $parte = $parte.Trim()
            if ($parte -match '^(\d+)-(\d+)$') { [int]$Matches[1]..[int]$Matches[2] }
            elseif ($parte -match '^\d+$')   { [int]$parte }
        }
        $selecionados = @($indices | Where-Object { $_ -lt $itens.Count } | Select-Object -Unique | ForEach-Object { $itens[$_] })
    }
}

if ($selecionados.Count -eq 0) { Write-Host "Nada selecionado." -ForegroundColor Yellow; return }

$alvo = if ($PastaDestino) { "pasta '$PastaDestino'" } else { 'pasta original de cada item' }
Write-Host "`nRestaurar $($selecionados.Count) item(ns) de $Caixa para a $alvo." -ForegroundColor Yellow
if (-not $Force -and (Read-Host "Confirmar? (S/N)").ToUpper() -ne 'S') { return }


# ==============================================================================
# PASSO 3: RESTAURAÇÃO (um a um, pelo EntryID, para não trazer itens a mais)
# ==============================================================================
$resultado = New-Object System.Collections.Generic.List[object]
$ok = 0
foreach ($it in $selecionados) {
    $status = 'OK'; $erro = ''
    try {
        $p = @{ Identity = $Caixa; EntryID = $it.EntryID; ErrorAction = 'Stop' }
        if ($PastaDestino) { $p.RestoreTargetFolder = $PastaDestino }
        $null = Restore-RecoverableItems @p
        $ok++
    } catch {
        $status = 'Falha'; $erro = $_.Exception.Message
        Write-Warning "Falha ao restaurar '$($it.Subject)': $erro"
    }
    $resultado.Add([pscustomobject]@{
        ExcluidoEm    = $it.LastModifiedTime
        Tipo          = $it.ItemClass
        Assunto       = $it.Subject
        Origem        = $it.SourceFolder
        PastaOriginal = $it.LastParentPath
        Status        = $status
        Erro          = $erro
    })
}

New-Item -ItemType Directory (Split-Path $Relatorio) -Force | Out-Null
$resultado | Export-Csv $Relatorio -NoTypeInformation -Encoding UTF8 -UseCulture

$cor = if ($ok -eq $selecionados.Count) { 'Green' } else { 'Yellow' }
Write-Host "`n$ok de $($selecionados.Count) item(ns) restaurados." -ForegroundColor $cor
Write-Host "Relatório: $Relatorio"
