<#
.SYNOPSIS
  Restauração granular a partir do backup gerado por Backup-Caixa.ps1.

.DESCRIPTION
  Filtra o índice do backup (assunto, remetente, pasta, período, só apagados),
  deixa escolher os itens e os reimporta com fidelidade total numa pasta nova
  ("Restaurados <data>"), recriando a estrutura de pastas original dentro dela.
  Nada existente na caixa é alterado ou sobrescrito.

.EXAMPLE
  # E-mails apagados com "boleto" no assunto em setembro
  .\Restaurar-Backup.ps1 -Caixa financeiro@empresa.com.br -Assunto boleto -De 2026-09-01 -Ate 2026-09-30 -SomenteRemovidos

.EXAMPLE
  # Restaurar para outra caixa (ex: caixa do gestor) tudo de um remetente, sem perguntar
  .\Restaurar-Backup.ps1 -Caixa ex.funcionario@empresa.com.br -CaixaDestino gestor@empresa.com.br -Remetente fornecedor.com -Todos -Force
#>
param(
    [Parameter(Mandatory = $true, HelpMessage = "Caixa de origem do backup (ex: financeiro@empresa.com.br)")]
    [string]$Caixa,

    # Caixa que vai receber os itens (padrão: a própria caixa de origem)
    [string]$CaixaDestino,

    [string]$Destino = 'C:\Backup\Exchange',

    # Filtros (combinados com E)
    [string]$Assunto,
    [string]$Remetente,
    [string]$Pasta,          # aceita curinga: 'Inbox*'
    [datetime]$De,
    [datetime]$Ate,
    [switch]$SomenteRemovidos,

    [string]$PastaRestauracao = "Restaurados $(Get-Date -Format 'yyyy-MM-dd HHmm')",

    # Coloca tudo direto na pasta de restauração, sem recriar as subpastas originais
    [switch]$SemEstrutura,

    # Restaura todos os itens filtrados sem perguntar quais
    [switch]$Todos,

    # Seleção pela janela interativa (Out-GridView) em vez da lista no console
    [switch]$GridView,

    # Não pede confirmação
    [switch]$Force,

    # Modo aplicativo (opcional)
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\GraphMailbox.ps1')

if (-not $CaixaDestino) { $CaixaDestino = $Caixa }
$dirCaixa  = Join-Path $Destino $Caixa.ToLower()
$arqIndice = Join-Path $dirCaixa 'indice.csv'
if (-not (Test-Path $arqIndice)) { throw "Backup não encontrado: $arqIndice" }


# ==============================================================================
# PASSO 1: FILTRAR O ÍNDICE
# ==============================================================================
function ConvertTo-Data([string]$Texto) {
    if (-not $Texto) { return $null }
    [datetime]::ParseExact($Texto, 'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
}

# Parâmetro [datetime] não informado vale DateTime.MinValue, por isso checamos se foi passado
$temDe  = $PSBoundParameters.ContainsKey('De')
$temAte = $PSBoundParameters.ContainsKey('Ate')
$fimAte = if ($temAte) { $Ate.Date.AddDays(1).AddTicks(-1) }

$itens = @((Import-Indice $arqIndice).Values | Where-Object {
    $d = ConvertTo-Data $_.Data
    (-not $Assunto   -or $_.Assunto -like "*$Assunto*") -and
    (-not $Remetente -or $_.Remetente -like "*$Remetente*" -or $_.Email -like "*$Remetente*") -and
    (-not $Pasta     -or $_.Pasta -like $Pasta) -and
    (-not $temDe     -or ($d -and $d -ge $De)) -and
    (-not $temAte    -or ($d -and $d -le $fimAte)) -and
    (-not $SomenteRemovidos -or $_.RemovidoEm) -and
    (Test-Path (Join-Path $dirCaixa $_.Arquivo))
} | Sort-Object Data -Descending)

if ($itens.Count -eq 0) {
    Write-Host "Nenhum item do backup corresponde aos filtros." -ForegroundColor Yellow
    return
}
Write-Host "$($itens.Count) item(ns) encontrados no backup de $Caixa." -ForegroundColor Cyan


# ==============================================================================
# PASSO 2: SELEÇÃO
# ==============================================================================
if ($Todos) {
    $selecionados = $itens
} elseif ($GridView -and (Get-Command Out-GridView -ErrorAction SilentlyContinue)) {
    $selecionados = @($itens | Select-Object Data, Pasta, Assunto, Remetente, Email, Anexo, TamanhoKB, RemovidoEm, Id |
        Out-GridView -Title "Selecione os itens para restaurar (Ctrl+clique para vários)" -PassThru |
        ForEach-Object { $id = $_.Id; $itens | Where-Object { $_.Id -ceq $id } })
} else {
    $max = [math]::Min($itens.Count, 200)
    for ($i = 0; $i -lt $max; $i++) {
        $it  = $itens[$i]
        $rem = if ($it.RemovidoEm) { ' [apagado]' } else { '' }
        Write-Host ("  [{0,3}] {1}  {2,-25} {3}{4}" -f $i, $it.Data, ($it.Remetente -replace '^(.{25}).+', '$1'), $it.Assunto, $rem)
    }
    if ($itens.Count -gt $max) { Write-Warning "Mostrando $max de $($itens.Count). Refine os filtros ou use -Todos / -GridView." }

    $resp = Read-Host "`nNúmeros (ex: 0,3,5-8), [T] para todos os $($itens.Count) ou [Enter] para cancelar"
    $selecionados = @()
    if ($resp.Trim().ToUpper() -eq 'T') {
        $selecionados = $itens
    } elseif ($resp) {
        $indices = foreach ($parte in $resp -split ',') {
            $parte = $parte.Trim()
            if ($parte -match '^(\d+)-(\d+)$') { [int]$Matches[1]..[int]$Matches[2] }
            elseif ($parte -match '^\d+$')   { [int]$parte }
        }
        $selecionados = @($indices | Where-Object { $_ -lt $max } | Select-Object -Unique | ForEach-Object { $itens[$_] })
    }
}

if ($selecionados.Count -eq 0) { Write-Host "Nada selecionado." -ForegroundColor Yellow; return }

Write-Host "`n--- RESUMO ---" -ForegroundColor Yellow
Write-Host "Itens:    $($selecionados.Count) ($([math]::Round(($selecionados | Measure-Object TamanhoKB -Sum).Sum / 1KB, 1)) MB)"
Write-Host "Destino:  $CaixaDestino \ $PastaRestauracao$(if (-not $SemEstrutura) { ' \ <pastas originais>' })"
if (-not $Force -and (Read-Host "Confirmar restauração? (S/N)").ToUpper() -ne 'S') { return }


# ==============================================================================
# PASSO 3: CONEXÃO E PASTAS DE DESTINO
# ==============================================================================
Connect-GraphBackup -Escopos 'User.Read.All', 'MailboxFolder.ReadWrite', 'MailboxItem.ImportExport' `
    -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint

$mbx = Get-MailboxId $CaixaDestino

# Procura uma subpasta pelo nome; cria se não existir
function Get-OrNewFolder([string]$ParentId, [string]$Nome, [string]$Tipo) {
    $base = if ($ParentId) { "$script:GraphBase/$mbx/folders/$(ConvertTo-UrlId $ParentId)/childFolders" } else { "$script:GraphBase/$mbx/folders" }
    $url  = "$($base)?`$top=100"
    while ($url) {
        $r = Invoke-GraphRetry -Uri $url
        $achou = $r.value | Where-Object { $_.displayName -eq $Nome } | Select-Object -First 1
        if ($achou) { return $achou.id }
        $url = $r.'@odata.nextLink'
    }
    if (-not $Tipo) { $Tipo = 'IPF.Note' }
    return (Invoke-GraphRetry -Method POST -Uri $base -Body @{ displayName = $Nome; type = $Tipo }).id
}

$raizId     = Get-OrNewFolder $null $PastaRestauracao 'IPF.Note'
$cachePasta = @{ '' = $raizId }

function Get-PastaDestino($Item) {
    if ($SemEstrutura -or -not $Item.Pasta) { return $raizId }
    if ($cachePasta.ContainsKey($Item.Pasta)) { return $cachePasta[$Item.Pasta] }

    $partes  = $Item.Pasta -split '/'
    $atual   = $raizId
    $caminho = ''
    for ($i = 0; $i -lt $partes.Count; $i++) {
        $caminho = if ($caminho) { "$caminho/$($partes[$i])" } else { $partes[$i] }
        if (-not $cachePasta.ContainsKey($caminho)) {
            $tipo = if ($i -eq $partes.Count - 1) { $Item.TipoPasta } else { 'IPF.Note' }
            $cachePasta[$caminho] = Get-OrNewFolder $atual $partes[$i] $tipo
        }
        $atual = $cachePasta[$caminho]
    }
    return $atual
}


# ==============================================================================
# PASSO 4: IMPORTAÇÃO
# ==============================================================================
$script:sessao = $null
function Get-ImportUrl {
    $expira = $null
    if ($script:sessao) {
        $e = $script:sessao.expirationDateTime
        $expira = if ($e -is [datetime]) { [datetimeoffset]$e.ToUniversalTime() } else { [datetimeoffset]::Parse("$e", [Globalization.CultureInfo]::InvariantCulture) }
    }
    if (-not $script:sessao -or $expira -lt [datetimeoffset]::UtcNow.AddMinutes(5)) {
        $script:sessao = Invoke-GraphRetry -Method POST -Uri "$script:GraphBase/$mbx/createImportSession"
    }
    return $script:sessao.importUrl
}

# A URL de importação já vem autenticada: não usa token do Graph
function Import-Item([string]$PastaId, [string]$Dados) {
    $body = @{ FolderId = $PastaId; Mode = 'create'; Data = $Dados } | ConvertTo-Json -Compress
    for ($t = 1; ; $t++) {
        try {
            return Invoke-RestMethod -Method Post -Uri (Get-ImportUrl) -Body $body -ContentType 'application/json' -ErrorAction Stop
        } catch {
            $status = Get-HttpStatus $_
            if ($status -eq 401) { $script:sessao = $null }   # sessão expirou antes do previsto
            if ($t -ge 5 -or $status -notin 401, 429, 500, 502, 503, 504) { throw }
            Start-Sleep -Seconds ([math]::Min(60, [math]::Pow(2, $t)))
        }
    }
}

$relatorio = New-Object System.Collections.Generic.List[object]
$ok = 0
$n  = 0
foreach ($it in $selecionados) {
    $n++
    Write-Progress -Activity "Restaurando" -Status "$n/$($selecionados.Count)" -PercentComplete (100 * $n / $selecionados.Count)
    $status = 'OK'; $erro = ''
    try {
        $dados = [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $dirCaixa $it.Arquivo)))
        $null  = Import-Item (Get-PastaDestino $it) $dados
        $ok++
    } catch {
        $status = 'Falha'; $erro = $_.Exception.Message
        Write-Warning "Falha ao restaurar '$($it.Assunto)': $erro"
    }
    $relatorio.Add([pscustomobject]@{
        Data = $it.Data; Pasta = $it.Pasta; Assunto = $it.Assunto; Remetente = $it.Remetente
        Status = $status; Erro = $erro; Id = $it.Id
    })
}
Write-Progress -Activity "Restaurando" -Completed

$dirRel = Join-Path $dirCaixa 'restauracoes'
New-Item -ItemType Directory $dirRel -Force | Out-Null
$arqRel = Join-Path $dirRel "restauracao-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$relatorio | Export-Csv $arqRel -Delimiter ';' -NoTypeInformation -Encoding UTF8

$cor = if ($ok -eq $selecionados.Count) { 'Green' } else { 'Yellow' }
Write-Host "`n$ok de $($selecionados.Count) item(ns) restaurados em '$CaixaDestino\$PastaRestauracao'." -ForegroundColor $cor
Write-Host "Relatório: $arqRel"
