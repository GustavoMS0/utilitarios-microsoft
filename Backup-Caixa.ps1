<#
.SYNOPSIS
  Backup incremental de caixas do Exchange Online, com fidelidade total, via Microsoft Graph.

.DESCRIPTION
  Usa a API de import/export de caixas do Graph (v1.0). Cada item (e-mail, compromisso,
  contato, tarefa), com anexos, é salvo como arquivo .fts. O .fts é um formato opaco que
  só serve para restaurar com Restaurar-Backup.ps1; ele não abre no Outlook.

  A primeira execução copia tudo. As seguintes copiam apenas o que mudou (delta query).
  Itens apagados na caixa NÃO são apagados do backup: eles são marcados com a data em
  que sumiram (coluna RemovidoEm do índice).

  Estrutura gerada em <Destino>\<caixa>\:
    itens\AAAA\MM\*.fts   conteúdo dos itens
    indice.csv            assunto, remetente, data e pasta de cada item (pesquisável no Excel)
    estado.json           tokens de sincronização incremental
    falhas.log            itens que não puderam ser exportados

.EXAMPLE
  # Login interativo (só acessa caixas em que sua conta tem permissão)
  .\Backup-Caixa.ps1 -Caixas financeiro@empresa.com.br

.EXAMPLE
  # Como aplicativo, para agendar (ver Criar-AppBackup.ps1)
  .\Backup-Caixa.ps1 -Caixas financeiro@empresa.com.br, boletos@empresa.com.br `
      -TenantId <tenant-id> -ClientId <app-id> -CertificateThumbprint <thumbprint>
#>
param(
    [Parameter(Mandatory = $true, HelpMessage = "E-mails das caixas (ex: financeiro@empresa.com.br)")]
    [string[]]$Caixas,

    [string]$Destino = 'C:\Backup\Exchange',

    # Caminhos de pasta a ignorar; aceita curinga (ex: 'Lixo Eletrônico', 'Inbox/Newsletter*')
    [string[]]$ExcluirPastas = @(),

    # Tamanho máximo somado dos itens em cada chamada de exportação
    [int]$LoteMB = 20,

    # Modo aplicativo (opcional)
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\GraphMailbox.ps1')

$stamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
$dirLog = Join-Path $Destino 'logs'
New-Item -ItemType Directory $dirLog -Force | Out-Null
Start-Transcript -Path (Join-Path $dirLog "backup-$stamp.log") | Out-Null


# ==============================================================================
# FUNÇÕES
# ==============================================================================
function Get-Hash([string]$Texto) {
    $sha = [Security.Cryptography.SHA1]::Create()
    -join ($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Texto))[0..7] | ForEach-Object { $_.ToString('x2') })
}

function Format-Data($Valor) {
    if (-not $Valor) { return '' }
    if ($Valor -is [datetime])       { return $Valor.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss') }
    if ($Valor -is [datetimeoffset]) { return $Valor.LocalDateTime.ToString('yyyy-MM-dd HH:mm:ss') }
    try   { return [datetimeoffset]::Parse("$Valor", [Globalization.CultureInfo]::InvariantCulture).LocalDateTime.ToString('yyyy-MM-dd HH:mm:ss') }
    catch { return "$Valor" }
}

function Add-Falha([string]$ItemId, [string]$Pasta, [string]$Mensagem) {
    $script:stats.Falhas++
    Add-Content -Path $script:arqFalhas -Value "$(Get-Date -Format s);$Pasta;$ItemId;$Mensagem" -Encoding UTF8
}

# Exporta até 20 itens numa chamada e grava os arquivos + entradas do índice
function Export-Lote($Itens, $Pasta, $Meta) {
    $porId = @{}
    foreach ($i in $Itens) { $porId[$i.id] = $i }

    try {
        $r = Invoke-GraphRetry -Method POST -Uri "$script:GraphBase/$script:mbx/exportItems" -Body @{ itemIds = @($Itens | ForEach-Object { $_.id }) }
    } catch {
        foreach ($i in $Itens) { Add-Falha $i.id $Pasta.Caminho $_.Exception.Message }
        return
    }

    foreach ($x in $r.value) {
        $it = $porId[$x.itemId]
        if (-not $x.data -or -not $it) {
            Add-Falha $x.itemId $Pasta.Caminho "$($x.error.message)"
            continue
        }

        $criado = try { [datetimeoffset]::Parse((Format-Data $it.createdDateTime), [Globalization.CultureInfo]::InvariantCulture) } catch { [datetimeoffset]::Now }
        $rel    = 'itens\{0:yyyy}\{0:MM}\{1}.fts' -f $criado, (Get-Hash $x.itemId)
        $abs    = Join-Path $script:dirCaixa $rel
        New-Item -ItemType Directory (Split-Path $abs) -Force | Out-Null
        $bytes = [Convert]::FromBase64String($x.data)
        [IO.File]::WriteAllBytes($abs, $bytes)

        $m = $Meta[$x.itemId]
        if (-not $m) { $m = @{} }
        $data = if ($m.Recebido) { Format-Data $m.Recebido } else { Format-Data $it.createdDateTime }

        if ($script:indice.ContainsKey($x.itemId)) { $script:stats.Atualizados++ } else { $script:stats.Novos++ }
        $script:stats.Bytes += $bytes.Length

        $script:indice[$x.itemId] = [pscustomobject]@{
            Id         = $x.itemId
            Data       = $data
            Pasta      = $Pasta.Caminho
            Tipo       = $it.type
            Assunto    = $m.Assunto
            Remetente  = $m.Remetente
            Email      = $m.Email
            Para       = $m.Para
            Anexo      = if ("$($m.Anexo)" -eq 'true') { 'Sim' } else { '' }   # o Graph devolve "true"/"false" como texto
            TamanhoKB  = [math]::Round($bytes.Length / 1KB, 1)
            Arquivo    = $rel
            PastaId    = $Pasta.Id
            TipoPasta  = $Pasta.Tipo
            ChangeKey  = $x.changeKey
            BackupEm   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            RemovidoEm = ''
        }
    }
}

function Sync-Pasta($Pasta) {
    $urlInicial = "$script:GraphBase/$script:mbx/folders/$(ConvertTo-UrlId $Pasta.Id)/items/delta"
    $url        = if ($script:deltas[$Pasta.Id]) { $script:deltas[$Pasta.Id] } else { $urlInicial }
    $headers    = @{ Prefer = 'odata.maxpagesize=200' }
    $alterados  = New-Object System.Collections.Generic.List[object]
    $removidos  = New-Object System.Collections.Generic.List[string]
    $deltaLink  = $null
    $reiniciado = $false

    while ($url) {
        try {
            $r = Invoke-GraphRetry -Uri $url -Headers $headers
        } catch {
            # Token de sincronização expirado/inválido: refaz a varredura completa da pasta
            if (-not $reiniciado -and $url -ne $urlInicial -and (Get-HttpStatus $_) -in 400, 404, 410) {
                Write-Warning "Token de sincronização expirado em '$($Pasta.Caminho)'; varrendo a pasta inteira."
                $url = $urlInicial; $reiniciado = $true
                $alterados.Clear(); $removidos.Clear()
                continue
            }
            throw
        }
        foreach ($it in $r.value) {
            if ($null -ne $it['@removed']) { $removidos.Add($it.id) } else { $alterados.Add($it) }
        }
        $url = $r.'@odata.nextLink'
        if (-not $url) { $deltaLink = $r.'@odata.deltaLink' }
    }

    # Apagados na origem: mantém no backup e registra quando sumiram
    foreach ($id in $removidos) {
        $e = $script:indice[$id]
        if ($e -and $e.PastaId -eq $Pasta.Id -and -not $e.RemovidoEm) {
            $e.RemovidoEm = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            $script:stats.Removidos++
        }
    }

    # Só exporta o que é novo ou mudou de versão (ChangeKey)
    $exportar = New-Object System.Collections.Generic.List[object]
    foreach ($it in $alterados) {
        $e = $script:indice[$it.id]
        if ($e -and $e.ChangeKey -eq $it.changeKey -and (Test-Path (Join-Path $script:dirCaixa $e.Arquivo))) {
            $e.PastaId = $Pasta.Id; $e.Pasta = $Pasta.Caminho; $e.TipoPasta = $Pasta.Tipo; $e.RemovidoEm = ''
        } else {
            $exportar.Add($it)
        }
    }

    if ($exportar.Count -gt 0) {
        Write-Host ("  {0,-50} {1,6} item(ns) para exportar" -f $Pasta.Caminho, $exportar.Count)
        $meta = Get-ItemMetadata $script:mbx $Pasta.Id @($exportar | ForEach-Object { $_.id })

        $lote = New-Object System.Collections.Generic.List[object]
        $tam  = 0
        foreach ($it in $exportar) {
            if ($lote.Count -ge 20 -or ($lote.Count -gt 0 -and $tam + [long]$it.size -gt $LoteMB * 1MB)) {
                Export-Lote $lote $Pasta $meta
                $lote.Clear(); $tam = 0
            }
            $lote.Add($it); $tam += [long]$it.size
        }
        if ($lote.Count -gt 0) { Export-Lote $lote $Pasta $meta }
    }

    if ($deltaLink) { $script:deltas[$Pasta.Id] = $deltaLink }
}

function Save-Estado {
    @{
        Caixa        = $script:caixa
        MailboxId    = $script:mbx
        UltimoBackup = Get-Date -Format 's'
        Deltas       = $script:deltas
    } | ConvertTo-Json -Depth 5 | Set-Content $script:arqEstado -Encoding UTF8
    Export-Indice $script:indice $script:arqIndice
}


# ==============================================================================
# EXECUÇÃO
# ==============================================================================
$caixasComErro = 0
try {
    Connect-GraphBackup -Escopos 'User.Read.All', 'MailboxFolder.Read', 'MailboxItem.Read', 'MailboxItem.Export' `
        -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint

    foreach ($script:caixa in $Caixas) {
        Write-Host "`n=== $script:caixa ===" -ForegroundColor Cyan
        $script:stats = @{ Novos = 0; Atualizados = 0; Removidos = 0; Falhas = 0; Bytes = 0 }

        try {
            $script:dirCaixa  = Join-Path $Destino $script:caixa.ToLower()
            $script:arqEstado = Join-Path $script:dirCaixa 'estado.json'
            $script:arqIndice = Join-Path $script:dirCaixa 'indice.csv'
            $script:arqFalhas = Join-Path $script:dirCaixa 'falhas.log'
            New-Item -ItemType Directory (Join-Path $script:dirCaixa 'itens') -Force | Out-Null

            $script:mbx    = Get-MailboxId $script:caixa
            $script:indice = Import-Indice $script:arqIndice
            $script:deltas = @{}
            if (Test-Path $script:arqEstado) {
                $estado = Get-Content $script:arqEstado -Raw | ConvertFrom-Json
                if ($estado.MailboxId -eq $script:mbx -and $estado.Deltas) {
                    foreach ($p in $estado.Deltas.PSObject.Properties) { $script:deltas[$p.Name] = $p.Value }
                }
            }
            $modo = if ($script:deltas.Count -gt 0) { 'incremental' } else { 'completo (primeira execução)' }

            $pastas = @(Get-MailboxFolders $script:mbx | Where-Object {
                $pasta = $_
                -not ($ExcluirPastas | Where-Object { $pasta.Caminho -like $_ })
            })
            Write-Host "Backup $modo de $($pastas.Count) pasta(s)..."

            foreach ($pasta in $pastas) {
                try {
                    Sync-Pasta $pasta
                } catch {
                    Add-Falha '' $pasta.Caminho "Pasta: $($_.Exception.Message)"
                    Write-Warning "Falha na pasta '$($pasta.Caminho)': $($_.Exception.Message)"
                }
                Save-Estado   # salva a cada pasta: se cair no meio, a próxima execução continua daqui
            }

            # Pastas que deixaram de existir: seus itens foram apagados na origem
            $idsPastas = @{}
            foreach ($p in $pastas) { $idsPastas[$p.Id] = $true }
            foreach ($e in $script:indice.Values) {
                if (-not $idsPastas[$e.PastaId] -and -not $e.RemovidoEm -and
                    -not ($ExcluirPastas | Where-Object { $e.Pasta -like $_ })) {
                    $e.RemovidoEm = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                    $script:stats.Removidos++
                }
            }
            foreach ($id in @($script:deltas.Keys)) { if (-not $idsPastas[$id]) { $script:deltas.Remove($id) } }
            Save-Estado

            $s = $script:stats
            Write-Host ("Novos: {0} | Atualizados: {1} | Removidos na origem: {2} | Falhas: {3} | {4:N1} MB | Itens no backup: {5}" -f `
                $s.Novos, $s.Atualizados, $s.Removidos, $s.Falhas, ($s.Bytes / 1MB), $script:indice.Count) -ForegroundColor Green
            if ($s.Falhas -gt 0) {
                Write-Warning "Detalhes das falhas em: $script:arqFalhas"
                $caixasComErro++
            }
        } catch {
            Write-Warning "Falha no backup de ${script:caixa}: $($_.Exception.Message)"
            $caixasComErro++
        }
    }
} finally {
    Stop-Transcript | Out-Null
}

# Código de saída diferente de zero para o Agendador de Tarefas sinalizar erro
exit $caixasComErro
