<#
.SYNOPSIS
  Menu interativo com todos os utilitários deste repositório.

.DESCRIPTION
  Escolha um utilitário, responda às perguntas e confirme. Antes de executar, o menu mostra o
  comando equivalente, que pode ser copiado para rodar direto depois (ou agendar).
  Os scripts continuam funcionando sozinhos; o menu só monta os parâmetros.

.EXAMPLE
  .\Menu.ps1
#>
param(
    # Pasta dos scripts (usado pelos testes)
    [string]$PastaScripts = $PSScriptRoot
)

# ==============================================================================
# CATÁLOGO
# ==============================================================================
# Tipos de pergunta: Texto, Caixas (lista de e-mails), Caixa (um e-mail), Numero, Data,
# SimNao (vira switch), Escolha (uma opção), Multipla (várias opções)
$regexEmail = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

$itens = @(
    @{ Grupo = 'Usuários'; Titulo = 'Criar usuário a partir de um modelo'; Script = 'NewClone.ps1'; Perguntas = @() }

    @{ Grupo = 'SharePoint'; Titulo = 'Auditoria de uso dos sites (e OneDrive)'; Script = 'Sharepoint.ps1'; Perguntas = @(
        @{ Nome = 'Periodo';         Tipo = 'Escolha'; Texto = 'Período do relatório'; Opcoes = 'D7', 'D30', 'D90', 'D180'; Padrao = 'D30' }
        @{ Nome = 'DiasInativo';     Tipo = 'Numero';  Texto = 'Dias sem uso para considerar inativo'; Padrao = 90 }
        @{ Nome = 'SpoAdminUrl';     Tipo = 'Texto';   Texto = 'URL do admin do SharePoint para dados de governança (ex: https://empresa-admin.sharepoint.com)' }
        @{ Nome = 'IncluirOneDrive'; Tipo = 'SimNao';  Texto = 'Incluir relatório do OneDrive?' }
    ) }

    @{ Grupo = 'Exchange: auditoria'; Titulo = 'Ativar auditoria em caixas'; Script = 'Auditoriaexchange.ps1'; Perguntas = @(
        @{ Nome = 'Caixas'; Tipo = 'Caixas'; Texto = 'Caixas (separadas por vírgula)'; Obrigatorio = $true }
    ) }
    @{ Grupo = 'Exchange: auditoria'; Titulo = 'Relatório: quem moveu/apagou e-mails'; Script = 'ExtracaoExchange.ps1'; Perguntas = @(
        @{ Nome = 'Caixas'; Tipo = 'Caixas'; Texto = 'Caixas (separadas por vírgula)'; Obrigatorio = $true }
        @{ Nome = 'Dias';   Tipo = 'Numero'; Texto = 'Quantos dias para trás'; Padrao = 7 }
    ) }

    @{ Grupo = 'Exchange: backup e recuperação'; Titulo = 'Recuperar itens excluídos (até 30 dias)'; Script = 'Recuperar-ItensExcluidos.ps1'; Perguntas = @(
        @{ Nome = 'Caixa';   Tipo = 'Caixa'; Texto = 'Caixa'; Obrigatorio = $true }
        @{ Nome = 'Assunto'; Tipo = 'Texto'; Texto = 'Trecho do assunto (Enter = todos)' }
        @{ Nome = 'De';      Tipo = 'Data';  Texto = 'Excluídos a partir de (AAAA-MM-DD, Enter = sem limite)' }
    ) }
    @{ Grupo = 'Exchange: backup e recuperação'; Titulo = 'Backup de caixas'; Script = 'Backup-Caixa.ps1'; Perguntas = @(
        @{ Nome = 'Caixas';  Tipo = 'Caixas'; Texto = 'Caixas (separadas por vírgula)'; Obrigatorio = $true }
        @{ Nome = 'Destino'; Tipo = 'Texto';  Texto = 'Pasta do backup'; Padrao = 'C:\Backup\Exchange' }
    ) }
    @{ Grupo = 'Exchange: backup e recuperação'; Titulo = 'Restaurar itens do backup'; Script = 'Restaurar-Backup.ps1'; Perguntas = @(
        @{ Nome = 'Caixa';            Tipo = 'Caixa';  Texto = 'Caixa do backup'; Obrigatorio = $true }
        @{ Nome = 'CaixaDestino';     Tipo = 'Caixa';  Texto = 'Restaurar em outra caixa? (Enter = na própria caixa)' }
        @{ Nome = 'Assunto';          Tipo = 'Texto';  Texto = 'Trecho do assunto (Enter = qualquer)' }
        @{ Nome = 'Remetente';        Tipo = 'Texto';  Texto = 'Remetente, nome ou e-mail (Enter = qualquer)' }
        @{ Nome = 'De';               Tipo = 'Data';   Texto = 'Recebidos a partir de (AAAA-MM-DD, Enter = sem limite)' }
        @{ Nome = 'Ate';              Tipo = 'Data';   Texto = 'Recebidos até (AAAA-MM-DD, Enter = sem limite)' }
        @{ Nome = 'SomenteRemovidos'; Tipo = 'SimNao'; Texto = 'Só itens apagados na caixa?' }
    ) }
    @{ Grupo = 'Exchange: backup e recuperação'; Titulo = 'Exportar caixa em PST (Purview)'; Script = 'Exportar-PST.ps1'; Perguntas = @(
        @{ Nome = 'Caixa';           Tipo = 'Caixa';    Texto = 'Caixa a exportar'; Obrigatorio = $true }
        @{ Nome = 'SomenteDownload'; Tipo = 'SimNao';   Texto = 'Só baixar a última exportação já pronta (sem exportar de novo)?' }
        @{ Nome = 'Conteudo';        Tipo = 'Multipla'; Texto = 'O que exportar'; Opcoes = 'Email', 'Calendario', 'Contatos', 'Tarefas', 'Notas', 'Teams'; Padrao = 'Tudo'; PularSe = 'SomenteDownload' }
        @{ Nome = 'De';              Tipo = 'Data';     Texto = 'Recebidos a partir de (AAAA-MM-DD, Enter = sem limite)'; PularSe = 'SomenteDownload' }
        @{ Nome = 'Ate';             Tipo = 'Data';     Texto = 'Recebidos até (AAAA-MM-DD, Enter = sem limite)'; PularSe = 'SomenteDownload' }
        @{ Nome = 'Destino';         Tipo = 'Texto';    Texto = 'Pasta de destino'; Padrao = 'C:\Backup\PST' }
    ) }
    @{ Grupo = 'Exchange: backup e recuperação'; Titulo = 'Criar app para backup agendado'; Script = 'Criar-AppBackup.ps1'; Perguntas = @(
        @{ Nome = 'PermitirRestauracao'; Tipo = 'SimNao'; Texto = 'Incluir permissões de restauração?' }
    ) }
)


# ==============================================================================
# FUNÇÕES
# ==============================================================================
function Read-Resposta($P) {
    $sufixo = if ($null -ne $P.Padrao) { " [$($P.Padrao)]" } else { '' }
    while ($true) {
        switch ($P.Tipo) {
            'SimNao' {
                $r = (Read-Host "$($P.Texto) (S/N) [N]").Trim().ToUpper()
                if ($r -in '', 'N') { return $null }
                if ($r -eq 'S') { return $true }
            }
            'Escolha' {
                $r = (Read-Host "$($P.Texto) ($($P.Opcoes -join '/'))$sufixo").Trim()
                if (-not $r) { return $P.Padrao }
                $op = $P.Opcoes | Where-Object { $_ -eq $r } | Select-Object -First 1
                if ($op) { return $op }
            }
            'Multipla' {
                Write-Host "$($P.Texto):"
                for ($i = 0; $i -lt $P.Opcoes.Count; $i++) { Write-Host ("  [{0}] {1}" -f ($i + 1), $P.Opcoes[$i]) }
                $r = (Read-Host "Números separados por vírgula (Enter = $($P.Padrao))").Trim()
                if (-not $r) { return $null }   # padrão do próprio script
                $nums = @($r -split ',' | ForEach-Object { $_.Trim() })
                if (-not ($nums | Where-Object { $_ -notmatch '^\d+$' -or [int]$_ -lt 1 -or [int]$_ -gt $P.Opcoes.Count })) {
                    return [string[]]@($nums | ForEach-Object { $P.Opcoes[[int]$_ - 1] } | Select-Object -Unique)
                }
            }
            'Numero' {
                $r = (Read-Host "$($P.Texto)$sufixo").Trim()
                if (-not $r) { return $null }
                if ($r -match '^\d+$') { return [int]$r }
            }
            'Data' {
                $r = (Read-Host $P.Texto).Trim()
                if (-not $r) { return $null }
                $d = [datetime]::MinValue
                if ([datetime]::TryParseExact($r, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$d)) { return $d }
            }
            'Caixa' {
                $r = (Read-Host $P.Texto).Trim()
                if (-not $r -and -not $P.Obrigatorio) { return $null }
                if ($r -match $regexEmail) { return $r }
            }
            'Caixas' {
                $lista = @((Read-Host $P.Texto) -split '[,;\s]+' | Where-Object { $_ })
                if ($lista.Count -gt 0 -and -not ($lista | Where-Object { $_ -notmatch $regexEmail })) { return [string[]]$lista }
            }
            default {
                $r = (Read-Host "$($P.Texto)$sufixo").Trim()
                if (-not $r) { return $null }
                return $r
            }
        }
        Write-Host "  Resposta inválida, tente de novo." -ForegroundColor Yellow
    }
}

# Comando equivalente, para mostrar e para reaproveitar
function Format-Comando([string]$Script, [System.Collections.Specialized.OrderedDictionary]$Params) {
    $partes = @(".\$Script")
    foreach ($k in $Params.Keys) {
        $v = $Params[$k]
        if ($v -is [bool])          { $partes += "-$k" }
        elseif ($v -is [datetime])  { $partes += "-$k $($v.ToString('yyyy-MM-dd'))" }
        elseif ($v -is [array])     { $partes += "-$k " + (($v | ForEach-Object { if ("$_" -match '\s') { "'$_'" } else { "$_" } }) -join ', ') }
        elseif ("$v" -match '\s')   { $partes += "-$k '$v'" }
        else                        { $partes += "-$k $v" }
    }
    $partes -join ' '
}

function Invoke-Item($Item) {
    Write-Host "`n=== $($Item.Titulo) ===" -ForegroundColor Cyan
    $params = [ordered]@{}
    foreach ($p in $Item.Perguntas) {
        if ($p.PularSe -and $params.Contains($p.PularSe)) { continue }
        $v = Read-Resposta $p
        if ($null -ne $v) { $params[$p.Nome] = $v }
    }

    $comando = Format-Comando $Item.Script $params
    Write-Host "`nComando: " -NoNewline; Write-Host $comando -ForegroundColor Green
    if ((Read-Host "Executar? (S/N) [S]").Trim().ToUpper() -notin '', 'S') { return }

    # Switches viram [switch]; o resto vai como está
    $splat = @{}
    foreach ($k in $params.Keys) { $splat[$k] = if ($params[$k] -is [bool]) { [switch]$true } else { $params[$k] } }

    $global:LASTEXITCODE = 0
    try {
        & (Join-Path $PastaScripts $Item.Script) @splat
        $cor = if ($LASTEXITCODE -eq 0) { 'Green' } else { 'Yellow' }
        Write-Host "`nConcluído (código $LASTEXITCODE)." -ForegroundColor $cor
    } catch {
        Write-Host "`nErro: $($_.Exception.Message)" -ForegroundColor Red
    }
    $null = Read-Host "`nEnter para voltar ao menu"
}


# ==============================================================================
# MENU
# ==============================================================================
while ($true) {
    try { Clear-Host } catch {}
    Write-Host "Utilitários Microsoft 365" -ForegroundColor Cyan
    $grupo = $null
    for ($i = 0; $i -lt $itens.Count; $i++) {
        if ($itens[$i].Grupo -ne $grupo) { $grupo = $itens[$i].Grupo; Write-Host "`n $grupo" -ForegroundColor Yellow }
        Write-Host ("  [{0}] {1}" -f ($i + 1), $itens[$i].Titulo)
    }
    Write-Host "`n  [0] Sair"
    $op = (Read-Host "`nEscolha").Trim()
    if ($op -eq '0') { break }
    if ($op -match '^\d+$' -and [int]$op -ge 1 -and [int]$op -le $itens.Count) {
        Invoke-Item $itens[[int]$op - 1]
    }
}
