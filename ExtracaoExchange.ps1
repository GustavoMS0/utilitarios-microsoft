<#
.SYNOPSIS
  Extrai do Unified Audit Log quem moveu/apagou e-mails das caixas informadas.

.DESCRIPTION
  Move, MoveToDeletedItems, SoftDelete e HardDelete são registrados com
  RecordType "ExchangeItemGroup" (ações em lote), não "ExchangeItem".
  Um evento pode afetar vários e-mails: o relatório gera uma linha por e-mail.
  Datas exibidas no horário local.

.EXAMPLE
  .\ExtracaoExchange.ps1 -Caixas financeiro@empresa.com.br, boletos@empresa.com.br -Dias 30
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "E-mails das caixas a auditar (ex: financeiro@empresa.com.br)")]
    [string[]]$Caixas,

    [int]$Dias = 7,

    [string[]]$Operacoes = @('Move', 'MoveToDeletedItems', 'SoftDelete', 'HardDelete'),

    # Se a pasta for criada pelo script, só o usuário atual, SYSTEM e Administradores têm acesso
    [string]$CsvPath = "C:\Temp\Auditoria-Exchange\Auditoria-Exchange-$(Get-Date -Format 'yyyyMMdd-HHmm').csv",

    # Abre também a janela interativa (só no Windows)
    [switch]$GridView,

    # Emite os registros estruturados de auditoria para o pipeline
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib/Common.ps1')

# ==============================================================================
# PASSO 0: CONEXÃO
# ==============================================================================
Connect-ExchangeSeNecessario

try {
    # O Unified Audit Log trabalha em UTC
    $dataFim    = (Get-Date).ToUniversalTime()
    $dataInicio = $dataFim.AddDays(-$Dias)

    $tiposLogon = @{ '0' = 'Proprietário'; '1' = 'Administrador'; '2' = 'Delegado' }

    function Get-AuditPath($folder) {
        if ($null -eq $folder) { return $null }
        if ($folder.Path) { return $folder.Path }
        return $folder.PathName
    }

    # O audit log grava o dono da caixa pelo UPN (MailboxOwnerUPN), que pode ser diferente do
    # e-mail informado (alias, endereço principal). Sem conseguir resolver, usa o texto informado.
    function Resolve-UpnCaixa([string]$Caixa) {
        try {
            $mbx = if (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue) {
                Get-EXOMailbox -Identity $Caixa -Properties UserPrincipalName, PrimarySmtpAddress -ErrorAction Stop
            } else {
                Get-Mailbox -Identity $Caixa -ErrorAction Stop
            }
            if ($mbx.UserPrincipalName) { return "$($mbx.UserPrincipalName)" }
        } catch {
            Write-Warning "Não foi possível resolver a caixa '$Caixa' (Get-EXOMailbox): $($_.Exception.Message). Usando o endereço informado como UPN."
        }
        return $Caixa
    }


    # ==============================================================================
    # PASSO 1: BUSCA PAGINADA
    # ==============================================================================
    # Sem paginação o Search-UnifiedAuditLog corta em 5000 registros sem avisar.
    function Search-AuditPaginado {
        param([string]$Caixa)

        $sessao = [guid]::NewGuid().ToString()
        $todos  = New-Object System.Collections.Generic.List[object]
        $total  = $null

        do {
            $pagina = @(Search-UnifiedAuditLog -StartDate $dataInicio -EndDate $dataFim `
                -RecordType ExchangeItemGroup -Operations $Operacoes -FreeText $Caixa `
                -SessionId $sessao -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)

            if ($pagina.Count -eq 0) { break }
            if ($null -eq $total) { $total = $pagina[0].ResultCount }
            $todos.AddRange($pagina)

            # Em alguns casos o serviço devolve ResultIndex -1 (erro interno); paramos para não entrar em loop
            if ($pagina[-1].ResultIndex -lt 0) {
                Write-Warning "O serviço de auditoria retornou erro de paginação. Resultados podem estar incompletos."
                break
            }
        } while ($todos.Count -lt $total)

        if ($total -ge 50000) {
            Write-Warning "Limite de 50.000 registros por busca atingido para $Caixa. Reduza o período (-Dias)."
        }

        # ReturnLargeSet pode devolver registros duplicados
        return $todos | Sort-Object Identity -Unique
    }


    # ==============================================================================
    # PASSO 2: PROCESSAMENTO
    # ==============================================================================
    $relatorio = New-Object System.Collections.Generic.List[object]

    foreach ($caixa in $Caixas) {
        $upn = Resolve-UpnCaixa $caixa
        $rotulo = if ($upn -ne $caixa) { "$caixa (UPN $upn)" } else { $caixa }
        Write-Host "Buscando logs no Unified Audit Log para: $rotulo..." -ForegroundColor Cyan

        try {
            $logs = Search-AuditPaginado -Caixa $upn
        } catch {
            Write-Warning "Falha na busca para ${caixa}: $($_.Exception.Message)"
            continue
        }

        $eventosEncontrados = 0

        foreach ($log in $logs) {
            $auditData = $log.AuditData | ConvertFrom-Json

            # -FreeText pode trazer eventos de outras caixas que apenas citam o endereço
            if ($auditData.MailboxOwnerUPN -ne $upn) { continue }
            $eventosEncontrados++

            $dataLocal    = [datetime]::SpecifyKind($log.CreationDate, 'Utc').ToLocalTime()
            $pastaOrigem  = Get-AuditPath $auditData.Folder
            $pastaDestino = Get-AuditPath $auditData.DestFolder
            $logon        = $tiposLogon["$($auditData.LogonType)"]

            # Um evento pode conter vários e-mails; sem AffectedItems geramos uma linha só
            $itens = @($auditData.AffectedItems)
            if ($itens.Count -eq 0) { $itens = @($null) }

            foreach ($item in $itens) {
                $relatorio.Add([pscustomobject]@{
                    DataHora           = $dataLocal
                    CaixaAfetada       = $caixa
                    AcaoRealizada      = $auditData.Operation
                    UsuarioResponsavel = $auditData.UserId
                    TipoAcesso         = if ($logon) { $logon } else { $auditData.LogonType }
                    AssuntoDoEmail     = $item.Subject
                    PastaDeOrigem      = if ($item.ParentFolder.Path) { $item.ParentFolder.Path } else { $pastaOrigem }
                    PastaDeDestino     = $pastaDestino
                    IPClient           = $auditData.ClientIPAddress
                    Cliente            = $auditData.ClientInfoString
                })
            }
        }

        if ($eventosEncontrados -gt 0) {
            Write-Host "Encontrados $eventosEncontrados eventos em $caixa." -ForegroundColor Green
        } else {
            Write-Host "Nenhum evento encontrado em $caixa no período." -ForegroundColor Yellow
        }
    }


    # ==============================================================================
    # PASSO 3: RESULTADOS
    # ==============================================================================
    if ($relatorio.Count -gt 0) {
        $ordenado = $relatorio | Sort-Object DataHora -Descending
        New-PastaSegura (Split-Path $CsvPath)
        $ordenado | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -UseCulture
        Write-Host "`nRelatório gerado: $CsvPath ($($relatorio.Count) e-mails afetados)" -ForegroundColor Green

        if ($GridView -and (Get-Command Out-GridView -ErrorAction SilentlyContinue)) {
            $ordenado | Out-GridView -Title "Relatório de Auditoria (Unified) - Emails Movidos/Apagados"
        }
        if ($PassThru) {
            $ordenado | Write-Output
        }
    } else {
        Write-Host "`nNão há dados para exibir no período selecionado." -ForegroundColor Yellow
    }
} finally {
    Disconnect-ExchangeSeConectado
}
