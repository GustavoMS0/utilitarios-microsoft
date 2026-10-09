<#
.SYNOPSIS
  Funções compartilhadas por Backup-Caixa.ps1 e Restaurar-Backup.ps1
  (API de import/export de caixas do Microsoft Graph, v1.0).
#>

$script:GraphBase = 'v1.0/admin/exchange/mailboxes'

# Conecta ao Graph em modo interativo (escopos delegados) ou como aplicativo (certificado)
function Connect-GraphBackup {
    param(
        [string[]]$Escopos,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint
    )

    if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
        Write-Host "Instalando módulo: Microsoft.Graph.Authentication ..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -Repository PSGallery
    }
    Import-Module Microsoft.Graph.Authentication

    $ctx = Get-MgContext
    if ($ClientId) {
        if (-not $TenantId -or -not $CertificateThumbprint) {
            throw "O modo aplicativo exige -TenantId, -ClientId e -CertificateThumbprint."
        }
        if ($null -eq $ctx -or $ctx.ClientId -ne $ClientId -or "$($ctx.AuthType)" -ne 'AppOnly') {
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
        }
        Write-Host "Conectado ao Graph como aplicativo ($ClientId)" -ForegroundColor Green
    } else {
        if ($null -eq $ctx -or ($Escopos | Where-Object { $_ -notin $ctx.Scopes })) {
            Connect-MgGraph -Scopes $Escopos -NoWelcome
        }
        Write-Host "Conectado ao Graph: $((Get-MgContext).Account)" -ForegroundColor Green
    }
}

function Get-HttpStatus($Erro) {
    try { return [int]$Erro.Exception.Response.StatusCode } catch { return 0 }
}

# Invoke-MgGraphRequest com nova tentativa em throttling (429) e falhas transitórias (5xx)
function Invoke-GraphRetry {
    param(
        [string]$Method = 'GET',
        [Parameter(Mandatory = $true)][string]$Uri,
        $Body,
        [hashtable]$Headers = @{},
        [int]$Tentativas = 6
    )

    for ($i = 1; ; $i++) {
        try {
            $p = @{ Method = $Method; Uri = $Uri; Headers = $Headers; ErrorAction = 'Stop' }
            if ($null -ne $Body) {
                $p.Body        = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 -Compress }
                $p.ContentType = 'application/json'
            }
            return Invoke-MgGraphRequest @p
        } catch {
            $status = Get-HttpStatus $_
            if ($i -ge $Tentativas -or $status -notin 429, 500, 502, 503, 504) { throw }

            $espera = [math]::Min(60, [math]::Pow(2, $i))
            try {
                $ra = $_.Exception.Response.Headers.RetryAfter.Delta
                if ($ra) { $espera = [math]::Ceiling($ra.TotalSeconds) }
            } catch {}
            Write-Warning "Graph respondeu $status; nova tentativa em ${espera}s ($i/$Tentativas)"
            Start-Sleep -Seconds $espera
        }
    }
}

function ConvertTo-UrlId([string]$Id) { [uri]::EscapeDataString($Id) }

# Id da caixa (MBX:...) usado pelas APIs /admin/exchange/mailboxes
function Get-MailboxId([string]$Upn) {
    $r = Invoke-GraphRetry -Uri "v1.0/users/$([uri]::EscapeDataString($Upn))/settings/exchange"
    if (-not $r.primaryMailboxId) { throw "Caixa não encontrada para $Upn" }
    return $r.primaryMailboxId
}

# Todas as pastas da caixa com o caminho completo (ex: "Inbox/Clientes/2026")
function Get-MailboxFolders([string]$MailboxId) {
    $pastas = New-Object System.Collections.Generic.List[object]
    $fila   = New-Object System.Collections.Generic.Queue[object]
    $fila.Enqueue(@{ Url = "$script:GraphBase/$MailboxId/folders?`$top=100"; Caminho = '' })

    while ($fila.Count -gt 0) {
        $atual = $fila.Dequeue()
        $url   = $atual.Url
        while ($url) {
            $r = Invoke-GraphRetry -Uri $url
            foreach ($f in $r.value) {
                $caminho = if ($atual.Caminho) { "$($atual.Caminho)/$($f.displayName)" } else { $f.displayName }
                $pastas.Add([pscustomobject]@{
                    Id            = $f.id
                    Nome          = $f.displayName
                    Caminho       = $caminho
                    Tipo          = $f.type
                    WellKnownName = $f.wellKnownName
                    Itens         = $f.totalItemCount
                })
                if ($f.childFolderCount -gt 0) {
                    $fila.Enqueue(@{ Url = "$script:GraphBase/$MailboxId/folders/$(ConvertTo-UrlId $f.id)/childFolders?`$top=100"; Caminho = $caminho })
                }
            }
            $url = $r.'@odata.nextLink'
        }
    }
    return $pastas
}

# Propriedades MAPI usadas para montar o índice pesquisável
$script:PropsMapi = [ordered]@{
    Assunto   = @{ Id = 'String 0x0037';     Tag = 0x0037 }   # PR_SUBJECT
    Remetente = @{ Id = 'String 0x0C1A';     Tag = 0x0C1A }   # PR_SENDER_NAME
    Email     = @{ Id = 'String 0x5D01';     Tag = 0x5D01 }   # PR_SENDER_SMTP_ADDRESS
    Para      = @{ Id = 'String 0x0E04';     Tag = 0x0E04 }   # PR_DISPLAY_TO
    Recebido  = @{ Id = 'SystemTime 0x0E06'; Tag = 0x0E06 }   # PR_MESSAGE_DELIVERY_TIME
    Anexo     = @{ Id = 'Boolean 0x0E1B';    Tag = 0x0E1B }   # PR_HASATTACH
}

# O Graph devolve os ids normalizados (ex: "String 0x37"), então comparamos pela tag numérica
function ConvertFrom-ExtendedProperties($Props) {
    $porTag = @{}
    foreach ($p in @($Props)) {
        if ($p.id -match '0x([0-9A-Fa-f]+)') { $porTag[[Convert]::ToInt32($Matches[1], 16)] = $p.value }
    }
    $saida = @{}
    foreach ($nome in $script:PropsMapi.Keys) { $saida[$nome] = $porTag[$script:PropsMapi[$nome].Tag] }
    return $saida
}

function Get-ExpandMetadados {
    $filtro = ($script:PropsMapi.Values | ForEach-Object { "id eq '$($_.Id)'" }) -join ' or '
    return '$expand=' + [uri]::EscapeDataString("singleValueExtendedProperties(`$filter=$filtro)")
}

# Metadados de até 20 itens por chamada ($batch). Falhas não interrompem o backup.
function Get-ItemMetadata([string]$MailboxId, [string]$FolderId, [string[]]$ItemIds) {
    $resultado = @{}
    $expand = Get-ExpandMetadados
    for ($i = 0; $i -lt $ItemIds.Count; $i += 20) {
        $lote = @($ItemIds[$i..([math]::Min($i + 19, $ItemIds.Count - 1))])
        $n = 0
        $reqs = foreach ($id in $lote) {
            @{ id = "$n"; method = 'GET'; url = "/admin/exchange/mailboxes/$MailboxId/folders/$(ConvertTo-UrlId $FolderId)/items/$(ConvertTo-UrlId $id)?$expand" }
            $n++
        }
        try {
            $r = Invoke-GraphRetry -Method POST -Uri 'v1.0/$batch' -Body @{ requests = @($reqs) }
            foreach ($resp in $r.responses) {
                if ($resp.status -eq 200) {
                    $resultado[$lote[[int]$resp.id]] = ConvertFrom-ExtendedProperties $resp.body.singleValueExtendedProperties
                }
            }
        } catch {
            Write-Warning "Não foi possível ler os metadados de $($lote.Count) item(ns): $($_.Exception.Message)"
        }
    }
    return $resultado
}

# Índice: um CSV por caixa, sempre com ';' para abrir direto no Excel em pt-BR
function Import-Indice([string]$Caminho) {
    $indice = @{}
    if (Test-Path $Caminho) {
        foreach ($l in Import-Csv $Caminho -Delimiter ';' -Encoding UTF8) { $indice[$l.Id] = $l }
    }
    return $indice
}

function Export-Indice([hashtable]$Indice, [string]$Caminho) {
    $tmp = "$Caminho.tmp"
    $Indice.Values | Sort-Object Data -Descending | Export-Csv $tmp -Delimiter ';' -NoTypeInformation -Encoding UTF8
    Move-Item $tmp $Caminho -Force
}
