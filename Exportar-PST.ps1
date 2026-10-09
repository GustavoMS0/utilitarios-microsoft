<#
.SYNOPSIS
  Exporta a caixa de um colaborador (ex: desligado) para PST pelo Microsoft Purview eDiscovery.

.DESCRIPTION
  Usa a API de eDiscovery do Microsoft Graph (v1.0), o mesmo processo do portal do Purview:
    1. cria um caso de eDiscovery e uma pesquisa da caixa do colaborador (inteira ou filtrada com -Conteudo)
    2. calcula a estimativa (quantidade de itens e tamanho)
    3. exporta em PST com a estrutura de pastas original
    4. baixa os arquivos gerados e extrai o(s) PST(s)
  A caixa do colaborador não é alterada. O caso fica no Purview como registro da exportação.

  Requisitos do tenant:
    - Conta do administrador no grupo de funções "eDiscovery Manager" (Purview > Funções e escopos)
    - Conforme a licença, a Microsoft pode exigir o Purview pay-as-you-go para a API de eDiscovery
      (Purview > Configurações > Faturamento); com os recursos premium do eDiscovery habilitados, não é preciso

  O download exige um segundo login no navegador: os arquivos ficam num serviço do Purview que
  não aceita o token do Graph. Se o download falhar, use -SomenteDownload para baixar a exportação
  já pronta, sem exportar de novo.

.EXAMPLE
  .\Exportar-PST.ps1 -Caixa ex.colaborador@empresa.com.br

.EXAMPLE
  # Só e-mails e calendário, de 2025 em diante
  .\Exportar-PST.ps1 -Caixa ex.colaborador@empresa.com.br -Conteudo Email, Calendario -De 2025-01-01

.EXAMPLE
  # Caixa inteira, menos os chats do Teams
  .\Exportar-PST.ps1 -Caixa ex.colaborador@empresa.com.br -Conteudo Email, Calendario, Contatos, Tarefas, Notas

.EXAMPLE
  # Baixa a exportação mais recente dessa caixa, sem exportar de novo
  .\Exportar-PST.ps1 -Caixa ex.colaborador@empresa.com.br -SomenteDownload

.EXAMPLE
  # Só prepara a exportação; o download é feito depois pelo portal
  .\Exportar-PST.ps1 -Caixa ex.colaborador@empresa.com.br -SemDownload
#>
param(
    [Parameter(Mandatory = $true, HelpMessage = "E-mail da caixa a exportar (ex: ex.colaborador@empresa.com.br)")]
    [ValidateNotNullOrEmpty()]
    [string]$Caixa,

    [string]$Destino = 'C:\Backup\PST',

    # O que exportar; combine vários (ex: -Conteudo Email, Calendario). Padrão: a caixa inteira.
    [ValidateSet('Tudo', 'Email', 'Calendario', 'Contatos', 'Tarefas', 'Notas', 'Teams')]
    [string[]]$Conteudo = @('Tudo'),

    # Período pela data de recebimento (itens sem essa data, como contatos, ficam de fora quando usado)
    [datetime]$De,
    [datetime]$Ate,

    # Consulta KQL adicional, combinada com os filtros acima
    [string]$Consulta = '',

    # Com filtros, itens não indexados (ex: anexos criptografados) ficam de fora por padrão, porque
    # não dá para saber se correspondem ao filtro. Use esta opção para incluí-los mesmo assim.
    [switch]$IncluirNaoIndexados,

    # Não baixa: deixa a exportação pronta no portal do Purview
    [switch]$SemDownload,

    # Não exporta: baixa a exportação concluída mais recente da caixa (ou do caso em -CasoId)
    [switch]$SomenteDownload,

    # Id de um caso de eDiscovery existente para -SomenteDownload
    [string]$CasoId,

    # Login do download por código de dispositivo, em vez de abrir o navegador
    [switch]$CodigoDispositivo,

    # Registra no tenant, sem perguntar, o serviço de download do Purview se ele faltar
    # (configuração única; exige Administrador Global ou Administrador de Aplicativos)
    [switch]$RegistrarServicoDownload,

    # Não extrai o PST do .zip baixado
    [switch]$ManterZip,

    # Tempo máximo de espera pela estimativa + exportação
    [int]$AguardarMinutos = 240,

    # Aplicativo usado no login do download (padrão: Microsoft Graph Command Line Tools)
    [string]$ClientIdDownload = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\GraphMailbox.ps1')

$appPurview   = 'b26e684c-5068-4120-a679-64a5d2c909d9'   # Microsoft Purview eDiscovery (serviço de download)
$escopoDown   = "$appPurview/eDiscovery.Download.Read"
$portal       = 'https://purview.microsoft.com/ediscovery'
# exportResult, includeFolderAndPath etc. são membros "evolvable" das enumerações: sem este cabeçalho
# o Graph os devolve como unknownFutureValue
$cabecalhos   = @{ Prefer = 'include-unknown-enum-members' }
$intervalo    = 30
$msgSemSP     = "O serviço de download do Purview (app $appPurview) não está registrado no tenant."
$cmdSemSP     = "Connect-MgGraph -Scopes Application.ReadWrite.All; Invoke-MgGraphRequest -Method POST -Uri v1.0/servicePrincipals -Body '{`"appId`":`"$appPurview`"}'"


# ==============================================================================
# FUNÇÕES
# ==============================================================================
# O ErrorDetails do Invoke-MgGraphRequest traz a requisição HTTP inteira; o que interessa é o JSON
# {"error":{"code":..., "message":...}} no final
function Get-DetalheErro($Erro) {
    $texto = "$($Erro.ErrorDetails.Message)"
    $ini   = $texto.IndexOf('{"error"')
    if ($ini -ge 0) {
        try {
            $j = $texto.Substring($ini) | ConvertFrom-Json
            if ($j.error.message) { return "$($j.error.code): $($j.error.message)" }
        } catch {}
    }
    return $Erro.Exception.Message
}

function Write-Orientacao([string]$Erro) {
    Write-Warning "A API de eDiscovery recusou a operação na etapa '$script:etapa': $Erro"
    if ($script:caso) {
        Write-Warning "O caso '$nome' foi criado no Purview e ficou incompleto; pode ser excluído no portal."
    }
    if ($Erro -match 'pay.?as.?you.?go|billing|subscription|PAYG|licen') {
        Write-Warning "O tenant precisa do Purview pay-as-you-go ativado para usar a API de eDiscovery:"
        Write-Warning "  Purview > Configurações > Faturamento > pay-as-you-go > vincular uma assinatura do Azure"
    } elseif ($Erro -match '\bForbidden\b|\b403\b|\bUnauthorized\b|Access ?denied|Authorization_RequestDenied') {
        Write-Warning "A conta precisa estar no grupo de funções 'eDiscovery Manager':"
        Write-Warning "  Purview > Configurações > Funções e escopos > Grupos de funções > eDiscovery Manager"
    }
    Write-Warning "Alternativa manual: $portal > Casos > criar caso > Pesquisa (fonte: $Caixa, sem condições) > Exportar > PST"
}

function Wait-Operacao([string]$Uri, [string]$Descricao, [datetime]$Limite) {
    while ($true) {
        $op = $null
        try {
            $op = Invoke-GraphRetry -Uri $Uri -Headers $cabecalhos
        } catch {
            if ((Get-HttpStatus $_) -ne 404) { throw }   # 404: a operação ainda está sendo criada
        }
        if ($op.status -in 'succeeded', 'partiallySucceeded') { return $op }
        if ($op.status -in 'failed', 'submissionFailed') {
            throw "$Descricao falhou: $($op.resultInfo.message)"
        }
        if ((Get-Date) -gt $Limite) { throw "$Descricao não terminou em $AguardarMinutos minutos (continua no portal: $portal)." }
        $pct = if ($null -ne $op.percentProgress) { " $($op.percentProgress)%" } else { '' }
        $st  = if ($op) { $op.status } else { 'aguardando' }
        Write-Host "  $Descricao em andamento$pct ($st)..." -ForegroundColor DarkGray
        Start-Sleep -Seconds $intervalo
    }
}

function ConvertTo-Base64Url([byte[]]$Bytes) { [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_') }

function Get-ParametrosQuery([string]$Query) {
    $r = @{}
    foreach ($par in $Query.TrimStart('?') -split '&') {
        if (-not $par) { continue }
        $k, $v = $par -split '=', 2
        $r[[uri]::UnescapeDataString($k)] = [uri]::UnescapeDataString(("$v" -replace '\+', ' '))
    }
    return $r
}

function Get-ErroToken($Erro) {
    try { return ($Erro.ErrorDetails.Message | ConvertFrom-Json) } catch { return $null }
}

# Login no navegador (código de autorização + PKCE, resposta recebida em http://localhost),
# o mesmo tipo de login do Connect-MgGraph
function Get-TokenNavegador([string]$TenantId, [string]$Conta) {
    $aleatorio = New-Object byte[] 32
    [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($aleatorio)
    $verificador = ConvertTo-Base64Url $aleatorio
    $desafio     = ConvertTo-Base64Url ([Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::ASCII.GetBytes($verificador)))
    $estado      = [guid]::NewGuid().ToString('N')

    $tcp = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    $tcp.Start(); $porta = $tcp.LocalEndpoint.Port; $tcp.Stop()
    $retorno = "http://localhost:$porta/"

    $params = [ordered]@{
        client_id             = $ClientIdDownload
        response_type         = 'code'
        redirect_uri          = $retorno
        response_mode         = 'query'
        scope                 = $escopoDown
        state                 = $estado
        code_challenge        = $desafio
        code_challenge_method = 'S256'
        login_hint            = $Conta
    }
    $url = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize?" +
           (($params.Keys | ForEach-Object { "$_=$([uri]::EscapeDataString("$($params[$_])"))" }) -join '&')

    $ouvinte = [System.Net.HttpListener]::new()
    $ouvinte.Prefixes.Add($retorno)
    $ouvinte.Start()
    try {
        Write-Host "Abrindo o navegador para o login do download. Se ele não abrir, acesse:" -ForegroundColor Yellow
        Write-Host $url -ForegroundColor DarkGray
        Start-Process $url
        $tarefa = $ouvinte.GetContextAsync()
        $limite = (Get-Date).AddMinutes(5)
        while (-not $tarefa.Wait(1000)) {
            if ((Get-Date) -gt $limite) { throw "Tempo esgotado aguardando o login no navegador." }
        }
        $ctx = $tarefa.Result
        $q   = Get-ParametrosQuery $ctx.Request.Url.Query
        $msg = if ($q.code) { 'Login concluído. Pode fechar esta janela e voltar ao PowerShell.' }
               else { "Falha no login: $([System.Net.WebUtility]::HtmlEncode($q.error_description))" }
        $bytes = [Text.Encoding]::UTF8.GetBytes("<html><head><meta charset='utf-8'></head><body style='font-family:sans-serif'><h3>$msg</h3></body></html>")
        $ctx.Response.ContentType = 'text/html; charset=utf-8'
        $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $ctx.Response.Close()
    } finally {
        $ouvinte.Stop(); $ouvinte.Close()
    }

    if ($q.state -ne $estado) { throw "Resposta de login inválida (state não confere)." }
    if (-not $q.code) {
        if ($q.error_description -match 'AADSTS500011|AADSTS650052') { throw $msgSemSP }
        throw "Login para o download falhou: $($q.error) $($q.error_description)"
    }
    try {
        $t = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body @{
            grant_type    = 'authorization_code'
            client_id     = $ClientIdDownload
            code          = $q.code
            redirect_uri  = $retorno
            code_verifier = $verificador
            scope         = $escopoDown
        }
        return $t.access_token
    } catch {
        $e = Get-ErroToken $_
        throw "Login para o download falhou: $(if ($e) { $e.error_description } else { $_.Exception.Message })"
    }
}

# Login por código de dispositivo (opção -CodigoDispositivo)
function Get-TokenCodigo([string]$TenantId) {
    $base = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0"
    try {
        $dc = Invoke-RestMethod -Method Post -Uri "$base/devicecode" -Body @{ client_id = $ClientIdDownload; scope = $escopoDown }
    } catch {
        if ("$($_.ErrorDetails.Message) $($_.Exception.Message)" -match 'AADSTS500011|AADSTS650052') { throw $msgSemSP }
        throw
    }
    Write-Host "`n$($dc.message)" -ForegroundColor Yellow
    $limite = (Get-Date).AddSeconds([int]$dc.expires_in)
    while ((Get-Date) -lt $limite) {
        Start-Sleep -Seconds ([math]::Max(1, [int]$dc.interval))
        try {
            $t = Invoke-RestMethod -Method Post -Uri "$base/token" -Body @{
                grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id   = $ClientIdDownload
                device_code = $dc.device_code
            }
            return $t.access_token
        } catch {
            $e = Get-ErroToken $_
            if ($e.error -in 'authorization_pending', 'slow_down') { continue }
            throw "Login para o download falhou: $(if ($e) { $e.error_description } else { $_.Exception.Message })"
        }
    }
    throw "O código de login expirou."
}

# Monta a consulta KQL a partir de -Conteudo, -De, -Ate e -Consulta (vazia = caixa inteira)
function Get-ConsultaKql {
    $tipos = @{
        Email      = 'kind:email'
        Calendario = 'kind:meetings'
        Contatos   = 'kind:contacts'
        Tarefas    = 'kind:tasks'
        Notas      = 'kind:notes'
        Teams      = 'kind:microsoftteams'
    }
    $partes = New-Object System.Collections.Generic.List[string]
    if ($Conteudo -notcontains 'Tudo') {
        $partes.Add('(' + (($Conteudo | Select-Object -Unique | ForEach-Object { $tipos[$_] }) -join ' OR ') + ')')
    }
    if ($script:temDe)  { $partes.Add("received>=$($De.ToString('yyyy-MM-dd'))") }
    if ($script:temAte) { $partes.Add("received<=$($Ate.ToString('yyyy-MM-dd'))") }
    if ($Consulta)      { $partes.Add("($Consulta)") }
    return ($partes -join ' AND ')
}

function Get-TokenDownload {
    if ($CodigoDispositivo) { return Get-TokenCodigo $tenantId }
    return Get-TokenNavegador $tenantId $conta
}

# Cria o service principal do serviço de download do Purview (o portal faz isso por baixo dos panos)
function Register-ServicoDownload {
    Write-Host "Registrando o serviço de download do Purview no tenant..." -ForegroundColor Cyan
    Connect-GraphBackup -Escopos 'eDiscovery.ReadWrite.All', 'Application.ReadWrite.All'
    try {
        $null = Invoke-GraphRetry -Method POST -Uri 'v1.0/servicePrincipals' -Body @{ appId = $appPurview }
    } catch {
        $d = Get-DetalheErro $_
        if ($d -notmatch 'already exists|conflict|MultipleObjects') { throw "Não foi possível registrar o serviço de download: $d" }
    }
    Write-Host "Registrado. Aguardando a replicação no Entra ID (1 min)..." -ForegroundColor DarkGray
    Start-Sleep -Seconds 60
}

# Exportação concluída mais recente de um caso, com os links de download
function Get-ExportacaoConcluida([string]$UriCaso) {
    $ops = @((Invoke-GraphRetry -Uri "$UriCaso/operations" -Headers $cabecalhos).value |
        Where-Object { $_.action -eq 'exportResult' -and $_.status -in 'succeeded', 'partiallySucceeded' } |
        Sort-Object { [datetimeoffset]"$($_.createdDateTime)" } -Descending)
    if ($ops.Count -eq 0) { return $null }
    return Invoke-GraphRetry -Uri "$UriCaso/operations/$($ops[0].id)" -Headers $cabecalhos
}


# ==============================================================================
# PASSO 1: CONEXÃO
# ==============================================================================
Connect-GraphBackup -Escopos 'eDiscovery.ReadWrite.All'
$tenantId = (Get-MgContext).TenantId
$conta    = (Get-MgContext).Account
$base     = 'v1.0/security/cases/ediscoveryCases'
$limite   = (Get-Date).AddMinutes($AguardarMinutos)
$prefixo  = "Exportação PST - $Caixa - "

$script:caso        = $null
$script:custodiante = $null
$script:etapa       = ''

# Parâmetro [datetime] não informado vale DateTime.MinValue, por isso checamos se foi passado
$script:temDe  = $PSBoundParameters.ContainsKey('De')
$script:temAte = $PSBoundParameters.ContainsKey('Ate')
if ($script:temDe -and $script:temAte -and $De -gt $Ate) { throw "-De ($($De.ToString('dd/MM/yyyy'))) é posterior a -Ate ($($Ate.ToString('dd/MM/yyyy')))." }
$kql      = Get-ConsultaKql
$selecao  = if ($kql) { "Seleção: $($Conteudo -join ', ')" } else { 'Caixa inteira' }

if ($SomenteDownload -or $CasoId) {
    # ==========================================================================
    # RETOMADA: usa uma exportação já concluída
    # ==========================================================================
    if ($CasoId) {
        $caso = Invoke-GraphRetry -Uri "$base/$CasoId"
    } else {
        $casos = New-Object System.Collections.Generic.List[object]
        $url = $base
        while ($url) {
            $r = Invoke-GraphRetry -Uri $url
            foreach ($c in $r.value) { $casos.Add($c) }
            $url = $r.'@odata.nextLink'
        }
        $caso = $casos | Where-Object { "$($_.displayName)".StartsWith($prefixo) } |
            Sort-Object { [datetimeoffset]"$($_.createdDateTime)" } -Descending | Select-Object -First 1
        if (-not $caso) { throw "Nenhum caso '$prefixo...' encontrado no Purview. Rode sem -SomenteDownload para exportar." }
    }
    $nome    = $caso.displayName
    $uriCaso = "$base/$($caso.id)"
    Write-Host "`nCaso: $nome" -ForegroundColor Cyan
    $export = Get-ExportacaoConcluida $uriCaso
    if (-not $export) { throw "O caso '$nome' não tem exportação concluída. Acompanhe em $portal" }
} else {
    $nome = "$prefixo$(Get-Date -Format 'yyyy-MM-dd HHmm')"

    # ==========================================================================
    # PASSO 2: CASO, PESQUISA E ESTIMATIVA
    # ==========================================================================
    try {
        $script:etapa = 'criar caso'
        Write-Host "`nCriando caso no Purview: $nome" -ForegroundColor Cyan
        Write-Host "Conteúdo: $selecao$(if ($kql) { " | consulta: $kql" })"
        $script:caso = Invoke-GraphRetry -Method POST -Uri $base -Body @{
            displayName = $nome
            description = "Exportação da caixa $Caixa em PST (Exportar-PST.ps1), por $conta. $selecao$(if ($kql) { " ($kql)" })"
        }
        $uriCaso = "$base/$($script:caso.id)"

        # A pesquisa precisa nascer com a fonte (a caixa do colaborador) já vinculada
        $script:etapa = 'criar pesquisa'
        $corpo = @{ displayName = "$selecao - $Caixa" }
        if ($kql) { $corpo.contentQuery = $kql }   # sem consulta = todo o conteúdo
        $pesquisa = $null
        try {
            # Tentativa 1: fonte embutida na própria pesquisa (não cria custodiante nem retenção)
            $pesquisa = Invoke-GraphRetry -Method POST -Uri "$uriCaso/searches" -Body ($corpo + @{
                additionalSources = @(@{ '@odata.type' = '#microsoft.graph.security.userSource'; email = $Caixa; includedSources = 'mailbox' })
            })
        } catch {
            Write-Host "  A API não aceitou a fonte embutida ($(Get-DetalheErro $_)); usando custodiante." -ForegroundColor DarkGray
        }
        if (-not $pesquisa) {
            # Tentativa 2 (caminho documentado): custodiante > fonte só da caixa de correio > pesquisa vinculada
            $script:etapa = 'adicionar custodiante'
            $script:custodiante = (Invoke-GraphRetry -Method POST -Uri "$uriCaso/custodians" -Body @{ email = $Caixa }).id
            $fonte = Invoke-GraphRetry -Method POST -Uri "$uriCaso/custodians/$($script:custodiante)/userSources" -Body @{
                email = $Caixa; includedSources = 'mailbox'
            }
            $script:etapa = 'criar pesquisa'
            $vinculo  = "https://graph.microsoft.com/v1.0/security/cases/ediscoveryCases/$($script:caso.id)/custodians/$($script:custodiante)/userSources/$($fonte.id)"
            $pesquisa = Invoke-GraphRetry -Method POST -Uri "$uriCaso/searches" -Body ($corpo + @{ 'custodianSources@odata.bind' = @($vinculo) })
        }
        $uriPesquisa = "$uriCaso/searches/$($pesquisa.id)"

        $script:etapa = 'estimativa'
        Write-Host "Calculando a estimativa..." -ForegroundColor Cyan
        $null = Invoke-GraphRetry -Method POST -Uri "$uriPesquisa/estimateStatistics"
        $est  = Wait-Operacao "$uriPesquisa/lastEstimateStatisticsOperation" 'Estimativa' $limite
        $gb   = [double]$est.indexedItemsSize / 1GB
        Write-Host ("Estimativa: {0} itens ({1:N2} GB) + {2} itens não indexados" -f $est.indexedItemCount, $gb, $est.unindexedItemCount) -ForegroundColor Green
        if ([long]$est.indexedItemCount -eq 0 -and [long]$est.unindexedItemCount -eq 0) {
            Write-Warning "A pesquisa não encontrou nada. Confira o endereço da caixa e a consulta."
            exit 1
        }


        # ======================================================================
        # PASSO 3: EXPORTAÇÃO EM PST
        # ======================================================================
        $script:etapa = 'exportação'
        Write-Host "`nIniciando a exportação em PST..." -ForegroundColor Cyan
        $null = Invoke-GraphRetry -Method POST -Uri "$uriPesquisa/exportResult" -Headers $cabecalhos -Body @{
            displayName       = "PST - $Caixa"
            # Não indexados (ex: anexos criptografados) não dá para filtrar: com filtro, só entram com -IncluirNaoIndexados
            exportCriteria    = $(if ($kql -and -not $IncluirNaoIndexados) { 'searchHits' } else { 'searchHits, partiallyIndexed' })
            exportLocation    = 'responsiveLocations'
            additionalOptions = 'includeFolderAndPath, splitSource'   # estrutura de pastas original; um PST por caixa
            exportFormat      = 'pst'
        }

        # A operação criada é a exportação mais recente do caso
        $opExport = $null
        for ($t = 0; $t -lt 10 -and -not $opExport; $t++) {
            $ops = (Invoke-GraphRetry -Uri "$uriCaso/operations" -Headers $cabecalhos).value
            $opExport = $ops | Where-Object { $_.action -eq 'exportResult' } |
                Sort-Object { [datetimeoffset]"$($_.createdDateTime)" } -Descending | Select-Object -First 1
            if (-not $opExport) { Start-Sleep -Seconds 5 }
        }
        if (-not $opExport) { throw "A exportação foi solicitada, mas a operação não apareceu no caso. Acompanhe em $portal" }

        $export = Wait-Operacao "$uriCaso/operations/$($opExport.id)" 'Exportação' $limite
    } catch {
        Write-Orientacao (Get-DetalheErro $_)
        exit 1
    } finally {
        # Adicionar custodiante pode colocar a caixa em retenção (hold): libera ao terminar
        if ($script:custodiante) {
            try {
                $null = Invoke-GraphRetry -Method POST -Uri "$uriCaso/custodians/$($script:custodiante)/release"
                Write-Host "Custodiante liberado no caso (sem retenção na caixa)." -ForegroundColor DarkGray
            } catch {
                Write-Warning "Não foi possível liberar o custodiante do caso: libere no portal para retirar a retenção da caixa ($(Get-DetalheErro $_))."
            }
        }
    }
}

$arquivos = @($export.exportFileMetadata | Where-Object { $_.downloadUrl })
if ($arquivos.Count -eq 0) {
    throw "A exportação do caso '$nome' não tem arquivos disponíveis para download (as exportações expiram após alguns dias). Rode sem -SomenteDownload para exportar de novo."
}
Write-Host "Exportação concluída ($($export.status)): $($arquivos.Count) arquivo(s)." -ForegroundColor Green
if ($export.status -eq 'partiallySucceeded') {
    Write-Warning "A exportação terminou parcialmente: confira o relatório no pacote ou no portal."
}
$retomar = ".\Exportar-PST.ps1 -Caixa $Caixa -SomenteDownload"

if ($SemDownload) {
    Write-Host "`nPara baixar: $retomar" -ForegroundColor Yellow
    Write-Host "Ou pelo portal: $portal > Casos > '$nome' > Exportações. Baixe nos próximos dias (disponibilidade limitada)."
    exit 0
}


# ==============================================================================
# PASSO 4: DOWNLOAD
# ==============================================================================
$dirSaida = Join-Path $Destino ($nome -replace '[\\/:*?"<>|]', '_')
New-Item -ItemType Directory $dirSaida -Force | Out-Null

$baixados = New-Object System.Collections.Generic.List[string]
try {
    Write-Host "`nO download exige um segundo login (serviço de arquivos do Purview)." -ForegroundColor Cyan
    try {
        $token = Get-TokenDownload
    } catch {
        if ("$($_.Exception.Message)" -ne $msgSemSP) { throw }
        Write-Warning $msgSemSP
        Write-Warning "É uma configuração única do tenant (o portal do Purview faz o mesmo ao baixar pela primeira vez)."
        $registrar = $RegistrarServicoDownload -or
                     ((Read-Host "Registrar agora? Exige Administrador Global ou de Aplicativos (S/N)").Trim().ToUpper() -eq 'S')
        if (-not $registrar) { throw "Serviço de download não registrado. Para registrar manualmente: $cmdSemSP" }
        Register-ServicoDownload
        $token = Get-TokenDownload
    }
    $header = @{ Authorization = "Bearer $token"; 'X-AllowWithAADToken' = 'true' }

    foreach ($a in $arquivos) {
        $arq = Join-Path $dirSaida ($a.fileName -replace '[\\/:*?"<>|]', '_')
        Write-Host ("  Baixando {0} ({1:N1} MB)..." -f $a.fileName, ([double]$a.size / 1MB))
        $p = @{ Uri = $a.downloadUrl; Headers = $header; OutFile = $arq; ErrorAction = 'Stop' }
        if ($PSVersionTable.PSEdition -ne 'Core') { $p.UseBasicParsing = $true }
        $ProgressPreference = 'SilentlyContinue'   # a barra de progresso deixa downloads grandes muito lentos
        Invoke-WebRequest @p
        $baixados.Add($arq)
    }
} catch {
    Write-Warning "O download falhou: $($_.Exception.Message)"
    Write-Warning "A exportação continua pronta no Purview. Para tentar de novo SEM exportar outra vez:"
    Write-Warning "  $retomar        (ou acrescente -CodigoDispositivo para outro tipo de login)"
    Write-Warning "Ou baixe pelo portal: $portal > Casos > '$nome' > Exportações"
    exit 1
}

# Extrai os .zip para chegar aos PSTs
if (-not $ManterZip) {
    foreach ($z in $baixados | Where-Object { $_ -like '*.zip' }) {
        Expand-Archive -Path $z -DestinationPath ([IO.Path]::Combine($dirSaida, [IO.Path]::GetFileNameWithoutExtension($z))) -Force
        Remove-Item $z
    }
}


# ==============================================================================
# RESUMO
# ==============================================================================
$psts = @(Get-ChildItem $dirSaida -Recurse -Filter *.pst)
Write-Host "`nArquivos em: $dirSaida" -ForegroundColor Green
foreach ($p in $psts) { Write-Host ("  {0} ({1:N1} MB)" -f $p.FullName.Substring($dirSaida.Length + 1), ($p.Length / 1MB)) }
if ($psts.Count -eq 0 -and -not $ManterZip) {
    Write-Warning "Nenhum .pst encontrado no pacote baixado. Confira o conteúdo da pasta e o relatório da exportação."
    exit 1
}
Write-Host "Caso no Purview: '$nome' (fica como registro; pode ser fechado no portal)." -ForegroundColor DarkGray
exit 0
