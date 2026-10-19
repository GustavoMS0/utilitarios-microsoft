<#
.SYNOPSIS
  Auditoria de uso e governança dos sites do SharePoint Online.

.DESCRIPTION
  Usa o relatório oficial de uso (Microsoft Graph) e classifica cada site por
  atividade, ocupação de cota e dono. Opcionalmente (com -SpoAdminUrl) enriquece
  com dados do módulo SharePoint Online: Hub, compartilhamento externo e bloqueio.

.EXAMPLE
  .\Sharepoint.ps1
  .\Sharepoint.ps1 -Periodo D90 -DiasInativo 180
  .\Sharepoint.ps1 -SpoAdminUrl https://tenant-admin.sharepoint.com -IncluirOneDrive

.EXAMPLE
  # Revela temporariamente os nomes nos relatórios do tenant (pede confirmação; o valor original
  # é restaurado no final)
  .\Sharepoint.ps1 -RevelarNomes
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [ValidateSet('D7','D30','D90','D180')]
    [string]$Periodo = 'D30',

    # Sites sem atividade há mais dias que isso são marcados como "Inativo"
    [int]$DiasInativo = 90,

    # Percentual de cota a partir do qual o site gera alerta
    [int]$AlertaCotaPct = 90,

    # URL do admin center do SharePoint (ex: https://tenant-admin.sharepoint.com).
    # Se informado, traz Hub, compartilhamento externo e bloqueio via módulo SPO.
    [string]$SpoAdminUrl,

    # Gera também o relatório de uso do OneDrive (por usuário)
    [switch]$IncluirOneDrive,

    # Desativa no tenant a ocultação de nomes nos relatórios durante a execução (altera configuração
    # global! pede confirmação e restaura o valor original no final)
    [switch]$RevelarNomes,

    # Com -RevelarNomes: não restaura a ocultação no final (a mudança pode levar alguns minutos
    # para refletir nos relatórios)
    [switch]$ManterNomesRevelados,

    # Pasta de saída; se for criada pelo script, só o usuário atual, SYSTEM e Administradores têm acesso
    [string]$Out = "C:\Temp\SP-Auditoria",

    # Emite os registros de auditoria dos sites para o pipeline
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
# -WhatIf vale só para a alteração no tenant ($PSCmdlet.ShouldProcess); o relatório é gerado normalmente
$WhatIfPreference = $false
. (Join-Path $PSScriptRoot 'lib/GraphMailbox.ps1')
New-PastaSegura $Out
$stamp = Get-Date -Format 'yyyyMMdd-HHmm'

# Limpeza garantida no finally
$brutos             = New-Object System.Collections.Generic.List[string]
$restaurarOcultacao = $false
$spoConectado       = $false

# ==============================================================================
# PASSO 0: MÓDULOS E CONEXÃO
# ==============================================================================
Write-Host "Verificando pré-requisitos..." -ForegroundColor Cyan

$scopes = @('Reports.Read.All', 'Sites.Read.All', 'ReportSettings.Read.All')
if ($RevelarNomes) { $scopes += 'ReportSettings.ReadWrite.All' }
Connect-GraphBackup -Escopos $scopes


try {
    # ==============================================================================
    # PASSO 1: OCULTAÇÃO DE NOMES NOS RELATÓRIOS
    # ==============================================================================
    # Desde 2021 o M365 oculta por padrão URLs/nomes nos relatórios. Com isso a coluna
    # "Site URL" vem vazia e a auditoria fica inútil.
    $ocultos = $null
    try {
        $ocultos = (Invoke-GraphRetry -Uri 'v1.0/admin/reportSettings').displayConcealedNames
    } catch {
        Write-Warning "Não foi possível ler a configuração de ocultação de relatórios: $($_.Exception.Message)"
    }

    if ($ocultos) {
        if ($RevelarNomes) {
            if ($PSCmdlet.ShouldProcess('Configurações de relatórios do tenant (todos os relatórios do Microsoft 365)', 'Desativar a ocultação de nomes (displayConcealedNames = false)')) {
                $null = Invoke-GraphRetry -Method PATCH -Uri 'v1.0/admin/reportSettings' -Body @{ displayConcealedNames = $false }
                $restaurarOcultacao = -not $ManterNomesRevelados
                Write-Host "Ocultação de nomes desativada no tenant. Pode levar alguns minutos para refletir nos relatórios." -ForegroundColor Yellow
                if ($restaurarOcultacao) { Write-Host "O valor original será restaurado no final (use -ManterNomesRevelados para manter)." -ForegroundColor DarkGray }
            }
        } else {
            Write-Warning "Os relatórios do tenant estão com nomes OCULTOS: as URLs dos sites virão vazias."
            Write-Warning "Desative em: Admin Center > Configurações > Configurações da organização > Relatórios,"
            Write-Warning "ou rode novamente com -RevelarNomes."
        }
    }


    # ==============================================================================
    # PASSO 2: RELATÓRIO DE USO (GRAPH)
    # ==============================================================================
    Write-Host "`nBaixando relatório de uso do SharePoint ($Periodo)..." -ForegroundColor Cyan
    $csvUso = Join-Path $Out "uso-bruto-$stamp.csv"
    $brutos.Add($csvUso)
    Invoke-MgGraphRequest -Method GET -Uri "v1.0/reports/getSharePointSiteUsageDetail(period='$Periodo')" -OutputFilePath $csvUso
    $uso = @(Import-Csv $csvUso)

    $excluidos = @($uso | Where-Object { $_.'Is Deleted' -eq 'True' })
    $uso       = @($uso | Where-Object { $_.'Is Deleted' -ne 'True' })
    $semUrl    = @($uso | Where-Object { -not $_.'Site URL' })

    Write-Host "Sites no relatório: $($uso.Count) ativos, $($excluidos.Count) excluídos (ignorados)." -ForegroundColor Green
    if ($semUrl.Count -gt 0) {
        Write-Warning "$($semUrl.Count) site(s) sem URL (nomes ocultos). Eles aparecerão identificados apenas pelo Site Id."
    }


    # ==============================================================================
    # PASSO 3: DADOS DO SHAREPOINT ONLINE (OPCIONAL)
    # ==============================================================================
    $spoPorUrl = @{}
    $hubPorId  = @{}

    if ($SpoAdminUrl) {
        Write-Host "`nColetando dados do SharePoint Online ($SpoAdminUrl)..." -ForegroundColor Cyan
        try {
            Initialize-Modulo -Nome Microsoft.Online.SharePoint.PowerShell -VersaoMinima 16.0.0 -SemImportar
            # No PowerShell 7 o módulo SPO só funciona de forma confiável via compatibilidade com o Windows PowerShell
            if ($PSVersionTable.PSEdition -eq 'Core') {
                Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -WarningAction SilentlyContinue
            } else {
                Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
            }
            Connect-SPOService -Url $SpoAdminUrl
            $spoConectado = $true

            foreach ($h in Get-SPOHubSite) { $hubPorId["$($h.ID)"] = $h.Title }
            foreach ($s in Get-SPOSite -Limit All) { $spoPorUrl[$s.Url.TrimEnd('/').ToLower()] = $s }

            Write-Host "SPO: $($spoPorUrl.Count) sites, $($hubPorId.Count) hubs." -ForegroundColor Green
        } catch {
            Write-Warning "Falha ao coletar dados do SPO, seguindo só com o Graph: $($_.Exception.Message)"
        }
    }


    # ==============================================================================
    # PASSO 4: MONTAGEM DA AUDITORIA
    # ==============================================================================
    $tiposTemplate = @{
        'GROUP#0'              = 'Teams / Grupo M365'
        'SITEPAGEPUBLISHING#0' = 'Comunicação'
        'STS#3'                = 'Equipe (sem grupo)'
        'STS#0'                = 'Equipe (clássico)'
        'TEAMCHANNEL#0'        = 'Canal privado Teams'
        'TEAMCHANNEL#1'        = 'Canal compartilhado Teams'
        'APPCATALOG#0'         = 'Catálogo de apps'
        'SRCHCEN#0'            = 'Central de pesquisa'
        'SPSMSITEHOST#0'       = 'Host de OneDrive'
        'EHS#1'                = 'Equipe (raiz)'
        'BLANKINTERNET#0'      = 'Publicação (clássico)'
    }

    $compartilhamento = @{
        'Disabled'                        = 'Somente interno'
        'ExistingExternalUserSharingOnly' = 'Convidados existentes'
        'ExternalUserSharingOnly'         = 'Convidados novos e existentes'
        'ExternalUserAndGuestSharing'     = 'Qualquer pessoa (link anônimo)'
    }

    function ConvertTo-Long([string]$v) { if ($v) { [long]$v } else { 0 } }

    function Get-SubsitesDoSite([string]$SiteId, [string]$SitePaiUrl, [string]$SitePaiTitulo) {
        $lista = New-Object System.Collections.Generic.List[object]
        try {
            $resp = Invoke-GraphRetry -Uri "v1.0/sites/$($SiteId)/sites?`$select=id,displayName,webUrl,createdDateTime"
            $filhos = if ($resp.value) { $resp.value } else { @($resp) }
            foreach ($sub in $filhos) {
                if (-not $sub -or -not $sub.id) { continue }
                $lista.Add([pscustomobject]@{
                    TituloSubsite = $sub.displayName
                    UrlSubsite    = $sub.webUrl
                    SitePaiTitulo = $SitePaiTitulo
                    SitePaiUrl    = $SitePaiUrl
                    CriadoEm      = $sub.createdDateTime
                })
                $netos = Get-SubsitesDoSite -SiteId $sub.id -SitePaiUrl $sub.webUrl -SitePaiTitulo $sub.displayName
                foreach ($n in $netos) { $lista.Add($n) }
            }
        } catch {
            Write-Verbose "Sem subsites ou erro ao buscar subsites de $($SiteId): $($_.Exception.Message)"
        }
        return $lista
    }

    $hoje          = Get-Date
    $resultado     = New-Object System.Collections.Generic.List[object]
    $todosSubsites = New-Object System.Collections.Generic.List[object]
    $i = 0

    foreach ($u in $uso) {
        $i++
        $url = $u.'Site URL'
        Write-Progress -Activity "Processando sites" -Status "$i/$($uso.Count)" -PercentComplete (100 * $i / [math]::Max($uso.Count, 1))

        # Recupera URL e Título reais via Graph caso venham ocultos no relatório
        $titulo = $null
        if (-not $url -and $u.'Site Id') {
            try {
                $siteInfo = Invoke-GraphRetry -Uri "v1.0/sites/$($u.'Site Id')?`$select=displayName,webUrl"
                if ($siteInfo.webUrl)      { $url = $siteInfo.webUrl }
                if ($siteInfo.displayName) { $titulo = $siteInfo.displayName }
            } catch {
                Write-Verbose "Não foi possível resolver site por ID ($($u.'Site Id')): $($_.Exception.Message)"
            }
        }

        $spo = if ($url) { $spoPorUrl[$url.TrimEnd('/').ToLower()] }

        # Título real: SPO > Graph > último trecho da URL
        if ($spo -and -not $titulo) {
            $titulo = $spo.Title
        } elseif ($url -and -not $titulo) {
            try {
                $uri  = [uri]$url
                $path = $uri.AbsolutePath.TrimEnd('/')
                $graphUri = if ($path) { "v1.0/sites/$($uri.Host):$path" } else { "v1.0/sites/$($uri.Host)" }
                $titulo = (Invoke-GraphRetry -Uri "$($graphUri)?`$select=displayName").displayName
            } catch {
                Write-Warning "Título do site não obtido pelo Graph ($url), usando o nome da URL: $($_.Exception.Message)"
            }
        }
        if (-not $titulo) { $titulo = if ($url) { ([uri]$url).Segments[-1].Trim('/') } else { $u.'Site Id' } }
        if (-not $titulo) { $titulo = '(raiz)' }

        # Coleta de subsites via Graph
        $subsites = @()
        if ($u.'Site Id') {
            $subsites = Get-SubsitesDoSite -SiteId $u.'Site Id' -SitePaiUrl $url -SitePaiTitulo $titulo
            foreach ($s in $subsites) { $todosSubsites.Add($s) }
        }
        $resumoSubsites = if ($subsites.Count -gt 0) {
            ($subsites | ForEach-Object { "$($_.TituloSubsite) ($($_.UrlSubsite))" }) -join '; '
        } else {
            ''
        }

        # Atividade
        $ultimaAtividade = $null
        $diasSemAtividade = $null
        if ($u.'Last Activity Date') {
            $ultimaAtividade  = [datetime]::ParseExact($u.'Last Activity Date', 'yyyy-MM-dd', $null)
            $diasSemAtividade = [int]($hoje - $ultimaAtividade).TotalDays
        }
        $status = if ($null -eq $diasSemAtividade)          { 'Nunca usado' }
                  elseif ($diasSemAtividade -gt $DiasInativo) { 'Inativo' }
                  elseif ($diasSemAtividade -gt 30)           { 'Pouco usado' }
                  else                                        { 'Ativo' }

        # Armazenamento
        $usadoB  = ConvertTo-Long $u.'Storage Used (Byte)'
        $cotaB   = ConvertTo-Long $u.'Storage Allocated (Byte)'
        $pctCota = if ($cotaB -gt 0) { [math]::Round(100 * $usadoB / $cotaB, 1) } else { $null }

        $arquivos       = ConvertTo-Long $u.'File Count'
        $arquivosAtivos = ConvertTo-Long $u.'Active File Count'

        $template = $u.'Root Web Template'
        $tipo     = if ($tiposTemplate[$template]) { $tiposTemplate[$template] } elseif ($template) { $template } else { '' }

        $sharing  = if ($spo) { "$($spo.SharingCapability)" }
        $hub      = if ($spo -and $spo.HubSiteId -and "$($spo.HubSiteId)" -ne [guid]::Empty.ToString()) { $hubPorId["$($spo.HubSiteId)"] }

        # Alertas de governança
        $alertas = New-Object System.Collections.Generic.List[string]
        if (-not $u.'Owner Principal Name' -and -not $u.'Owner Display Name') { $alertas.Add('Sem dono') }
        if ($status -in 'Inativo', 'Nunca usado')                              { $alertas.Add($status) }
        if ($null -ne $pctCota -and $pctCota -ge $AlertaCotaPct)              { $alertas.Add("Cota $pctCota%") }
        if ($sharing -eq 'ExternalUserAndGuestSharing')                       { $alertas.Add('Link anônimo permitido') }
        if ($spo -and $spo.LockState -and "$($spo.LockState)" -ne 'Unlock')   { $alertas.Add("Bloqueado ($($spo.LockState))") }

        $resultado.Add([pscustomobject]@{
            Titulo            = $titulo
            Url               = if ($url) { $url } else { "(oculto) $($u.'Site Id')" }
            QtdSubsites       = $subsites.Count
            Subsites          = $resumoSubsites
            Tipo              = $tipo
            Hub               = $hub
            Dono              = $u.'Owner Display Name'
            DonoUPN           = $u.'Owner Principal Name'
            Status            = $status
            UltimaAtividade   = if ($ultimaAtividade) { $ultimaAtividade.ToString('dd/MM/yyyy') } else { '' }
            DiasSemAtividade  = $diasSemAtividade
            Arquivos          = $arquivos
            ArquivosAtivos    = $arquivosAtivos
            PctArquivosAtivos = if ($arquivos -gt 0) { [math]::Round(100 * $arquivosAtivos / $arquivos, 1) } else { 0 }
            PaginasVistas     = ConvertTo-Long $u.'Page View Count'
            PaginasVisitadas  = ConvertTo-Long $u.'Visited Page Count'
            UsadoGB           = [math]::Round($usadoB / 1GB, 2)
            CotaGB            = [math]::Round($cotaB / 1GB, 2)
            PctCota           = $pctCota
            Compartilhamento  = if ($sharing) { $compartilhamento[$sharing] } else { '' }
            Alertas           = $alertas -join '; '
        })
    }
    Write-Progress -Activity "Processando sites" -Completed

    $csvSaida = Join-Path $Out "SharePoint-Auditoria-$stamp.csv"
    $resultado | Sort-Object UsadoGB -Descending | Export-Csv $csvSaida -Delimiter ';' -NoTypeInformation -Encoding UTF8

    if ($todosSubsites.Count -gt 0) {
        $csvSubsitesSaida = Join-Path $Out "SharePoint-Subsites-$stamp.csv"
        $todosSubsites | Export-Csv $csvSubsitesSaida -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-Host "Subsites: $($todosSubsites.Count) encontrados -> $csvSubsitesSaida" -ForegroundColor Green
    }


    # ==============================================================================
    # PASSO 5: ONEDRIVE (OPCIONAL)
    # ==============================================================================
    if ($IncluirOneDrive) {
        Write-Host "`nBaixando relatório de uso do OneDrive ($Periodo)..." -ForegroundColor Cyan
        $csvOd = Join-Path $Out "od-bruto-$stamp.csv"
        $brutos.Add($csvOd)
        Invoke-MgGraphRequest -Method GET -Uri "v1.0/reports/getOneDriveUsageAccountDetail(period='$Periodo')" -OutputFilePath $csvOd

        $od = Import-Csv $csvOd | Where-Object { $_.'Is Deleted' -ne 'True' } | ForEach-Object {
            $dias = if ($_.'Last Activity Date') { [int]($hoje - [datetime]::ParseExact($_.'Last Activity Date', 'yyyy-MM-dd', $null)).TotalDays }
            [pscustomobject]@{
                Usuario          = $_.'Owner Display Name'
                UPN              = $_.'Owner Principal Name'
                Url              = $_.'Site URL'
                UltimaAtividade  = $_.'Last Activity Date'
                DiasSemAtividade = $dias
                Arquivos         = ConvertTo-Long $_.'File Count'
                ArquivosAtivos   = ConvertTo-Long $_.'Active File Count'
                UsadoGB          = [math]::Round((ConvertTo-Long $_.'Storage Used (Byte)') / 1GB, 2)
                CotaGB           = [math]::Round((ConvertTo-Long $_.'Storage Allocated (Byte)') / 1GB, 2)
            }
        }
        $csvOdSaida = Join-Path $Out "OneDrive-Uso-$stamp.csv"
        $od | Sort-Object UsadoGB -Descending | Export-Csv $csvOdSaida -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-Host "OneDrive: $(@($od).Count) contas -> $csvOdSaida" -ForegroundColor Green
    }


    # ==============================================================================
    # PASSO 6: RESUMO
    # ==============================================================================
    Write-Host "`n--- RESUMO ---" -ForegroundColor Yellow
    Write-Host ("Sites analisados:   {0}" -f $resultado.Count)
    $resultado | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Host ("  {0,-12} {1}" -f $_.Name, $_.Count) }
    Write-Host ("Armazenamento total: {0:N2} GB" -f ($resultado | Measure-Object UsadoGB -Sum).Sum)
    Write-Host ("Sites sem dono:      {0}" -f @($resultado | Where-Object Alertas -match 'Sem dono').Count)
    Write-Host ("Cota >= {0}%:         {1}" -f $AlertaCotaPct, @($resultado | Where-Object Alertas -match 'Cota').Count)
    if ($SpoAdminUrl) {
        Write-Host ("Link anônimo:        {0}" -f @($resultado | Where-Object Alertas -match 'anônimo').Count)
    }
    Write-Host ("Subsites encontrados:{0}" -f $todosSubsites.Count)

    Write-Host "`nTop 10 por armazenamento:" -ForegroundColor Cyan
    $resultado | Sort-Object UsadoGB -Descending | Select-Object -First 10 Titulo, Status, UsadoGB, PctCota | Format-Table -AutoSize | Out-Host

    Write-Host "Relatório gerado em: $csvSaida" -ForegroundColor Green
    if ($PassThru) {
        $resultado | Write-Output
    }
} finally {
    # Relatórios brutos (com UPNs de todos os donos) não ficam no disco
    foreach ($bruto in $brutos) { Remove-Item -LiteralPath $bruto -Force -ErrorAction SilentlyContinue }

    # Devolve a configuração de ocultação ao valor original
    if ($restaurarOcultacao) {
        try {
            $null = Invoke-GraphRetry -Method PATCH -Uri v1.0/admin/reportSettings -Body @{ displayConcealedNames = $true }
            Write-Host "Ocultação de nomes nos relatórios restaurada no tenant." -ForegroundColor Yellow
        } catch {
            Write-Warning "Não foi possível restaurar a ocultação de nomes: $($_.Exception.Message). Reative em Admin Center > Configurações > Configurações da organização > Relatórios."
        }
    }
    if ($spoConectado) {
        try { Disconnect-SPOService -ErrorAction Stop } catch { Write-Verbose "Falha ao desconectar do SPO: $($_.Exception.Message)" }
    }
    Disconnect-GraphSeConectado
}
