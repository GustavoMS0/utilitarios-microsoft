<#
.SYNOPSIS
  Utilitário avançado de manutenção, diagnóstico e liberação de espaço em disco no Windows.

.DESCRIPTION
  Executa varredura profunda e remediação automatizada para cenários críticos de esgotamento
  de disco (disco 100% cheio) em estações de trabalho e servidores Windows.
  
  Módulos de limpeza e otimização:
    1. Arquivos Temporários de Sistema (C:\Windows\Temp) e de Todos os Perfis de Usuário.
    2. Cache de Atualizações do Windows (SoftwareDistribution\Download e DeliveryOptimization).
    3. Limpeza do Component Store (WinSxS / DISM Component Cleanup).
    4. Logs Corrompidos ou Gigantes do Windows (CBS.log, CbsPersist_*.log, DISM.log, Panther).
    5. Relatórios de Erro e Dumps de Memória (WER, Minidump, MEMORY.DMP, CrashDumps).
    6. Cópias de Sombra de Volume (VSS / Shadow Copies acumuladas).
    7. Arquivo de Hibernação (hiberfil.sys - opcional/Deep para liberar RAM equivalente).
    8. Lixeiras de Todos os Usuários e Unidades (Clear-RecycleBin).
    9. Caches de Aplicativos Corporativos (Microsoft Teams clássico e novo, Edge, Chrome, Outlook).
   10. Diagnóstico dos 20 maiores arquivos e maiores diretórios da unidade C:.

.PARAMETER Mode
  Modo de execução:
    - Safe: Limpeza recomendada e segura sem impacto operacional (padrão).
    - Deep: Limpeza agressiva (inclui WinSxS ResetBase, ajuste de Shadow Copies e logs antigos).
    - Analyze: Apenas audita e lista os maiores consumidores sem apagar nada.

.PARAMETER DisableHibernation
  Desativa a hibernação (powercfg /h off), eliminando o hiberfil.sys e liberando o equivalente
  à memória RAM instalada (ex: 8GB a 64GB). Recomendado para desktops fixos e VMs.

.PARAMETER PurgeShadowCopies
  Reduz a área reservada de cópias de sombra (VSS) para 5GB e purga sombras antigas.

.PARAMETER SkipDism
  Pula a limpeza do WinSxS via DISM (útil se precisar de liberação de emergência em segundos).

.PARAMETER TopFilesCount
  Quantidade de arquivos gigantes a listar no relatório final (padrão: 20).

.EXAMPLE
  .\Otimizar-DiscoWindows.ps1
  Executa a limpeza padrão (Safe) com relatório de espaço recuperado.

.EXAMPLE
  .\Otimizar-DiscoWindows.ps1 -Mode Deep -DisableHibernation -PurgeShadowCopies
  Limpeza profunda e agressiva para máquinas com disco crítico/lotado.

.EXAMPLE
  .\Otimizar-DiscoWindows.ps1 -Mode Analyze
  Apenas audita e gera diagnóstico dos maiores arquivos e pastas sem deletar nada.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('Safe', 'Deep', 'Analyze')]
    [string]$Mode = 'Safe',

    [switch]$DisableHibernation,
    [switch]$PurgeShadowCopies,
    [switch]$SkipDism,
    [int]$TopFilesCount = 20
)

$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# -----------------------------------------------------------------------------
# 1. Elevação de Privilégio (Administrador para limpeza; Análise roda direto)
# -----------------------------------------------------------------------------
$ehAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $ehAdmin) {
    if ($Mode -eq 'Analyze') {
        Write-Host "[MODO ANÁLISE] Executando no terminal atual sem elevação." -ForegroundColor Yellow
        Write-Host "Nota: Apenas leitura de disco. Para limpeza de pastas protegidas, execute como Administrador.`n" -ForegroundColor Gray
    } else {
        if ($Host.Name -match 'ConsoleHost') {
            Write-Warning "Este script requer privilégios de Administrador para efetuar limpezas no sistema."
            Write-Host "DICA: Você pode abrir o PowerShell diretamente como Administrador ou confirmar a janela UAC.`n" -ForegroundColor Yellow
            $argumentos = @('-NoExit', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
            foreach ($k in $PSBoundParameters.Keys) {
                if ($PSBoundParameters[$k] -is [switch] -and $PSBoundParameters[$k].IsPresent) {
                    $argumentos += "-$k"
                } elseif ($PSBoundParameters[$k] -isnot [switch]) {
                    $argumentos += "-$k"; $argumentos += "`"$($PSBoundParameters[$k])`""
                }
            }
            try {
                Start-Process powershell.exe -ArgumentList $argumentos -Verb RunAs -ErrorAction Stop
                Write-Host "Janela elevada iniciada. Ela permanecerá aberta até você pressionar Enter ao final." -ForegroundColor Green
                exit 0
            } catch {
                Write-Error "Não foi possível abrir janela elevada automaticamente: $($_.Exception.Message)"
                Write-Host "Para executar, abra o PowerShell como Administrador e execute: `"$PSCommandPath`"" -ForegroundColor Yellow
                exit 1
            }
        } else {
            throw "Execute este script em uma janela do PowerShell como Administrador."
        }
    }
}

# -----------------------------------------------------------------------------
# 2. Funções Auxiliares e Métricas
# -----------------------------------------------------------------------------
function Format-Tamanho([double]$Bytes) {
    if ($Bytes -ge 1TB) { return ("{0:N2} TB" -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    return ("{0:N0} B" -f $Bytes)
}

function Get-EspacoDisco([string]$DriveLetter = 'C') {
    $drive = Get-PSDrive -Name $DriveLetter -ErrorAction SilentlyContinue
    if (-not $drive) { return $null }
    return [PSCustomObject]@{
        Drive      = $DriveLetter
        TotalBytes = ($drive.Used + $drive.Free)
        UsedBytes  = $drive.Used
        FreeBytes  = $drive.Free
        FreeGB     = [math]::Round($drive.Free / 1GB, 2)
        UsedGB     = [math]::Round($drive.Used / 1GB, 2)
        TotalGB    = [math]::Round(($drive.Used + $drive.Free) / 1GB, 2)
        PctFree    = [math]::Round(($drive.Free / ($drive.Used + $drive.Free)) * 100, 1)
    }
}

function Remove-ItemSeguro([string]$Path, [switch]$Recurse, [int]$DiasMinimos = 0) {
    if (-not (Test-Path -LiteralPath $Path)) { return 0 }
    $bytesLiberados = 0

    try {
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            $f = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
            if ($f -and ($DiasMinimos -eq 0 -or $f.LastWriteTime -lt (Get-Date).AddDays(-$DiasMinimos))) {
                $tam = $f.Length
                Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
                return $tam
            }
            return 0
        }

        # Diretório
        $itens = Get-ChildItem -LiteralPath $Path -Force -Recurse:$Recurse -ErrorAction SilentlyContinue
        foreach ($item in $itens) {
            if (-not $item.PSIsContainer) {
                if ($DiasMinimos -eq 0 -or $item.LastWriteTime -lt (Get-Date).AddDays(-$DiasMinimos)) {
                    $tam = $item.Length
                    try {
                        Remove-Item -LiteralPath $item.FullName -Force -ErrorAction Stop
                        $bytesLiberados += $tam
                    } catch {
                        # Arquivo em uso / travado pelo SO
                    }
                }
            } elseif ($Recurse) {
                try {
                    if ((Get-ChildItem -LiteralPath $item.FullName -Force -ErrorAction SilentlyContinue).Count -eq 0) {
                        Remove-Item -LiteralPath $item.FullName -Force -Recurse -ErrorAction SilentlyContinue
                    }
                } catch {}
            }
        }
    } catch {}

    return $bytesLiberados
}

# -----------------------------------------------------------------------------
# 3. Cabeçalho e Diagnóstico Inicial
# -----------------------------------------------------------------------------
Clear-Host
Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "         OTIMIZAÇÃO E MANUTENÇÃO AVANÇADA DE DISCO - WINDOWS                  " -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "Modo de Operação : $Mode" -ForegroundColor Yellow
Write-Host "Máquina Alvo     : $env:COMPUTERNAME ($([System.Environment]::OSVersion.VersionString))" -ForegroundColor Gray
Write-Host "Data/Hora        : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray

$discoAntes = Get-EspacoDisco 'C'
Write-Host "`n[+] Status Atual da Unidade C:" -ForegroundColor Cyan
Write-Host "    Espaço Total : $($discoAntes.TotalGB) GB"
Write-Host "    Espaço Usado : $($discoAntes.UsedGB) GB ($([math]::Round(100 - $discoAntes.PctFree, 1))%)"
Write-Host "    Espaço Livre : $($discoAntes.FreeGB) GB ($($discoAntes.PctFree)% livre)" -ForegroundColor $(if ($discoAntes.FreeGB -lt 5) { 'Red' } elseif ($discoAntes.FreeGB -lt 15) { 'Yellow' } else { 'Green' })

if ($discoAntes.FreeGB -lt 5) {
    Write-Host "    [ALERTA CRÍTICO] Menos de 5 GB disponíveis. Sistema sob risco de travamento." -ForegroundColor Red
}

if ($Mode -eq 'Analyze') {
    Write-Host "`n[MODO ANÁLISE] Nenhuma exclusão será efetuada. Coletando consumidores..." -ForegroundColor Yellow
}

# -----------------------------------------------------------------------------
# 4. Execução dos Módulos de Otimização
# -----------------------------------------------------------------------------
$resumoLiberado = [ordered]@{}

if ($Mode -ne 'Analyze') {

    # MÓDULO 1: Lixeira
    Write-Host "`n[1/9] Esvaziando Lixeira de todas as unidades..." -ForegroundColor White
    $antesLixeira = (Get-EspacoDisco 'C').FreeBytes
    try {
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue | Out-Null
    } catch {}
    $resumoLiberado['Lixeiras'] = [math]::Max(0, (Get-EspacoDisco 'C').FreeBytes - $antesLixeira)
    Write-Host "      Liberado: $(Format-Tamanho $resumoLiberado['Lixeiras'])" -ForegroundColor Green

    # MÓDULO 2: Arquivos Temporários de Sistema e Usuários
    Write-Host "`n[2/9] Limpando pastas temporárias (Sistema e todos os perfis)..." -ForegroundColor White
    $liberadoTemp = 0

    # Temp do Sistema
    $liberadoTemp += Remove-ItemSeguro "C:\Windows\Temp" -Recurse
    $liberadoTemp += Remove-ItemSeguro "$env:SystemRoot\Prefetch" -Recurse -DiasMinimos 14

    # Temp de todos os perfis de usuários em C:\Users
    $perfis = Get-ChildItem "C:\Users" -Directory -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -notin @('Public', 'Default', 'Default User', 'All Users') }
    foreach ($p in $perfis) {
        $userTemp = Join-Path $p.FullName "AppData\Local\Temp"
        $userINet = Join-Path $p.FullName "AppData\Local\Microsoft\Windows\INetCache"
        $userHistory = Join-Path $p.FullName "AppData\Local\Microsoft\Windows\History"

        $liberadoTemp += Remove-ItemSeguro $userTemp -Recurse
        $liberadoTemp += Remove-ItemSeguro $userINet -Recurse
        $liberadoTemp += Remove-ItemSeguro $userHistory -Recurse
    }
    $resumoLiberado['Temporários'] = $liberadoTemp
    Write-Host "      Liberado: $(Format-Tamanho $liberadoTemp)" -ForegroundColor Green

    # MÓDULO 3: Relatórios de Erro e Dumps (WER / Minidump / MEMORY.DMP)
    Write-Host "`n[3/9] Limpando relatórios de erro e despejos de memória (Dumps/WER)..." -ForegroundColor White
    $liberadoWer = 0
    $liberadoWer += Remove-ItemSeguro "C:\Windows\MEMORY.DMP"
    $liberadoWer += Remove-ItemSeguro "C:\Windows\Minidump" -Recurse
    $liberadoWer += Remove-ItemSeguro "C:\ProgramData\Microsoft\Windows\WER\ReportArchive" -Recurse
    $liberadoWer += Remove-ItemSeguro "C:\ProgramData\Microsoft\Windows\WER\ReportQueue" -Recurse
    $liberadoWer += Remove-ItemSeguro "C:\ProgramData\Microsoft\Windows\WER\Temp" -Recurse

    foreach ($p in $perfis) {
        $crash = Join-Path $p.FullName "AppData\Local\CrashDumps"
        $liberadoWer += Remove-ItemSeguro $crash -Recurse
    }
    $resumoLiberado['Dumps_e_WER'] = $liberadoWer
    Write-Host "      Liberado: $(Format-Tamanho $liberadoWer)" -ForegroundColor Green

    # MÓDULO 4: Logs Gigantes e CBS Corrompidos
    Write-Host "`n[4/9] Purgando logs corrompidos ou acumulados (CBS / DISM / Panther)..." -ForegroundColor White
    $liberadoLogs = 0
    $liberadoLogs += Remove-ItemSeguro "C:\Windows\Logs\CBS\CbsPersist_*.log"
    $liberadoLogs += Remove-ItemSeguro "C:\Windows\Logs\CBS\CbsPersist_*.cab"
    $liberadoLogs += Remove-ItemSeguro "C:\Windows\Panther" -Recurse -DiasMinimos 30

    # Trunca CBS.log se for anômalo (> 200MB)
    $cbsFile = "C:\Windows\Logs\CBS\CBS.log"
    if (Test-Path $cbsFile) {
        $itemCbs = Get-Item $cbsFile -ErrorAction SilentlyContinue
        if ($itemCbs -and $itemCbs.Length -gt 200MB) {
            Write-Host "      Truncando CBS.log anômalo ($([math]::Round($itemCbs.Length/1MB,1)) MB)..." -ForegroundColor Yellow
            $tamOriginal = $itemCbs.Length
            try {
                Stop-Service -Name 'TrustedInstaller' -Force -ErrorAction SilentlyContinue
                Set-Content -Path $cbsFile -Value "" -Force -ErrorAction SilentlyContinue
                $liberadoLogs += $tamOriginal
            } catch {}
        }
    }
    $resumoLiberado['Logs_CBS_DISM'] = $liberadoLogs
    Write-Host "      Liberado: $(Format-Tamanho $liberadoLogs)" -ForegroundColor Green

    # MÓDULO 5: Cache do Windows Update e Otimização de Entrega
    Write-Host "`n[5/9] Limpando backlog do Windows Update (SoftwareDistribution)..." -ForegroundColor White
    $liberadoUpdate = 0
    try {
        Stop-Service -Name 'wuauserv' -Force -ErrorAction SilentlyContinue
        Stop-Service -Name 'bits' -Force -ErrorAction SilentlyContinue
        Stop-Service -Name 'dosvc' -Force -ErrorAction SilentlyContinue

        $liberadoUpdate += Remove-ItemSeguro "C:\Windows\SoftwareDistribution\Download" -Recurse
        $liberadoUpdate += Remove-ItemSeguro "C:\Windows\SoftwareDistribution\DeliveryOptimization" -Recurse
        $liberadoUpdate += Remove-ItemSeguro "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache" -Recurse
    } finally {
        Start-Service -Name 'bits' -ErrorAction SilentlyContinue
        Start-Service -Name 'wuauserv' -ErrorAction SilentlyContinue
    }
    $resumoLiberado['WindowsUpdate'] = $liberadoUpdate
    Write-Host "      Liberado: $(Format-Tamanho $liberadoUpdate)" -ForegroundColor Green

    # MÓDULO 6: Caches Corporativos (Teams, Browsers e Office)
    Write-Host "`n[6/9] Limpando caches de navegadores e Teams de todos os usuários..." -ForegroundColor White
    $liberadoAppCaches = 0
    foreach ($p in $perfis) {
        # Teams Clássico e Novo
        $liberadoAppCaches += Remove-ItemSeguro (Join-Path $p.FullName "AppData\Roaming\Microsoft\Teams\Cache") -Recurse
        $liberadoAppCaches += Remove-ItemSeguro (Join-Path $p.FullName "AppData\Roaming\Microsoft\Teams\blob_storage") -Recurse
        $liberadoAppCaches += Remove-ItemSeguro (Join-Path $p.FullName "AppData\Roaming\Microsoft\Teams\IndexedDB") -Recurse
        $liberadoAppCaches += Remove-ItemSeguro (Join-Path $p.FullName "AppData\Roaming\Microsoft\Teams\GPUCache") -Recurse
        $liberadoAppCaches += Remove-ItemSeguro (Join-Path $p.FullName "AppData\Local\Packages\MSTeams_*\LocalCache") -Recurse

        # Edge & Chrome Caches
        $liberadoAppCaches += Remove-ItemSeguro (Join-Path $p.FullName "AppData\Local\Microsoft\Edge\User Data\Default\Cache") -Recurse
        $liberadoAppCaches += Remove-ItemSeguro (Join-Path $p.FullName "AppData\Local\Microsoft\Edge\User Data\Default\Code Cache") -Recurse
        $liberadoAppCaches += Remove-ItemSeguro (Join-Path $p.FullName "AppData\Local\Google\Chrome\User Data\Default\Cache") -Recurse
        $liberadoAppCaches += Remove-ItemSeguro (Join-Path $p.FullName "AppData\Local\Google\Chrome\User Data\Default\Code Cache") -Recurse

        # Anexos temporários do Outlook
        $liberadoAppCaches += Remove-ItemSeguro (Join-Path $p.FullName "AppData\Local\Microsoft\Windows\INetCache\Content.Outlook") -Recurse
    }
    $resumoLiberado['AppCaches'] = $liberadoAppCaches
    Write-Host "      Liberado: $(Format-Tamanho $liberadoAppCaches)" -ForegroundColor Green

    # MÓDULO 7: Hibernação (Desativação para recuperação imediata de RAM em disco)
    if ($DisableHibernation -or ($Mode -eq 'Deep' -and -not (Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue))) {
        Write-Host "`n[7/9] Verificando arquivo de hibernação (hiberfil.sys)..." -ForegroundColor White
        $hiberFile = "C:\hiberfil.sys"
        if (Test-Path $hiberFile) {
            $tamHiber = (Get-Item $hiberFile -Force -ErrorAction SilentlyContinue).Length
            Write-Host "      Desativando hibernação para liberar $(Format-Tamanho $tamHiber)..." -ForegroundColor Yellow
            & powercfg.exe /hibernate off
            $resumoLiberado['Hibernacao'] = $tamHiber
            Write-Host "      Liberado: $(Format-Tamanho $tamHiber)" -ForegroundColor Green
        } else {
            Write-Host "      Hibernação já está desativada." -ForegroundColor Gray
        }
    } else {
        Write-Host "`n[7/9] Hibernação mantida (use -DisableHibernation para forçar remoção do hiberfil.sys)." -ForegroundColor Gray
    }

    # MÓDULO 8: Cópias de Sombra (VSS)
    if ($PurgeShadowCopies -or $Mode -eq 'Deep') {
        Write-Host "`n[8/9] Gerenciando Cópias de Sombra de Volume (VSS)..." -ForegroundColor White
        try {
            $vssOutput = & vssadmin.exe list shadowstorage /for=C: 2>&1
            $vssUsed = ($vssOutput | Select-String "Espaço de armazenamento de cópia de sombra usado|Used Shadow Copy Storage space")
            Write-Host "      Status VSS: $vssUsed" -ForegroundColor Gray
            Write-Host "      Ajustando cota máxima de shadow storage para 5GB e expurgando antigas..." -ForegroundColor Yellow
            & vssadmin.exe resize shadowstorage /for=C: /on=C: /maxsize=5GB 2>&1 | Out-Null
            & vssadmin.exe delete shadows /for=C: /oldest /quiet 2>&1 | Out-Null
        } catch {}
    } else {
        Write-Host "`n[8/9] VSS mantido (use -PurgeShadowCopies para podar sombras antigas)." -ForegroundColor Gray
    }

    # MÓDULO 9: Component Store DISM (WinSxS)
    if (-not $SkipDism) {
        Write-Host "`n[9/9] Executando DISM Component Cleanup no repositório WinSxS..." -ForegroundColor White
        Write-Host "      (Pode levar de 1 a 4 minutos, aguarde...)" -ForegroundColor Gray
        $antesDism = (Get-EspacoDisco 'C').FreeBytes
        $dismArgs = @('/online', '/Cleanup-Image', '/StartComponentCleanup')
        if ($Mode -eq 'Deep') {
            Write-Host "      Modo Deep ativo: aplicando /ResetBase para expurgo definitivo..." -ForegroundColor Yellow
            $dismArgs += '/ResetBase'
        }
        $procDism = Start-Process -FilePath "dism.exe" -ArgumentList $dismArgs -Wait -NoNewWindow -PassThru
        $depoisDism = (Get-EspacoDisco 'C').FreeBytes
        $resumoLiberado['WinSxS_DISM'] = [math]::Max(0, $depoisDism - $antesDism)
        Write-Host "      Código de saída DISM: $($procDism.ExitCode) | Liberado: $(Format-Tamanho $resumoLiberado['WinSxS_DISM'])" -ForegroundColor Green
    } else {
        Write-Host "`n[9/9] DISM pulado por parâmetro (-SkipDism)." -ForegroundColor Gray
    }
}

# -----------------------------------------------------------------------------
# 5. Auditoria dos Maiores Arquivos e Pastas de C: (Diagnóstico Proativo)
# -----------------------------------------------------------------------------
Write-Host "`n"
Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "         DIAGNÓSTICO: TOP $TopFilesCount MAIORES ARQUIVOS NA UNIDADE C:       " -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "Localizando grandes arquivos soltos (> 100 MB)..." -ForegroundColor Gray

# Varre pastas principais evitando varrer recursivamente a pasta Windows inteira
$pastasAlvo = @("C:\Users", "C:\ProgramData", "C:\Windows\Temp", "C:\Windows\Logs")
$pastasRaiz = Get-ChildItem -Path "C:\" -Directory -Attributes !ReparsePoint -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notin @('Windows', 'Users', 'ProgramData', 'Program Files', 'Program Files (x86)', '$Recycle.Bin', 'System Volume Information', '$WINDOWS.~BT', 'Recovery') } |
    ForEach-Object { $_.FullName }
if ($pastasRaiz) { $pastasAlvo += $pastasRaiz }

$grandesArquivos = [System.Collections.Generic.List[PSObject]]::new()

# Arquivos na raiz de C:
Get-ChildItem -Path "C:\" -File -Force -ErrorAction SilentlyContinue |
    Where-Object { $_.Length -gt 100MB } | ForEach-Object { $grandesArquivos.Add($_) }

# Dumps no Windows
Get-ChildItem -Path "C:\Windows" -File -Force -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match '\.(dmp|log)$' -and $_.Length -gt 100MB } | ForEach-Object { $grandesArquivos.Add($_) }

foreach ($pasta in $pastasAlvo) {
    if (Test-Path -LiteralPath $pasta) {
        Get-ChildItem -LiteralPath $pasta -File -Recurse -Attributes !ReparsePoint -ErrorAction SilentlyContinue |
            Where-Object { $_.Length -gt 100MB } | ForEach-Object { $grandesArquivos.Add($_) }
    }
}

$grandesArquivos = $grandesArquivos |
    Sort-Object Length -Descending |
    Select-Object -First $TopFilesCount

if ($grandesArquivos -and $grandesArquivos.Count -gt 0) {
    $grandesArquivos | Select-Object @{Name="Tamanho"; Expression={ Format-Tamanho $_.Length }},
                                     @{Name="Última Modificação"; Expression={ $_.LastWriteTime.ToString('dd/MM/yyyy HH:mm') }},
                                     @{Name="Caminho Completo"; Expression={ $_.FullName }} |
        Format-Table -AutoSize
} else {
    Write-Host "Nenhum arquivo avulso acima de 100 MB encontrado fora do WinSxS." -ForegroundColor Green
}

# -----------------------------------------------------------------------------
# 6. Painel Comparativo Final
# -----------------------------------------------------------------------------
$discoDepois = Get-EspacoDisco 'C'
$totalLiberado = $discoDepois.FreeBytes - $discoAntes.FreeBytes

Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "                         RELATÓRIO CONSOLIDADO                                " -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan

if ($Mode -ne 'Analyze') {
    Write-Host "`nDetalhamento por Módulo:" -ForegroundColor Yellow
    foreach ($m in $resumoLiberado.Keys) {
        Write-Host ("  {0,-20} : {1}" -f $m, (Format-Tamanho $resumoLiberado[$m]))
    }

    Write-Host "`nResultado na Unidade C:" -ForegroundColor Cyan
    Write-Host "  Espaço Livre ANTES   : $($discoAntes.FreeGB) GB ($($discoAntes.PctFree)% livre)"
    Write-Host "  Espaço Livre DEPOIS  : $($discoDepois.FreeGB) GB ($($discoDepois.PctFree)% livre)"
    Write-Host "  TOTAL RECUPERADO     : $(Format-Tamanho $totalLiberado)" -ForegroundColor $(if ($totalLiberado -gt 0) { 'Green' } else { 'Yellow' })
} else {
    Write-Host "`nModo Análise finalizado. Nenhuma alteração foi realizada em disco." -ForegroundColor Yellow
}
Write-Host "==============================================================================" -ForegroundColor Cyan

if ($Host.Name -match 'ConsoleHost' -and [Environment]::UserInteractive -and -not [Console]::IsInputRedirected) {
    Write-Host "`nPressione Enter para fechar esta janela..." -ForegroundColor Gray
    Read-Host | Out-Null
}

