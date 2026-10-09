<#
.SYNOPSIS
  Ferramenta interativa para clonar usuários no Microsoft Entra ID / Intune / Exchange Online.
#>

# ==============================================================================
# PASSO 0: PRÉ-REQUISITOS E MÓDULOS
# ==============================================================================
Write-Host "Verificando pré-requisitos..." -ForegroundColor Cyan

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
    Install-PackageProvider -Name NuGet -Force | Out-Null
}

$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.Identity.DirectoryManagement",
    "ExchangeOnlineManagement"
)

foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable $m)) {
        Write-Host "Instalando módulo: $m ..." -ForegroundColor Yellow
        Install-Module $m -Scope CurrentUser -AllowClobber -Force -Repository PSGallery | Out-Null
    }
}

Import-Module Microsoft.Graph.Authentication -Force
Import-Module ExchangeOnlineManagement -Force

Write-Host "Módulos carregados com sucesso." -ForegroundColor Green


# ==============================================================================
# PASSO 1: CONEXÃO COM MS GRAPH E EXCHANGE ONLINE
# ==============================================================================
$requiredScopes = @("User.ReadWrite.All", "Group.ReadWrite.All", "Directory.Read.All")
Write-Host "`nVerificando conexões..." -ForegroundColor Cyan

# Conexão Microsoft Graph
$ctx = Get-MgContext -ErrorAction SilentlyContinue
if ($null -eq $ctx -or ($requiredScopes | Where-Object { $_ -notin $ctx.Scopes })) {
    Write-Host "Iniciando login no Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome | Out-Null
}
Write-Host "Conectado ao Graph: $((Get-MgContext).Account)" -ForegroundColor Green

# Conexão Exchange Online (Necessário para Distribution Lists)
# NOTA: -Device força device code flow, evitando o broker nativo (WAM), que costuma
# falhar com "RuntimeBroker NullReferenceException" em Windows Server / sessões RDP.
# Só existe em ExchangeOnlineManagement >= 3.2.0, por isso checamos antes de usar.
$script:exoConectado = $false
$exoSuportaDevice = (Get-Command Connect-ExchangeOnline).Parameters.Keys -contains 'Device'

try {
    $exoAccount = (Get-ConnectionInformation -ErrorAction SilentlyContinue).UserPrincipalName
    if ([string]::IsNullOrWhiteSpace($exoAccount)) {
        Write-Host "Conectando ao Exchange Online (necessário para Listas de Distribuição)..." -ForegroundColor Yellow
        if ($exoSuportaDevice) {
            Connect-ExchangeOnline -ShowBanner:$false -Device
        } else {
            Write-Warning "Módulo ExchangeOnlineManagement desatualizado (sem suporte a -Device). Considere rodar: Update-Module ExchangeOnlineManagement -Force"
            Connect-ExchangeOnline -ShowBanner:$false
        }
    }
    $script:exoConectado = $true
    Write-Host "Conectado ao Exchange Online." -ForegroundColor Green
} catch {
    Write-Warning "Não foi possível conectar ao Exchange Online. Listas de Distribuição podem falhar."
    Write-Warning "Detalhe: $($_.Exception.Message)"
}


# ==============================================================================
# PASSO 2: COLETA DE DADOS
# ==============================================================================
$modeloUser = $null
while ($null -eq $modeloUser) {
    $modeloUPN = Read-Host "`n[1/9] Informe o E-mail (UPN) do usuário MODELO"
    if ([string]::IsNullOrWhiteSpace($modeloUPN)) { continue }

    # UsageLocation não vem por padrão e é obrigatório para atribuir licença
    $modeloUser = Get-MgUser -UserId $modeloUPN -Property Id,DisplayName,UserPrincipalName,UsageLocation -ErrorAction SilentlyContinue
    if ($null -eq $modeloUser) { Write-Warning "Usuário não encontrado." }
}
$usageLocation = if ($modeloUser.UsageLocation) { $modeloUser.UsageLocation } else { 'BR' }

Write-Host "`nDados do NOVO usuário:" -ForegroundColor Cyan
$fn    = Read-Host "[2/9] Primeiro nome"
$ln    = Read-Host "[3/9] Sobrenome"
$dept  = Read-Host "[4/9] Departamento"
$cargo = Read-Host "[5/9] Cargo (Job Title)"
$upn   = Read-Host "[6/9] E-mail (UPN) novo"

if ([string]::IsNullOrWhiteSpace($fn) -or [string]::IsNullOrWhiteSpace($ln) -or [string]::IsNullOrWhiteSpace($upn)) {
    Write-Error "Erro: Nome, Sobrenome e E-mail são obrigatórios."
    exit 1
}

if (Get-MgUser -UserId $upn -ErrorAction SilentlyContinue) {
    Write-Error "Erro: Já existe um usuário com o e-mail $upn"
    exit 1
}

# --- Gerente ---
$gerenteUser = $null
Write-Host "`n[7/9] Gerente do novo usuário" -ForegroundColor Cyan
$gerenteInput = Read-Host "Informe o UPN, e-mail ou parte do nome do gerente ([Enter] para pular)"

if (-not [string]::IsNullOrWhiteSpace($gerenteInput)) {
    $gerenteUser = Get-MgUser -UserId $gerenteInput -ErrorAction SilentlyContinue

    if ($null -eq $gerenteUser) {
        $candidatos = @(Get-MgUser -Filter "startswith(displayName,'$($gerenteInput -replace "'", "''")')" -ConsistencyLevel eventual -CountVariable ct -All -ErrorAction SilentlyContinue)

        if ($null -eq $candidatos -or $candidatos.Count -eq 0) {
            Write-Warning "Nenhum gerente encontrado com '$gerenteInput'. Prosseguindo sem gerente."
        }
        elseif ($candidatos.Count -eq 1) {
            $gerenteUser = $candidatos[0]
        }
        else {
            Write-Host "Múltiplos usuários encontrados:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $candidatos.Count; $i++) {
                Write-Host "  [$i] $($candidatos[$i].DisplayName) - $($candidatos[$i].UserPrincipalName)"
            }
            $idx = Read-Host "Escolha o número correspondente ao gerente ([Enter] para pular)"
            if ($idx -match '^\d+$' -and [int]$idx -lt $candidatos.Count) {
                $gerenteUser = $candidatos[[int]$idx]
            }
        }
    }

    if ($gerenteUser) {
        Write-Host "Gerente selecionado: $($gerenteUser.DisplayName) ($($gerenteUser.UserPrincipalName))" -ForegroundColor Green
    }
}

# --- Senha ---
$passOK = $false
while (-not $passOK) {
    $s1 = Read-Host "[8/9] Senha temporária" -AsSecureString
    $s2 = Read-Host "Confirme a senha" -AsSecureString
    $p1 = [Net.NetworkCredential]::new('', $s1).Password
    $p2 = [Net.NetworkCredential]::new('', $s2).Password

    if ($p1 -eq $p2 -and -not [string]::IsNullOrWhiteSpace($p1)) {
        $newPass = $p1
        $passOK = $true
    } else { Write-Warning "As senhas não coincidem." }
    $p1 = $null; $p2 = $null
}


# ==============================================================================
# PASSO 3: EXECUÇÃO DA CRIAÇÃO
# ==============================================================================
Write-Host "`n--- RESUMO ---" -ForegroundColor Yellow
Write-Host "Modelo:  $($modeloUser.DisplayName)"
Write-Host "Novo:    $fn $ln ($upn)"
Write-Host "Depto:   $dept"
Write-Host "Cargo:   $cargo"
Write-Host "Gerente: $(if ($gerenteUser) { $gerenteUser.DisplayName } else { '(nenhum)' })"

if ((Read-Host "`nConfirmar criação? (S/N)").ToUpper() -ne 'S') { exit 0 }

Write-Host "`nCriando conta..." -ForegroundColor Cyan
try {
    $nick = $upn.Split('@')[0]
    $pwProfile = @{ Password = $newPass; ForceChangePasswordNextSignIn = $true }

    $novoUsuario = New-MgUser -GivenName $fn -Surname $ln -DisplayName "$fn $ln" -Department $dept -JobTitle $cargo -UserPrincipalName $upn -MailNickname $nick -UsageLocation $usageLocation -AccountEnabled -PasswordProfile $pwProfile
    Write-Host "✔️ Usuário criado com sucesso." -ForegroundColor Green
} catch {
    Write-Error "Falha ao criar usuário: $($_.Exception.Message)"
    exit 1
} finally {
    $newPass = $null; $pwProfile = $null
}

# Atribuir Gerente
if ($gerenteUser) {
    try {
        $managerRef = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($gerenteUser.Id)" }
        Set-MgUserManagerByRef -UserId $novoUsuario.Id -BodyParameter $managerRef
        Write-Host "✔️ Gerente atribuído: $($gerenteUser.DisplayName)" -ForegroundColor Green
    } catch {
        Write-Warning "Falha ao atribuir gerente: $($_.Exception.Message)"
    }
}


# ==============================================================================
# PASSO 3.2: LICENCIAMENTO
# ==============================================================================
# Feito antes da sincronização com o Exchange: sem licença com Exchange o usuário
# não ganha mailbox e nunca vira recipient, fazendo as DLs falharem.

# Nomes amigáveis dos SKUs (CSV oficial da Microsoft, cache local de 30 dias)
function Get-SkuFriendlyNames {
    $cache = Join-Path $env:TEMP 'M365_SkuFriendlyNames.csv'
    $url   = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
    $map   = @{}
    try {
        if (-not (Test-Path $cache) -or (Get-Item $cache).LastWriteTime -lt (Get-Date).AddDays(-30)) {
            Invoke-WebRequest -Uri $url -OutFile $cache -UseBasicParsing -ErrorAction Stop
        }
        foreach ($row in Import-Csv $cache) {
            if (-not $map.ContainsKey($row.GUID)) { $map[$row.GUID] = $row.Product_Display_Name }
        }
    } catch {
        Write-Warning "Não foi possível obter os nomes amigáveis das licenças. Exibindo SkuPartNumber."
    }
    return $map
}

Write-Host "`n--- LICENÇAS DISPONÍVEIS ---" -ForegroundColor Yellow
$nomesSku    = Get-SkuFriendlyNames
$skusModelo  = @(Get-MgUserLicenseDetail -UserId $modeloUser.Id -ErrorAction SilentlyContinue | ForEach-Object { "$($_.SkuId)" })

$skus = @(Get-MgSubscribedSku | Where-Object { $_.CapabilityStatus -eq 'Enabled' -and $_.PrepaidUnits.Enabled -gt 0 } | ForEach-Object {
    $id = "$($_.SkuId)"
    [pscustomobject]@{
        Nome          = if ($nomesSku[$id]) { $nomesSku[$id] } else { $_.SkuPartNumber }
        SkuPartNumber = $_.SkuPartNumber
        SkuId         = $id
        Total         = [int]$_.PrepaidUnits.Enabled
        Livres        = [int]$_.PrepaidUnits.Enabled - [int]$_.ConsumedUnits
        DoModelo      = $skusModelo -contains $id
    }
} | Sort-Object Nome)

# Write-Host (em vez de Format-Table) garante que a lista apareça ANTES do Read-Host
for ($i = 0; $i -lt $skus.Count; $i++) {
    $s = $skus[$i]
    $cor   = if ($s.Livres -le 0) { 'DarkGray' } elseif ($s.DoModelo) { 'Green' } else { 'White' }
    $marca = if ($s.DoModelo) { '*' } else { ' ' }
    Write-Host ("  [{0,2}]{1} {2,-55} {3,7} livres de {4,-8} ({5})" -f $i, $marca, $s.Nome, $s.Livres, $s.Total, $s.SkuPartNumber) -ForegroundColor $cor
}
Write-Host "  (* = licença do usuário modelo | cinza = sem licenças livres)" -ForegroundColor DarkGray

$skuChoice = Read-Host "`nNúmero(s) separados por vírgula, [M] para copiar as do modelo ou [Enter] para pular"
$skusSel = @()
if ($skuChoice.Trim().ToUpper() -eq 'M') {
    $skusSel = @($skus | Where-Object DoModelo)
} elseif (-not [string]::IsNullOrWhiteSpace($skuChoice)) {
    $skusSel = @($skuChoice -split ',' | ForEach-Object { $_.Trim() } |
        Where-Object { $_ -match '^\d+$' -and [int]$_ -lt $skus.Count } |
        ForEach-Object { $skus[[int]$_] })
}

$semSaldo = @($skusSel | Where-Object { $_.Livres -le 0 })
foreach ($s in $semSaldo) { Write-Warning "Sem licenças livres de '$($s.Nome)'. Ignorada." }
$skusSel = @($skusSel | Where-Object { $_.Livres -gt 0 })

if ($skusSel.Count -gt 0) {
    try {
        # Chamada direta à API: Set-MgUserLicense exigiria o módulo Microsoft.Graph.Users.Actions
        $body = @{
            addLicenses    = @($skusSel | ForEach-Object { @{ skuId = $_.SkuId; disabledPlans = @() } })
            removeLicenses = @()
        } | ConvertTo-Json -Depth 5
        Invoke-MgGraphRequest -Method POST -Uri "v1.0/users/$($novoUsuario.Id)/assignLicense" -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
        foreach ($s in $skusSel) { Write-Host "✔️ Licença atribuída: $($s.Nome)" -ForegroundColor Green }
    } catch {
        Write-Warning "Erro ao atribuir licença: $($_.Exception.Message)"
    }
}


# ==============================================================================
# PASSO 3.5: AGUARDAR SINCRONIZAÇÃO ENTRA ID -> EXCHANGE ONLINE
# ==============================================================================
# Usuários recém-criados via Graph podem levar alguns minutos para aparecer como
# recipient no Exchange Online. Sem isso, Add-DistributionGroupMember falha com
# "Não foi possível localizar o objeto" mesmo o usuário existindo corretamente.
function Wait-ForExchangeRecipient {
    param(
        [Parameter(Mandatory=$true)] [string]$UPN,
        [int]$MaxAttempts = 10,
        [int]$DelaySeconds = 15
    )

    if (-not $script:exoConectado) { return $false }

    Write-Host "`nAguardando sincronização do usuário com o Exchange Online..." -ForegroundColor Cyan
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        $recipient = Get-Recipient -Identity $UPN -ErrorAction SilentlyContinue
        if ($recipient) {
            Write-Host "✔️ Objeto sincronizado no Exchange Online." -ForegroundColor Green
            return $true
        }
        Write-Host "  ... ainda não sincronizado (tentativa $i/$MaxAttempts), aguardando ${DelaySeconds}s" -ForegroundColor DarkGray
        Start-Sleep -Seconds $DelaySeconds
    }

    Write-Warning "Objeto não sincronizou com o Exchange Online dentro do tempo esperado. Grupos mail-enabled podem falhar — rode a etapa de grupos adicionais novamente daqui a alguns minutos."
    return $false
}

$script:exoRecipientPronto = Wait-ForExchangeRecipient -UPN $upn


# ==============================================================================
# FUNÇÃO AUXILIAR: ADICIONAR AO GRUPO (GRAPH vs EXCHANGE)
# ==============================================================================
function Add-UserToGroupHybrid {
    param(
        [Parameter(Mandatory=$true)] $GroupObj,
        [Parameter(Mandatory=$true)] $UserObj
    )

    $gid = $GroupObj.Id

    # Objetos de Get-MgUserMemberOf trazem as propriedades em AdditionalProperties;
    # objetos de Get-MgGroup trazem como propriedades tipadas.
    $ap          = $GroupObj.AdditionalProperties
    $groupTypes  = @($GroupObj.GroupTypes) + @($ap.groupTypes) | Where-Object { $_ }
    $mailEnabled = [bool]($GroupObj.MailEnabled -or $ap.mailEnabled)
    $onPrem      = [bool]($GroupObj.OnPremisesSyncEnabled -or $ap.onPremisesSyncEnabled)

    if ($groupTypes -contains 'DynamicMembership') {
        return @{ Success = $false; Error = "Dinâmico" }
    }
    if ($onPrem) {
        return @{ Success = $false; Error = "Grupo sincronizado do AD local (altere no AD on-premises)" }
    }

    # DL e Security mail-enabled só aceitam membros via Exchange Online
    # (grupos M365 / Unified são mail-enabled, mas funcionam via Graph)
    if ($mailEnabled -and $groupTypes -notcontains 'Unified') {
        if (-not $script:exoConectado) {
            return @{ Success = $false; Error = "Grupo mail-enabled requer Exchange Online (sessão não conectada)" }
        }
        if (-not $script:exoRecipientPronto) {
            return @{ Success = $false; Pending = $true; Error = "Usuário ainda não sincronizado com Exchange Online (nova tentativa no final)" }
        }
        try {
            Add-DistributionGroupMember -Identity $gid -Member $UserObj.UserPrincipalName -BypassSecurityGroupManagerCheck -ErrorAction Stop
            return @{ Success = $true; Mode = "Exchange" }
        } catch {
            return @{ Success = $false; Error = "Exchange: $($_.Exception.Message)" }
        }
    }

    try {
        New-MgGroupMember -GroupId $gid -DirectoryObjectId $UserObj.Id -ErrorAction Stop | Out-Null
        return @{ Success = $true; Mode = "Graph" }
    } catch {
        return @{ Success = $false; Error = "Graph: $($_.Exception.Message)" }
    }
}


# ==============================================================================
# PASSO 4: CLONAGEM DE GRUPOS (do usuário MODELO)
# ==============================================================================
Write-Host "`nClonando grupos do usuário modelo..." -ForegroundColor Cyan

$added   = New-Object System.Collections.Generic.List[string]
$skipped = New-Object System.Collections.Generic.List[string]
$failed  = New-Object System.Collections.Generic.List[string]
$jaAdicionadosIds = New-Object System.Collections.Generic.HashSet[string]
# Grupos mail-enabled que falharam só porque o usuário ainda não chegou ao Exchange
$pendentesExo = New-Object System.Collections.Generic.List[object]

$memberships = Get-MgUserMemberOf -UserId $modeloUser.Id -All -ErrorAction SilentlyContinue
$groups = $memberships | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' }

foreach ($g in $groups) {
    $gid = $g.Id
    $nomeG = $g.AdditionalProperties.displayName
    if ([string]::IsNullOrWhiteSpace($nomeG)) {
        try { $nomeG = (Get-MgGroup -GroupId $gid).DisplayName } catch { $nomeG = $gid }
    }

    $res = Add-UserToGroupHybrid -GroupObj $g -UserObj $novoUsuario

    if ($res.Success) {
        $added.Add("$nomeG ($($res.Mode))") | Out-Null
        $jaAdicionadosIds.Add($gid) | Out-Null
        Write-Host "  + Adicionado: $nomeG" -ForegroundColor Green
    } else {
        if ($res.Error -eq "Dinâmico") {
            $skipped.Add("$nomeG (Dinâmico)") | Out-Null
            Write-Host "  - Ignorado (Dinâmico): $nomeG" -ForegroundColor DarkGray
        } elseif ($res.Pending) {
            $pendentesExo.Add([pscustomobject]@{ Grupo = $g; Nome = $nomeG }) | Out-Null
            Write-Host "  ~ Pendente (aguardando Exchange): $nomeG" -ForegroundColor Yellow
        } else {
            $failed.Add("$nomeG") | Out-Null
            Write-Warning ("  ! Erro no grupo {0}: {1}" -f $nomeG, $res.Error)
        }
    }
}

Write-Host "`nResumo de grupos clonados:" -ForegroundColor Yellow
Write-Host "  Sucesso:   $($added.Count)"
Write-Host "  Pendentes: $($pendentesExo.Count)"
Write-Host "  Ignorados: $($skipped.Count)"
Write-Host "  Falhas:    $($failed.Count)"


# ==============================================================================
# PASSO 4.5: GRUPOS ADICIONAIS
# ==============================================================================
Write-Host "`n--- GRUPOS ADICIONAIS (M365 / Security / Distribuição) ---" -ForegroundColor Yellow

while ($true) {
    $respAdd = Read-Host "`nDeseja adicionar um grupo adicional a este usuário? (S/N)"
    if ($respAdd.ToUpper() -ne 'S') { break }

    $termoBusca = Read-Host "Digite parte do nome do grupo para buscar"
    if ([string]::IsNullOrWhiteSpace($termoBusca)) { continue }

    try {
        $gruposEncontrados = @(Get-MgGroup -Filter "startswith(displayName,'$($termoBusca -replace "'", "''")')" -ConsistencyLevel eventual -CountVariable gct -All -ErrorAction Stop)
    } catch {
        Write-Warning "Erro na busca: $($_.Exception.Message)"
        continue
    }

    if ($null -eq $gruposEncontrados -or $gruposEncontrados.Count -eq 0) {
        Write-Warning "Nenhum grupo encontrado com '$termoBusca'."
        continue
    }

    Write-Host "`nGrupos encontrados:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $gruposEncontrados.Count; $i++) {
        $tipo = if ($gruposEncontrados[$i].MailEnabled -and $gruposEncontrados[$i].SecurityEnabled) { "Security (Mail-Enabled)" }
                elseif ($gruposEncontrados[$i].MailEnabled -and -not $gruposEncontrados[$i].SecurityEnabled) { "Distribuição / M365 Group" }
                elseif ($gruposEncontrados[$i].SecurityEnabled) { "Security Group" }
                else { "Outro" }
        Write-Host "  [$i] $($gruposEncontrados[$i].DisplayName)  -  Tipo: $tipo"
    }

    $escolha = Read-Host "Digite o(s) número(s) do(s) grupo(s) a adicionar, separados por vírgula ([Enter] para cancelar)"
    if ([string]::IsNullOrWhiteSpace($escolha)) { continue }

    $indices = $escolha -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }

    foreach ($idx in $indices) {
        if ($idx -ge $gruposEncontrados.Count) { continue }
        $grupoSel = $gruposEncontrados[$idx]

        if ($jaAdicionadosIds.Contains($grupoSel.Id)) {
            Write-Host "  - Já é membro de: $($grupoSel.DisplayName)" -ForegroundColor DarkGray
            continue
        }

        $res = Add-UserToGroupHybrid -GroupObj $grupoSel -UserObj $novoUsuario

        if ($res.Success) {
            $added.Add("$($grupoSel.DisplayName) ($($res.Mode))") | Out-Null
            $jaAdicionadosIds.Add($grupoSel.Id) | Out-Null
            Write-Host "  + Adicionado: $($grupoSel.DisplayName)" -ForegroundColor Green
        } elseif ($res.Pending) {
            if (-not ($pendentesExo | Where-Object { $_.Grupo.Id -eq $grupoSel.Id })) {
                $pendentesExo.Add([pscustomobject]@{ Grupo = $grupoSel; Nome = $grupoSel.DisplayName }) | Out-Null
            }
            Write-Host "  ~ Pendente (aguardando Exchange): $($grupoSel.DisplayName)" -ForegroundColor Yellow
        } else {
            $failed.Add($grupoSel.DisplayName) | Out-Null
            Write-Warning ("  ! Falha ao adicionar '{0}': {1}" -f $grupoSel.DisplayName, $res.Error)
        }
    }
}



# ==============================================================================
# PASSO 4.8: NOVA TENTATIVA DOS GRUPOS PENDENTES (EXCHANGE)
# ==============================================================================
# A criação da mailbox após o licenciamento pode passar de 5 minutos. Em vez de
# desistir, espera mais e tenta de novo enquanto o operador quiser.
while ($pendentesExo.Count -gt 0) {
    Write-Host "`n$($pendentesExo.Count) grupo(s) aguardando sincronização com o Exchange:" -ForegroundColor Yellow
    $pendentesExo | ForEach-Object { Write-Host "  - $($_.Nome)" }

    $script:exoRecipientPronto = Wait-ForExchangeRecipient -UPN $upn -MaxAttempts 20 -DelaySeconds 15

    if ($script:exoRecipientPronto) {
        foreach ($p in @($pendentesExo)) {
            $res = Add-UserToGroupHybrid -GroupObj $p.Grupo -UserObj $novoUsuario
            if ($res.Success) {
                $added.Add("$($p.Nome) ($($res.Mode))") | Out-Null
                $jaAdicionadosIds.Add($p.Grupo.Id) | Out-Null
                Write-Host "  + Adicionado: $($p.Nome)" -ForegroundColor Green
            } else {
                $failed.Add($p.Nome) | Out-Null
                Write-Warning ("  ! Erro no grupo {0}: {1}" -f $p.Nome, $res.Error)
            }
        }
        $pendentesExo.Clear()
        break
    }

    if ((Read-Host "Usuário ainda não apareceu no Exchange. Continuar aguardando? (S/N)").ToUpper() -ne 'S') {
        foreach ($p in $pendentesExo) { $failed.Add("$($p.Nome) (não sincronizado)") | Out-Null }
        Write-Warning "Adicione manualmente depois (Exchange Admin Center ou Add-DistributionGroupMember):"
        $pendentesExo | ForEach-Object { Write-Warning "  - $($_.Nome)" }
        $pendentesExo.Clear()
    }
}

Write-Host "`nResumo final de grupos:" -ForegroundColor Yellow
Write-Host "  Total adicionados: $($added.Count)"
Write-Host "  Ignorados:         $($skipped.Count)"
Write-Host "  Falhas:            $($failed.Count)"


Clear-History
Write-Host "`n✅ PROCESSO FINALIZADO." -ForegroundColor Green