<#
.SYNOPSIS
  Ferramenta interativa para clonar usuários no Microsoft Entra ID / Intune (Microsoft Graph).
  Versão Final Corrigida - Estrutura Limpa + Cargo/Gerente + Grupos Adicionais M365.
#>

# ==============================================================================
# PASSO 0: PRÉ-REQUISITOS
# ==============================================================================
Write-Host "Verificando pre-requisitos..." -ForegroundColor Cyan

# Protocolos de segurança
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# Provedor NuGet
try {
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -Force | Out-Null
    }
} catch {}

# Módulos necessários
$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.Identity.DirectoryManagement"
)

function Install-GraphModules {
    param([string[]]$Modules)
    foreach ($m in $Modules) {
        if (-not (Get-Module -ListAvailable $m)) {
            Write-Host "Instalando modulo: $m ..." -ForegroundColor Yellow
            Install-Module $m -Scope CurrentUser -AllowClobber -Force -Repository PSGallery
        }
    }
}

function Test-ImportGraphAuth {
    try {
        Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop
        $null = Get-Command Get-MgContext -ErrorAction Stop
        $null = Get-Command Connect-MgGraph -ErrorAction Stop
        return $true
    } catch { return $false }
}

Install-GraphModules -Modules $requiredModules

if (-not (Test-ImportGraphAuth)) {
    Write-Host "Tentando reparar instalacao..." -ForegroundColor Yellow
    foreach ($m in $requiredModules) { try { Uninstall-Module $m -AllVersions -Force -ErrorAction SilentlyContinue } catch {} }
    try { Uninstall-Module Microsoft.Graph -AllVersions -Force -ErrorAction SilentlyContinue } catch {}

    Install-GraphModules -Modules $requiredModules

    if (-not (Test-ImportGraphAuth)) {
        Write-Error "ERRO CRITICO: Ocorreu uma falha no carregamento dos modulos. Reinicie o PowerShell."
        exit 1
    }
}

Write-Host "Módulos carregados." -ForegroundColor Green


# ==============================================================================
# PASSO 1: CONEXÃO
# ==============================================================================
$requiredScopes = @("User.ReadWrite.All", "Group.ReadWrite.All", "Directory.Read.All")
Write-Host "`nVerificando conexão..." -ForegroundColor Cyan

function Ensure-GraphConnection {
    param([string[]]$Scopes)
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    $needLogin = $true

    if ($null -ne $ctx -and $ctx.Scopes) {
        $missing = $Scopes | Where-Object { $_ -notin $ctx.Scopes }
        if (-not $missing) { $needLogin = $false }
    }

    if ($needLogin) {
        Write-Host "Iniciando login..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null
    }
    Write-Host "Conectado: $((Get-MgContext).Account)" -ForegroundColor Green
}

Ensure-GraphConnection -Scopes $requiredScopes


# ==============================================================================
# PASSO 2: COLETA DE DADOS
# ==============================================================================
$modeloUser = $null
while ($null -eq $modeloUser) {
    $modeloUPN = Read-Host "`n[1/9] Informe o E-mail (UPN) do usuário MODELO"
    if ([string]::IsNullOrWhiteSpace($modeloUPN)) { continue }

    $modeloUser = Get-MgUser -UserId $modeloUPN -ErrorAction SilentlyContinue
    if ($null -eq $modeloUser) { Write-Warning "Usuario não encontrado." }
}

Write-Host "`nDados do NOVO usuário:" -ForegroundColor Cyan
$fn   = Read-Host "[2/9] Primeiro nome"
$ln   = Read-Host "[3/9] Sobrenome"
$dept = Read-Host "[4/9] Departamento"
$cargo = Read-Host "[5/9] Cargo (Job Title)"
$upn  = Read-Host "[6/9] E-mail (UPN) novo"

if ([string]::IsNullOrWhiteSpace($fn) -or [string]::IsNullOrWhiteSpace($ln) -or [string]::IsNullOrWhiteSpace($upn)) {
    Write-Error "Erro: Nome, Sobrenome e E-mail são obrigatórios."
    exit 1
}

if (Get-MgUser -UserId $upn -ErrorAction SilentlyContinue) {
    Write-Error "Erro: Já existe um usuário com o e-mail $upn"
    exit 1
}

# --- Gerente (opcional, com busca) ---
$gerenteUser = $null
Write-Host "`n[7/9] Gerente do novo usuário" -ForegroundColor Cyan
$gerenteInput = Read-Host "Informe o UPN, e-mail ou parte do nome do gerente ([Enter] para pular)"

if (-not [string]::IsNullOrWhiteSpace($gerenteInput)) {
    # Tenta achar direto por UPN/e-mail
    $gerenteUser = Get-MgUser -UserId $gerenteInput -ErrorAction SilentlyContinue

    if ($null -eq $gerenteUser) {
        # Busca por nome (displayName)
        $candidatos = Get-MgUser -Filter "startswith(displayName,'$gerenteInput')" -ConsistencyLevel eventual -CountVariable ct -All -ErrorAction SilentlyContinue

        if ($null -eq $candidatos -or $candidatos.Count -eq 0) {
            Write-Warning "Nenhum gerente encontrado com '$gerenteInput'. Prosseguindo sem gerente."
        }
        elseif ($candidatos.Count -eq 1) {
            $gerenteUser = $candidatos
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

$passOK = $false
while (-not $passOK) {
    $s1 = Read-Host "[8/9] Senha temporária" -AsSecureString
    $s2 = Read-Host "Confirme a senha" -AsSecureString
    $p1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s1))
    $p2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s2))

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

    $novoUsuario = New-MgUser -GivenName $fn -Surname $ln -DisplayName "$fn $ln" -Department $dept -JobTitle $cargo -UserPrincipalName $upn -MailNickname $nick -AccountEnabled -PasswordProfile $pwProfile
    Write-Host "✔️ Usuário criado com sucesso." -ForegroundColor Green
} catch {
    Write-Error "Falha ao criar usuário: $($_.Exception.Message)"
    exit 1
}

# --- Atribuir gerente (se selecionado) ---
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
# PASSO 4: CLONAGEM DE GRUPOS (do usuário MODELO)
# ==============================================================================
Write-Host "`nClonando grupos do usuário modelo..." -ForegroundColor Cyan

$added   = New-Object System.Collections.Generic.List[string]
$skipped = New-Object System.Collections.Generic.List[string]
$failed  = New-Object System.Collections.Generic.List[string]
$jaAdicionadosIds = New-Object System.Collections.Generic.HashSet[string]

$memberships = Get-MgUserMemberOf -UserId $modeloUser.Id -All -ErrorAction SilentlyContinue
$groups = $memberships | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' }

foreach ($g in $groups) {
    $gid = $g.Id
    $nomeG = $g.AdditionalProperties.displayName
    if ([string]::IsNullOrWhiteSpace($nomeG)) {
        try { $nomeG = (Get-MgGroup -GroupId $gid).DisplayName } catch { $nomeG = $gid }
    }

    try {
        New-MgGroupMember -GroupId $gid -DirectoryObjectId $novoUsuario.Id -ErrorAction Stop | Out-Null
        $added.Add($nomeG) | Out-Null
        $jaAdicionadosIds.Add($gid) | Out-Null
        Write-Host "  + Adicionado: $nomeG"
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match "DynamicMembership|dynamic") {
            $skipped.Add("$nomeG (Dinamico)") | Out-Null
            Write-Host "  - Ignorado (Dinâmico): $nomeG" -ForegroundColor DarkGray
        } elseif ($msg -match "Insufficient privileges") {
            $failed.Add("$nomeG (Permissao)") | Out-Null
            Write-Warning ("  ! Sem permissão para o grupo: {0}" -f $nomeG)
        } else {
            $failed.Add("$nomeG (Erro)") | Out-Null
            Write-Warning ("  ! Erro no grupo {0}: {1}" -f $nomeG, $msg)
        }
    }
}

Write-Host "`nResumo de grupos clonados:" -ForegroundColor Yellow
Write-Host "  Sucesso:   $($added.Count)"
Write-Host "  Ignorados: $($skipped.Count)"
Write-Host "  Falhas:    $($failed.Count)"


# ==============================================================================
# PASSO 4.5: GRUPOS ADICIONAIS (Security Groups / Listas de Distribuição M365)
# ==============================================================================
# Observação: isso cobre Security Groups, Microsoft 365 Groups e grupos
# habilitados para e-mail (mail-enabled security). Distribution Lists
# clássicas (não mail-enabled security) em alguns tenants não são
# gerenciáveis via Microsoft Graph e exigem o módulo ExchangeOnlineManagement
# (cmdlet Add-DistributionGroupMember). Se a adição abaixo falhar para um
# grupo desse tipo, use o Exchange Online.

Write-Host "`n--- GRUPOS ADICIONAIS (M365 / Security / Distribuição) ---" -ForegroundColor Yellow
$continuarAdicionandoGrupos = $true

while ($continuarAdicionandoGrupos) {
    $respAdd = Read-Host "`nDeseja adicionar um grupo adicional a este usuário? (S/N)"
    if ($respAdd.ToUpper() -ne 'S') { $continuarAdicionandoGrupos = $false; break }

    $termoBusca = Read-Host "Digite parte do nome do grupo para buscar"
    if ([string]::IsNullOrWhiteSpace($termoBusca)) { continue }

    try {
        $gruposEncontrados = Get-MgGroup -Filter "startswith(displayName,'$termoBusca')" -ConsistencyLevel eventual -CountVariable gct -All -ErrorAction Stop
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

        try {
            New-MgGroupMember -GroupId $grupoSel.Id -DirectoryObjectId $novoUsuario.Id -ErrorAction Stop | Out-Null
            $added.Add($grupoSel.DisplayName) | Out-Null
            $jaAdicionadosIds.Add($grupoSel.Id) | Out-Null
            Write-Host "  + Adicionado: $($grupoSel.DisplayName)" -ForegroundColor Green
        } catch {
            $msgG = $_.Exception.Message
            $failed.Add("$($grupoSel.DisplayName) (Erro)") | Out-Null
            Write-Warning ("  ! Falha ao adicionar '{0}': {1}" -f $grupoSel.DisplayName, $msgG)
            Write-Warning "    Se for uma Distribution List classica, use o modulo ExchangeOnlineManagement (Add-DistributionGroupMember)."
        }
    }
}

Write-Host "`nResumo final de grupos:" -ForegroundColor Yellow
Write-Host "  Total adicionados: $($added.Count)"
Write-Host "  Ignorados:         $($skipped.Count)"
Write-Host "  Falhas:            $($failed.Count)"


# ==============================================================================
# PASSO 5: LICENCIAMENTO
# ==============================================================================
Write-Host "`n--- LICENÇAS DISPONÍVEIS ---" -ForegroundColor Yellow
$skus = Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId
$skus | Format-Table -AutoSize

$skuChoice = Read-Host "`nDigite o SkuPartNumber (ou SkuId). [Enter] para pular"
if (-not [string]::IsNullOrWhiteSpace($skuChoice)) {
    $skuObj = $skus | Where-Object { $_.SkuId.Guid -eq $skuChoice -or $_.SkuId -eq $skuChoice -or $_.SkuPartNumber -eq $skuChoice } | Select-Object -First 1

    if ($null -eq $skuObj) {
        Write-Warning "Licença não encontrada."
    } else {
        try {
            Set-MgUserLicense -UserId $novoUsuario.Id -AddLicenses @{ SkuId = $skuObj.SkuId } -RemoveLicenses @() | Out-Null
            Write-Host "✔️ Licença atribuída: $($skuObj.SkuPartNumber)" -ForegroundColor Green
        } catch {
            Write-Error "Erro ao atribuir licença: $($_.Exception.Message)"
        }
    }
}

Clear-History
Write-Host "`n✅ PROCESSO FINALIZADO." -ForegroundColor Green