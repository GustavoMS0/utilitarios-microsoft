<#
.SYNOPSIS
  Funções compartilhadas por todos os scripts do repositório: módulos, conexão com o
  Exchange Online, pastas de saída protegidas e leitura da seleção de itens.
#>

# Instala (só se faltar) e importa um módulo com versão mínima fixada.
# Sem -Force/-AllowClobber: o PowerShell pede confirmação para instalar da PSGallery
# e não sobrescreve comandos de outros módulos sem avisar.
function Initialize-Modulo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Nome,
        [Parameter(Mandatory = $true)][version]$VersaoMinima,
        # Só garante a instalação (quem chama importa do jeito que precisa)
        [switch]$SemImportar
    )

    $instalado = Get-Module -ListAvailable -Name $Nome | Where-Object { $_.Version -ge $VersaoMinima } | Select-Object -First 1
    if (-not $instalado) {
        Write-Host "Instalando módulo: $Nome (>= $VersaoMinima) no escopo do usuário..." -ForegroundColor Yellow
        Install-Module -Name $Nome -MinimumVersion $VersaoMinima -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
    }
    if (-not $SemImportar) {
        Import-Module -Name $Nome -MinimumVersion $VersaoMinima -ErrorAction Stop
    }
}

# Conecta ao Exchange Online só se ainda não houver sessão. Guarda se a conexão foi aberta
# por este script, para Disconnect-ExchangeSeConectado não derrubar a sessão de quem chamou.
# NOTA: O broker WAM nativo do Windows costuma falhar com "RuntimeBroker NullReferenceException"
# em Windows Server, sessões RDP ou terminais não-interativos. Caso ocorra, alternamos
# automaticamente para o Device Code flow (-Device).
$script:ExoConectadoPorScript = $false
function Connect-ExchangeSeNecessario {
    [CmdletBinding()]
    param(
        # Força autenticação por Device Code (evita falhas de WAM / RuntimeBroker em Windows Server / RDP)
        [switch]$Device,
        [switch]$DisableWAM,
        [string]$UserPrincipalName
    )

    Initialize-Modulo -Nome ExchangeOnlineManagement -VersaoMinima 3.2.0
    if (Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected') { return }

    $cmd = Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue
    $suportaDevice = $cmd -and ($cmd.Parameters.Keys -contains 'Device')
    $suportaDisableWAM = $cmd -and ($cmd.Parameters.Keys -contains 'DisableWAM')

    $splat = @{
        ShowBanner = $false
    }
    if ($UserPrincipalName) {
        $splat['UserPrincipalName'] = $UserPrincipalName
    }

    if ($Device) {
        if ($suportaDevice) {
            $splat['Device'] = $true
        } else {
            Write-Warning "O parâmetro -Device não é suportado nesta versão do ExchangeOnlineManagement."
        }
        Connect-ExchangeOnline @splat
        $script:ExoConectadoPorScript = $true
        return
    }

    if ($DisableWAM) {
        if ($suportaDisableWAM) {
            $splat['DisableWAM'] = $true
        } else {
            Write-Warning "O parâmetro -DisableWAM não é suportado nesta versão do ExchangeOnlineManagement."
        }
        Connect-ExchangeOnline @splat
        $script:ExoConectadoPorScript = $true
        return
    }

    # Tentativa interativa padrão com fallback automático para -Device se o broker WAM falhar
    try {
        Connect-ExchangeOnline @splat -ErrorAction Stop
        $script:ExoConectadoPorScript = $true
    } catch {
        $exMsg = $_.Exception.ToString()
        $ehErroBroker = $exMsg -match 'RuntimeBroker|NullReferenceException|WAM|BrokerExtension'
        if ($suportaDevice -and $ehErroBroker) {
            Write-Warning "Falha no broker de autenticação WAM do Windows (RuntimeBroker / NullReferenceException)."
            Write-Host "Alternando automaticamente para autenticação via Device Code (-Device)..." -ForegroundColor Yellow
            $splat['Device'] = $true
            Connect-ExchangeOnline @splat
            $script:ExoConectadoPorScript = $true
        } else {
            throw
        }
    }
}

function Disconnect-ExchangeSeConectado {
    if (-not $script:ExoConectadoPorScript) { return }
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
    } catch {
        Write-Verbose "Falha ao desconectar do Exchange Online: $($_.Exception.Message)"
    }
    $script:ExoConectadoPorScript = $false
}

# Cria a pasta de saída. Se ela for nova, desativa a herança e deixa acesso só para o usuário
# atual, SYSTEM e Administradores (relatórios e backups têm e-mails, nomes e UPNs; pastas como
# C:\Temp e C:\Backup herdam leitura para BUILTIN\Users). Pastas que já existiam não são
# alteradas (podem ter permissões configuradas de propósito), mas geram aviso se forem amplas.
function New-PastaSegura {
    # Criar a pasta de saída não é a alteração que -WhatIf do script deve simular
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [CmdletBinding()]
    param([string]$Caminho)

    if ([string]::IsNullOrWhiteSpace($Caminho)) { return }
    $existia = Test-Path -LiteralPath $Caminho -PathType Container
    $null = New-Item -ItemType Directory -Path $Caminho -Force -WhatIf:$false

    # ACLs NTFS só no Windows ($IsWindows não existe no Windows PowerShell 5.1)
    if ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows) { return }

    if ($existia) {
        $amplos = 'S-1-1-0', 'S-1-5-11', 'S-1-5-32-545'   # Everyone, Authenticated Users, Users
        $acl    = Get-Acl -LiteralPath $Caminho
        $abertos = foreach ($r in $acl.Access) {
            if ($r.AccessControlType -ne 'Allow') { continue }
            try { $sid = $r.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value } catch { continue }
            if ($sid -in $amplos) { "$($r.IdentityReference)" }
        }
        if ($abertos) {
            Write-Warning "A pasta '$Caminho' já existia e dá acesso a: $((@($abertos) | Select-Object -Unique) -join ', ')."
            Write-Warning "  Para restringir: icacls `"$Caminho`" /inheritance:r /grant:r `"$($env:USERDOMAIN)\$($env:USERNAME):(OI)(CI)F`" *S-1-5-18:(OI)(CI)F *S-1-5-32-544:(OI)(CI)F"
        }
        return
    }

    $heranca = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
    $acl = New-Object System.Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true, $false)   # sem herança e sem copiar as regras herdadas
    $sids = @(
        [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        [System.Security.Principal.SecurityIdentifier]::new('S-1-5-18')       # SYSTEM
        [System.Security.Principal.SecurityIdentifier]::new('S-1-5-32-544')   # Administradores
    )
    foreach ($sid in $sids) {
        $acl.AddAccessRule([System.Security.AccessControl.FileSystemAccessRule]::new(
            $sid, 'FullControl', $heranca, 'None', 'Allow'))
    }
    # Set-Acl do Windows PowerShell tenta resolver dono/grupo e falha em máquinas com problema de
    # confiança com o domínio; a API do .NET grava só a DACL
    $dir = [System.IO.DirectoryInfo]::new((Resolve-Path -LiteralPath $Caminho).ProviderPath)
    if ($dir.PSObject.Methods.Name -contains 'SetAccessControl') {
        $dir.SetAccessControl($acl)                                     # .NET Framework (5.1)
    } else {
        [System.IO.FileSystemAclExtensions]::SetAccessControl($dir, $acl)   # .NET (PowerShell 7)
    }
}

# Converte a resposta "0,3,5-8" em índices únicos e válidos (0 .. Maximo-1)
function ConvertFrom-Selecao {
    param([string]$Texto, [int]$Maximo)

    $indices = foreach ($parte in "$Texto" -split ',') {
        $parte = $parte.Trim()
        if ($parte -match '^(\d{1,9})-(\d{1,9})$') {
            $ini = [int]$Matches[1]
            $fim = [math]::Min([int]$Matches[2], $Maximo - 1)
            if ($ini -le $fim) { $ini..$fim }
        } elseif ($parte -match '^\d{1,9}$') {
            [int]$parte
        }
    }
    return @($indices | Where-Object { $_ -lt $Maximo } | Select-Object -Unique)
}
