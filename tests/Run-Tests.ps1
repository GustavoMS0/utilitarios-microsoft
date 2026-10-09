<#
.SYNOPSIS
  Testes offline dos scripts (sem tenant). Os cmdlets do Graph/Exchange são
  substituídos por funções simuladas; nada é alterado no Microsoft 365.

.EXAMPLE
  pwsh -NoProfile -File .\tests\Run-Tests.ps1
#>
$ErrorActionPreference = 'Stop'
$raiz = Split-Path $PSScriptRoot
$tmp  = Join-Path ([IO.Path]::GetTempPath()) "ms-utils-tests-$(Get-Random)"
New-Item -ItemType Directory $tmp | Out-Null

$script:falhas = 0
$script:total  = 0

function Assert($condicao, [string]$descricao) {
    $script:total++
    if ($condicao) { Write-Host "  [OK]   $descricao" -ForegroundColor Green }
    else           { Write-Host "  [FALHA] $descricao" -ForegroundColor Red; $script:falhas++ }
}

# Nunca instalar nada durante os testes
function global:Install-Module { }

function Remove-Mocks([string[]]$Nomes) {
    foreach ($n in $Nomes) { Remove-Item "Function:\$n" -ErrorAction SilentlyContinue }
}


# ==============================================================================
Write-Host "`n== Sintaxe de todos os scripts ==" -ForegroundColor Cyan
# ==============================================================================
foreach ($f in Get-ChildItem $raiz -Filter *.ps1) {
    $erros = $null
    [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$erros) | Out-Null
    Assert ($erros.Count -eq 0) "$($f.Name) sem erros de sintaxe"
}


# ==============================================================================
Write-Host "`n== Sharepoint.ps1 ==" -ForegroundColor Cyan
# ==============================================================================
$global:spCsv = @"
Report Refresh Date,Site Id,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Page View Count,Visited Page Count,Storage Used (Byte),Storage Allocated (Byte),Root Web Template,Owner Principal Name,Report Period
2026-10-08,a1,https://t.sharepoint.com/sites/Financeiro,Fulano,False,$((Get-Date).AddDays(-2).ToString('yyyy-MM-dd')),1000,200,50,10,107374182400,107374182400,GROUP#0,fulano@x.com,30
2026-10-08,a2,https://t.sharepoint.com/,,False,,10,0,0,0,1048576,27487790694400,SITEPAGEPUBLISHING#0,,30
2026-10-08,a3,,Ciclano,False,$((Get-Date).AddDays(-200).ToString('yyyy-MM-dd')),5,0,0,0,0,1099511627776,STS#3,c@x.com,30
2026-10-08,a4,https://t.sharepoint.com/sites/Velho,X,True,2025-01-01,5,0,0,0,0,0,STS#3,x@x.com,30
"@
$global:graphPatch = 0
function global:Get-MgContext { [pscustomobject]@{ Scopes = @('Reports.Read.All','Sites.Read.All','ReportSettings.Read.All'); Account = 'teste' } }
function global:Connect-MgGraph { }
function global:Import-Module { }
function global:Invoke-MgGraphRequest { param($Method, $Uri, $OutputFilePath, $Body, $ContentType)
    if ($OutputFilePath) { $global:spCsv | Set-Content $OutputFilePath -Encoding utf8; return }
    if ($Uri -like '*reportSettings*') {
        if ($Method -eq 'PATCH') { $global:graphPatch++; return }
        return @{ displayConcealedNames = $true }
    }
    if ($Uri -like '*Financeiro*') { return @{ displayName = 'Financeiro Corp' } }
    throw 'not found'
}

$saidaSp = Join-Path $tmp 'sp'
$avisos  = & (Join-Path $raiz 'Sharepoint.ps1') -Out $saidaSp -IncluirOneDrive 3>&1 6>$null | Where-Object { $_ -is [System.Management.Automation.WarningRecord] }
$csvSp   = Get-ChildItem $saidaSp -Filter 'SharePoint-Auditoria-*.csv' | Select-Object -First 1
$linhas  = @(Import-Csv $csvSp.FullName -Delimiter (Get-Culture).TextInfo.ListSeparator)

Assert ($null -ne $csvSp)                                         "gera o CSV de auditoria"
Assert ($linhas.Count -eq 3)                                      "ignora o site excluído (3 de 4)"
Assert (@($avisos | Where-Object Message -match 'OCULTOS').Count -gt 0) "avisa quando os nomes estão ocultos"
Assert ($global:graphPatch -eq 0)                                 "não altera o tenant sem -RevelarNomes"
$fin = $linhas | Where-Object Url -like '*Financeiro'
Assert ($fin.Titulo -eq 'Financeiro Corp')                        "busca o título real via Graph"
Assert ($fin.Status -eq 'Ativo' -and $fin.Alertas -match 'Cota')  "site cheio: Ativo + alerta de cota"
Assert ($fin.Tipo -eq 'Teams / Grupo M365')                       "traduz o template"
$raizSite = $linhas | Where-Object Url -eq 'https://t.sharepoint.com/'
Assert ($raizSite.Titulo -eq '(raiz)' -and $raizSite.Status -eq 'Nunca usado' -and $raizSite.Alertas -match 'Sem dono') "site raiz: título, 'Nunca usado' e 'Sem dono'"
$oculto = $linhas | Where-Object Url -like '(oculto)*'
Assert ($oculto.Status -eq 'Inativo')                             "site sem URL é mantido e classificado como Inativo"
Assert ((Get-ChildItem $saidaSp -Filter 'OneDrive-Uso-*.csv').Count -eq 1) "gera o CSV do OneDrive com -IncluirOneDrive"

& (Join-Path $raiz 'Sharepoint.ps1') -Out $saidaSp -RevelarNomes 3>$null 6>$null | Out-Null
Assert ($global:graphPatch -eq 1)                                 "-RevelarNomes desativa a ocultação (PATCH)"

Remove-Mocks 'Get-MgContext','Connect-MgGraph','Import-Module','Invoke-MgGraphRequest'


# ==============================================================================
Write-Host "`n== Auditoriaexchange.ps1 ==" -ForegroundColor Cyan
# ==============================================================================
$global:setMailbox = New-Object System.Collections.Generic.List[hashtable]
function global:Get-ConnectionInformation { [pscustomobject]@{ State = 'Connected' } }
function global:Connect-ExchangeOnline { throw 'não deveria reconectar' }
function global:Get-OrganizationConfig { [pscustomobject]@{ AuditDisabled = $true } }
function global:Get-Mailbox { param($Identity, $ErrorAction)
    if ($Identity -like 'inexistente*') { throw "Couldn't find object $Identity" }
    [pscustomobject]@{ AuditEnabled = $true; AuditDelegate = @('Create','UpdateInboxRules','Move'); AuditOwner = @('Move') }
}
function global:Set-Mailbox { param($Identity, $AuditEnabled, $DefaultAuditSet, $AuditDelegate, $AuditOwner, $ErrorAction)
    $global:setMailbox.Add($PSBoundParameters) }
function global:Get-MailboxAuditBypassAssociation { param($Identity, $ErrorAction)
    [pscustomobject]@{ AuditBypassEnabled = ($Identity -like 'bypass*') } }

$avisos = & (Join-Path $raiz 'Auditoriaexchange.ps1') -Caixas 'a@x.com','bypass@x.com','inexistente@x.com' 3>&1 6>$null |
    Where-Object { $_ -is [System.Management.Automation.WarningRecord] }

$resets = @($global:setMailbox | Where-Object { $_.ContainsKey('DefaultAuditSet') })
$adds   = @($global:setMailbox | Where-Object { $_.ContainsKey('AuditDelegate') })
Assert ($resets.Count -eq 2)                                         "restaura o DefaultAuditSet nas 2 caixas válidas"
Assert (($resets[0].DefaultAuditSet -join ',') -eq 'Delegate,Owner') "DefaultAuditSet = Delegate, Owner"
Assert ($adds.Count -eq 2 -and $adds[0].AuditDelegate -is [hashtable] -and $adds[0].AuditDelegate.Add -contains 'Move') "AuditDelegate usa @{Add='Move'} (não substitui a lista)"
Assert (@($avisos | Where-Object Message -match 'DESATIVADA').Count -eq 1) "avisa se a auditoria está desligada na organização"
Assert (@($avisos | Where-Object Message -match 'bypass@x.com.*AuditBypassEnabled').Count -eq 1) "avisa sobre caixa com bypass"
Assert (@($avisos | Where-Object Message -match 'Falha ao configurar inexistente').Count -eq 1) "erro numa caixa é reportado e não interrompe as demais"

Remove-Mocks 'Get-ConnectionInformation','Connect-ExchangeOnline','Get-OrganizationConfig','Get-Mailbox','Set-Mailbox','Get-MailboxAuditBypassAssociation'


# ==============================================================================
Write-Host "`n== ExtracaoExchange.ps1 ==" -ForegroundColor Cyan
# ==============================================================================
function New-FakeLog($id, $idx, $total, $owner, $op, $itens) {
    [pscustomobject]@{
        Identity     = $id
        ResultIndex  = $idx
        ResultCount  = $total
        CreationDate = [datetime]'2026-10-05 15:00:00'
        AuditData    = (@{
            Operation       = $op
            MailboxOwnerUPN = $owner
            UserId          = 'joao@x.com'
            LogonType       = 2
            ClientIPAddress = '10.0.0.1'
            Folder          = @{ Path = '\Caixa de Entrada' }
            DestFolder      = @{ Path = '\Itens Excluídos' }
            AffectedItems   = $itens
        } | ConvertTo-Json -Depth 5)
    }
}
$global:ualChamadas = New-Object System.Collections.Generic.List[hashtable]
function global:Get-ConnectionInformation { [pscustomobject]@{ State = 'Connected' } }
function global:Connect-ExchangeOnline { throw 'não deveria reconectar' }
function global:Search-UnifiedAuditLog { param($StartDate, $EndDate, $RecordType, $Operations, $FreeText, $SessionId, $SessionCommand, $ResultSize, $ErrorAction)
    $global:ualChamadas.Add($PSBoundParameters)
    $pagina = @($global:ualChamadas | Where-Object { $_.SessionId -eq $SessionId }).Count
    if ($FreeText -ne 'boletos@x.com') { return }
    switch ($pagina) {
        1 { @(
                New-FakeLog 'e1' 1 3 'boletos@x.com' 'MoveToDeletedItems' @(@{ Subject = 'Boleto 1' }, @{ Subject = 'Boleto 2' })
                New-FakeLog 'e2' 2 3 'outra@x.com'   'HardDelete'         @(@{ Subject = 'Outra caixa' })
            ) }
        2 { @(
                New-FakeLog 'e2' 2 3 'outra@x.com'   'HardDelete'         @(@{ Subject = 'Outra caixa' })   # duplicado
                New-FakeLog 'e3' 3 3 'boletos@x.com' 'SoftDelete'         @()
            ) }
        default { }
    }
}

$csvEx = Join-Path $tmp 'exchange\rel.csv'
& (Join-Path $raiz 'ExtracaoExchange.ps1') -Caixas 'boletos@x.com','vazia@x.com' -CsvPath $csvEx 3>$null 6>$null | Out-Null
$linhas = @(Import-Csv $csvEx -Delimiter (Get-Culture).TextInfo.ListSeparator)

Assert ($global:ualChamadas[0].RecordType -eq 'ExchangeItemGroup')        "busca com RecordType ExchangeItemGroup"
Assert ($global:ualChamadas[0].SessionCommand -eq 'ReturnLargeSet')       "usa paginação (ReturnLargeSet)"
Assert (@($global:ualChamadas | Where-Object FreeText -eq 'boletos@x.com').Count -ge 2) "busca mais de uma página"
Assert ($global:ualChamadas[0].StartDate.Kind -eq 'Utc')                  "datas da busca em UTC"
Assert ($linhas.Count -eq 3)                                              "1 linha por e-mail afetado (2) + evento sem itens (1)"
Assert (-not ($linhas | Where-Object AssuntoDoEmail -eq 'Outra caixa'))   "descarta eventos de outras caixas"
Assert (($linhas | Where-Object AssuntoDoEmail -eq 'Boleto 1').PastaDeDestino -eq '\Itens Excluídos') "lê Folder/DestFolder .Path"
Assert ($linhas[0].TipoAcesso -eq 'Delegado')                             "traduz LogonType"

Remove-Mocks 'Get-ConnectionInformation','Connect-ExchangeOnline','Search-UnifiedAuditLog'


# ==============================================================================
Write-Host "`n== NewClone.ps1 (função Add-UserToGroupHybrid) ==" -ForegroundColor Cyan
# ==============================================================================
# O script é interativo; testamos só a função de decisão Graph x Exchange extraída via AST.
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $raiz 'NewClone.ps1'), [ref]$null, [ref]$null)
$fn  = $ast.Find({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Add-UserToGroupHybrid' }, $true)
. ([scriptblock]::Create($fn.Extent.Text))

$global:modo = $null
function global:Add-DistributionGroupMember { $global:modo = 'Exchange' }
function global:New-MgGroupMember { $global:modo = 'Graph' }
$user = [pscustomobject]@{ Id = 'u1'; UserPrincipalName = 'novo@x.com' }
$script:exoConectado = $true; $script:exoRecipientPronto = $true

$r = Add-UserToGroupHybrid -GroupObj ([pscustomobject]@{ Id = 'g'; GroupTypes = @('DynamicMembership') }) -UserObj $user
Assert ($r.Error -eq 'Dinâmico')                                 "grupo dinâmico é ignorado"
$r = Add-UserToGroupHybrid -GroupObj ([pscustomobject]@{ Id = 'g'; OnPremisesSyncEnabled = $true }) -UserObj $user
Assert (-not $r.Success -and $r.Error -match 'AD local')         "grupo sincronizado do AD é recusado"
$r = Add-UserToGroupHybrid -GroupObj ([pscustomobject]@{ Id = 'g'; MailEnabled = $true; GroupTypes = @() }) -UserObj $user
Assert ($r.Success -and $global:modo -eq 'Exchange')             "lista de distribuição vai pelo Exchange"
$r = Add-UserToGroupHybrid -GroupObj ([pscustomobject]@{ Id = 'g'; MailEnabled = $true; GroupTypes = @('Unified') }) -UserObj $user
Assert ($r.Success -and $global:modo -eq 'Graph')                "grupo M365 vai pelo Graph"
$r = Add-UserToGroupHybrid -GroupObj ([pscustomobject]@{ Id = 'g'; AdditionalProperties = @{ mailEnabled = $true; groupTypes = @() } }) -UserObj $user
Assert ($r.Success -and $global:modo -eq 'Exchange')             "lê propriedades de AdditionalProperties (Get-MgUserMemberOf)"
$script:exoRecipientPronto = $false
$r = Add-UserToGroupHybrid -GroupObj ([pscustomobject]@{ Id = 'g'; MailEnabled = $true }) -UserObj $user
Assert ($r.Pending)                                              "DL fica pendente enquanto o usuário não sincroniza"

Remove-Mocks 'Add-DistributionGroupMember','New-MgGroupMember','Install-Module'


# ==============================================================================
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
$cor = if ($script:falhas) { 'Red' } else { 'Green' }
Write-Host "`n$($script:total - $script:falhas)/$($script:total) testes passaram." -ForegroundColor $cor
exit $script:falhas
