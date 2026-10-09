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

Remove-Mocks 'Add-DistributionGroupMember','New-MgGroupMember'


# ==============================================================================
Write-Host "`n== Recuperar-ItensExcluidos.ps1 ==" -ForegroundColor Cyan
# ==============================================================================
$global:getRI     = $null
$global:restoreRI = New-Object System.Collections.Generic.List[hashtable]
function global:Get-ConnectionInformation { [pscustomobject]@{ State = 'Connected' } }
function global:Connect-ExchangeOnline { throw 'não deveria reconectar' }
function global:Get-RecoverableItems { param($Identity, $ResultSize, $SubjectContains, $FilterItemType, $SourceFolder, $FilterStartTime, $FilterEndTime)
    $global:getRI = $PSBoundParameters
    1..3 | ForEach-Object {
        [pscustomobject]@{ EntryID = "E$_"; Subject = "Boleto $_"; ItemClass = 'IPM.Note'; SourceFolder = 'RecoverableItems'
                           LastParentPath = 'Inbox'; LastModifiedTime = (Get-Date).AddHours(-$_) }
    }
}
function global:Restore-RecoverableItems { param($Identity, $EntryID, $RestoreTargetFolder, $ErrorAction)
    $global:restoreRI.Add($PSBoundParameters)
    if ($EntryID -eq 'E3') { throw 'Item não encontrado' }
}

$relRI = Join-Path $tmp 'ri\rel.csv'
& (Join-Path $raiz 'Recuperar-ItensExcluidos.ps1') -Caixa 'fin@x.com' -Assunto boleto -De (Get-Date).AddDays(-2) `
    -Todos -Force -PastaDestino 'Recuperados' -Relatorio $relRI 3>$null 6>$null | Out-Null
$linhas = @(Import-Csv $relRI -Delimiter (Get-Culture).TextInfo.ListSeparator)

Assert ($global:getRI.SubjectContains -eq 'boleto' -and $global:getRI.FilterStartTime) "repassa filtros de assunto e data"
Assert (-not $global:getRI.ContainsKey('FilterEndTime'))                    "-Ate não informado não vira filtro (DateTime.MinValue)"
Assert (($global:restoreRI | ForEach-Object EntryID) -join ',' -eq 'E1,E2,E3') "restaura cada item pelo EntryID"
Assert (@($global:restoreRI | Where-Object RestoreTargetFolder -eq 'Recuperados').Count -eq 3) "usa -PastaDestino"
Assert (@($linhas | Where-Object Status -eq 'OK').Count -eq 2 -and @($linhas | Where-Object Status -eq 'Falha').Count -eq 1) "falha num item não interrompe os demais e vai pro relatório"

Remove-Mocks 'Get-RecoverableItems','Restore-RecoverableItems'
& (Join-Path $raiz 'Recuperar-ItensExcluidos.ps1') -Caixa 'fin@x.com' -Todos -Force 3>$null 6>$null | Out-Null
Assert ($LASTEXITCODE -eq 1)                                                "sem a função Mailbox Import Export: orienta e sai com erro"
Remove-Mocks 'Get-ConnectionInformation','Connect-ExchangeOnline'


# ==============================================================================
Write-Host "`n== Backup-Caixa.ps1 ==" -ForegroundColor Cyan
# ==============================================================================
# Caixa simulada: Inbox (com subpasta Clientes) e Calendar
function New-HttpError([int]$Status) {
    $e = [Exception]::new("HTTP $Status")
    $e | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = $Status })
    $e
}
function New-FakeItem($Id, $Ck, $Tipo, $Criado, $Assunto, $Remetente, $Recebido, $Anexo, $Tamanho = 50KB) {
    $global:mbx.Itens[$Id] = @{
        Item = @{ id = $Id; changeKey = $Ck; type = $Tipo; size = $Tamanho; createdDateTime = $Criado }
        Meta = @{ Assunto = $Assunto; Remetente = $Remetente; Recebido = $Recebido; Anexo = $Anexo }
    }
    $global:mbx.Itens[$Id].Item
}
function global:Get-FakeFolders($Pai) {
    $todas = @($global:mbx.Pastas) + $global:mbx.Criadas.ToArray()
    foreach ($p in $todas | Where-Object { $_.pai -eq $Pai }) {
        $id = $p.id
        @{ id = $id; displayName = $p.displayName; type = $p.type; totalItemCount = 0
           childFolderCount = @($todas | Where-Object { $_.pai -eq $id }).Count }
    }
}
function global:New-FakeFolder($Nome, $Tipo, $Pai) {
    $p = @{ id = "N$($global:mbx.Criadas.Count + 1)"; displayName = $Nome; type = $Tipo; pai = $Pai }
    $global:mbx.Criadas.Add($p)
    @{ id = $p.id; displayName = $Nome; type = $Tipo }
}

$global:mbx = @{
    Pastas  = @(
        @{ id = 'F1'; displayName = 'Inbox';    type = 'IPF.Note';        pai = $null }
        @{ id = 'F2'; displayName = 'Clientes'; type = 'IPF.Note';        pai = 'F1' }
        @{ id = 'F3'; displayName = 'Calendar'; type = 'IPF.Appointment'; pai = $null }
    )
    Itens   = [hashtable]::new([StringComparer]::Ordinal)   # IDs do Exchange diferenciam maiúsculas
    Delta   = @{}
    Exports = New-Object System.Collections.Generic.List[object]
    Criadas = New-Object System.Collections.Generic.List[object]
    Falha   = @('mBig')
}
$m1   = New-FakeItem 'm1' 'ck1' 'IPM.Note' '2026-09-10T12:00:00Z' 'Boleto setembro' 'Fornecedor A' '2026-09-10T12:00:00Z' 'true'
$m2   = New-FakeItem 'm2' 'ck1' 'IPM.Note' '2026-09-20T12:00:00Z' 'Reunião' 'Fulano' '2026-09-20T12:00:00Z' 'false'
$m3   = New-FakeItem 'm3' 'ck1' 'IPM.Note' '2026-08-01T12:00:00Z' 'Contrato cliente X' 'Cliente X' '2026-08-01T12:00:00Z' 'false'
$mBig = New-FakeItem 'mBig' 'ck1' 'IPM.Note' '2026-08-02T12:00:00Z' 'Arquivo enorme' 'Cliente X' '2026-08-02T12:00:00Z' 'true' 50MB
$c1   = New-FakeItem 'c1' 'ck1' 'IPM.Appointment' '2026-07-01T10:00:00Z' $null $null $null $null
# IDs que só diferem em maiúsculas/minúsculas (aconteceu na caixa real: 2.157 de 9.812 itens)
$caso1 = New-FakeItem 'AAMkCaso1' 'ck1' 'IPM.Note' '2026-09-11T12:00:00Z' 'Caso maiúsculo' 'Sistema' '2026-09-11T12:00:00Z' 'false'
$caso2 = New-FakeItem 'aamKcASO1' 'ck1' 'IPM.Note' '2026-09-11T12:00:00Z' 'Caso minúsculo' 'Sistema' '2026-09-11T12:00:00Z' 'false'

# Execução 1: varredura completa, Inbox em 2 páginas
$global:mbx.Delta = @{
    'init-F1' = @{ value = @($m1); '@odata.nextLink' = 'p2-F1' }
    'p2-F1'   = @{ value = @($m2, $caso1, $caso2); '@odata.deltaLink' = 'd1-F1' }
    'init-F2' = @{ value = @($m3, $mBig, $m3); '@odata.deltaLink' = 'd1-F2' }   # m3 repetido no delta
    'init-F3' = @{ value = @($c1); '@odata.deltaLink' = 'd1-F3' }
}

function global:Get-MgContext {
    [pscustomobject]@{ Account = 'teste'; TenantId = 'tenant-1'; AuthType = 'Delegated'
        Scopes = @('User.Read.All','MailboxFolder.Read','MailboxItem.Read','MailboxItem.Export','MailboxFolder.ReadWrite',
                   'MailboxItem.ImportExport','Application.ReadWrite.All','AppRoleAssignment.ReadWrite.All') }
}
function global:Connect-MgGraph { }
function global:Import-Module { }
function global:Invoke-MgGraphRequest { param($Method = 'GET', $Uri, $Body, $Headers, $ContentType, $ErrorAction)
    $f = $global:mbx
    $b = if ($Body) { $Body | ConvertFrom-Json } else { $null }

    $chave = if ($Uri -match '/folders/([^/]+)/items/delta$') { "init-$([uri]::UnescapeDataString($Matches[1]))" } else { $Uri }
    if ($f.Delta.ContainsKey($chave)) {
        $resp = $f.Delta[$chave]
        if ($resp -is [scriptblock]) { return & $resp }
        return $resp
    }
    switch -Regex ($Uri) {
        'settings/exchange$' { return @{ primaryMailboxId = 'MBX:1@2' } }
        '\$batch$' {
            $resps = foreach ($r in $b.requests) {
                $id = [uri]::UnescapeDataString(($r.url -split '/items/')[1].Split('?')[0])
                $m  = $f.Itens[$id].Meta
                $props = @(
                    @{ id = 'String 0x37';      value = $m.Assunto }
                    @{ id = 'String 0xc1a';     value = $m.Remetente }
                    @{ id = 'SystemTime 0xe06'; value = $m.Recebido }
                    @{ id = 'Boolean 0xe1b';    value = $m.Anexo }
                ) | Where-Object { $null -ne $_.value }
                @{ id = $r.id; status = 200; body = @{ singleValueExtendedProperties = @($props) } }
            }
            return @{ responses = @($resps) }
        }
        'exportItems$' {
            $f.Exports.Add(@($b.itemIds))
            return @{ value = @($b.itemIds | ForEach-Object {
                if ($_ -in $f.Falha) { @{ itemId = $_; error = @{ message = 'ItemTooLarge' } } }
                else {
                    $ck = $f.Itens[$_].Item.changeKey
                    @{ itemId = $_; changeKey = $ck; data = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("FTS:$($_):$ck")) }
                }
            }) }
        }
        'createImportSession$' {
            return @{ importUrl = 'https://outlook.office365.com/import?authtoken=x'; expirationDateTime = (Get-Date).ToUniversalTime().AddHours(1).ToString('o') }
        }
        '/childFolders(\?|$)' {
            $pai = [uri]::UnescapeDataString(($Uri -split '/folders/')[1].Split('/')[0])
            if ($Method -eq 'POST') { return New-FakeFolder $b.displayName $b.type $pai }
            return @{ value = @(Get-FakeFolders $pai) }
        }
        '/folders(\?|$)' {
            if ($Method -eq 'POST') { return New-FakeFolder $b.displayName $b.type $null }
            return @{ value = @(Get-FakeFolders $null) }
        }
    }
    throw "Rota não simulada: $Method $Uri"
}


$bkp = Join-Path $tmp 'bkp'
$cx  = Join-Path $bkp 'fin@x.com'
function Get-Indice {
    $t = [hashtable]::new([StringComparer]::Ordinal)
    Import-Csv (Join-Path $cx 'indice.csv') -Delimiter ';' | ForEach-Object { $t[$_.Id] = $_ }
    $t
}
function Get-Delta($PastaId) { (@((Get-Content (Join-Path $cx 'estado.json') -Raw | ConvertFrom-Json).Deltas) | Where-Object { $_.PastaId -ceq $PastaId }).Link }
function Get-Conteudo($Id) { [IO.File]::ReadAllText((Join-Path $cx $indice[$Id].Arquivo)) }

& (Join-Path $raiz 'Backup-Caixa.ps1') -Caixas 'Fin@X.com' -Destino $bkp 3>$null 6>$null | Out-Null
$exit1      = $LASTEXITCODE
$indice     = Get-Indice
$exportados = @($global:mbx.Exports | ForEach-Object { $_ })

Assert ($indice.Count -eq 6)                                           "1ª execução: 6 itens no índice (o que falhou fica de fora)"
Assert ($indice.ContainsKey('AAMkCaso1') -and $indice.ContainsKey('aamKcASO1')) "IDs que só diferem em maiúsculas/minúsculas são itens distintos no índice"
Assert ($indice['AAMkCaso1'].Assunto -eq 'Caso maiúsculo' -and $indice['aamKcASO1'].Assunto -eq 'Caso minúsculo') "...e cada um mantém seus próprios metadados"
Assert ((Get-Conteudo 'aamKcASO1') -eq 'FTS:aamKcASO1:ck1')            "...e aponta para o próprio arquivo"
Assert (@($exportados | Where-Object { $_ -ceq 'm3' }).Count -eq 1)    "item repetido no delta é exportado uma vez só"
Assert ((Get-Conteudo 'm1') -eq 'FTS:m1:ck1')                          "grava o conteúdo exportado (.fts)"
Assert ($indice['m1'].Arquivo -like 'itens\2026\09\*.fts')             "organiza os arquivos por ano/mês"
Assert ($indice['m1'].Assunto -eq 'Boleto setembro' -and $indice['m1'].Remetente -eq 'Fornecedor A' -and $indice['m1'].Anexo -eq 'Sim') "índice com assunto, remetente e anexo (ids MAPI normalizados)"
Assert ($indice['m2'].Anexo -eq '')                                    "'false' do Graph não vira anexo"
Assert ($indice['m3'].Pasta -eq 'Inbox/Clientes')                      "guarda o caminho completo da pasta"
Assert ($indice['c1'].Tipo -eq 'IPM.Appointment' -and $indice['c1'].TipoPasta -eq 'IPF.Appointment') "inclui itens de calendário"
Assert ($indice.ContainsKey('m2'))                                     "segue a paginação do delta (@odata.nextLink)"
Assert (@($global:mbx.Exports | Where-Object { $_ -contains 'mBig' -and $_.Count -eq 1 }).Count -eq 1) "item grande é exportado num lote separado"
Assert ((Get-Content (Join-Path $cx 'falhas.log') -Raw) -match 'mBig.*ItemTooLarge') "registra falhas em falhas.log"
Assert ($exit1 -eq 1)                                                  "código de saída 1 quando há falhas (para o agendador)"
Assert ((Get-Delta 'F1') -eq 'd1-F1' -and (Get-Delta 'F3') -eq 'd1-F3') "salva os tokens de sincronização"
Assert ((Get-ChildItem (Join-Path $bkp 'logs') -Filter 'backup-*.log').Count -ge 1) "gera log da execução"

# Execução 2: incremental (m2 apagado, m1 alterado, m4 novo, token do Calendar expirado)
$m1v2 = New-FakeItem 'm1' 'ck2' 'IPM.Note' '2026-09-10T12:00:00Z' 'Boleto setembro' 'Fornecedor A' '2026-09-10T12:00:00Z' 'true'
$m4   = New-FakeItem 'm4' 'ck1' 'IPM.Note' '2026-10-01T12:00:00Z' 'Novo pedido' 'Cliente Y' '2026-10-01T12:00:00Z' 'false'
$global:mbx.Delta = @{
    'd1-F1'   = @{ value = @(@{ id = 'm2'; '@removed' = @{ reason = 'deleted' } }, $m1v2, $m4); '@odata.deltaLink' = 'd2-F1' }
    'd1-F2'   = @{ value = @(); '@odata.deltaLink' = 'd2-F2' }
    'd1-F3'   = { throw (New-HttpError 410) }
    'init-F3' = @{ value = @($c1); '@odata.deltaLink' = 'd2-F3' }
}
$global:mbx.Exports.Clear()
& (Join-Path $raiz 'Backup-Caixa.ps1') -Caixas 'fin@x.com' -Destino $bkp 3>$null 6>$null | Out-Null
$indice     = Get-Indice
$exportados = @($global:mbx.Exports | ForEach-Object { $_ })

Assert ($indice.Count -eq 7)                                            "2ª execução: item novo entra no índice"
Assert ($indice['m2'].RemovidoEm -and (Test-Path (Join-Path $cx $indice['m2'].Arquivo))) "item apagado na origem fica no backup, marcado com RemovidoEm"
Assert ((Get-Conteudo 'm1') -eq 'FTS:m1:ck2')                           "item alterado (novo ChangeKey) é exportado de novo"
Assert ($exportados -notcontains 'c1' -and $exportados -notcontains 'm3') "itens sem mudança não são exportados de novo"
Assert ((Get-Delta 'F3') -eq 'd2-F3')                                   "token expirado (410): refaz a varredura da pasta"

# Execução 3: backup feito pela versão 1 (estado sem versão, índice com entradas perdidas)
$linhas = Import-Csv (Join-Path $cx 'indice.csv') -Delimiter ';'
$linhas | Where-Object { $_.Id -cnotin 'm3', 'AAMkCaso1' } | Export-Csv (Join-Path $cx 'indice.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8
@{ Caixa = 'fin@x.com'; MailboxId = 'MBX:1@2'; Deltas = @{ F1 = 'd2-F1'; F2 = 'd2-F2'; F3 = 'd2-F3' } } |
    ConvertTo-Json | Set-Content (Join-Path $cx 'estado.json') -Encoding UTF8
$global:mbx.Delta = @{
    'init-F1' = @{ value = @($m1v2, $m4, $caso1, $caso2); '@odata.deltaLink' = 'd3-F1' }
    'init-F2' = @{ value = @($m3, $mBig); '@odata.deltaLink' = 'd3-F2' }
    'init-F3' = @{ value = @($c1); '@odata.deltaLink' = 'd3-F3' }
}
$global:mbx.Exports.Clear()
& (Join-Path $raiz 'Backup-Caixa.ps1') -Caixas 'fin@x.com' -Destino $bkp 3>$null 6>$null | Out-Null
$indice     = Get-Indice
$exportados = @($global:mbx.Exports | ForEach-Object { $_ } | Sort-Object)

Assert ($indice.ContainsKey('m3') -and $indice.ContainsKey('AAMkCaso1') -and $indice.Count -eq 7) "estado da versão 1: varredura completa recoloca no índice o que faltava"
Assert (($exportados -join ',') -ceq 'AAMkCaso1,m3,mBig')               "...exportando só os itens ausentes (o resto não é baixado de novo)"
Assert ((Get-Delta 'F1') -eq 'd3-F1')                                   "...e grava o estado no formato novo"


# ==============================================================================
Write-Host "`n== Restaurar-Backup.ps1 ==" -ForegroundColor Cyan
# ==============================================================================
$global:imports = New-Object System.Collections.Generic.List[object]
function global:Invoke-RestMethod { param($Method, $Uri, $Body, $ContentType, $ErrorAction)
    $global:imports.Add(($Body | ConvertFrom-Json))
    @{ itemId = 'novo'; changeKey = 'x' }
}
function Get-B64([string]$Texto) { [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Texto)) }
$restaurar = Join-Path $raiz 'Restaurar-Backup.ps1'

& $restaurar -Caixa 'fin@x.com' -Destino $bkp -Assunto boleto -PastaRestauracao 'Restaurados teste' -Todos -Force 3>$null 6>$null | Out-Null
$raizR = $global:mbx.Criadas | Where-Object { $_.displayName -eq 'Restaurados teste' }
$inboxR = $global:mbx.Criadas | Where-Object { $_.displayName -eq 'Inbox' -and $_.pai -eq $raizR.id }
Assert ($global:imports.Count -eq 1 -and $global:imports[0].Data -eq (Get-B64 'FTS:m1:ck2')) "filtra por assunto e envia o conteúdo do backup"
Assert ($global:imports[0].Mode -eq 'create')                          "importa em modo 'create' (não sobrescreve nada)"
Assert ($inboxR -and $global:imports[0].FolderId -eq $inboxR.id)       "recria a pasta original dentro de 'Restaurados'"

$global:imports.Clear()
& $restaurar -Caixa 'fin@x.com' -Destino $bkp -SomenteRemovidos -SemEstrutura -PastaRestauracao 'Restaurados teste' -Todos -Force 3>$null 6>$null | Out-Null
Assert ($global:imports.Count -eq 1 -and $global:imports[0].Data -eq (Get-B64 'FTS:m2:ck1')) "-SomenteRemovidos traz só o que foi apagado na origem"
Assert ($global:imports[0].FolderId -eq $raizR.id)                     "-SemEstrutura restaura direto na pasta de restauração"
Assert (@($global:mbx.Criadas | Where-Object displayName -eq 'Restaurados teste').Count -eq 1) "reaproveita a pasta de restauração existente"

$global:imports.Clear()
& $restaurar -Caixa 'fin@x.com' -Destino $bkp -De '2026-09-01' -Ate '2026-09-30' -PastaRestauracao 'Restaurados teste' -Todos -Force 3>$null 6>$null | Out-Null
Assert ($global:imports.Count -eq 4)                                   "filtro por período (-De/-Ate) traz os 4 itens de setembro"

$global:imports.Clear()
& $restaurar -Caixa 'fin@x.com' -Destino $bkp -Assunto 'nao-existe' -Todos -Force 3>$null 6>$null | Out-Null
Assert ($global:imports.Count -eq 0)                                   "sem resultados: não conecta nem importa nada"

Remove-Mocks 'Invoke-RestMethod'


# ==============================================================================
Write-Host "`n== Criar-AppBackup.ps1 ==" -ForegroundColor Cyan
# ==============================================================================
$global:app = @{ Body = $null; Atribuicoes = New-Object System.Collections.Generic.List[object] }
function global:New-SelfSignedCertificate { [pscustomobject]@{ Thumbprint = 'ABC123'; RawData = [byte[]](1, 2, 3); NotAfter = (Get-Date).AddYears(2) } }
function global:Invoke-MgGraphRequest { param($Method = 'GET', $Uri, $Body, $Headers, $ContentType, $ErrorAction)
    $b = if ($Body) { $Body | ConvertFrom-Json } else { $null }
    switch -Regex ($Uri) {
        "servicePrincipals\(appId=" {
            $nomes = 'User.Read.All','MailboxFolder.Read.All','MailboxItem.Read.All','MailboxItem.Export.All','MailboxFolder.ReadWrite.All','MailboxItem.ImportExport.All','Mail.Read'
            return @{ id = 'graph-sp'; appRoles = @($nomes | ForEach-Object { @{ id = "role-$_"; value = $_; allowedMemberTypes = @('Application') } }) }
        }
        'v1.0/applications$'      { $global:app.Body = $b; return @{ appId = 'app-1'; id = 'obj-1' } }
        'v1.0/servicePrincipals$' { return @{ id = 'sp-1' } }
        'appRoleAssignments$'     { $global:app.Atribuicoes.Add($b); return @{} }
    }
    throw "Rota não simulada: $Method $Uri"
}
$criarApp = Join-Path $raiz 'Criar-AppBackup.ps1'

& $criarApp 6>$null | Out-Null
$acessos = @($global:app.Body.requiredResourceAccess[0].resourceAccess)
Assert ($acessos.Count -eq 4 -and @($acessos | Where-Object type -eq 'Role').Count -eq 4) "pede só as 4 permissões de leitura/exportação"
Assert ($global:app.Body.keyCredentials[0].key -eq 'AQID')              "registra o certificado (chave pública) no app"
Assert ($global:app.Atribuicoes.Count -eq 4 -and $global:app.Atribuicoes[0].principalId -eq 'sp-1' -and $global:app.Atribuicoes[0].resourceId -eq 'graph-sp') "concede o consentimento de administrador"

$global:app.Atribuicoes.Clear()
& $criarApp -PermitirRestauracao -SemConsentimento 3>$null 6>$null | Out-Null
Assert (@($global:app.Body.requiredResourceAccess[0].resourceAccess).Count -eq 6) "-PermitirRestauracao acrescenta as permissões de escrita"
Assert ($global:app.Atribuicoes.Count -eq 0)                            "-SemConsentimento não concede permissões"

Remove-Mocks 'New-SelfSignedCertificate','Invoke-MgGraphRequest','Get-MgContext','Connect-MgGraph','Import-Module','Get-FakeFolders','New-FakeFolder','Install-Module'





# ==============================================================================
Write-Host "`n== Exportar-PST.ps1 ==" -ForegroundColor Cyan
# ==============================================================================
# Purview simulado: caso, pesquisa, custodiante, estimativa, exportação e download
$zipFake = Join-Path $tmp 'pv-pacote.zip'
$dirZip  = Join-Path $tmp 'pv-pacote'
New-Item -ItemType Directory (Join-Path $dirZip 'Exchange') -Force | Out-Null
Set-Content (Join-Path $dirZip 'Exchange\ex@x.com.pst') 'PST simulado'
Set-Content (Join-Path $dirZip 'Summary.csv') 'relatorio'
Compress-Archive -Path (Join-Path $dirZip '*') -DestinationPath $zipFake -Force

function New-ErroToken([string]$Codigo) {
    $er = [System.Management.Automation.ErrorRecord]::new([Exception]::new('HTTP 400'), 'Token', 'InvalidOperation', $null)
    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new("{`"error`":`"$Codigo`"}")
    $er
}
# Mesmo formato do Invoke-MgGraphRequest real: requisição HTTP inteira + JSON no final
function New-ErroGraph([string]$Codigo, [string]$Mensagem) {
    $json = @{ error = @{ code = $Codigo; message = $Mensagem } } | ConvertTo-Json -Compress
    $dump = "POST https://graph.microsoft.com/v1.0/security/cases/ediscoveryCases`nHTTP/1.1 400 Bad Request`nclient-request-id: 163fac04-7c26-43d0-958c-403d23127f74`n`n$json"
    $er = [System.Management.Automation.ErrorRecord]::new([Exception]::new('Response status code does not indicate success: BadRequest (Bad Request).'), 'Graph', 'InvalidOperation', $null)
    $er.ErrorDetails = [System.Management.Automation.ErrorDetails]::new($dump)
    $er
}
function Reset-Purview {
    $global:pv = @{
        Chamadas = [System.Collections.Generic.List[object]]::new()
        Downloads = [System.Collections.Generic.List[object]]::new()
        Estimativa = 0; Export = 0; Token = 0
        Itens = 120; FalharExport = $false; StatusExport = 'succeeded'; FalharCaso = $false
        AceitaFonteEmbutida = $true; FalharDownload = $false; ErroLogin = $null; TokenBody = $null; UrlLogin = $null
    }
}

function global:Get-MgContext { [pscustomobject]@{ Account = 'admin@x.com'; TenantId = 'tenant-1'; AuthType = 'Delegated'; Scopes = @('eDiscovery.ReadWrite.All') } }
function global:Connect-MgGraph { }
function global:Import-Module { }
function global:Install-Module { }
function global:Start-Sleep { }
function global:Invoke-MgGraphRequest { param($Method = 'GET', $Uri, $Body, $Headers, $ContentType, $ErrorAction)
    $b = if ($Body) { $Body | ConvertFrom-Json } else { $null }
    $global:pv.Chamadas.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $b; Headers = $Headers })
    switch -Regex ($Uri) {
        'ediscoveryCases$' {
            if ($Method -eq 'GET') {
                return @{ value = @(
                    @{ id = 'outro'; displayName = 'Outro caso';                              createdDateTime = '2026-10-09T12:00:00Z' }
                    @{ id = 'caso0'; displayName = 'Exportação PST - ex@x.com - 2026-10-01 0900'; createdDateTime = '2026-10-01T09:00:00Z' }
                    @{ id = 'caso2'; displayName = 'Exportação PST - ex@x.com - 2026-10-09 1215'; createdDateTime = '2026-10-09T12:15:00Z' }
                ) }
            }
            if ($global:pv.FalharCaso) { throw (New-ErroGraph 'BadRequest' 'Invalid case displayName') }
            return @{ id = 'caso1' }
        }
        '/searches$' {
            $temFonte = ($b.PSObject.Properties.Name -contains 'additionalSources' -and $global:pv.AceitaFonteEmbutida) -or
                        ($b.PSObject.Properties.Name -contains 'custodianSources@odata.bind')
            if (-not $temFonte) { throw (New-ErroGraph 'BadRequest' 'At least one data source is required.') }
            return @{ id = 'pesq1' }
        }
        'v1.0/servicePrincipals$'           { $global:pv.SPRegistrado = $true; $global:pv.SPBody = $b; return @{ id = 'sp1' } }
        '/custodians$'                       { return @{ id = 'cust1' } }
        '/custodians/cust1/userSources$'     { return @{ id = 'us1'; includedSources = 'mailbox' } }
        '/custodians/cust1/release$'         { return $null }
        '/estimateStatistics$'               { return @{} }
        '/lastEstimateStatisticsOperation$' {
            $global:pv.Estimativa++
            if ($global:pv.Estimativa -eq 1) { throw (New-HttpError 404) }   # ainda sendo criada
            if ($global:pv.Estimativa -eq 2) { return @{ status = 'running'; percentProgress = 40 } }
            return @{ status = 'succeeded'; indexedItemCount = $global:pv.Itens; indexedItemsSize = 1073741824; unindexedItemCount = $(if ($global:pv.Itens) { 2 } else { 0 }) }
        }
        '/exportResult$' {
            if ($global:pv.FalharExport) { throw (New-ErroGraph 'Forbidden' 'Usage of eDiscovery APIs requires a subscription to Purview pay-as-you-go billing.') }
            return $null
        }
        '/operations$' {
            return @{ value = @(
                @{ id = 'opEst'; action = 'estimateStatistics'; createdDateTime = '2026-10-09T10:00:00Z' }
                @{ id = 'opExp'; action = 'exportResult';       createdDateTime = '2026-10-09T10:05:00Z'; status = 'succeeded' }
            ) }
        }
        '/operations/opExp$' {
            $global:pv.Export++
            if ($global:pv.Export -eq 1) { return @{ status = 'running'; percentProgress = 50 } }
            return @{ status = $global:pv.StatusExport; resultInfo = @{ message = 'Falha simulada' }
                      exportFileMetadata = $(if ($global:pv.SemArquivos) { @() } else { @(@{ fileName = 'Exports.zip'; downloadUrl = 'https://proxy.purview/exp/abc'; size = 2048 }) }) }
        }
    }
    throw "Rota não simulada: $Method $Uri"
}
function global:Invoke-RestMethod { param($Method, $Uri, $Body, $ErrorAction)
    if ($Uri -like '*/devicecode') {
        $global:pv.Escopo = $Body.scope
        return [pscustomobject]@{ message = 'Acesse https://microsoft.com/devicelogin e use o código ABC'; device_code = 'dc1'; interval = 1; expires_in = 900 }
    }
    if ($Uri -like '*/token' -and $Body.grant_type -eq 'authorization_code') {
        $global:pv.TokenBody = $Body
        return [pscustomobject]@{ access_token = 'tok123' }
    }
    if ($Uri -like '*/token') {
        $global:pv.Token++
        if ($global:pv.Token -eq 1) { throw (New-ErroToken 'authorization_pending') }
        return [pscustomobject]@{ access_token = 'tok123' }
    }
    throw "Rota não simulada: $Uri"
}
function global:Start-Process { param($FilePath)
    $global:pv.UrlLogin = $FilePath
    $q = @{}
    foreach ($par in ($FilePath -split '\?', 2)[1] -split '&') { $k, $v = $par -split '=', 2; $q[$k] = [uri]::UnescapeDataString($v) }
    if (-not $global:clienteHttp) { Add-Type -AssemblyName System.Net.Http; $global:clienteHttp = [System.Net.Http.HttpClient]::new() }
    $resposta = if ($global:pv.ErroLogin -and -not $global:pv.SPRegistrado) { "error=access_denied&error_description=$([uri]::EscapeDataString($global:pv.ErroLogin))&state=$($q.state)" }
                else { "code=abc&state=$($q.state)" }
    $global:pv.RespostaLogin = $global:clienteHttp.GetAsync("$($q.redirect_uri)?$resposta")
}
function global:Invoke-WebRequest { param($Uri, $Headers, $OutFile, $ErrorAction, [switch]$UseBasicParsing)
    if ($global:pv.FalharDownload) { throw 'Falha simulada no download' }
    $global:pv.Downloads.Add([pscustomobject]@{ Uri = $Uri; Headers = $Headers })
    Copy-Item $zipFake $OutFile
}

$exportar = Join-Path $raiz 'Exportar-PST.ps1'
function Get-Chamada([string]$Padrao) { $global:pv.Chamadas | Where-Object { $_.Uri -match $Padrao } | Select-Object -First 1 }
function Get-Avisos([scriptblock]$Bloco) { & $Bloco 3>&1 6>$null | Where-Object { $_ -is [System.Management.Automation.WarningRecord] } }

# 1) Exportação completa com download (fonte embutida na pesquisa)
Reset-Purview
$dirPst = Join-Path $tmp 'pv1'
& $exportar -Caixa 'ex@x.com' -Destino $dirPst 3>$null 6>$null | Out-Null
$exitPv   = $LASTEXITCODE
$caso     = Get-Chamada 'ediscoveryCases$'
$pesquisa = Get-Chamada '/searches$'
$exp      = Get-Chamada '/exportResult$'
$down     = $global:pv.Downloads | Select-Object -First 1
$psts     = @(Get-ChildItem $dirPst -Recurse -Filter *.pst)

Assert ($caso.Body.displayName -like 'Exportação PST - ex@x.com - *')    "cria um caso de eDiscovery identificado com a caixa"
Assert (-not ($pesquisa.Body.PSObject.Properties.Name -contains 'contentQuery')) "pesquisa sem condições: não envia contentQuery (caixa inteira)"
Assert ($pesquisa.Body.additionalSources[0].email -eq 'ex@x.com' -and $pesquisa.Body.additionalSources[0].includedSources -eq 'mailbox') "a pesquisa já nasce com a caixa do colaborador como fonte (só caixa de correio)"
Assert (-not (Get-Chamada '/custodians'))                                "com fonte embutida aceita: não cria custodiante (nem retenção)"
Assert ($global:pv.Estimativa -ge 3)                                     "aguarda a estimativa (404 inicial e 'running' não são erro)"
Assert ($exp.Body.exportFormat -eq 'pst' -and $exp.Body.additionalOptions -match 'includeFolderAndPath') "exporta em PST com a estrutura de pastas original"
Assert ($exp.Body.exportCriteria -match 'partiallyIndexed')              "inclui itens não indexados (ex: anexos criptografados)"
Assert ($exp.Headers.Prefer -eq 'include-unknown-enum-members')          "envia o cabeçalho das enumerações novas (exportResult, includeFolderAndPath)"
Assert ($global:pv.UrlLogin -match 'scope=b26e684c-5068-4120-a679-64a5d2c909d9%2FeDiscovery.Download.Read' -and $global:pv.UrlLogin -match 'code_challenge_method=S256' -and $global:pv.UrlLogin -match 'login_hint=admin%40x.com') "login do download no navegador, no serviço do Purview (PKCE, conta sugerida)"
Assert ($global:pv.TokenBody.code -eq 'abc' -and $global:pv.TokenBody.code_verifier -and $global:pv.TokenBody.redirect_uri -like 'http://localhost:*/' -and $global:pv.Token -eq 0) "troca o código recebido em localhost pelo token (sem código de dispositivo)"
Assert ($down.Uri -eq 'https://proxy.purview/exp/abc' -and $down.Headers.Authorization -eq 'Bearer tok123' -and $down.Headers.'X-AllowWithAADToken' -eq 'true') "baixa o arquivo com o token e o cabeçalho exigidos"
Assert ($psts.Count -eq 1 -and $psts[0].Name -eq 'ex@x.com.pst')         "extrai o PST do pacote baixado"
Assert (@(Get-ChildItem $dirPst -Recurse -Filter *.zip).Count -eq 0)    "remove o .zip depois de extrair"
Assert ($exitPv -eq 0)                                                   "código de saída 0"

# 1b) API exige a fonte por custodiante (o erro visto no tenant real)
Reset-Purview
$global:pv.AceitaFonteEmbutida = $false
& $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv1b') -SemDownload 3>$null 6>$null | Out-Null
$exitCust = $LASTEXITCODE
$us       = Get-Chamada '/custodians/cust1/userSources$'
$vinculo  = @($global:pv.Chamadas | Where-Object { $_.Uri -match '/searches$' })[-1].Body.'custodianSources@odata.bind'
Assert ((Get-Chamada '/custodians$').Body.email -eq 'ex@x.com' -and $us.Body.includedSources -eq 'mailbox') "plano B: custodiante com fonte só da caixa de correio"
Assert ($vinculo -contains 'https://graph.microsoft.com/v1.0/security/cases/ediscoveryCases/caso1/custodians/cust1/userSources/us1') "plano B: pesquisa vinculada à fonte do custodiante"
Assert ((Get-Chamada '/custodians/cust1/release$') -and $exitCust -eq 0) "plano B: libera o custodiante no final (sem retenção na caixa)"

# 1c) Erro no meio com custodiante: libera mesmo assim
Reset-Purview
$global:pv.AceitaFonteEmbutida = $false; $global:pv.FalharExport = $true
& $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv1c') 3>$null 6>$null | Out-Null
Assert ($LASTEXITCODE -eq 1 -and (Get-Chamada '/custodians/cust1/release$')) "erro na exportação: o custodiante é liberado mesmo assim"

# 1d) Erro do Graph: mostra a etapa e a mensagem, sem falso alarme de permissão
Reset-Purview
$global:pv.FalharCaso = $true
$avisos = Get-Avisos { & $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv1d') }
Assert ($LASTEXITCODE -eq 1 -and @($avisos | Where-Object { $_.Message -match "etapa 'criar caso': BadRequest: Invalid case displayName" }).Count -eq 1) "erro do Graph: mostra a etapa e o código/mensagem extraídos do JSON"
Assert (-not ($avisos | Where-Object Message -match 'eDiscovery Manager'))  "'403' dentro do client-request-id não vira falso alarme de permissão"

# 1e) -SomenteDownload: pega o caso mais recente da caixa e não exporta de novo
Reset-Purview
$global:pv.Export = 1   # exportação já concluída no Purview
$dirRet = Join-Path $tmp 'pv1e'
& $exportar -Caixa 'ex@x.com' -Destino $dirRet -SomenteDownload 3>$null 6>$null | Out-Null
Assert ($LASTEXITCODE -eq 0 -and -not ($global:pv.Chamadas | Where-Object { $_.Method -eq 'POST' }))   "-SomenteDownload: não cria caso nem exportação"
Assert ((Get-Chamada '/caso2/operations$') -and -not (Get-Chamada '/caso0/'))                          "-SomenteDownload: usa o caso mais recente daquela caixa"
Assert (@(Get-ChildItem $dirRet -Recurse -Filter *.pst).Count -eq 1 -and (Test-Path (Join-Path $dirRet 'Exportação PST - ex@x.com - 2026-10-09 1215'))) "-SomenteDownload: baixa e extrai na pasta com o nome do caso"

# 1f) -CodigoDispositivo
Reset-Purview
$global:pv.Export = 1
& $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv1f') -SomenteDownload -CodigoDispositivo 3>$null 6>$null | Out-Null
Assert ($LASTEXITCODE -eq 0 -and $global:pv.Escopo -eq 'b26e684c-5068-4120-a679-64a5d2c909d9/eDiscovery.Download.Read' -and $global:pv.Token -eq 2 -and -not $global:pv.UrlLogin) "-CodigoDispositivo: login por código, aguardando a autorização"

# 1g) Download falha: orienta a retomar sem exportar de novo
Reset-Purview
$global:pv.FalharDownload = $true
$avisos = Get-Avisos { & $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv1g') }
Assert ($LASTEXITCODE -eq 1 -and @($avisos | Where-Object Message -match '-SomenteDownload').Count -ge 1) "download com falha: código 1 e instrução para retomar com -SomenteDownload"

# 1h) Serviço de download não registrado no tenant: administrador recusa o registro
Reset-Purview
$global:pv.Export = 1
$global:pv.ErroLogin = 'AADSTS650052: The app needs access to a service that your organization has not subscribed to or enabled.'
function global:Read-Host { 'N' }
$avisos = Get-Avisos { & $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv1h') -SomenteDownload }
Assert ($LASTEXITCODE -eq 1 -and -not $global:pv.SPRegistrado -and @($avisos | Where-Object Message -match 'Invoke-MgGraphRequest -Method POST -Uri v1.0/servicePrincipals').Count -eq 1) "serviço de download ausente e registro recusado: não altera o tenant e mostra o comando manual"
Remove-Mocks 'Read-Host'

# 1j) -RegistrarServicoDownload: registra, faz o login de novo e baixa
Reset-Purview
$global:pv.Export = 1
$global:pv.ErroLogin = 'AADSTS650052: The app needs access to a service that your organization has not subscribed to or enabled.'
$dirReg = Join-Path $tmp 'pv1j'
& $exportar -Caixa 'ex@x.com' -Destino $dirReg -SomenteDownload -RegistrarServicoDownload 3>$null 6>$null | Out-Null
Assert ($LASTEXITCODE -eq 0 -and $global:pv.SPBody.appId -eq 'b26e684c-5068-4120-a679-64a5d2c909d9' -and @(Get-ChildItem $dirReg -Recurse -Filter *.pst).Count -eq 1) "-RegistrarServicoDownload: registra o serviço do Purview e conclui o download"

# 1i) Exportação antiga, já sem arquivos
Reset-Purview
$global:pv.Export = 1; $global:pv.SemArquivos = $true
$erro = $null
try { & $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv1i') -SomenteDownload 3>$null 6>$null | Out-Null } catch { $erro = $_ }
Assert ($erro -and "$erro" -match 'expiram' -and $global:pv.Downloads.Count -eq 0) "exportação expirada (sem arquivos): avisa em vez de baixar arquivo vazio"

# 1k) -Conteudo, -De, -Ate e -Consulta viram uma consulta KQL na pesquisa
Reset-Purview
& $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv1k') -Conteudo Email, Calendario -De '2026-01-01' -Ate '2026-06-30' -Consulta 'from:fornecedor.com' -SemDownload 3>$null 6>$null | Out-Null
$pesqK = @($global:pv.Chamadas | Where-Object { $_.Uri -match '/searches$' })[-1]
Assert ($LASTEXITCODE -eq 0 -and $pesqK.Body.contentQuery -eq '(kind:email OR kind:meetings) AND received>=2026-01-01 AND received<=2026-06-30 AND (from:fornecedor.com)') "-Conteudo/-De/-Ate/-Consulta: monta a consulta KQL da pesquisa"
Assert ($pesqK.Body.displayName -eq 'Seleção: Email, Calendario - ex@x.com')                    "pesquisa identifica a seleção no nome"
Assert ((Get-Chamada '/exportResult$').Body.exportCriteria -eq 'searchHits')                    "com filtro: não inclui itens não indexados (não dá para saber se correspondem)"

# 1l) -Conteudo Teams com -IncluirNaoIndexados
Reset-Purview
& $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv1l') -Conteudo Teams -IncluirNaoIndexados -SemDownload 3>$null 6>$null | Out-Null
Assert ((Get-Chamada '/searches$').Body.contentQuery -eq '(kind:microsoftteams)' -and (Get-Chamada '/exportResult$').Body.exportCriteria -match 'partiallyIndexed') "-Conteudo Teams + -IncluirNaoIndexados"

# 1m) Período invertido
Reset-Purview
$erro = $null
try { & $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv1m') -De '2026-07-01' -Ate '2026-01-01' 3>$null 6>$null | Out-Null } catch { $erro = $_ }
Assert ($erro -and "$erro" -match 'posterior' -and -not (Get-Chamada 'ediscoveryCases$'))     "-De depois de -Ate: erro antes de criar qualquer coisa no Purview"

# 2) -SemDownload
Reset-Purview
& $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv2') -SemDownload 3>$null 6>$null | Out-Null
Assert ($LASTEXITCODE -eq 0 -and $global:pv.Downloads.Count -eq 0 -and $global:pv.Token -eq 0) "-SemDownload: prepara a exportação e não baixa"

# 3) Tenant sem pay-as-you-go
Reset-Purview
$global:pv.FalharExport = $true
$avisos = Get-Avisos { & $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv3') }
Assert ($LASTEXITCODE -eq 1 -and @($avisos | Where-Object Message -match 'pay-as-you-go').Count -ge 1 -and @($avisos | Where-Object Message -match 'Alternativa manual').Count -eq 1) "sem pay-as-you-go: explica o que ativar e o caminho manual pelo portal"

# 4) Pesquisa vazia
Reset-Purview
$global:pv.Itens = 0
& $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv4') 3>$null 6>$null | Out-Null
Assert ($LASTEXITCODE -eq 1 -and -not (Get-Chamada '/exportResult$'))    "pesquisa sem resultados: não exporta"

# 5) Exportação falha no Purview
Reset-Purview
$global:pv.StatusExport = 'failed'
& $exportar -Caixa 'ex@x.com' -Destino (Join-Path $tmp 'pv5') 3>$null 6>$null | Out-Null
Assert ($LASTEXITCODE -eq 1 -and $global:pv.Downloads.Count -eq 0)       "exportação com falha: código de saída 1, sem download"

Remove-Mocks 'Get-MgContext','Connect-MgGraph','Import-Module','Install-Module','Start-Sleep','Invoke-MgGraphRequest','Invoke-RestMethod','Invoke-WebRequest','Start-Process'




# ==============================================================================
Write-Host "`n== Menu.ps1 ==" -ForegroundColor Cyan
# ==============================================================================
# Scripts falsos com os MESMOS parâmetros dos reais: registram a chamada em vez de executar
$dirFalsos = Join-Path $tmp 'menu-scripts'
New-Item -ItemType Directory $dirFalsos -Force | Out-Null
foreach ($s in Get-ChildItem $raiz -Filter *.ps1 | Where-Object Name -ne 'Menu.ps1') {
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($s.FullName, [ref]$null, [ref]$null)
    $params = @($(if ($ast.ParamBlock) { $ast.ParamBlock.Parameters }) | ForEach-Object {
        $tipo = if ($_.StaticType -eq [switch]) { '[switch]' } else { '' }
        "$tipo`$$($_.Name.VariablePath.UserPath)"
    })
    @"
param($($params -join ', '))
`$global:menuChamadas.Add([pscustomobject]@{ Script = '$($s.Name)'; Params = [hashtable]::new(`$PSBoundParameters) })
if (`$global:menuFalhar) { throw 'falha simulada' }
exit 0
"@ | Set-Content (Join-Path $dirFalsos $s.Name) -Encoding UTF8
}

function global:Read-Host { param($Prompt) if ($global:respostas.Count -eq 0) { throw "Sem resposta para: $Prompt" }; $global:respostas.Dequeue() }
function Invoke-Menu([string[]]$Respostas) {
    $global:respostas    = [System.Collections.Generic.Queue[string]]::new([string[]]$Respostas)
    $global:menuChamadas = [System.Collections.Generic.List[object]]::new()
    & (Join-Path $raiz 'Menu.ps1') -PastaScripts $dirFalsos 6>&1 | Out-String
}
$menuItens = @{ NewClone = '1'; Sharepoint = '2'; Auditoria = '3'; Extracao = '4'; Recuperar = '5'; Backup = '6'; Restaurar = '7'; PST = '8'; App = '9' }

# 1) Exportar PST com seleção de conteúdo e período (e um e-mail inválido no meio)
$saida = Invoke-Menu @($menuItens.PST, 'abc', 'ex@x.com', 'N', '1,2', '2026-01-01', '', '', '', '', '0')
$c = $global:menuChamadas[0]
Assert ($global:menuChamadas.Count -eq 1 -and $c.Script -eq 'Exportar-PST.ps1')                     "menu: chama o script escolhido"
Assert ($c.Params.Caixa -eq 'ex@x.com' -and $saida -match 'Resposta inválida')                     "menu: rejeita e-mail inválido e pergunta de novo"
Assert ((@($c.Params.Conteudo) -join ',') -eq 'Email,Calendario' -and $c.Params.De -eq [datetime]'2026-01-01') "menu: escolha múltipla e data viram parâmetros"
Assert (-not $c.Params.ContainsKey('Ate') -and -not $c.Params.ContainsKey('Destino') -and -not $c.Params.ContainsKey('SomenteDownload')) "menu: respostas em branco usam o padrão do script"
Assert ($saida -match [regex]::Escape(".\Exportar-PST.ps1 -Caixa ex@x.com -Conteudo Email, Calendario -De 2026-01-01")) "menu: mostra o comando equivalente"

# 2) PST só download: não pergunta conteúdo nem período
$saida = Invoke-Menu @($menuItens.PST, 'ex@x.com', 'S', '', '', '', '0')
$c = $global:menuChamadas[0]
Assert ($c.Params.SomenteDownload -and -not $c.Params.ContainsKey('Conteudo') -and -not $c.Params.ContainsKey('De')) "menu: -SomenteDownload pula as perguntas de conteúdo e período"

# 3) Lista de caixas (vírgula, ponto e vírgula ou espaço)
$saida = Invoke-Menu @($menuItens.Auditoria, 'a@x.com; b@x.com', '', '', '0')
Assert ((@($global:menuChamadas[0].Params.Caixas) -join ',') -eq 'a@x.com,b@x.com')                   "menu: lista de caixas vira array"

# 4) Cancelar na confirmação
$saida = Invoke-Menu @($menuItens.Backup, 'a@x.com', '', 'N', '0')
Assert ($global:menuChamadas.Count -eq 0)                                                          "menu: 'N' na confirmação não executa"

# 5) Erro no script: mostra e volta ao menu
$global:menuFalhar = $true
$saida = Invoke-Menu @($menuItens.NewClone, '', '', '0')
$global:menuFalhar = $false
Assert ($saida -match 'Erro: falha simulada' -and $global:respostas.Count -eq 0)                    "menu: erro no script é mostrado e o menu continua"

# 6) Opção inválida e todos os itens apontando para scripts que existem
$saida = Invoke-Menu @('99', 'x', '0')
Assert ($global:respostas.Count -eq 0)                                                             "menu: ignora opção inválida"
$menuAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $raiz 'Menu.ps1'), [ref]$null, [ref]$null)
$scriptsMenu = @($menuAst.FindAll({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $args[0].Value -like '*.ps1' }, $true) | ForEach-Object Value | Select-Object -Unique)
Assert (-not ($scriptsMenu | Where-Object { -not (Test-Path (Join-Path $raiz $_)) }) -and $scriptsMenu.Count -eq 9) "menu: os 9 itens apontam para scripts existentes"

# 7) Cada pergunta do menu corresponde a um parâmetro real do script
$problemas = foreach ($s in $scriptsMenu) {
    $real = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $raiz $s), [ref]$null, [ref]$null)
    $nomes = @($real.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
    $bloco = $menuAst.FindAll({ $args[0] -is [System.Management.Automation.Language.HashtableAst] -and $args[0].Extent.Text -match "Script = '$([regex]::Escape($s))'" }, $true) | Select-Object -First 1
    foreach ($m in [regex]::Matches($bloco.Extent.Text, "Nome = '(\w+)'")) {
        if ($nomes -notcontains $m.Groups[1].Value) { "$s -$($m.Groups[1].Value)" }
    }
}
Assert (@($problemas).Count -eq 0)                                                                 "menu: toda pergunta vira um parâmetro que existe no script ($(@($problemas) -join ', '))"

Remove-Mocks 'Read-Host'

# ==============================================================================
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
$cor = if ($script:falhas) { 'Red' } else { 'Green' }
Write-Host "`n$($script:total - $script:falhas)/$($script:total) testes passaram." -ForegroundColor $cor
exit $script:falhas
