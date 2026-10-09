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
    Itens   = @{}
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

# Execução 1: varredura completa, Inbox em 2 páginas
$global:mbx.Delta = @{
    'init-F1' = @{ value = @($m1); '@odata.nextLink' = 'p2-F1' }
    'p2-F1'   = @{ value = @($m2); '@odata.deltaLink' = 'd1-F1' }
    'init-F2' = @{ value = @($m3, $mBig); '@odata.deltaLink' = 'd1-F2' }
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
& (Join-Path $raiz 'Backup-Caixa.ps1') -Caixas 'Fin@X.com' -Destino $bkp 3>$null 6>$null | Out-Null
$exit1  = $LASTEXITCODE
$indice = @{}; Import-Csv (Join-Path $cx 'indice.csv') -Delimiter ';' | ForEach-Object { $indice[$_.Id] = $_ }
$estado = Get-Content (Join-Path $cx 'estado.json') -Raw | ConvertFrom-Json
function Get-Conteudo($Id) { [IO.File]::ReadAllText((Join-Path $cx $indice[$Id].Arquivo)) }

Assert ($indice.Count -eq 4)                                           "1ª execução: 4 itens no índice (o que falhou fica de fora)"
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
Assert ($estado.Deltas.F1 -eq 'd1-F1' -and $estado.Deltas.F3 -eq 'd1-F3') "salva os tokens de sincronização"
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
$indice = @{}; Import-Csv (Join-Path $cx 'indice.csv') -Delimiter ';' | ForEach-Object { $indice[$_.Id] = $_ }
$estado = Get-Content (Join-Path $cx 'estado.json') -Raw | ConvertFrom-Json
$exportados = @($global:mbx.Exports | ForEach-Object { $_ })

Assert ($indice.Count -eq 5)                                            "2ª execução: item novo entra no índice"
Assert ($indice['m2'].RemovidoEm -and (Test-Path (Join-Path $cx $indice['m2'].Arquivo))) "item apagado na origem fica no backup, marcado com RemovidoEm"
Assert ((Get-Conteudo 'm1') -eq 'FTS:m1:ck2')                           "item alterado (novo ChangeKey) é exportado de novo"
Assert ($exportados -notcontains 'c1' -and $exportados -notcontains 'm3') "itens sem mudança não são exportados de novo"
Assert ($estado.Deltas.F3 -eq 'd2-F3')                                  "token expirado (410): refaz a varredura da pasta"


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
Assert ($global:imports.Count -eq 2)                                   "filtro por período (-De/-Ate) traz os 2 itens de setembro"

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
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
$cor = if ($script:falhas) { 'Red' } else { 'Green' }
Write-Host "`n$($script:total - $script:falhas)/$($script:total) testes passaram." -ForegroundColor $cor
exit $script:falhas
