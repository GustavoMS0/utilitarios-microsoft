# Utilitários Microsoft 365

Scripts PowerShell para administração e auditoria de **Entra ID**, **Exchange Online** e **SharePoint Online**.

| Script | Para que serve | Conecta em |
|---|---|---|
| [NewClone.ps1](NewClone.ps1) | Cria um usuário clonando grupos, licenças e listas de distribuição de um usuário modelo | Graph + Exchange Online |
| [Sharepoint.ps1](Sharepoint.ps1) | Auditoria de uso e governança dos sites do SharePoint (e OneDrive) | Graph (+ SPO opcional) |
| [Auditoriaexchange.ps1](Auditoriaexchange.ps1) | Liga/ajusta a auditoria em caixas de correio específicas | Exchange Online |
| [ExtracaoExchange.ps1](ExtracaoExchange.ps1) | Relatório de quem moveu/apagou e-mails nessas caixas | Exchange Online |
| [Recuperar-ItensExcluidos.ps1](Recuperar-ItensExcluidos.ps1) | Restaura e-mails/itens apagados (dentro do prazo de retenção), escolhendo um a um | Exchange Online |
| [Backup-Caixa.ps1](Backup-Caixa.ps1) | Backup incremental de caixas com fidelidade total + índice pesquisável | Graph |
| [Restaurar-Backup.ps1](Restaurar-Backup.ps1) | Restauração granular a partir do backup (por assunto, remetente, data, pasta) | Graph |
| [Criar-AppBackup.ps1](Criar-AppBackup.ps1) | Cria o app registration com certificado para o backup rodar agendado | Graph |
| [Exportar-PST.ps1](Exportar-PST.ps1) | Exporta a caixa de um colaborador (ex: desligado) para PST com as pastas originais | Purview (Graph) |

## Requisitos

- PowerShell 7+ (recomendado) ou Windows PowerShell 5.1
- Módulos (os scripts instalam automaticamente no escopo do usuário se faltarem):
  - `Microsoft.Graph.Authentication` (+ `Users`, `Groups`, `Identity.DirectoryManagement` para o NewClone)
  - `ExchangeOnlineManagement` 3.2+
  - `Microsoft.Online.SharePoint.PowerShell` — só para `Sharepoint.ps1 -SpoAdminUrl`
- Conta com as funções adequadas (ver cada script).

---

## NewClone.ps1 — criar usuário a partir de um modelo

Interativo. Pede o usuário modelo, os dados do novo usuário, gerente e senha temporária, e então:

1. Cria a conta no Entra ID (com `UsageLocation` do modelo, necessário para licença).
2. Atribui gerente.
3. Atribui licenças — escolha por número ou `M` para copiar as do modelo (mostra saldo livre).
4. Aguarda o usuário aparecer no Exchange Online.
5. Copia os grupos do modelo:
   - Security / Microsoft 365 → via Graph
   - Listas de distribuição e security mail-enabled → via Exchange (`Add-DistributionGroupMember`)
   - Dinâmicos e sincronizados do AD local → ignorados (com aviso)
6. Permite buscar e adicionar grupos extras.
7. Repete as DLs pendentes até o mailbox ficar pronto.

**Permissões:** `User.ReadWrite.All`, `Group.ReadWrite.All`, `Directory.Read.All` no Graph e uma função do Exchange que permita gerenciar membros de grupos.

```powershell
.\NewClone.ps1
```

---

## Sharepoint.ps1 — auditoria do SharePoint

Usa o relatório oficial de uso do Microsoft 365 e gera `SharePoint-Auditoria-<data>.csv` (abre direto no Excel, separador `;`).

**Colunas:** título, URL, tipo do site (Teams, Comunicação…), hub, dono, status (*Ativo / Pouco usado / Inativo / Nunca usado*), última atividade, arquivos e % ativos, páginas vistas, GB usados, cota e % da cota, compartilhamento externo e **Alertas** (sem dono, inativo, cota alta, link anônimo, bloqueado).

```powershell
# Básico (últimos 30 dias)
.\Sharepoint.ps1

# 90 dias, inativo = sem uso há mais de 180 dias, com OneDrive
.\Sharepoint.ps1 -Periodo D90 -DiasInativo 180 -IncluirOneDrive

# Com Hub, compartilhamento externo e bloqueio (módulo SPO)
.\Sharepoint.ps1 -SpoAdminUrl https://<tenant>-admin.sharepoint.com
```

| Parâmetro | Padrão | Descrição |
|---|---|---|
| `-Periodo` | `D30` | `D7`, `D30`, `D90` ou `D180` |
| `-DiasInativo` | `90` | Dias sem atividade para marcar como Inativo |
| `-AlertaCotaPct` | `90` | % de cota que gera alerta |
| `-SpoAdminUrl` | — | Admin center do SharePoint, para dados de governança |
| `-IncluirOneDrive` | — | Gera também `OneDrive-Uso-<data>.csv` |
| `-RevelarNomes` | — | Desativa a ocultação de nomes nos relatórios **do tenant inteiro** |
| `-Out` | `C:\Temp\SP-Auditoria` | Pasta de saída |

> ⚠️ **URLs vazias?** Por padrão o Microsoft 365 oculta nomes nos relatórios. Desative em
> *Admin Center → Configurações → Configurações da organização → Relatórios* ou rode com `-RevelarNomes`.

**Permissões:** `Reports.Read.All`, `Sites.Read.All`, `ReportSettings.Read.All` (e `ReportSettings.ReadWrite.All` com `-RevelarNomes`). A função *Leitor de Relatórios* ou *Administrador do SharePoint* basta.

---

## Auditoriaexchange.ps1 — ligar auditoria nas caixas

Restaura o conjunto padrão de ações auditadas pela Microsoft e **acrescenta** `Move` para proprietário e delegados.
Avisa se a auditoria estiver desligada na organização ou se a caixa tiver *bypass* de auditoria.

```powershell
.\Auditoriaexchange.ps1 -Caixas financeiro@empresa.com.br, boletos@empresa.com.br
```

Sem `-Caixas`, o PowerShell pergunta os endereços interativamente.

> Por que não `Set-Mailbox -AuditDelegate Move, SoftDelete, ...`? Esse formato **substitui** a lista padrão
> e desliga ações importantes como `UpdateInboxRules` (regras de encaminhamento). Use sempre `@{Add = ...}`.

---

## ExtracaoExchange.ps1 — quem moveu/apagou e-mails

Consulta o Unified Audit Log e gera um CSV com **uma linha por e-mail afetado**: data/hora local, caixa, ação, usuário responsável, tipo de acesso (proprietário/delegado/admin), assunto, pasta de origem e destino, IP e cliente.

```powershell
$caixas = 'financeiro@empresa.com.br', 'boletos@empresa.com.br'

.\ExtracaoExchange.ps1 -Caixas $caixas                      # últimos 7 dias
.\ExtracaoExchange.ps1 -Caixas $caixas -Dias 30 -GridView   # 30 dias e abre janela interativa
.\ExtracaoExchange.ps1 -Caixas $caixas -CsvPath C:\Temp\rel.csv
```

Notas:
- Ações de mover/excluir são do tipo `ExchangeItemGroup` no log (não `ExchangeItem`).
- A busca é paginada; o limite do serviço é 50.000 registros por consulta — reduza `-Dias` se aparecer o aviso.
- Retenção do log: 180 dias (Audit Standard) ou mais com Audit Premium.
- **Permissão:** função *View-Only Audit Logs* ou *Audit Logs*.

---

## Backup e recuperação de caixas

O Exchange Online não faz backup das caixas para você. Estas ferramentas cobrem dois cenários:

| Situação | Ferramenta |
|---|---|
| Alguém apagou e-mails **nos últimos 14–30 dias** | `Recuperar-ItensExcluidos.ps1`: usa a lixeira interna do Exchange, sem precisar de backup |
| Apagado há mais tempo, caixa comprometida, funcionário desligado, cópia fora do Microsoft 365 | `Backup-Caixa.ps1` + `Restaurar-Backup.ps1` |

> Para aumentar a lixeira interna para o máximo de 30 dias:
> `Get-Mailbox -ResultSize Unlimited | Set-Mailbox -RetainDeletedItemsFor 30`

### Recuperar-ItensExcluidos.ps1

Lista os itens das pastas *Itens Excluídos* e *Itens Recuperáveis*, deixa escolher quais restaurar
(números, intervalos como `5-8`, ou `T` para todos) e os devolve à pasta original, ou a `-PastaDestino`.

```powershell
.\Recuperar-ItensExcluidos.ps1 -Caixa financeiro@empresa.com.br -Assunto boleto
.\Recuperar-ItensExcluidos.ps1 -Caixa financeiro@empresa.com.br -Tipo IPM.Note -De (Get-Date).AddDays(-3) -PastaDestino Recuperados
```

**Permissão:** função *Mailbox Import Export* (não vem atribuída a ninguém por padrão):

```powershell
New-ManagementRoleAssignment -Role "Mailbox Import Export" -User admin@empresa.com.br
```

### Backup-Caixa.ps1

Usa a [API de import/export de caixas do Microsoft Graph](https://learn.microsoft.com/graph/mailbox-import-export-concept-overview), que substitui o EWS (desativado no Exchange Online em outubro de 2026).

- **Fidelidade total:** e-mails, calendário, contatos e tarefas, com anexos e propriedades.
- **Incremental:** a 1ª execução copia tudo; as seguintes, só o que mudou.
- **Nada é perdido:** itens apagados na caixa continuam no backup (coluna `RemovidoEm`).
- **Índice pesquisável** (`indice.csv`, abre no Excel): data, pasta, assunto, remetente, destinatários, anexo e tamanho.

```
C:\Backup\Exchange\
├─ logs\backup-<data>.log
└─ financeiro@empresa.com.br\
   ├─ itens\2026\09\*.fts      ← conteúdo (formato opaco, só para restaurar)
   ├─ indice.csv
   ├─ estado.json             ← controle do incremental
   └─ falhas.log
```

```powershell
# Login interativo
.\Backup-Caixa.ps1 -Caixas financeiro@empresa.com.br

# Como aplicativo (agendado)
.\Backup-Caixa.ps1 -Caixas financeiro@empresa.com.br, boletos@empresa.com.br `
    -TenantId <tenant-id> -ClientId <app-id> -CertificateThumbprint <thumbprint>
```

| Parâmetro | Padrão | Descrição |
|---|---|---|
| `-Caixas` | — | Caixas a copiar (usuário ou compartilhada) |
| `-Destino` | `C:\Backup\Exchange` | Pasta do backup; use um disco/compartilhamento fora do servidor |
| `-ExcluirPastas` | — | Caminhos a ignorar, com curinga (`'Lixo Eletrônico'`, `'Inbox/Newsletter*'`) |
| `-TenantId` `-ClientId` `-CertificateThumbprint` | — | Modo aplicativo |

O script sai com código ≠ 0 se alguma caixa tiver falhas, e o Agendador de Tarefas mostra isso como erro.

### Restaurar-Backup.ps1

Filtra o índice, deixa escolher os itens e os reimporta numa pasta nova (**Restaurados &lt;data&gt;**) na caixa,
recriando dentro dela as pastas originais. Nada que já existe na caixa é alterado.

```powershell
# E-mails apagados com "boleto" no assunto em setembro
.\Restaurar-Backup.ps1 -Caixa financeiro@empresa.com.br -Assunto boleto -De 2026-09-01 -Ate 2026-09-30 -SomenteRemovidos

# Tudo de um fornecedor, da caixa de um ex-funcionário para a caixa do gestor
.\Restaurar-Backup.ps1 -Caixa ex.funcionario@empresa.com.br -CaixaDestino gestor@empresa.com.br -Remetente fornecedor.com.br -Todos
```

Filtros: `-Assunto`, `-Remetente` (nome ou e-mail), `-Pasta` (curinga), `-De`, `-Ate`, `-SomenteRemovidos`.
Opções: `-CaixaDestino`, `-PastaRestauracao`, `-SemEstrutura`, `-GridView` (seleção em janela), `-Todos`, `-Force`.

### Login: interativo ou aplicativo

| | Interativo | Aplicativo (`Criar-AppBackup.ps1`) |
|---|---|---|
| Configuração | Nenhuma | Cria um app no Entra ID com certificado |
| Agendamento | Não | Sim |
| Alcance | Caixas em que sua conta tem permissão | **Todas** as caixas do tenant |
| Permissões Graph | `MailboxFolder.Read`, `MailboxItem.Read`, `MailboxItem.Export`, `User.Read.All` (restaurar: `MailboxFolder.ReadWrite`, `MailboxItem.ImportExport`) | Mesmas, na versão `.All` de aplicativo |

```powershell
# Só backup
.\Criar-AppBackup.ps1
# Backup + restauração; certificado no repositório da máquina (tarefa agendada com outra conta)
.\Criar-AppBackup.ps1 -PermitirRestauracao -Repositorio LocalMachine
```

O script mostra o TenantId, o ClientId e o Thumbprint, além do comando pronto para agendar. Exige Administrador Global
(ou Administrador de Função Privilegiada) para o consentimento; com `-SemConsentimento`, conceda depois no portal.

> ⚠️ Com permissão de aplicativo, quem tiver o certificado lê **qualquer** caixa do tenant. Mantenha o
> certificado só no servidor de backup e restrinja o acesso à pasta do backup.

### Exportar-PST.ps1: PST de colaborador desligado (Purview)

Para administradores. Exporta a caixa de um colaborador (inteira ou só o que você escolher) em **PST com a estrutura de pastas original**, pelo
**Microsoft Purview eDiscovery** (API do Graph). É o mesmo processo do portal, sem Outlook e sem acesso à caixa.

```powershell
.\Exportar-PST.ps1 -Caixa ex.colaborador@empresa.com.br                                    # caixa inteira
.\Exportar-PST.ps1 -Caixa ex.colaborador@empresa.com.br -Conteudo Email, Calendario -De 2025-01-01
.\Exportar-PST.ps1 -Caixa ex.colaborador@empresa.com.br -SemDownload                       # baixar depois
.\Exportar-PST.ps1 -Caixa ex.colaborador@empresa.com.br -SomenteDownload                   # baixa a última exportação, sem exportar de novo
```

**Escolher o que exportar** (o filtro é aplicado na pesquisa do Purview: só o selecionado é exportado):

| Parâmetro | Opções |
|---|---|
| `-Conteudo` | `Tudo` (padrão), `Email`, `Calendario`, `Contatos`, `Tarefas`, `Notas`, `Teams` (chats e reuniões). Combine com vírgula: `-Conteudo Email, Calendario` |
| `-De` / `-Ate` | Período pela data de recebimento, no formato **AAAA-MM-DD** (`01/07/2026` seria lido como 7 de janeiro). Itens sem essa data, como contatos, ficam de fora quando usado |
| `-Consulta` | KQL livre, combinada com os filtros acima (ex: `-Consulta 'from:fornecedor.com.br'`) |
| `-IncluirNaoIndexados` | Com filtros, itens não indexados (ex: anexos criptografados) ficam de fora porque não dá para saber se correspondem; esta opção os inclui |

Exemplo: caixa inteira **sem os chats do Teams**: `-Conteudo Email, Calendario, Contatos, Tarefas, Notas`.

O que ele faz:
1. Cria um **caso de eDiscovery** ("Exportação PST - &lt;caixa&gt; - &lt;data&gt;") e uma pesquisa com o conteúdo escolhido
2. Calcula a **estimativa** (itens e GB) antes de exportar
3. **Exporta em PST** com a estrutura de pastas, incluindo os Itens Recuperáveis (excluídos/expurgados)
4. **Baixa** o pacote e extrai o(s) PST(s) em `<Destino>\<nome do caso>\`

O download pede um **segundo login no navegador** (a conta já vem sugerida): os arquivos ficam num serviço do Purview que
não aceita o login do Graph. Se o download falhar, a exportação continua pronta no Purview: rode de novo com
`-SomenteDownload` para baixar **sem exportar outra vez** (ou acrescente `-CodigoDispositivo` para entrar com um código em
https://microsoft.com/devicelogin). O caso fica no Purview como registro da exportação e pode ser fechado no portal.

Sobre o PST gerado pelo Purview:
- Os nomes das pastas vêm com **hífen no lugar de espaço** (`Caixa-de-Entrada`) e ficam dentro de
  `Início-do-Repositório-de-Informações`; é o padrão do Purview, inclusive pelo portal.
- Com `-Conteudo Tudo`, entram também os **chats do Teams** (`TeamsMessagesData`), que ficam guardados na caixa de correio.
- O pacote pode trazer alguns arquivos de sistema (`SubstrateFiles`, `.json`) que não são conteúdo do usuário.

**Requisitos do tenant:**
- Conta do administrador no grupo de funções **eDiscovery Manager** (*Purview > Funções e escopos*).
- **Serviço de download do Purview registrado no tenant** (configuração única). Se faltar, o script pergunta e registra
  na hora (exige Administrador Global ou de Aplicativos); use `-RegistrarServicoDownload` para não perguntar.
- Conforme a licença, a Microsoft pode exigir o **Purview pay-as-you-go** para usar a API de eDiscovery
  (*Purview > Configurações > Faturamento*; [cobrança](https://learn.microsoft.com/purview/edisc-billing) pelo volume
  exportado, com franquia mensal). Em tenants com os recursos premium do eDiscovery habilitados, funciona sem ele.

Sem esses requisitos, o script explica o que falta e mostra o caminho manual pelo portal do Purview.

> **Antes de remover a licença do colaborador**, converta a caixa em compartilhada:
> `Set-Mailbox -Identity ex.colaborador@empresa.com.br -Type Shared`. Uma caixa de usuário sem licença é **excluída
> após 30 dias**. A compartilhada é preservada sem custo (até 50 GB).

> A Microsoft indica o [Microsoft 365 Backup](https://learn.microsoft.com/microsoft-365/backup/backup-overview)
> como solução oficial de backup. Estes scripts são uma alternativa sem custo, com cópia fora do Microsoft 365.

---

## Testes

Testes offline: os cmdlets do Graph/Exchange são simulados, nada é alterado no tenant.

```powershell
pwsh -NoProfile -File .\tests\Run-Tests.ps1
```

Cobrem sintaxe de todos os scripts, a lógica do relatório do SharePoint, a configuração de auditoria, a extração/paginação do audit log, a decisão Graph × Exchange do NewClone e, com uma caixa simulada, o backup incremental (paginação, itens apagados/alterados, token expirado, falhas), a restauração com filtros, a criação do app e, com um Purview simulado, a exportação em PST (filtros, custodiante, login e retomada do download).
