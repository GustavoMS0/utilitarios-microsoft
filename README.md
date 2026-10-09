# Utilitários Microsoft 365

Scripts PowerShell para administração e auditoria de **Entra ID**, **Exchange Online** e **SharePoint Online**.

| Script | Para que serve | Conecta em |
|---|---|---|
| [NewClone.ps1](NewClone.ps1) | Cria um usuário clonando grupos, licenças e listas de distribuição de um usuário modelo | Graph + Exchange Online |
| [Sharepoint.ps1](Sharepoint.ps1) | Auditoria de uso e governança dos sites do SharePoint (e OneDrive) | Graph (+ SPO opcional) |
| [Auditoriaexchange.ps1](Auditoriaexchange.ps1) | Liga/ajusta a auditoria em caixas de correio específicas | Exchange Online |
| [ExtracaoExchange.ps1](ExtracaoExchange.ps1) | Relatório de quem moveu/apagou e-mails nessas caixas | Exchange Online |
| [Criauser.ps1](Criauser.ps1) | **Obsoleto** — versão antiga do NewClone, mantida só como referência | Graph |

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

## Testes

Testes offline: os cmdlets do Graph/Exchange são simulados, nada é alterado no tenant.

```powershell
pwsh -NoProfile -File .\tests\Run-Tests.ps1
```

Cobrem sintaxe de todos os scripts, a lógica do relatório do SharePoint, a configuração de auditoria, a extração/paginação do audit log e a decisão Graph × Exchange do NewClone.
