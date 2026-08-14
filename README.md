# Microsoft 365 E5 — Assessment de Adoção

![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-5391FE?logo=powershell&logoColor=white)
![Somente leitura](https://img.shields.io/badge/modo-somente%20leitura-107C10)
![Licença](https://img.shields.io/badge/licen%C3%A7a-MIT-0f6cbd)

Avalia se os workloads de segurança e Purview cobertos pela licença **Microsoft 365 E5** estão
sendo efetivamente utilizados e por quantos usuários. Para cada workload o assessment calcula:

| Métrica | Significado |
|---|---|
| **Entitled** | Usuários com o service plan habilitado na licença |
| **Coberto** | Usuários efetivamente alcançados por configuração ou telemetria real |
| **Gap** | `Entitled − Coberto`, com % de adoção e nível de maturidade |

Responde à pergunta que todo cliente com E5 faz: *"estou pagando por tudo isso, mas quanto disso
está realmente em uso e por quantas pessoas?"*

São dois artefatos, sem instalação e sem dependência de nuvem própria:

1. Um script PowerShell único, feito para copiar e colar no console.
2. Um dashboard HTML standalone que lê os CSVs gerados, direto no navegador.

O assessment é **somente leitura** — nenhuma configuração do tenant é alterada.

---

## Índice

- [Arquivos](#arquivos)
- [Pré-requisitos](#pré-requisitos)
- [Permissões necessárias](#permissões-necessárias)
- [Como executar](#como-executar)
- [Metodologia](#metodologia)
- [Workloads avaliados](#workloads-avaliados)
- [Arquivos gerados](#arquivos-gerados)
- [Como interpretar o resultado](#como-interpretar-o-resultado)
- [Troubleshooting](#troubleshooting)
- [Limitações conhecidas](#limitações-conhecidas)
- [Privacidade](#privacidade)
- [Aviso](#aviso)
- [Licença](#licença)

---

## Arquivos

| Arquivo | Descrição |
|---|---|
| `Invoke-M365E5Assessment.ps1` | Script único, autocontido, feito para copiar e colar no PowerShell |
| `E5-Assessment-Dashboard.html` | Dashboard standalone que lê a pasta de CSVs gerada |

---

## Pré-requisitos

### PowerShell

**PowerShell 7 é fortemente recomendado.** O script foi escrito para ser colado no console, e o
PS7 usa *bracketed paste*, que preserva blocos multi-linha. No Windows PowerShell 5.1 o paste de
scripts grandes pode truncar — nesse caso salve o conteúdo em um `.ps1` e execute o arquivo.

### Módulos

O script instala sozinho o que faltar (escopo `CurrentUser`), desde que `$InstalarModulos = $true`:

- `Microsoft.Graph.Authentication`
- `ExchangeOnlineManagement` (>= 3.0)
- `ImportExcel` (apenas para o `.xlsx`; desligue com `$SkipExcel = $true`)

Se o `Install-Module` falhar por causa do repositório, rode antes:

```powershell
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
```

---

## Permissões necessárias

O assessment é **100% somente leitura**. Nenhuma role de escrita é necessária, e o script nunca
cria, altera ou remove objetos no tenant.

### 1. Roles do Microsoft Entra ID

Atribua à conta que vai executar. O conjunto abaixo é o **mínimo com privilégio reduzido**:

| Role | Por que é necessária | O que quebra sem ela |
|---|---|---|
| **Global Reader** | Ler usuários, licenças, grupos, políticas de Acesso Condicional, PIM, Access Reviews | Fases 1 e 2 ficam vazias — sem baseline, o assessment inteiro não fecha |
| **Security Reader** | Ler Secure Score, alertas e incidentes do Defender XDR, Identity Protection, Advanced Hunting | Fase 3 e Secure Score ficam sem evidência |
| **Compliance Administrator** ou **Compliance Data Administrator** | Ler políticas de rótulos, DLP, retenção, Insider Risk, Communication Compliance e eDiscovery | Fases 4 e 5 (todo o Purview) ficam sem evidência |

Alternativa: **Global Administrator** cobre tudo, mas não é recomendado para um assessment de
leitura. Prefira as três roles acima.

> Observação: `Global Reader` **não** dá acesso a todos os cmdlets do Purview. Para as fases 4 e 5
> a conta precisa mesmo de uma das roles de compliance no portal do Microsoft Purview.

### 2. Escopos delegados do Microsoft Graph

Solicitados automaticamente no primeiro `Connect-MgGraph`. Se a organização exigir aprovação de
admin para o app **Microsoft Graph Command Line Tools**, um Global Administrator precisa aprovar
o consentimento uma única vez.

| Escopo | Usado para |
|---|---|
| `Organization.Read.All` | SKUs e service plans do tenant (`/subscribedSkus`) |
| `User.Read.All` | Inventário de usuários, licenças e planos atribuídos |
| `Directory.Read.All` | Objetos de diretório e resolução de principals |
| `Group.Read.All` | Expansão transitiva de grupos para resolver o escopo das políticas |
| `Policy.Read.All` | Políticas de Acesso Condicional e Security Defaults |
| `AuditLog.Read.All` | `signInActivity` (última autenticação) e relatório de métodos de autenticação |
| `UserAuthenticationMethod.Read.All` | Registro de MFA e passwordless por usuário |
| `IdentityRiskyUser.Read.All` | Usuários de risco do Identity Protection (Entra ID P2) |
| `IdentityRiskEvent.Read.All` | Detecções de risco no período analisado |
| `SecurityEvents.Read.All` | Microsoft Secure Score e perfis de controle |
| `SecurityAlert.Read.All` | Alertas por produto Defender (evidência objetiva de uso) |
| `ThreatHunting.Read.All` | Advanced Hunting: dispositivos MDE, sinais MDI e MDA |
| `Reports.Read.All` | Relatórios de uso e de registro de autenticação |
| `AccessReview.Read.All` | Definições de Access Review (Entra ID P2) |
| `RoleManagement.Read.Directory` | Atribuições permanentes vs. elegíveis no PIM |
| `DeviceManagementManagedDevices.Read.All` | Correlação de dispositivos com o Intune |

Para conferir o que o token atual já tem:

```powershell
(Get-MgContext).Scopes | Sort-Object
```

Para forçar um novo consentimento com todos os escopos:

```powershell
Disconnect-MgGraph
```

### 3. Exchange Online PowerShell

Usado para as políticas do Defender for Office 365 (Safe Links, Safe Attachments, Anti-phishing e
presets), `Get-OrganizationConfig` (Customer Lockbox e auditoria) e a amostra de mailboxes do
Audit Premium. `Global Reader` já é suficiente — ela mapeia para **View-Only Organization
Management** no Exchange Online.

### 4. Security & Compliance PowerShell (IPPSSession)

Usado para Purview: rótulos, políticas de rótulo, auto-labeling, DLP, retenção, Insider Risk,
Communication Compliance, eDiscovery Premium e Information Barriers. Exige uma role de compliance
(Compliance Administrator ou Compliance Data Administrator).

Alguns cmdlets têm requisitos próprios mesmo para leitura:

| Cmdlet | Requisito adicional |
|---|---|
| `Get-InsiderRiskPolicy` | Role **Insider Risk Management Admin** ou **Insider Risk Management Analyst** |
| `Get-SupervisoryReviewPolicyV2` | Role **Supervisory Review Administrator** |
| `Get-ComplianceCase` | Role **eDiscovery Manager** ou **Compliance Administrator** |
| `Get-InformationBarrierPolicy` | Role **Information Barriers Admin** |

Se alguma faltar, o script registra o workload como `Sem evidencia` e continua — nunca aborta.

### 5. Rede e ambiente

- Acesso de saída a `graph.microsoft.com`, `outlook.office365.com` e `ps.compliance.protection.outlook.com`.
- Autenticação interativa (o navegador é aberto). Em servidor sem navegador, use `Connect-MgGraph -UseDeviceCode` antes de rodar o script.
- Se houver Acesso Condicional exigindo dispositivo compatível para admins, execute a partir de uma máquina em conformidade.

### 6. O que o assessment NÃO precisa

- Nenhuma role de escrita, nenhuma permissão de aplicação (app-only), nenhum App Registration.
- Nenhum agente, extensão ou software instalado nas estações do cliente.
- Nenhum dado sai do ambiente: os CSVs ficam na máquina e o dashboard lê os arquivos localmente
  no navegador.

---

## Como executar

1. Abra `Invoke-M365E5Assessment.ps1` e edite o bloco `CONFIGURACAO` no topo:

```powershell
$AdminUPN            = 'admin@seutenant.onmicrosoft.com'   # obrigatorio
$DiasPeriodo         = 90        # janela de analise de alertas e deteccoes de risco
$DiasInatividade     = 90        # dias sem sign-in para considerar a licenca ociosa
$OutputFolder        = ''        # vazio = subpasta ao lado do script / no diretorio atual
$MaxUsuarios         = 0         # 0 = todos; use 500 para um teste rapido
$MaxMailboxesAudit   = 500       # amostra de mailboxes para avaliar Audit Premium
$SkipEntra           = $false
$SkipDefender        = $false
$SkipPurview         = $false
$SkipHunting         = $false    # desliga Advanced Hunting (MDE/MDI/MDA)
$SkipExchange        = $false
$SkipExcel           = $false
$Anonimizar          = $false    # hasheia UPNs nos CSVs de saida
$InstalarModulos     = $true
$AuthGraph           = 'Auto'    # Auto | DeviceCode | Modulo (ver Troubleshooting)
```

2. Selecione o arquivo inteiro, copie e cole no PowerShell. Serão solicitados até três logins:
   Microsoft Graph, Exchange Online e Security & Compliance.
3. Ao final o console imprime o scorecard e o caminho da pasta gerada.
4. Abra `E5-Assessment-Dashboard.html` no navegador, clique em **Selecionar pasta** e aponte para
   a pasta do assessment. Use **Imprimir / PDF** para gerar o entregável.

### Teste rápido antes do run completo

```powershell
$MaxUsuarios = 300; $SkipDefender = $true; $SkipPurview = $true
```

Valida conexão, licenciamento e baseline de usuários em poucos minutos.

---

## Metodologia

O cálculo tem três etapas.

**1. Entitled — quem tem direito ao recurso.**
O script lê `/subscribedSkus` para montar o mapa `servicePlanId → servicePlanName` e depois lê
`assignedPlans` de cada usuário. Um usuário é *entitled* a um workload quando o service plan
correspondente está com `capabilityStatus = Enabled`. Isso é mais preciso que olhar só o SKU,
porque planos individuais podem estar desabilitados dentro de uma licença E5.

**2. Coberto — quem é realmente alcançado.**
Para cada workload existe uma evidência concreta, nunca uma suposição:

- **Políticas** (CA, Safe Links, rótulos, DLP, retenção): o escopo é resolvido de verdade.
  Grupos são expandidos por `/groups/{id}/transitiveMembers` com cache, e inclusões e exclusões
  são aplicadas na ordem correta. Um usuário só conta como coberto se estiver no conjunto final.
- **Telemetria** (MDE, MDI, MDA): consultas de Advanced Hunting mostram quem gerou sinal de
  verdade nos últimos 7 a 30 dias. Licença atribuída sem telemetria significa produto pago e
  não implantado.
- **Nível de tenant**: workloads cujo escopo por usuário não é exposto pela API (Insider Risk,
  Communication Compliance, eDiscovery Premium, Audit Premium, Information Barriers, Customer
  Lockbox) são avaliados como ativo ou inativo, e a coluna `Detalhe` explica a base usada.

Alguns workloads combinam mais de uma fonte. **Entra ID P2**, por exemplo, considera coberto o
usuário que atende a qualquer um destes quatro critérios, porque o P2 se paga em várias frentes:

| Fonte | Evidência |
|---|---|
| Acesso Condicional | Estar no escopo de política habilitada com `userRiskLevels` ou `signInRiskLevels` |
| PIM | Ter atribuição **elegível** de role (atribuição permanente não conta, pois não usa o P2) |
| Identity Protection | Aparecer em `riskyUsers` ou ter detecção de risco no período |
| Access Reviews | Estar no escopo de uma revisão de acesso ainda ativa |

A coluna `Detalhe` do `08b_Cobertura_PorWorkload.csv` mostra quantos usuários vieram de cada
fonte, o que permite ver de onde a adoção está vindo.

**Evidência disponível vs. adoção zero.** Cada chamada ao Graph registra se realmente teve
sucesso. Se todas as fontes de um workload falharem — por falta de permissão, por exemplo — ele
é marcado como `Sem evidencia` e fica fora do score, em vez de ser reportado como adoção zero.
Basta uma fonte responder para o workload voltar a ser avaliado.

**3. Gap e maturidade.**

| % de adoção | Maturidade |
|---|---|
| Sem dado coletado | `Sem evidencia` |
| Sem usuários licenciados | `Nao licenciado` |
| 0% | `Nao implantado` |
| 1–49% | `Inicial` |
| 50–79% | `Parcial` |
| 80–100% | `Implantado` |

O **score de adoção E5** é a média das % de adoção dos workloads que têm evidência coletada e ao
menos um usuário licenciado. Workloads sem evidência ficam fora da média em vez de contar como
zero — assim uma permissão faltando não distorce o resultado para baixo.

---

## Workloads avaliados

| Pilar | Workload | Service plan | Evidência de cobertura |
|---|---|---|---|
| Identidade | Entra ID P1 | `AAD_PREMIUM` | Usuário no escopo de CA habilitada |
| Identidade | Entra ID P2 | `AAD_PREMIUM_P2` | CA de risco, elegibilidade PIM, Identity Protection ou Access Review ativa |
| Identidade | MFA | `AAD_PREMIUM*` | Método de MFA registrado |
| Defender | Defender for Endpoint P2 | `WINDEFATP` | Logon em dispositivo onboarded (Advanced Hunting) |
| Defender | Defender for Office 365 P2 | `THREAT_INTELLIGENCE` | Escopo de Safe Links / Safe Attachments / Anti-phishing |
| Defender | Defender for Identity | `ATA` | Eventos de sensores MDI |
| Defender | Defender for Cloud Apps | `ADALLOM_S_STANDARD` | Atividade em `CloudAppEvents` |
| Purview | Sensitivity Labels | `MIP_S_CLP*`, `RMS_S_PREMIUM*` | Escopo de label policy publicada |
| Purview | Auto-labeling | `MIP_S_CLP2` | Políticas de auto-labeling ativas (tenant) |
| Purview | DLP | `MIP_S_CLP*` | Escopo de política DLP ativa |
| Purview | Endpoint DLP | `MICROSOFTENDPOINTDLP` | Política DLP com location Endpoint |
| Governança | Retenção / Records | `RECORDS_MANAGEMENT` | Escopo de política de retenção |
| Governança | Insider Risk | `INSIDER_RISK*` | Políticas configuradas (tenant) |
| Governança | Communication Compliance | `COMMUNICATIONS_COMPLIANCE` | Políticas configuradas (tenant) |
| Governança | eDiscovery Premium | `EQUIVIO_ANALYTICS` | Casos Premium criados (tenant) |
| Governança | Audit Premium | `M365_ADVANCED_AUDITING` | Mailboxes com eventos premium (amostra) |
| Governança | Information Barriers | `INFORMATION_BARRIERS` | Políticas ativas (tenant) |
| Governança | Customer Lockbox | `LOCKBOX_ENTERPRISE` | `CustomerLockBoxEnabled` |

Workloads marcados como **tenant** não têm escopo por usuário exposto pela API: quando ativos,
todos os usuários licenciados são contados como cobertos, e a coluna `Detalhe` explica a base.

---

## Arquivos gerados

| Arquivo | Conteúdo |
|---|---|
| `00_Sumario.csv` | KPIs consolidados do assessment |
| `01_SKUs_Tenant.csv` | SKUs, licenças compradas, consumidas e disponíveis |
| `01b_ServicePlans.csv` | Service plan → workload → status |
| `02_Usuarios.csv` | 1 linha por usuário: licenças, planos ativos, sign-in, licença ociosa |
| `03_CA_Policies.csv` | Políticas de Acesso Condicional com escopo resolvido |
| `03b_MFA_Usuarios.csv` | Registro de métodos de autenticação |
| `03c_PIM.csv` | Atribuições permanentes vs. elegíveis |
| `03d_IdentityProtection.csv` | Usuários de risco |
| `03e_AccessReviews.csv` | Definições de Access Review |
| `04_MDE_Dispositivos.csv` | Dispositivos vistos pelo Defender for Endpoint |
| `04b_MDO_Politicas.csv` | Regras do Defender for Office 365 com escopo resolvido |
| `04c_MDI_MDA.csv` | Contas com sinais de MDI e MDA |
| `04d_Alertas.csv` | Alertas por produto Defender no período |
| `05_Labels.csv`, `05b_LabelPolicies.csv`, `05c_AutoLabeling.csv` | Information Protection |
| `05d_DLP_Politicas.csv` | Políticas DLP, modo e escopo |
| `06_Retention.csv` … `06f_InformationBarriers.csv` | Governança e risco |
| `07_SecureScore.csv` | Secure Score e controles |
| `07b_Checklist_Manual.csv` | Itens sem API — preencher `StatusManual` |
| `08_Cobertura_PorUsuario.csv` | Matriz usuário × workload |
| `08b_Cobertura_PorWorkload.csv` | Entitled, Coberto, Gap, % adoção, maturidade |
| `09_Recomendacoes.csv` | Gaps priorizados com próximo passo |
| `M365_E5_Assessment.xlsx` | Consolidado com uma aba por CSV |

---

## Como interpretar o resultado

Abra o dashboard e leia nesta ordem:

1. **Score de adoção E5** — quanto do que já foi comprado está de fato em uso.
2. **Adoção por pilar** — mostra onde está o desequilíbrio (é comum Identidade alta e
   Governança perto de zero).
3. **Cobertura por workload** — a tabela principal. Linhas em vermelho com `Entitled` alto e
   `Coberto` zero são o argumento mais forte do assessment: recurso pago e não implantado.
4. **Licenciamento** — licenças compradas e não atribuídas, mais licenças em contas desabilitadas.
   Costuma virar economia imediata.
5. **Recomendações** — já vêm priorizadas: prioridade 1 é workload licenciado com adoção zero ou
   usuários sem MFA; prioridade 2 é cobertura parcial; prioridade 4 é falta de evidência,
   que se resolve com permissão e não com projeto.

Distinção importante: `Nao implantado` significa que o dado foi coletado e a adoção é realmente
zero. `Sem evidencia` significa que o script não conseguiu coletar — verifique a coluna `Detalhe`
e a seção de permissões antes de reportar isso ao cliente como um gap.

---

## Troubleshooting

### `Method 'GetTokenAsync' ... does not have an implementation`

Conflito de assembly do Microsoft Graph: o `Microsoft.Graph.Authentication.Core` encontra no
processo um `Azure.Core` de versão incompatível. Costuma acontecer quando há versões misturadas
no disco, quando outro módulo (Az, Exchange) já carregou a DLL na sessão, ou no Windows
PowerShell 5.1.

**O script se resolve sozinho.** Com `$AuthGraph = 'Auto'` (padrão), ao detectar a falha ele troca
para **autenticação por device code em HTTP puro**, sem usar o módulo. Você recebe um código para
colar em [microsoft.com/devicelogin](https://microsoft.com/devicelogin) e o assessment segue
normalmente.

Para pular a tentativa com o módulo e ir direto ao device code:

```powershell
$AuthGraph = 'DeviceCode'
```

Esse modo não exige `Microsoft.Graph.Authentication` instalado — usa apenas `Invoke-RestMethod`
contra o endpoint OAuth e o app público **Microsoft Graph Command Line Tools**, o mesmo usado pelo
módulo oficial (o consentimento já dado continua valendo).

Se ainda assim quiser corrigir o módulo, feche **todas** as janelas do PowerShell e, em uma janela
nova:

```powershell
Get-InstalledModule Microsoft.Graph.Authentication -AllVersions | Uninstall-Module -Force
Remove-Item "$HOME\Documents\PowerShell\Modules\Microsoft.Graph*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$HOME\Documents\WindowsPowerShell\Modules\Microsoft.Graph*" -Recurse -Force -ErrorAction SilentlyContinue
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
```

O `Uninstall-Module` sozinho costuma não bastar: ele deixa para trás arquivos em `Dependencies`
que estavam em uso, e cópias instaladas em outro escopo ou na outra edição do PowerShell
continuam no `PSModulePath`. Para conferir o que sobrou:

```powershell
Get-Module -ListAvailable Microsoft.Graph.Authentication | Select-Object Version, ModuleBase
```

### `Nenhum usuario coletado`

Verifique `(Get-MgContext).Scopes` — o token precisa de `User.Read.All` e `Directory.Read.All`.
Se a sessão já estava conectada com menos escopos, rode `Disconnect-MgGraph` e execute de novo
para forçar um consentimento novo.

---

## Limitações conhecidas

- **Advanced Hunting** exige `ThreatHunting.Read.All` e Defender XDR provisionado. Sem ele,
  MDE, MDI e MDA aparecem como `Sem evidencia` em vez de zerados — a diferença é intencional.
- **Defender for Cloud Apps**: a tabela `CloudAppEvents` também recebe sinais do Office 365, então
  a cobertura do MDA é uma aproximação. A validação fina fica no `07b_Checklist_Manual.csv`.
- **Priva** e **Compliance Manager** não têm API pública de leitura — entram no checklist manual.
- **Insider Risk** e **Communication Compliance** não expõem o escopo de usuários; são avaliados
  no nível do tenant.
- **Audit Premium** é avaliado por amostragem de mailboxes (`$MaxMailboxesAudit`).
- Tenants grandes: a expansão de grupos usa cache, mas políticas com muitos grupos aninhados
  aumentam bastante o tempo. Use `$MaxUsuarios` para um primeiro run.

---

## Privacidade

O `02_Usuarios.csv` e o `08_Cobertura_PorUsuario.csv` contêm dados pessoais (UPN e nome).
Para entregar ao cliente sem expor identidades, use `$Anonimizar = $true` — os UPNs viram hashes
SHA-256 truncados e os nomes são omitidos. O dashboard lê os arquivos localmente no navegador;
nada é enviado para a internet.

Os arquivos gerados ficam apenas na máquina de quem executa. O `.gitignore` deste repositório já
bloqueia `*.csv`, `*.xlsx` e as pastas `M365E5_Assessment_*` para evitar commit acidental de dados
de cliente.

---

## Aviso

Este projeto é pessoal e **não é um produto oficial da Microsoft**, não sendo coberto por
suporte, SLA ou garantia da Microsoft. Nomes de produtos e APIs são marcas de seus respectivos
proprietários. Valide os resultados no portal antes de tomar decisões de licenciamento.

APIs do Microsoft Graph e cmdlets do Purview mudam com frequência: se um cmdlet ou endpoint for
alterado, o workload correspondente aparece como `Sem evidencia` em vez de gerar dado incorreto.

---

## Licença

[MIT](LICENSE)

---

## Autor

**Leandro Lima — 4P**

Contribuições são bem-vindas via issues e pull requests.

