# ==============================================================================
# Invoke-M365E5Assessment.ps1
#
# Assessment de adocao dos workloads de seguranca e Purview cobertos pela
# licenca Microsoft 365 E5. Para cada workload calcula:
#   Entitled = usuarios com o service plan habilitado
#   Coberto  = usuarios efetivamente alcancados por configuracao/telemetria real
#   Gap      = Entitled - Coberto (com % de adocao e nivel de maturidade)
#
# SOMENTE LEITURA. Nenhuma configuracao do tenant e alterada.
#
# COMO USAR (feito para copiar e colar no PowerShell):
#   1. Edite o bloco CONFIGURACAO abaixo (no minimo o $AdminUPN).
#   2. Selecione o arquivo inteiro, copie e cole em uma janela do PowerShell 7.
#   3. Ao final, abra o E5-Assessment-Dashboard.html e aponte para a pasta gerada.
#
# Regras de estilo deste arquivo (nao quebre, ele e feito para paste):
#   - sem bloco param(), sem backtick de continuacao de linha, sem aliases
#   - sem linhas em branco (o console legado do PS 5.1 encerra blocos nelas)
#   - somente ASCII
#
# Permissoes necessarias na conta usada:
#   Global Reader + Security Reader + Compliance Administrator
# ==============================================================================
#
# ############################ CONFIGURACAO ####################################
$AdminUPN            = 'admin@seutenant.onmicrosoft.com'
$DiasPeriodo         = 90
$DiasInatividade     = 90
# Vazio = cria a subpasta M365E5_Assessment_<data> ao lado do script (ou no diretorio atual, se colado no console).
# Para escolher o destino, informe o caminho completo, ex.: 'C:\temp\assessment-e5\saida'
$OutputFolder        = ''
$MaxUsuarios         = 0
$MaxMailboxesAudit   = 500
$SkipEntra           = $false
$SkipDefender        = $false
$SkipPurview         = $false
$SkipHunting         = $false
$SkipExchange        = $false
$SkipExcel           = $false
$Anonimizar          = $false
$InstalarModulos     = $true
# ########################## FIM DA CONFIGURACAO ###############################
#
$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'
#
# ------------------------------------------------------------------ helpers --
function Write-Etapa {
    param([string]$Texto)
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host ("  " + $Texto) -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor Cyan
}
function Write-Info {
    param([string]$Msg)
    Write-Host ('  [' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Msg) -ForegroundColor Gray
}
function Write-Ok {
    param([string]$Msg)
    Write-Host ('  [' + (Get-Date -Format 'HH:mm:ss') + '] ' + $Msg) -ForegroundColor Green
}
function Write-Aviso {
    param([string]$Msg)
    Write-Host ('  [AVISO] ' + $Msg) -ForegroundColor Yellow
}
function Write-Falha {
    param([string]$Msg)
    Write-Host ('  [ERRO]  ' + $Msg) -ForegroundColor Red
}
function Get-Prop {
    param($Obj, [string]$Nome, $Padrao = '')
    if ($null -eq $Obj) { return $Padrao }
    if ($Obj -is [System.Collections.IDictionary]) {
        foreach ($k in $Obj.Keys) {
            if ([string]$k -eq $Nome) {
                $v = $Obj[$k]
                if ($null -eq $v) { return $Padrao }
                return $v
            }
        }
        return $Padrao
    }
    $p = $Obj.PSObject.Properties[$Nome]
    if ($null -eq $p) { return $Padrao }
    if ($null -eq $p.Value) { return $Padrao }
    return $p.Value
}
function Get-NomesPropriedades {
    param($Obj)
    if ($null -eq $Obj) { return '' }
    if ($Obj -is [System.Collections.IDictionary]) { return (@($Obj.Keys) -join ', ') }
    return ((@($Obj.PSObject.Properties) | ForEach-Object { $_.Name }) -join ', ')
}
function Format-Lista {
    param($Valor, [int]$Max = 12)
    if ($null -eq $Valor) { return '' }
    $itens = New-Object System.Collections.ArrayList
    foreach ($v in @($Valor)) {
        if ($null -eq $v) { continue }
        $d = Get-Prop -Obj $v -Nome 'DisplayName' -Padrao $null
        if (-not $d) { $d = Get-Prop -Obj $v -Nome 'Name' -Padrao $null }
        if ($d) { [void]$itens.Add([string]$d) } else { [void]$itens.Add([string]$v) }
    }
    if ($itens.Count -eq 0) { return '' }
    if ($itens.Count -gt $Max) {
        $corte = $itens.GetRange(0, $Max) -join '; '
        return ($corte + '; +' + ($itens.Count - $Max) + ' outros')
    }
    return ($itens -join '; ')
}
function Get-UpnSaida {
    param([string]$Upn)
    if (-not $Upn) { return '' }
    if (-not $Anonimizar) { return $Upn }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Upn.ToLower()))
    $sha.Dispose()
    $hex = ([System.BitConverter]::ToString($bytes)).Replace('-', '')
    return ('user-' + $hex.Substring(0, 12).ToLower())
}
function Save-CsvRel {
    param($Dados, [string]$Arquivo)
    $caminho = Join-Path $script:PastaSaida $Arquivo
    $linhas = @($Dados)
    if ($linhas.Count -gt 0 -and $null -ne $linhas[0]) {
        $linhas | Export-Csv -Path $caminho -NoTypeInformation -Encoding UTF8
        Write-Ok ('CSV: ' + $Arquivo + ' (' + $linhas.Count + ' linhas)')
    }
    else {
        Set-Content -Path $caminho -Value '' -Encoding UTF8
        Write-Aviso ('CSV: ' + $Arquivo + ' (sem dados)')
    }
    if (-not $SkipExcel -and $linhas.Count -gt 0 -and $null -ne $linhas[0]) {
        $aba = [System.IO.Path]::GetFileNameWithoutExtension($Arquivo)
        if ($aba.Length -gt 30) { $aba = $aba.Substring(0, 30) }
        try {
            $linhas | Export-Excel -Path $script:ArquivoExcel -WorksheetName $aba -AutoSize -FreezeTopRow -Append -ErrorAction Stop
        }
        catch {
            Write-Aviso ('Aba Excel ' + $aba + ' falhou: ' + $_.Exception.Message)
        }
    }
}
# ------------------------------------------------------------------- Graph ---
function Invoke-GraphGet {
    param([string]$Uri, [switch]$Silencioso)
    $tentativa = 0
    while ($tentativa -lt 5) {
        $tentativa++
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
        }
        catch {
            $msg = $_.Exception.Message
            $throttle = ($msg -match '429' -or $msg -match 'Too Many Requests' -or $msg -match 'ServiceUnavailable' -or $msg -match '503')
            if ($throttle -and $tentativa -lt 5) {
                $espera = [math]::Min(60, [math]::Pow(2, $tentativa) * 2)
                Write-Aviso ('Throttling do Graph. Aguardando ' + $espera + 's...')
                Start-Sleep -Seconds $espera
                continue
            }
            if (-not $Silencioso) { Write-Falha ('Graph GET falhou (' + $Uri + '): ' + $msg) }
            return $null
        }
    }
    return $null
}
function Invoke-GraphPaged {
    param([string]$Uri, [int]$MaxItens = 0, [switch]$Silencioso)
    $itens = New-Object System.Collections.ArrayList
    $proximo = $Uri
    $paginas = 0
    $script:UltimaChamadaOk = $true
    while ($proximo) {
        $paginas++
        $r = Invoke-GraphGet -Uri $proximo -Silencioso:$Silencioso
        if ($null -eq $r) {
            if ($paginas -eq 1) { $script:UltimaChamadaOk = $false }
            break
        }
        $marcador = [object]'__sem_valor__'
        $valor = Get-Prop -Obj $r -Nome 'value' -Padrao $marcador
        if ([object]::ReferenceEquals($valor, $marcador)) { [void]$itens.Add($r); break }
        foreach ($i in @($valor)) {
            if ($null -ne $i) { [void]$itens.Add($i) }
        }
        $proximo = [string](Get-Prop -Obj $r -Nome '@odata.nextLink')
        if (-not $proximo) { $proximo = $null }
        if ($MaxItens -gt 0 -and $itens.Count -ge $MaxItens) { break }
        if ($paginas % 20 -eq 0) { Write-Info ('  ... ' + $itens.Count + ' registros coletados') }
    }
    return $itens.ToArray()
}
function Invoke-Hunting {
    param([string]$Consulta, [string]$Rotulo)
    if ($SkipHunting) { return $null }
    $corpo = @{ Query = $Consulta } | ConvertTo-Json -Depth 3
    try {
        $r = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/security/runHuntingQuery' -Body $corpo -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
        $marcador = [object]'__sem_valor__'
        $res = Get-Prop -Obj $r -Nome 'results' -Padrao $marcador
        if ([object]::ReferenceEquals($res, $marcador)) { return , @() }
        return , @($res)
    }
    catch {
        Write-Aviso ('Advanced Hunting indisponivel para ' + $Rotulo + ': ' + $_.Exception.Message)
        return $null
    }
}
function Get-MembrosGrupo {
    param([string]$GrupoId)
    if (-not $GrupoId) { return @() }
    $chave = $GrupoId.ToLower()
    if ($script:CacheGrupos.ContainsKey($chave)) { return $script:CacheGrupos[$chave] }
    $uri = 'https://graph.microsoft.com/v1.0/groups/' + $GrupoId + '/transitiveMembers/microsoft.graph.user?$select=userPrincipalName&$top=999'
    $membros = Invoke-GraphPaged -Uri $uri -Silencioso
    $upns = New-Object System.Collections.ArrayList
    foreach ($m in @($membros)) {
        $u = [string](Get-Prop -Obj $m -Nome 'userPrincipalName')
        if ($u) { [void]$upns.Add($u.ToLower()) }
    }
    $script:CacheGrupos[$chave] = $upns.ToArray()
    return $script:CacheGrupos[$chave]
}
function Get-MembrosGrupoPorTexto {
    param([string]$Texto)
    if (-not $Texto) { return @() }
    $chave = $Texto.ToLower()
    if ($script:CacheGrupoNome.ContainsKey($chave)) { return $script:CacheGrupoNome[$chave] }
    $id = $null
    $guid = [guid]::Empty
    if ([guid]::TryParse($Texto, [ref]$guid)) {
        $id = $Texto
    }
    else {
        $filtro = $Texto.Replace("'", "''")
        $expr = [uri]::EscapeDataString("mail eq '$filtro' or displayName eq '$filtro'")
        $uri = 'https://graph.microsoft.com/v1.0/groups?$filter=' + $expr + '&$select=id&$top=1'
        $g = @(Invoke-GraphPaged -Uri $uri -Silencioso)
        if ($g.Count -gt 0) { $id = [string](Get-Prop -Obj $g[0] -Nome 'id') }
    }
    $res = @()
    if ($id) { $res = Get-MembrosGrupo -GrupoId $id }
    $script:CacheGrupoNome[$chave] = $res
    return $res
}
# -------------------------------------------------------- motor de cobertura -
function Set-Cobertura {
    param([string]$Workload, $Upns)
    if (-not $script:Cobertura.ContainsKey($Workload)) { $script:Cobertura[$Workload] = @{} }
    foreach ($u in @($Upns)) {
        if (-not $u) { continue }
        $script:Cobertura[$Workload][([string]$u).ToLower()] = $true
    }
}
function Set-Evidencia {
    param([string]$Workload, [bool]$Disponivel, [string]$Detalhe = '')
    $script:EvidenciaOk[$Workload] = $Disponivel
    if ($Detalhe) { $script:EvidenciaDetalhe[$Workload] = $Detalhe }
}
function Set-CoberturaTenant {
    param([string]$Workload, [bool]$Ativo, [string]$Detalhe = '')
    $script:TenantAtivo[$Workload] = $Ativo
    Set-Evidencia -Workload $Workload -Disponivel $true -Detalhe $Detalhe
}
function Resolve-EntradasEscopo {
    param($Entradas)
    $set = @{}
    foreach ($e in @($Entradas)) {
        if ($null -eq $e) { continue }
        $txt = [string](Get-Prop -Obj $e -Nome 'Name' -Padrao $null)
        if (-not $txt) { $txt = [string](Get-Prop -Obj $e -Nome 'DisplayName' -Padrao $null) }
        if (-not $txt) { $txt = [string]$e }
        $txt = $txt.Trim()
        if (-not $txt) { continue }
        if ($txt -eq 'All' -or $txt -eq 'All Users') {
            foreach ($u in $script:Usuarios.Keys) { $set[$u] = $true }
            continue
        }
        $chave = $txt.ToLower()
        if ($script:Usuarios.ContainsKey($chave)) { $set[$chave] = $true; continue }
        if ($script:UsuarioPorId.ContainsKey($chave)) { $set[$script:UsuarioPorId[$chave]] = $true; continue }
        foreach ($m in @(Get-MembrosGrupoPorTexto -Texto $txt)) { $set[$m] = $true }
    }
    return $set
}
function Resolve-EscopoCA {
    param($Politica)
    $cond = Get-Prop -Obj $Politica -Nome 'conditions' -Padrao $null
    $usr  = Get-Prop -Obj $cond -Nome 'users' -Padrao $null
    if ($null -eq $usr) { return @{} }
    $set = @{}
    $incU = @(Get-Prop -Obj $usr -Nome 'includeUsers' -Padrao @())
    $incG = @(Get-Prop -Obj $usr -Nome 'includeGroups' -Padrao @())
    $incR = @(Get-Prop -Obj $usr -Nome 'includeRoles' -Padrao @())
    $excU = @(Get-Prop -Obj $usr -Nome 'excludeUsers' -Padrao @())
    $excG = @(Get-Prop -Obj $usr -Nome 'excludeGroups' -Padrao @())
    foreach ($i in $incU) {
        $t = ([string]$i).ToLower()
        if ($t -eq 'all') { foreach ($u in $script:Usuarios.Keys) { $set[$u] = $true }; continue }
        if ($t -eq 'none' -or $t -eq 'guestsorexternalusers') { continue }
        if ($script:UsuarioPorId.ContainsKey($t)) { $set[$script:UsuarioPorId[$t]] = $true }
    }
    foreach ($i in $incG) { foreach ($m in @(Get-MembrosGrupo -GrupoId ([string]$i))) { $set[$m] = $true } }
    foreach ($i in $incR) {
        $t = ([string]$i).ToLower()
        if ($script:MembrosPorRole.ContainsKey($t)) { foreach ($m in $script:MembrosPorRole[$t]) { $set[$m] = $true } }
    }
    foreach ($i in $excU) {
        $t = ([string]$i).ToLower()
        if ($script:UsuarioPorId.ContainsKey($t)) { [void]$set.Remove($script:UsuarioPorId[$t]) }
    }
    foreach ($i in $excG) { foreach ($m in @(Get-MembrosGrupo -GrupoId ([string]$i))) { [void]$set.Remove($m) } }
    return $set
}
function Resolve-EscopoAccessReview {
    param($Definicao)
    $set = @{}
    $escopo = Get-Prop -Obj $Definicao -Nome 'scope' -Padrao $null
    $consulta = [string](Get-Prop -Obj $escopo -Nome 'query')
    if ($consulta) {
        $mg = [regex]::Match($consulta, '/groups/([0-9A-Fa-f-]{36})')
        if ($mg.Success) {
            foreach ($u in @(Get-MembrosGrupo -GrupoId $mg.Groups[1].Value)) { $set[$u] = $true }
            return $set
        }
        $mr = [regex]::Match($consulta, "roleDefinitionId eq '([0-9A-Fa-f-]{36})'")
        if ($mr.Success) {
            $chave = $mr.Groups[1].Value.ToLower()
            if ($script:MembrosPorRole.ContainsKey($chave)) {
                foreach ($u in $script:MembrosPorRole[$chave]) { $set[$u] = $true }
            }
            return $set
        }
        if ($consulta -match '/users') {
            foreach ($u in $script:Usuarios.Keys) { $set[$u] = $true }
            return $set
        }
    }
    $id = [string](Get-Prop -Obj $Definicao -Nome 'id')
    if (-not $id) { return $set }
    $inst = @(Invoke-GraphPaged -Uri ('https://graph.microsoft.com/v1.0/identityGovernance/accessReviews/definitions/' + $id + '/instances?$top=5') -Silencioso)
    if ($inst.Count -eq 0) { return $set }
    $iid = [string](Get-Prop -Obj $inst[0] -Nome 'id')
    if (-not $iid) { return $set }
    $decisoes = @(Invoke-GraphPaged -Uri ('https://graph.microsoft.com/v1.0/identityGovernance/accessReviews/definitions/' + $id + '/instances/' + $iid + '/decisions?$top=500') -MaxItens 2000 -Silencioso)
    foreach ($d in $decisoes) {
        $pr = Get-Prop -Obj $d -Nome 'principal' -Padrao $null
        $u = ([string](Get-Prop -Obj $pr -Nome 'userPrincipalName')).ToLower()
        if ($u) { $set[$u] = $true }
    }
    return $set
}
# ================================================================= FASE 0 ====
function Initialize-Ambiente {
    Write-Etapa 'Fase 0 - Pre-flight e conexoes'
    $destino = $OutputFolder
    if (-not $destino) {
        $raiz = $PSScriptRoot
        if (-not $raiz) { $raiz = (Get-Location).Path }
        $destino = Join-Path $raiz ('M365E5_Assessment_' + (Get-Date -Format 'yyyyMMdd_HHmm'))
    }
    $script:PastaSaida   = $destino
    $script:Usuarios        = @{}
    $script:UsuarioPorId    = @{}
    $script:Cobertura       = @{}
    $script:EvidenciaOk     = @{}
    $script:EvidenciaDetalhe = @{}
    $script:TenantAtivo     = @{}
    $script:CacheGrupos     = @{}
    $script:CacheGrupoNome  = @{}
    $script:MembrosPorRole  = @{}
    $script:PlanoNomePorId  = @{}
    $script:SkuNomePorId    = @{}
    $script:E5Skus          = @()
    $script:SecureScorePct  = ''
    $script:UltimaChamadaOk = $true
    $script:ExoOk           = $false
    $script:IppsOk          = $false
    New-Item -ItemType Directory -Path $script:PastaSaida -Force | Out-Null
    $script:PastaSaida = (Resolve-Path $script:PastaSaida).Path
    $script:ArquivoExcel = Join-Path $script:PastaSaida 'M365_E5_Assessment.xlsx'
    Write-Info ('Pasta de saida: ' + $script:PastaSaida)
    $modulos = @('Microsoft.Graph.Authentication')
    if (-not $SkipExchange) { $modulos += 'ExchangeOnlineManagement' }
    if (-not $SkipExcel)    { $modulos += 'ImportExcel' }
    foreach ($m in $modulos) {
        $jaCarregado = Get-Module -Name $m
        $todasVersoes = @(Get-Module -ListAvailable -Name $m | Sort-Object Version -Descending)
        $disp = $todasVersoes | Select-Object -First 1
        if (-not $disp) {
            if (-not $InstalarModulos) {
                Write-Falha ('Modulo ' + $m + ' ausente. Instale com: Install-Module ' + $m + ' -Scope CurrentUser')
                return $false
            }
            Write-Info ('Instalando modulo ' + $m + ' (CurrentUser)...')
            try {
                Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                $todasVersoes = @(Get-Module -ListAvailable -Name $m | Sort-Object Version -Descending)
                $disp = $todasVersoes | Select-Object -First 1
            }
            catch {
                Write-Falha ('Falha ao instalar ' + $m + ': ' + $_.Exception.Message)
                return $false
            }
        }
        if ($todasVersoes.Count -gt 1) {
            $lista = (@($todasVersoes | ForEach-Object { $_.Version.ToString() }) -join ', ')
            Write-Aviso ($m + ' tem ' + $todasVersoes.Count + ' versoes instaladas (' + $lista + '). Versoes multiplas causam conflito de assembly.')
        }
        if ($jaCarregado) { continue }
        try {
            Import-Module -Name $m -RequiredVersion $disp.Version -ErrorAction Stop
        }
        catch {
            $erro = $_.Exception.Message
            try {
                Import-Module -Name $m -ErrorAction Stop
            }
            catch {
                $erro = $_.Exception.Message
                Write-Falha ('Falha ao importar ' + $m + ': ' + $erro)
                if ($erro -match 'does not have an implementation' -or $erro -match 'GetTokenAsync' -or $erro -match 'Could not load file or assembly') {
                    Write-Aviso 'Isso e conflito de assembly do Microsoft.Graph (versoes misturadas no disco ou DLL ja carregada nesta sessao).'
                    Write-Aviso 'Feche TODAS as janelas do PowerShell e, em uma janela nova, execute:'
                    Write-Aviso '  Get-InstalledModule Microsoft.Graph.Authentication -AllVersions | Uninstall-Module -Force'
                    Write-Aviso '  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force'
                    Write-Aviso 'Depois abra outra janela nova e rode o assessment antes de carregar qualquer outro modulo.'
                }
                return $false
            }
        }
    }
    $escopos = @(
        'Organization.Read.All', 'User.Read.All', 'Directory.Read.All', 'Group.Read.All',
        'Policy.Read.All', 'AuditLog.Read.All', 'UserAuthenticationMethod.Read.All',
        'IdentityRiskyUser.Read.All', 'IdentityRiskEvent.Read.All', 'SecurityEvents.Read.All',
        'SecurityAlert.Read.All', 'ThreatHunting.Read.All', 'Reports.Read.All',
        'AccessReview.Read.All', 'RoleManagement.Read.Directory',
        'DeviceManagementManagedDevices.Read.All'
    )
    $ctx = $null
    try { $ctx = Get-MgContext } catch { $ctx = $null }
    $reconectar = $true
    if ($ctx -and $ctx.Scopes) {
        $faltando = @($escopos | Where-Object { $ctx.Scopes -notcontains $_ })
        if ($faltando.Count -eq 0) {
            $reconectar = $false
            Write-Ok ('Graph ja conectado como ' + $ctx.Account)
        }
        else {
            Write-Info ('Reconectando ao Graph (escopos faltando: ' + ($faltando -join ', ') + ')')
        }
    }
    if ($reconectar) {
        Write-Info 'Conectando ao Microsoft Graph (consentimento pode ser solicitado)...'
        try {
            Connect-MgGraph -Scopes $escopos -NoWelcome -ErrorAction Stop
        }
        catch {
            try {
                Connect-MgGraph -Scopes $escopos -ErrorAction Stop
            }
            catch {
                Write-Falha ('Nao foi possivel conectar ao Graph: ' + $_.Exception.Message)
                return $false
            }
        }
        $ctx = Get-MgContext
        Write-Ok ('Graph conectado como ' + $ctx.Account + ' (tenant ' + $ctx.TenantId + ')')
    }
    $script:TenantId = (Get-MgContext).TenantId
    if (-not $SkipExchange) {
        $jaExo = $false
        try {
            $conns = Get-ConnectionInformation -ErrorAction SilentlyContinue
            foreach ($c in @($conns)) {
                if ([string](Get-Prop -Obj $c -Nome 'ConnectionUri') -like '*outlook.office*') { $jaExo = $true }
            }
        }
        catch { $jaExo = $false }
        if ($jaExo) {
            $script:ExoOk = $true
            Write-Ok 'Exchange Online ja conectado.'
        }
        else {
            Write-Info 'Conectando ao Exchange Online...'
            try {
                Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -ErrorAction Stop
                $script:ExoOk = $true
                Write-Ok 'Exchange Online conectado.'
            }
            catch {
                Write-Aviso ('Exchange Online indisponivel: ' + $_.Exception.Message)
            }
        }
    }
    if (-not $SkipPurview) {
        $jaIpps = $null -ne (Get-Command Get-DlpCompliancePolicy -ErrorAction SilentlyContinue)
        if ($jaIpps) {
            $script:IppsOk = $true
            Write-Ok 'Security & Compliance ja conectado.'
        }
        else {
            Write-Info 'Conectando ao Security & Compliance (IPPSSession)...'
            try {
                Connect-IPPSSession -UserPrincipalName $AdminUPN -ShowBanner:$false -ErrorAction Stop
                $script:IppsOk = $true
                Write-Ok 'Security & Compliance conectado.'
            }
            catch {
                Write-Aviso ('Security & Compliance indisponivel: ' + $_.Exception.Message)
            }
        }
    }
    return $true
}
# ================================================================= FASE 1 ====
function Get-DadosLicencas {
    Write-Etapa 'Fase 1 - Licenciamento e baseline de usuarios'
    $skus = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus'
    $padroesE5 = @(
        'SPE_E5', 'SPE_E5_NOPSTNCONF', 'SPE_E5_CALLINGMINUTES', 'Microsoft_365_E5',
        'M365EDU_A5_FACULTY', 'M365EDU_A5_STUDENT', 'M365EDU_A5_STUUSEBNFT',
        'IDENTITY_THREAT_PROTECTION', 'INFORMATION_PROTECTION_COMPLIANCE'
    )
    $linhasSku = New-Object System.Collections.ArrayList
    $linhasPlano = New-Object System.Collections.ArrayList
    foreach ($s in @($skus)) {
        $skuId   = [string](Get-Prop -Obj $s -Nome 'skuId')
        $partNum = [string](Get-Prop -Obj $s -Nome 'skuPartNumber')
        $script:SkuNomePorId[$skuId.ToLower()] = $partNum
        $prepaid = Get-Prop -Obj $s -Nome 'prepaidUnits' -Padrao $null
        $comprados = [int](Get-Prop -Obj $prepaid -Nome 'enabled' -Padrao 0)
        $consumidos = [int](Get-Prop -Obj $s -Nome 'consumedUnits' -Padrao 0)
        $ehE5 = $false
        foreach ($p in $padroesE5) { if ($partNum -like ($p + '*')) { $ehE5 = $true } }
        if ($ehE5) { $script:E5Skus += $skuId.ToLower() }
        $pctUso = 0
        if ($comprados -gt 0) { $pctUso = [math]::Round(($consumidos / $comprados) * 100, 1) }
        [void]$linhasSku.Add([PSCustomObject]@{
            SkuPartNumber = $partNum
            SkuId         = $skuId
            EhE5          = $ehE5
            Comprados     = $comprados
            Consumidos    = $consumidos
            Disponiveis   = ($comprados - $consumidos)
            PctUso        = $pctUso
        })
        foreach ($sp in @(Get-Prop -Obj $s -Nome 'servicePlans' -Padrao @())) {
            $spId   = [string](Get-Prop -Obj $sp -Nome 'servicePlanId')
            $spNome = [string](Get-Prop -Obj $sp -Nome 'servicePlanName')
            if ($spId) { $script:PlanoNomePorId[$spId.ToLower()] = $spNome }
            [void]$linhasPlano.Add([PSCustomObject]@{
                SkuPartNumber   = $partNum
                ServicePlanName = $spNome
                ServicePlanId   = $spId
                Workload        = (Get-WorkloadDoPlano -Plano $spNome)
                Status          = [string](Get-Prop -Obj $sp -Nome 'provisioningStatus')
            })
        }
    }
    Save-CsvRel -Dados $linhasSku.ToArray()   -Arquivo '01_SKUs_Tenant.csv'
    Save-CsvRel -Dados $linhasPlano.ToArray() -Arquivo '01b_ServicePlans.csv'
    $script:SkusE5Linhas = @($linhasSku.ToArray() | Where-Object { $_.EhE5 })
    Write-Info 'Coletando usuarios (pode demorar em tenants grandes)...'
    $camposCompletos = 'id,userPrincipalName,displayName,accountEnabled,userType,department,jobTitle,createdDateTime,assignedLicenses,assignedPlans,signInActivity'
    $camposBasicos   = 'id,userPrincipalName,displayName,accountEnabled,userType,department,jobTitle,createdDateTime,assignedLicenses,assignedPlans'
    $camposMinimos   = 'id,userPrincipalName,accountEnabled,assignedLicenses,assignedPlans'
    $brutos = @(Invoke-GraphPaged -Uri ('https://graph.microsoft.com/v1.0/users?$select=' + $camposCompletos + '&$top=100') -MaxItens $MaxUsuarios -Silencioso)
    $validos = @($brutos | Where-Object { (Get-Prop -Obj $_ -Nome 'userPrincipalName') })
    if ($validos.Count -eq 0) {
        Write-Aviso 'Consulta com signInActivity nao retornou usuarios. Repetindo sem esse campo.'
        $brutos = @(Invoke-GraphPaged -Uri ('https://graph.microsoft.com/v1.0/users?$select=' + $camposBasicos + '&$top=999') -MaxItens $MaxUsuarios)
        $validos = @($brutos | Where-Object { (Get-Prop -Obj $_ -Nome 'userPrincipalName') })
    }
    if ($validos.Count -eq 0) {
        Write-Aviso 'Repetindo com o conjunto minimo de campos.'
        $brutos = @(Invoke-GraphPaged -Uri ('https://graph.microsoft.com/v1.0/users?$select=' + $camposMinimos + '&$top=999') -MaxItens $MaxUsuarios)
        $validos = @($brutos | Where-Object { (Get-Prop -Obj $_ -Nome 'userPrincipalName') })
    }
    if ($validos.Count -eq 0 -and $brutos.Count -gt 0) {
        Write-Aviso ('O Graph retornou ' + $brutos.Count + ' objetos sem userPrincipalName.')
        Write-Aviso ('Tipo do primeiro objeto: ' + $brutos[0].GetType().FullName)
        Write-Aviso ('Propriedades: ' + (Get-NomesPropriedades -Obj $brutos[0]))
    }
    Write-Info ('Usuarios retornados pelo Graph: ' + $validos.Count)
    $agora = Get-Date
    $linhasUsr = New-Object System.Collections.ArrayList
    foreach ($u in $validos) {
        $upn = [string](Get-Prop -Obj $u -Nome 'userPrincipalName')
        if (-not $upn) { continue }
        $id = ([string](Get-Prop -Obj $u -Nome 'id')).ToLower()
        $planos = @{}
        foreach ($ap in @(Get-Prop -Obj $u -Nome 'assignedPlans' -Padrao @())) {
            if ([string](Get-Prop -Obj $ap -Nome 'capabilityStatus') -ne 'Enabled') { continue }
            $spid = ([string](Get-Prop -Obj $ap -Nome 'servicePlanId')).ToLower()
            if ($script:PlanoNomePorId.ContainsKey($spid)) { $planos[$script:PlanoNomePorId[$spid]] = $true }
        }
        $skusUsr = New-Object System.Collections.ArrayList
        $temE5 = $false
        foreach ($al in @(Get-Prop -Obj $u -Nome 'assignedLicenses' -Padrao @())) {
            $sid = ([string](Get-Prop -Obj $al -Nome 'skuId')).ToLower()
            if ($script:SkuNomePorId.ContainsKey($sid)) { [void]$skusUsr.Add($script:SkuNomePorId[$sid]) }
            if ($script:E5Skus -contains $sid) { $temE5 = $true }
        }
        $sia = Get-Prop -Obj $u -Nome 'signInActivity' -Padrao $null
        $ultimo = [string](Get-Prop -Obj $sia -Nome 'lastSignInDateTime')
        $ultimoNI = [string](Get-Prop -Obj $sia -Nome 'lastNonInteractiveSignInDateTime')
        $dias = ''
        $refData = $null
        foreach ($d in @($ultimo, $ultimoNI)) {
            if (-not $d) { continue }
            try {
                $dt = [datetime]$d
                if ($null -eq $refData -or $dt -gt $refData) { $refData = $dt }
            }
            catch { }
        }
        if ($refData) { $dias = [int]($agora - $refData).TotalDays }
        $habilitado = [bool](Get-Prop -Obj $u -Nome 'accountEnabled' -Padrao $false)
        $ociosa = $false
        if ($temE5 -and (-not $habilitado)) { $ociosa = $true }
        if ($temE5 -and $habilitado -and $null -ne $refData -and ($agora - $refData).TotalDays -gt $DiasInatividade) { $ociosa = $true }
        $nomeExib = [string](Get-Prop -Obj $u -Nome 'displayName')
        $nomeSaida = $nomeExib
        if ($Anonimizar) { $nomeSaida = '' }
        $obj = [PSCustomObject]@{
            Upn           = $upn
            DisplayName   = $nomeExib
            Habilitado    = $habilitado
            Tipo          = [string](Get-Prop -Obj $u -Nome 'userType')
            Departamento  = [string](Get-Prop -Obj $u -Nome 'department')
            Cargo         = [string](Get-Prop -Obj $u -Nome 'jobTitle')
            TemE5         = $temE5
            SKUs          = ($skusUsr -join '; ')
            PlanosAtivos  = (($planos.Keys | Sort-Object) -join '; ')
            UltimoSignIn  = $ultimo
            UltimoSignInNaoInterativo = $ultimoNI
            DiasSemSignIn = $dias
            LicencaOciosa = $ociosa
        }
        $script:Usuarios[$upn.ToLower()] = [PSCustomObject]@{
            Upn = $upn; TemE5 = $temE5; Habilitado = $habilitado; Planos = $planos; Id = $id
            DisplayName = $obj.DisplayName
        }
        if ($id) { $script:UsuarioPorId[$id] = $upn.ToLower() }
        [void]$linhasUsr.Add([PSCustomObject]@{
            Upn           = (Get-UpnSaida -Upn $obj.Upn)
            DisplayName   = $nomeSaida
            Habilitado    = $obj.Habilitado
            Tipo          = $obj.Tipo
            Departamento  = $obj.Departamento
            Cargo         = $obj.Cargo
            TemE5         = $obj.TemE5
            SKUs          = $obj.SKUs
            PlanosAtivos  = $obj.PlanosAtivos
            UltimoSignIn  = $obj.UltimoSignIn
            UltimoSignInNaoInterativo = $obj.UltimoSignInNaoInterativo
            DiasSemSignIn = $obj.DiasSemSignIn
            LicencaOciosa = $obj.LicencaOciosa
        })
    }
    Save-CsvRel -Dados $linhasUsr.ToArray() -Arquivo '02_Usuarios.csv'
    Write-Ok ('Usuarios coletados: ' + $script:Usuarios.Count + ' | com E5: ' + (@($script:Usuarios.Values | Where-Object { $_.TemE5 }).Count))
}
function Get-WorkloadDoPlano {
    param([string]$Plano)
    foreach ($w in $script:Catalogo) {
        if ($w.Planos -contains $Plano) { return $w.Nome }
    }
    return ''
}
# ================================================================= FASE 2 ====
function Get-DadosEntra {
    Write-Etapa 'Fase 2 - Entra ID (Acesso Condicional, MFA, PIM, Identity Protection)'
    $roleAssign = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?$expand=roleDefinition,principal&$top=500' -Silencioso
    $linhasPim = New-Object System.Collections.ArrayList
    $adminsUpn = @{}
    foreach ($ra in @($roleAssign)) {
        $rd = Get-Prop -Obj $ra -Nome 'roleDefinition' -Padrao $null
        $pr = Get-Prop -Obj $ra -Nome 'principal' -Padrao $null
        $upn = ([string](Get-Prop -Obj $pr -Nome 'userPrincipalName')).ToLower()
        $tpl = ([string](Get-Prop -Obj $rd -Nome 'templateId')).ToLower()
        $rid = ([string](Get-Prop -Obj $rd -Nome 'id')).ToLower()
        foreach ($k in @($tpl, $rid)) {
            if (-not $k) { continue }
            if (-not $script:MembrosPorRole.ContainsKey($k)) { $script:MembrosPorRole[$k] = New-Object System.Collections.ArrayList }
            if ($upn) { [void]$script:MembrosPorRole[$k].Add($upn) }
        }
        if ($upn) { $adminsUpn[$upn] = $true }
        $principalTxt = [string](Get-Prop -Obj $pr -Nome 'displayName')
        if ($upn) { $principalTxt = (Get-UpnSaida -Upn $upn) }
        [void]$linhasPim.Add([PSCustomObject]@{
            Role        = [string](Get-Prop -Obj $rd -Nome 'displayName')
            Principal   = $principalTxt
            TipoAtribuicao = 'Permanente (Active)'
            Origem      = 'roleAssignments'
        })
    }
    $eleg = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?$expand=roleDefinition,principal&$top=500' -Silencioso
    $okPim = $script:UltimaChamadaOk
    $upnsPim = @{}
    foreach ($e in @($eleg)) {
        $rd = Get-Prop -Obj $e -Nome 'roleDefinition' -Padrao $null
        $pr = Get-Prop -Obj $e -Nome 'principal' -Padrao $null
        $upn = ([string](Get-Prop -Obj $pr -Nome 'userPrincipalName')).ToLower()
        if ($upn) { $upnsPim[$upn] = $true; $adminsUpn[$upn] = $true }
        $principalTxt = [string](Get-Prop -Obj $pr -Nome 'displayName')
        if ($upn) { $principalTxt = (Get-UpnSaida -Upn $upn) }
        [void]$linhasPim.Add([PSCustomObject]@{
            Role        = [string](Get-Prop -Obj $rd -Nome 'displayName')
            Principal   = $principalTxt
            TipoAtribuicao = 'Elegivel (PIM / Entra ID P2)'
            Origem      = 'roleEligibilityScheduleInstances'
        })
    }
    Save-CsvRel -Dados $linhasPim.ToArray() -Arquivo '03c_PIM.csv'
    $caPols = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'
    $okCa = $script:UltimaChamadaOk
    $linhasCa = New-Object System.Collections.ArrayList
    $cobP1 = @{}
    $cobP2 = @{}
    foreach ($p in @($caPols)) {
        $estado = [string](Get-Prop -Obj $p -Nome 'state')
        $escopo = Resolve-EscopoCA -Politica $p
        $cond = Get-Prop -Obj $p -Nome 'conditions' -Padrao $null
        $riscoUsr = @(Get-Prop -Obj $cond -Nome 'userRiskLevels' -Padrao @())
        $riscoSes = @(Get-Prop -Obj $cond -Nome 'signInRiskLevels' -Padrao @())
        $usaRisco = ($riscoUsr.Count -gt 0 -or $riscoSes.Count -gt 0)
        $gc = Get-Prop -Obj $p -Nome 'grantControls' -Padrao $null
        $controles = @(Get-Prop -Obj $gc -Nome 'builtInControls' -Padrao @())
        if ($estado -eq 'enabled') {
            foreach ($u in $escopo.Keys) { $cobP1[$u] = $true }
            if ($usaRisco) { foreach ($u in $escopo.Keys) { $cobP2[$u] = $true } }
        }
        [void]$linhasCa.Add([PSCustomObject]@{
            Politica        = [string](Get-Prop -Obj $p -Nome 'displayName')
            Estado          = $estado
            UsuariosNoEscopo = $escopo.Count
            BaseadaEmRisco  = $usaRisco
            RiscoUsuario    = ($riscoUsr -join '; ')
            RiscoSessao     = ($riscoSes -join '; ')
            Controles       = ($controles -join '; ')
            CriadaEm        = [string](Get-Prop -Obj $p -Nome 'createdDateTime')
        })
    }
    $usuariosCaRisco = $cobP2.Count
    foreach ($u in $upnsPim.Keys) { $cobP2[$u] = $true }
    Save-CsvRel -Dados $linhasCa.ToArray() -Arquivo '03_CA_Policies.csv'
    Set-Cobertura -Workload 'Entra ID P1' -Upns @($cobP1.Keys)
    Set-Evidencia -Workload 'Entra ID P1' -Disponivel $okCa -Detalhe ('Politicas de Acesso Condicional habilitadas: ' + (@($linhasCa.ToArray() | Where-Object { $_.Estado -eq 'enabled' }).Count))
    $mfa = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?$top=500' -Silencioso
    $okMfa = $script:UltimaChamadaOk
    $linhasMfa = New-Object System.Collections.ArrayList
    $cobMfa = @{}
    foreach ($m in @($mfa)) {
        $upn = ([string](Get-Prop -Obj $m -Nome 'userPrincipalName')).ToLower()
        if (-not $upn) { continue }
        $capaz = [bool](Get-Prop -Obj $m -Nome 'isMfaCapable' -Padrao $false)
        if ($capaz) { $cobMfa[$upn] = $true }
        [void]$linhasMfa.Add([PSCustomObject]@{
            Upn             = (Get-UpnSaida -Upn $upn)
            MfaRegistrado   = [bool](Get-Prop -Obj $m -Nome 'isMfaRegistered' -Padrao $false)
            MfaCapaz        = $capaz
            SsprRegistrado  = [bool](Get-Prop -Obj $m -Nome 'isSsprRegistered' -Padrao $false)
            PasswordlessCapaz = [bool](Get-Prop -Obj $m -Nome 'isPasswordlessCapable' -Padrao $false)
            EhAdmin         = [bool](Get-Prop -Obj $m -Nome 'isAdmin' -Padrao $false)
            Metodos         = ((@(Get-Prop -Obj $m -Nome 'methodsRegistered' -Padrao @())) -join '; ')
        })
    }
    Save-CsvRel -Dados $linhasMfa.ToArray() -Arquivo '03b_MFA_Usuarios.csv'
    Set-Cobertura -Workload 'MFA / Autenticacao forte' -Upns @($cobMfa.Keys)
    Set-Evidencia -Workload 'MFA / Autenticacao forte' -Disponivel $okMfa -Detalhe 'Relatorio de registro de metodos de autenticacao'
    $script:AdminsSemMfa = @($linhasMfa.ToArray() | Where-Object { $_.EhAdmin -and -not $_.MfaCapaz }).Count
    $script:UsuariosSemMfa = @($script:Usuarios.Keys | Where-Object { -not $cobMfa.ContainsKey($_) }).Count
    $desde = (Get-Date).AddDays(-$DiasPeriodo).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $risky = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/identityProtection/riskyUsers?$top=500' -Silencioso
    $okIp = $script:UltimaChamadaOk
    $cobIp = @{}
    $linhasRisco = New-Object System.Collections.ArrayList
    foreach ($r in @($risky)) {
        $upn = ([string](Get-Prop -Obj $r -Nome 'userPrincipalName')).ToLower()
        if ($upn) { $cobIp[$upn] = $true }
        [void]$linhasRisco.Add([PSCustomObject]@{
            Upn           = (Get-UpnSaida -Upn $upn)
            NivelRisco    = [string](Get-Prop -Obj $r -Nome 'riskLevel')
            EstadoRisco   = [string](Get-Prop -Obj $r -Nome 'riskState')
            Detalhe       = [string](Get-Prop -Obj $r -Nome 'riskDetail')
            AtualizadoEm  = [string](Get-Prop -Obj $r -Nome 'riskLastUpdatedDateTime')
        })
    }
    $deteccoes = Invoke-GraphPaged -Uri ('https://graph.microsoft.com/v1.0/identityProtection/riskDetections?$filter=detectedDateTime%20ge%20' + $desde + '&$top=500') -Silencioso
    if ($script:UltimaChamadaOk) { $okIp = $true }
    foreach ($d in @($deteccoes)) {
        $upn = ([string](Get-Prop -Obj $d -Nome 'userPrincipalName')).ToLower()
        if ($upn) { $cobIp[$upn] = $true }
    }
    $script:RiskDetections = @($deteccoes).Count
    Save-CsvRel -Dados $linhasRisco.ToArray() -Arquivo '03d_IdentityProtection.csv'
    $revisoes = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/identityGovernance/accessReviews/definitions?$top=100' -Silencioso
    $okAr = $script:UltimaChamadaOk
    $cobAr = @{}
    $linhasRev = New-Object System.Collections.ArrayList
    foreach ($r in @($revisoes)) {
        $escopoRev = Resolve-EscopoAccessReview -Definicao $r
        $status = [string](Get-Prop -Obj $r -Nome 'status')
        if ($status -ne 'Completed' -and $status -ne 'Applied') {
            foreach ($u in $escopoRev.Keys) { $cobAr[$u] = $true }
        }
        [void]$linhasRev.Add([PSCustomObject]@{
            Nome       = [string](Get-Prop -Obj $r -Nome 'displayName')
            Status     = $status
            UsuariosNoEscopo = $escopoRev.Count
            CriadaEm   = [string](Get-Prop -Obj $r -Nome 'createdDateTime')
            Descricao  = [string](Get-Prop -Obj $r -Nome 'descriptionForAdmins')
        })
    }
    Save-CsvRel -Dados $linhasRev.ToArray() -Arquivo '03e_AccessReviews.csv'
    $script:AccessReviews = @($revisoes).Count
    foreach ($u in $cobIp.Keys) { $cobP2[$u] = $true }
    foreach ($u in $cobAr.Keys) { $cobP2[$u] = $true }
    Set-Cobertura -Workload 'Entra ID P2' -Upns @($cobP2.Keys)
    $fontesP2 = New-Object System.Collections.ArrayList
    if ($okCa)  { [void]$fontesP2.Add('CA baseada em risco: ' + $usuariosCaRisco) }
    if ($okPim) { [void]$fontesP2.Add('elegiveis no PIM: ' + $upnsPim.Count) }
    if ($okIp)  { [void]$fontesP2.Add('Identity Protection: ' + $cobIp.Count) }
    if ($okAr)  { [void]$fontesP2.Add('Access Reviews ativas: ' + $cobAr.Count) }
    $detalheP2 = 'Nenhuma fonte de evidencia de P2 pode ser consultada'
    if ($fontesP2.Count -gt 0) { $detalheP2 = 'Usuarios por fonte - ' + ($fontesP2 -join '; ') }
    Set-Evidencia -Workload 'Entra ID P2' -Disponivel ($okCa -or $okPim -or $okIp -or $okAr) -Detalhe $detalheP2
    $script:TotalAdmins = $adminsUpn.Count
    $script:AdminsComPim = $upnsPim.Count
}
# ================================================================= FASE 3 ====
function Get-DadosDefender {
    Write-Etapa 'Fase 3 - Defender XDR (Endpoint, Office 365, Identity, Cloud Apps)'
    $qDisp = 'DeviceInfo | where Timestamp > ago(30d) | summarize arg_max(Timestamp, DeviceName, OSPlatform, OnboardingStatus, JoinType) by DeviceId | take 20000'
    $disp = Invoke-Hunting -Consulta $qDisp -Rotulo 'DeviceInfo (MDE)'
    $linhasDisp = New-Object System.Collections.ArrayList
    foreach ($d in @($disp)) {
        [void]$linhasDisp.Add([PSCustomObject]@{
            DeviceId    = [string](Get-Prop -Obj $d -Nome 'DeviceId')
            DeviceName  = [string](Get-Prop -Obj $d -Nome 'DeviceName')
            OSPlatform  = [string](Get-Prop -Obj $d -Nome 'OSPlatform')
            Onboarding  = [string](Get-Prop -Obj $d -Nome 'OnboardingStatus')
            JoinType    = [string](Get-Prop -Obj $d -Nome 'JoinType')
            VistoEm     = [string](Get-Prop -Obj $d -Nome 'Timestamp')
        })
    }
    Save-CsvRel -Dados $linhasDisp.ToArray() -Arquivo '04_MDE_Dispositivos.csv'
    $script:DispositivosOnboarded = @($linhasDisp.ToArray() | Where-Object { $_.Onboarding -eq 'Onboarded' }).Count
    $qLogon = 'DeviceLogonEvents | where Timestamp > ago(30d) and ActionType == "LogonSuccess" and isnotempty(AccountUpn) | summarize Dispositivos=dcount(DeviceId) by AccountUpn | take 50000'
    $logons = Invoke-Hunting -Consulta $qLogon -Rotulo 'DeviceLogonEvents (MDE)'
    if ($null -ne $logons) {
        $cob = @()
        foreach ($l in @($logons)) { $cob += ([string](Get-Prop -Obj $l -Nome 'AccountUpn')).ToLower() }
        Set-Cobertura -Workload 'Defender for Endpoint P2' -Upns $cob
        Set-Evidencia -Workload 'Defender for Endpoint P2' -Disponivel $true -Detalhe ('Dispositivos onboarded: ' + $script:DispositivosOnboarded + '; usuarios com logon em dispositivo protegido nos ultimos 30 dias')
    }
    else {
        Set-Evidencia -Workload 'Defender for Endpoint P2' -Disponivel $false -Detalhe 'Advanced Hunting indisponivel (falta consentimento ThreatHunting.Read.All ou Defender XDR nao provisionado)'
    }
    $qIdent = 'IdentityLogonEvents | where Timestamp > ago(30d) and isnotempty(AccountUpn) | summarize Eventos=count() by AccountUpn | take 50000'
    $idents = Invoke-Hunting -Consulta $qIdent -Rotulo 'IdentityLogonEvents (MDI)'
    $linhasMdi = New-Object System.Collections.ArrayList
    if ($null -ne $idents) {
        $cob = @()
        foreach ($i in @($idents)) {
            $u = ([string](Get-Prop -Obj $i -Nome 'AccountUpn')).ToLower()
            $cob += $u
            [void]$linhasMdi.Add([PSCustomObject]@{
                Produto = 'Defender for Identity'
                Conta   = (Get-UpnSaida -Upn $u)
                Eventos = [int](Get-Prop -Obj $i -Nome 'Eventos' -Padrao 0)
            })
        }
        Set-Cobertura -Workload 'Defender for Identity' -Upns $cob
        Set-Evidencia -Workload 'Defender for Identity' -Disponivel $true -Detalhe 'Eventos de logon capturados por sensores MDI nos ultimos 30 dias'
    }
    else {
        Set-Evidencia -Workload 'Defender for Identity' -Disponivel $false -Detalhe 'Advanced Hunting indisponivel'
    }
    $qMda = 'CloudAppEvents | where Timestamp > ago(7d) and isnotempty(AccountObjectId) | summarize Eventos=count() by AccountObjectId, AccountDisplayName | take 50000'
    $mdas = Invoke-Hunting -Consulta $qMda -Rotulo 'CloudAppEvents (MDA)'
    if ($null -ne $mdas) {
        $cob = @()
        foreach ($a in @($mdas)) {
            $oid = ([string](Get-Prop -Obj $a -Nome 'AccountObjectId')).ToLower()
            if ($script:UsuarioPorId.ContainsKey($oid)) {
                $u = $script:UsuarioPorId[$oid]
                $cob += $u
                [void]$linhasMdi.Add([PSCustomObject]@{
                    Produto = 'Defender for Cloud Apps'
                    Conta   = (Get-UpnSaida -Upn $u)
                    Eventos = [int](Get-Prop -Obj $a -Nome 'Eventos' -Padrao 0)
                })
            }
        }
        Set-Cobertura -Workload 'Defender for Cloud Apps' -Upns $cob
        Set-Evidencia -Workload 'Defender for Cloud Apps' -Disponivel $true -Detalhe 'Atividade em CloudAppEvents nos ultimos 7 dias (aproximacao: a tabela tambem recebe sinais do Office 365)'
    }
    else {
        Set-Evidencia -Workload 'Defender for Cloud Apps' -Disponivel $false -Detalhe 'Advanced Hunting indisponivel'
    }
    Save-CsvRel -Dados $linhasMdi.ToArray() -Arquivo '04c_MDI_MDA.csv'
    $linhasMdo = New-Object System.Collections.ArrayList
    $cobMdo = @{}
    if ($script:ExoOk) {
        $regras = New-Object System.Collections.ArrayList
        foreach ($par in @(
            @{ Cmd = 'Get-SafeLinksRule';           Tipo = 'Safe Links' },
            @{ Cmd = 'Get-SafeAttachmentRule';      Tipo = 'Safe Attachments' },
            @{ Cmd = 'Get-AntiPhishRule';           Tipo = 'Anti-phishing' },
            @{ Cmd = 'Get-ATPProtectionPolicyRule'; Tipo = 'Preset (Standard/Strict)' }
        )) {
            if (-not (Get-Command $par.Cmd -ErrorAction SilentlyContinue)) { continue }
            try {
                foreach ($r in @(& $par.Cmd -ErrorAction Stop)) { [void]$regras.Add(@{ Regra = $r; Tipo = $par.Tipo }) }
            }
            catch {
                Write-Aviso ($par.Cmd + ' falhou: ' + $_.Exception.Message)
            }
        }
        foreach ($item in $regras) {
            $r = $item.Regra
            $estado = [string](Get-Prop -Obj $r -Nome 'State' -Padrao 'Enabled')
            $sentTo    = @(Get-Prop -Obj $r -Nome 'SentTo' -Padrao @())
            $sentToGrp = @(Get-Prop -Obj $r -Nome 'SentToMemberOf' -Padrao @())
            $dominios  = @(Get-Prop -Obj $r -Nome 'RecipientDomainIs' -Padrao @())
            $escopo = Resolve-EntradasEscopo -Entradas ($sentTo + $sentToGrp)
            if ($dominios.Count -gt 0) {
                foreach ($u in $script:Usuarios.Keys) {
                    foreach ($d in $dominios) {
                        if ($u -like ('*@' + ([string]$d).ToLower())) { $escopo[$u] = $true }
                    }
                }
            }
            if ($sentTo.Count -eq 0 -and $sentToGrp.Count -eq 0 -and $dominios.Count -eq 0) {
                foreach ($u in $script:Usuarios.Keys) { $escopo[$u] = $true }
            }
            foreach ($ex in @(Get-Prop -Obj $r -Nome 'ExceptIfSentTo' -Padrao @())) {
                $rem = Resolve-EntradasEscopo -Entradas @($ex)
                foreach ($k in $rem.Keys) { [void]$escopo.Remove($k) }
            }
            foreach ($ex in @(Get-Prop -Obj $r -Nome 'ExceptIfSentToMemberOf' -Padrao @())) {
                $rem = Resolve-EntradasEscopo -Entradas @($ex)
                foreach ($k in $rem.Keys) { [void]$escopo.Remove($k) }
            }
            if ($estado -eq 'Enabled') { foreach ($k in $escopo.Keys) { $cobMdo[$k] = $true } }
            [void]$linhasMdo.Add([PSCustomObject]@{
                Tipo             = $item.Tipo
                Regra            = [string](Get-Prop -Obj $r -Nome 'Name')
                Estado           = $estado
                Prioridade       = [string](Get-Prop -Obj $r -Nome 'Priority')
                UsuariosNoEscopo = $escopo.Count
                Destinatarios    = (Format-Lista -Valor $sentTo)
                Grupos           = (Format-Lista -Valor $sentToGrp)
                Dominios         = (Format-Lista -Valor $dominios)
            })
        }
        Set-Cobertura -Workload 'Defender for Office 365 P2' -Upns @($cobMdo.Keys)
        Set-Evidencia -Workload 'Defender for Office 365 P2' -Disponivel $true -Detalhe ('Regras de Safe Links/Safe Attachments/Anti-phishing ativas: ' + (@($linhasMdo.ToArray() | Where-Object { $_.Estado -eq 'Enabled' }).Count))
    }
    else {
        Set-Evidencia -Workload 'Defender for Office 365 P2' -Disponivel $false -Detalhe 'Exchange Online PowerShell nao conectado'
    }
    Save-CsvRel -Dados $linhasMdo.ToArray() -Arquivo '04b_MDO_Politicas.csv'
    $desde = (Get-Date).AddDays(-$DiasPeriodo).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $alertas = Invoke-GraphPaged -Uri ('https://graph.microsoft.com/v1.0/security/alerts_v2?$filter=createdDateTime%20ge%20' + $desde + '&$top=1000') -MaxItens 5000 -Silencioso
    $porFonte = @{}
    foreach ($a in @($alertas)) {
        $f = [string](Get-Prop -Obj $a -Nome 'serviceSource' -Padrao 'desconhecido')
        if (-not $porFonte.ContainsKey($f)) { $porFonte[$f] = 0 }
        $porFonte[$f] = $porFonte[$f] + 1
    }
    $linhasAl = New-Object System.Collections.ArrayList
    foreach ($k in ($porFonte.Keys | Sort-Object)) {
        [void]$linhasAl.Add([PSCustomObject]@{
            Produto      = $k
            Alertas      = $porFonte[$k]
            PeriodoDias  = $DiasPeriodo
        })
    }
    Save-CsvRel -Dados $linhasAl.ToArray() -Arquivo '04d_Alertas.csv'
    $script:TotalAlertas = @($alertas).Count
}
# ================================================================= FASE 4 ====
function Get-DadosPurview {
    Write-Etapa 'Fase 4 - Purview Information Protection e DLP'
    if (-not $script:IppsOk) {
        Write-Aviso 'Security & Compliance nao conectado. Fase 4 pulada.'
        foreach ($w in @('Sensitivity Labels (MIP)', 'Auto-labeling (E5)', 'DLP', 'Endpoint DLP')) {
            Set-Evidencia -Workload $w -Disponivel $false -Detalhe 'Security & Compliance PowerShell nao conectado'
        }
        return
    }
    $linhasLabel = New-Object System.Collections.ArrayList
    try {
        foreach ($l in @(Get-Label -ErrorAction Stop)) {
            [void]$linhasLabel.Add([PSCustomObject]@{
                Label       = [string](Get-Prop -Obj $l -Nome 'DisplayName')
                Nome        = [string](Get-Prop -Obj $l -Nome 'Name')
                Desabilitado = [string](Get-Prop -Obj $l -Nome 'Disabled')
                Prioridade  = [string](Get-Prop -Obj $l -Nome 'Priority')
                ContentType = (Format-Lista -Valor (Get-Prop -Obj $l -Nome 'ContentType' -Padrao @()))
                Pai         = [string](Get-Prop -Obj $l -Nome 'ParentLabelDisplayName')
            })
        }
    }
    catch { Write-Aviso ('Get-Label falhou: ' + $_.Exception.Message) }
    Save-CsvRel -Dados $linhasLabel.ToArray() -Arquivo '05_Labels.csv'
    $cobLabel = @{}
    $linhasLp = New-Object System.Collections.ArrayList
    try {
        foreach ($p in @(Get-LabelPolicy -ErrorAction Stop)) {
            $locs = @()
            foreach ($n in @('ExchangeLocation', 'SharePointLocation', 'OneDriveLocation', 'ModernGroupLocation')) {
                $locs += @(Get-Prop -Obj $p -Nome $n -Padrao @())
            }
            $escopo = Resolve-EntradasEscopo -Entradas $locs
            foreach ($k in $escopo.Keys) { $cobLabel[$k] = $true }
            [void]$linhasLp.Add([PSCustomObject]@{
                Politica         = [string](Get-Prop -Obj $p -Nome 'Name')
                Modo             = [string](Get-Prop -Obj $p -Nome 'Mode')
                Labels           = (Format-Lista -Valor (Get-Prop -Obj $p -Nome 'Labels' -Padrao @()))
                UsuariosNoEscopo = $escopo.Count
                Locais           = (Format-Lista -Valor $locs)
            })
        }
        Set-Cobertura -Workload 'Sensitivity Labels (MIP)' -Upns @($cobLabel.Keys)
        Set-Evidencia -Workload 'Sensitivity Labels (MIP)' -Disponivel $true -Detalhe ('Politicas de publicacao de rotulos: ' + $linhasLp.Count)
    }
    catch {
        Write-Aviso ('Get-LabelPolicy falhou: ' + $_.Exception.Message)
        Set-Evidencia -Workload 'Sensitivity Labels (MIP)' -Disponivel $false -Detalhe 'Get-LabelPolicy indisponivel'
    }
    Save-CsvRel -Dados $linhasLp.ToArray() -Arquivo '05b_LabelPolicies.csv'
    $linhasAuto = New-Object System.Collections.ArrayList
    try {
        foreach ($a in @(Get-AutoSensitivityLabelPolicy -ErrorAction Stop)) {
            [void]$linhasAuto.Add([PSCustomObject]@{
                Politica = [string](Get-Prop -Obj $a -Nome 'Name')
                Modo     = [string](Get-Prop -Obj $a -Nome 'Mode')
                Label    = [string](Get-Prop -Obj $a -Nome 'ApplySensitivityLabel')
                Workloads = (Format-Lista -Valor (Get-Prop -Obj $a -Nome 'Workload' -Padrao @()))
                Habilitada = [string](Get-Prop -Obj $a -Nome 'Enabled')
            })
        }
        $ativas = @($linhasAuto.ToArray() | Where-Object { $_.Modo -like 'Enable*' -or $_.Habilitada -eq 'True' }).Count
        Set-CoberturaTenant -Workload 'Auto-labeling (E5)' -Ativo ($ativas -gt 0) -Detalhe ('Politicas de auto-labeling ativas: ' + $ativas)
    }
    catch {
        Write-Aviso ('Get-AutoSensitivityLabelPolicy falhou: ' + $_.Exception.Message)
        Set-Evidencia -Workload 'Auto-labeling (E5)' -Disponivel $false -Detalhe 'Cmdlet indisponivel'
    }
    Save-CsvRel -Dados $linhasAuto.ToArray() -Arquivo '05c_AutoLabeling.csv'
    $cobDlp = @{}
    $cobEndpoint = @{}
    $linhasDlp = New-Object System.Collections.ArrayList
    try {
        foreach ($p in @(Get-DlpCompliancePolicy -ErrorAction Stop)) {
            $modo = [string](Get-Prop -Obj $p -Nome 'Mode')
            $locs = @()
            foreach ($n in @('ExchangeLocation', 'SharePointLocation', 'OneDriveLocation', 'TeamsLocation')) {
                $locs += @(Get-Prop -Obj $p -Nome $n -Padrao @())
            }
            $endpointLocs = @(Get-Prop -Obj $p -Nome 'EndpointDlpLocation' -Padrao @())
            $escopo = Resolve-EntradasEscopo -Entradas ($locs + $endpointLocs)
            $escopoEp = Resolve-EntradasEscopo -Entradas $endpointLocs
            $ativa = ($modo -like 'Enable*')
            if ($ativa) {
                foreach ($k in $escopo.Keys) { $cobDlp[$k] = $true }
                foreach ($k in $escopoEp.Keys) { $cobEndpoint[$k] = $true }
            }
            [void]$linhasDlp.Add([PSCustomObject]@{
                Politica         = [string](Get-Prop -Obj $p -Nome 'Name')
                Modo             = $modo
                Ativa            = $ativa
                UsuariosNoEscopo = $escopo.Count
                EndpointDlp      = ($endpointLocs.Count -gt 0)
                UsuariosEndpoint = $escopoEp.Count
                Locais           = (Format-Lista -Valor $locs)
                CriadaEm         = [string](Get-Prop -Obj $p -Nome 'WhenCreated')
            })
        }
        Set-Cobertura -Workload 'DLP' -Upns @($cobDlp.Keys)
        Set-Evidencia -Workload 'DLP' -Disponivel $true -Detalhe ('Politicas DLP ativas: ' + (@($linhasDlp.ToArray() | Where-Object { $_.Ativa }).Count))
        Set-Cobertura -Workload 'Endpoint DLP' -Upns @($cobEndpoint.Keys)
        Set-Evidencia -Workload 'Endpoint DLP' -Disponivel $true -Detalhe ('Politicas DLP com location Endpoint: ' + (@($linhasDlp.ToArray() | Where-Object { $_.EndpointDlp }).Count))
    }
    catch {
        Write-Aviso ('Get-DlpCompliancePolicy falhou: ' + $_.Exception.Message)
        Set-Evidencia -Workload 'DLP' -Disponivel $false -Detalhe 'Cmdlet indisponivel'
        Set-Evidencia -Workload 'Endpoint DLP' -Disponivel $false -Detalhe 'Cmdlet indisponivel'
    }
    Save-CsvRel -Dados $linhasDlp.ToArray() -Arquivo '05d_DLP_Politicas.csv'
}
# ================================================================= FASE 5 ====
function Get-DadosGovernanca {
    Write-Etapa 'Fase 5 - Purview Governanca e Risco'
    $todos = @('Retencao / Records Management', 'Insider Risk Management', 'Communication Compliance', 'eDiscovery Premium', 'Audit Premium', 'Information Barriers', 'Customer Lockbox')
    if (-not $script:IppsOk) {
        Write-Aviso 'Security & Compliance nao conectado. Fase 5 pulada.'
        foreach ($w in $todos) { Set-Evidencia -Workload $w -Disponivel $false -Detalhe 'Security & Compliance PowerShell nao conectado' }
        return
    }
    $cobRet = @{}
    $linhasRet = New-Object System.Collections.ArrayList
    try {
        foreach ($p in @(Get-RetentionCompliancePolicy -ErrorAction Stop)) {
            $locs = @()
            foreach ($n in @('ExchangeLocation', 'SharePointLocation', 'OneDriveLocation', 'ModernGroupLocation', 'TeamsChatLocation', 'TeamsChannelLocation')) {
                $locs += @(Get-Prop -Obj $p -Nome $n -Padrao @())
            }
            $escopo = Resolve-EntradasEscopo -Entradas $locs
            $habilitada = [string](Get-Prop -Obj $p -Nome 'Enabled') -eq 'True'
            if ($habilitada) { foreach ($k in $escopo.Keys) { $cobRet[$k] = $true } }
            [void]$linhasRet.Add([PSCustomObject]@{
                Politica         = [string](Get-Prop -Obj $p -Nome 'Name')
                Habilitada       = $habilitada
                Modo             = [string](Get-Prop -Obj $p -Nome 'Mode')
                UsuariosNoEscopo = $escopo.Count
                Locais           = (Format-Lista -Valor $locs)
            })
        }
        Set-Cobertura -Workload 'Retencao / Records Management' -Upns @($cobRet.Keys)
        Set-Evidencia -Workload 'Retencao / Records Management' -Disponivel $true -Detalhe ('Politicas de retencao habilitadas: ' + (@($linhasRet.ToArray() | Where-Object { $_.Habilitada }).Count))
    }
    catch {
        Write-Aviso ('Get-RetentionCompliancePolicy falhou: ' + $_.Exception.Message)
        Set-Evidencia -Workload 'Retencao / Records Management' -Disponivel $false -Detalhe 'Cmdlet indisponivel'
    }
    Save-CsvRel -Dados $linhasRet.ToArray() -Arquivo '06_Retention.csv'
    $linhasIrm = New-Object System.Collections.ArrayList
    try {
        foreach ($p in @(Get-InsiderRiskPolicy -ErrorAction Stop)) {
            [void]$linhasIrm.Add([PSCustomObject]@{
                Politica  = [string](Get-Prop -Obj $p -Nome 'Name')
                Tipo      = [string](Get-Prop -Obj $p -Nome 'InsiderRiskScenario')
                Habilitada = [string](Get-Prop -Obj $p -Nome 'Enabled')
                Status    = [string](Get-Prop -Obj $p -Nome 'Status')
                CriadaEm  = [string](Get-Prop -Obj $p -Nome 'WhenCreated')
            })
        }
        Set-CoberturaTenant -Workload 'Insider Risk Management' -Ativo ($linhasIrm.Count -gt 0) -Detalhe ('Politicas de Insider Risk configuradas: ' + $linhasIrm.Count + ' (escopo por usuario nao exposto pela API - avaliado no nivel do tenant)')
    }
    catch {
        Write-Aviso ('Get-InsiderRiskPolicy falhou: ' + $_.Exception.Message)
        Set-Evidencia -Workload 'Insider Risk Management' -Disponivel $false -Detalhe 'Cmdlet indisponivel para a conta usada'
    }
    Save-CsvRel -Dados $linhasIrm.ToArray() -Arquivo '06b_InsiderRisk.csv'
    $linhasCc = New-Object System.Collections.ArrayList
    try {
        foreach ($p in @(Get-SupervisoryReviewPolicyV2 -ErrorAction Stop)) {
            [void]$linhasCc.Add([PSCustomObject]@{
                Politica   = [string](Get-Prop -Obj $p -Nome 'Name')
                Habilitada = [string](Get-Prop -Obj $p -Nome 'Enabled')
                Revisores  = (Format-Lista -Valor (Get-Prop -Obj $p -Nome 'Reviewers' -Padrao @()))
                CriadaEm   = [string](Get-Prop -Obj $p -Nome 'WhenCreated')
            })
        }
        Set-CoberturaTenant -Workload 'Communication Compliance' -Ativo ($linhasCc.Count -gt 0) -Detalhe ('Politicas de Communication Compliance: ' + $linhasCc.Count + ' (avaliado no nivel do tenant)')
    }
    catch {
        Write-Aviso ('Get-SupervisoryReviewPolicyV2 falhou: ' + $_.Exception.Message)
        Set-Evidencia -Workload 'Communication Compliance' -Disponivel $false -Detalhe 'Cmdlet indisponivel para a conta usada'
    }
    Save-CsvRel -Dados $linhasCc.ToArray() -Arquivo '06c_CommCompliance.csv'
    $linhasEd = New-Object System.Collections.ArrayList
    try {
        foreach ($c in @(Get-ComplianceCase -CaseType AdvancedEdiscovery -ErrorAction Stop)) {
            [void]$linhasEd.Add([PSCustomObject]@{
                Caso     = [string](Get-Prop -Obj $c -Nome 'Name')
                Status   = [string](Get-Prop -Obj $c -Nome 'Status')
                CriadoEm = [string](Get-Prop -Obj $c -Nome 'CreatedDateTime')
                CriadoPor = [string](Get-Prop -Obj $c -Nome 'CreatedBy')
            })
        }
        Set-CoberturaTenant -Workload 'eDiscovery Premium' -Ativo ($linhasEd.Count -gt 0) -Detalhe ('Casos de eDiscovery Premium: ' + $linhasEd.Count)
    }
    catch {
        Write-Aviso ('Get-ComplianceCase falhou: ' + $_.Exception.Message)
        Set-Evidencia -Workload 'eDiscovery Premium' -Disponivel $false -Detalhe 'Cmdlet indisponivel'
    }
    Save-CsvRel -Dados $linhasEd.ToArray() -Arquivo '06d_eDiscovery.csv'
    $linhasIb = New-Object System.Collections.ArrayList
    try {
        foreach ($p in @(Get-InformationBarrierPolicy -ErrorAction Stop)) {
            [void]$linhasIb.Add([PSCustomObject]@{
                Politica = [string](Get-Prop -Obj $p -Nome 'Name')
                Estado   = [string](Get-Prop -Obj $p -Nome 'State')
                Segmento = [string](Get-Prop -Obj $p -Nome 'AssignedSegment')
            })
        }
        Set-CoberturaTenant -Workload 'Information Barriers' -Ativo ($linhasIb.Count -gt 0) -Detalhe ('Politicas de Information Barriers: ' + $linhasIb.Count)
    }
    catch {
        Set-Evidencia -Workload 'Information Barriers' -Disponivel $false -Detalhe 'Cmdlet indisponivel'
    }
    Save-CsvRel -Dados $linhasIb.ToArray() -Arquivo '06f_InformationBarriers.csv'
    $linhasAudit = New-Object System.Collections.ArrayList
    if ($script:ExoOk) {
        $auditOrg = ''
        $lockbox = ''
        try {
            $org = Get-OrganizationConfig -ErrorAction Stop
            $auditOrg = [string](Get-Prop -Obj $org -Nome 'AuditDisabled')
            $lockbox = [string](Get-Prop -Obj $org -Nome 'CustomerLockBoxEnabled')
        }
        catch { Write-Aviso ('Get-OrganizationConfig falhou: ' + $_.Exception.Message) }
        $comAvancado = 0
        $amostra = 0
        try {
            foreach ($mb in @(Get-Mailbox -ResultSize $MaxMailboxesAudit -ErrorAction Stop)) {
                $amostra++
                $eventos = @(Get-Prop -Obj $mb -Nome 'AuditOwner' -Padrao @())
                if ($eventos -contains 'MailItemsAccessed' -or $eventos -contains 'Send') { $comAvancado++ }
            }
        }
        catch { Write-Aviso ('Get-Mailbox falhou: ' + $_.Exception.Message) }
        $pctPremium = 0
        if ($amostra -gt 0) { $pctPremium = [math]::Round(($comAvancado / $amostra) * 100, 1) }
        [void]$linhasAudit.Add([PSCustomObject]@{
            AuditoriaDesabilitadaNoTenant = $auditOrg
            CustomerLockboxHabilitado     = $lockbox
            MailboxesAmostradas           = $amostra
            MailboxesComEventosPremium    = $comAvancado
            PctAmostraComPremium          = $pctPremium
        })
        Set-CoberturaTenant -Workload 'Audit Premium' -Ativo ($comAvancado -gt 0) -Detalhe ('Amostra de ' + $amostra + ' mailboxes: ' + $comAvancado + ' com eventos de Audit Premium (MailItemsAccessed)')
        Set-CoberturaTenant -Workload 'Customer Lockbox' -Ativo ($lockbox -eq 'True') -Detalhe ('CustomerLockBoxEnabled = ' + $lockbox)
    }
    else {
        Set-Evidencia -Workload 'Audit Premium' -Disponivel $false -Detalhe 'Exchange Online PowerShell nao conectado'
        Set-Evidencia -Workload 'Customer Lockbox' -Disponivel $false -Detalhe 'Exchange Online PowerShell nao conectado'
    }
    Save-CsvRel -Dados $linhasAudit.ToArray() -Arquivo '06e_Audit.csv'
}
# ================================================================= FASE 6 ====
function Get-DadosSecureScore {
    Write-Etapa 'Fase 6 - Secure Score e checklist manual'
    $ss = @(Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/security/secureScores?$top=1' -Silencioso)
    $linhas = New-Object System.Collections.ArrayList
    if ($ss.Count -gt 0) {
        $atual = $ss[0]
        $cur = [double](Get-Prop -Obj $atual -Nome 'currentScore' -Padrao 0)
        $max = [double](Get-Prop -Obj $atual -Nome 'maxScore' -Padrao 0)
        if ($max -gt 0) { $script:SecureScorePct = [math]::Round(($cur / $max) * 100, 1) }
        Write-Ok ('Secure Score: ' + $cur + ' / ' + $max + ' (' + $script:SecureScorePct + '%)')
        foreach ($c in @(Get-Prop -Obj $atual -Nome 'controlScores' -Padrao @())) {
            [void]$linhas.Add([PSCustomObject]@{
                Controle   = [string](Get-Prop -Obj $c -Nome 'controlName')
                Categoria  = [string](Get-Prop -Obj $c -Nome 'controlCategory')
                Pontos     = [string](Get-Prop -Obj $c -Nome 'score')
                Estado     = [string](Get-Prop -Obj $c -Nome 'implementationStatus')
                Descricao  = [string](Get-Prop -Obj $c -Nome 'description')
                ScoreAtual = $cur
                ScoreMaximo = $max
            })
        }
    }
    else {
        Write-Aviso 'Secure Score indisponivel (verifique SecurityEvents.Read.All).'
    }
    Save-CsvRel -Dados $linhas.ToArray() -Arquivo '07_SecureScore.csv'
    $checklist = @(
        [PSCustomObject]@{ Item = 'Microsoft Priva - Privacy Risk Management'; Portal = 'https://purview.microsoft.com/priva'; Motivo = 'Sem API publica de leitura; validar manualmente se ha politicas de risco de privacidade ativas'; StatusManual = '' },
        [PSCustomObject]@{ Item = 'Microsoft Priva - Subject Rights Requests'; Portal = 'https://purview.microsoft.com/priva'; Motivo = 'Verificar se ha solicitacoes de titular sendo tratadas na ferramenta'; StatusManual = '' },
        [PSCustomObject]@{ Item = 'Compliance Manager - score e avaliacoes'; Portal = 'https://purview.microsoft.com/compliancemanager'; Motivo = 'Score e avaliacoes premium do E5 nao expostos via Graph'; StatusManual = '' },
        [PSCustomObject]@{ Item = 'Defender for Cloud Apps - app governance e politicas'; Portal = 'https://security.microsoft.com/cloudapps'; Motivo = 'API do MDA exige URL de tenant dedicada; validar politicas e conectores de app'; StatusManual = '' },
        [PSCustomObject]@{ Item = 'Defender for Office 365 - Attack Simulation Training'; Portal = 'https://security.microsoft.com/attacksimulator'; Motivo = 'Recurso exclusivo do P2; confirmar campanhas executadas nos ultimos 12 meses'; StatusManual = '' },
        [PSCustomObject]@{ Item = 'Defender for Endpoint - Vulnerability Management'; Portal = 'https://security.microsoft.com/tvm_dashboard'; Motivo = 'Confirmar uso de recomendacoes de seguranca e exposure score'; StatusManual = '' }
    )
    Save-CsvRel -Dados $checklist -Arquivo '07b_Checklist_Manual.csv'
}
# ================================================================= FASE 7 ====
function New-Cobertura {
    Write-Etapa 'Fase 7 - Motor de cobertura, scorecard e recomendacoes'
    $usuarios = @($script:Usuarios.Values)
    $linhasWl = New-Object System.Collections.ArrayList
    $colunas = New-Object System.Collections.ArrayList
    foreach ($w in $script:Catalogo) {
        $nome = $w.Nome
        $entitled = New-Object System.Collections.ArrayList
        foreach ($u in $usuarios) {
            $tem = $false
            foreach ($p in $w.Planos) { if ($u.Planos.ContainsKey($p)) { $tem = $true } }
            if ($tem) { [void]$entitled.Add($u.Upn.ToLower()) }
        }
        $evOk = $false
        if ($script:EvidenciaOk.ContainsKey($nome)) { $evOk = [bool]$script:EvidenciaOk[$nome] }
        $tenantLevel = $script:TenantAtivo.ContainsKey($nome)
        $cobertos = New-Object System.Collections.ArrayList
        if ($evOk) {
            if ($tenantLevel) {
                if ($script:TenantAtivo[$nome]) { foreach ($u in $entitled) { [void]$cobertos.Add($u) } }
            }
            elseif ($script:Cobertura.ContainsKey($nome)) {
                foreach ($u in $entitled) { if ($script:Cobertura[$nome].ContainsKey($u)) { [void]$cobertos.Add($u) } }
            }
        }
        $pct = 0
        if ($entitled.Count -gt 0) { $pct = [math]::Round(($cobertos.Count / $entitled.Count) * 100, 1) }
        $maturidade = 'Sem evidencia'
        if ($evOk) {
            if ($entitled.Count -eq 0) { $maturidade = 'Nao licenciado' }
            elseif ($pct -eq 0) { $maturidade = 'Nao implantado' }
            elseif ($pct -lt 50) { $maturidade = 'Inicial' }
            elseif ($pct -lt 80) { $maturidade = 'Parcial' }
            else { $maturidade = 'Implantado' }
        }
        $detalhe = ''
        if ($script:EvidenciaDetalhe.ContainsKey($nome)) { $detalhe = [string]$script:EvidenciaDetalhe[$nome] }
        $colCoberto = ''
        $colGap = ''
        $colPct = ''
        if ($evOk) {
            $colCoberto = $cobertos.Count
            $colGap = ($entitled.Count - $cobertos.Count)
            $colPct = $pct
        }
        [void]$linhasWl.Add([PSCustomObject]@{
            Workload            = $nome
            Pilar               = $w.Pilar
            ServicePlans        = ($w.Planos -join '; ')
            Entitled            = $entitled.Count
            Coberto             = $colCoberto
            Gap                 = $colGap
            PctAdocao           = $colPct
            Maturidade          = $maturidade
            EscopoTenant        = $tenantLevel
            EvidenciaDisponivel = $evOk
            Evidencia           = $w.Evidencia
            Detalhe             = $detalhe
        })
        [void]$colunas.Add([PSCustomObject]@{
            Nome = $nome
            Coluna = ($nome -replace '[^A-Za-z0-9]', '_')
            Entitled = @{}
            Cobertos = @{}
            EvOk = $evOk
            TenantLevel = $tenantLevel
        })
        $ultimo = $colunas[$colunas.Count - 1]
        foreach ($u in $entitled) { $ultimo.Entitled[$u] = $true }
        foreach ($u in $cobertos) { $ultimo.Cobertos[$u] = $true }
    }
    Save-CsvRel -Dados $linhasWl.ToArray() -Arquivo '08b_Cobertura_PorWorkload.csv'
    $linhasUsr = New-Object System.Collections.ArrayList
    foreach ($u in $usuarios) {
        $chave = $u.Upn.ToLower()
        $nomeSaida = $u.DisplayName
        if ($Anonimizar) { $nomeSaida = '' }
        $linha = [ordered]@{
            Upn         = (Get-UpnSaida -Upn $u.Upn)
            DisplayName = $nomeSaida
            Habilitado  = $u.Habilitado
            TemE5       = $u.TemE5
        }
        foreach ($c in $colunas) {
            $valor = 'NaoLicenciado'
            if ($c.Entitled.ContainsKey($chave)) {
                if (-not $c.EvOk) { $valor = 'SemEvidencia' }
                elseif ($c.Cobertos.ContainsKey($chave)) { $valor = 'Coberto' }
                else { $valor = 'NaoCoberto' }
            }
            $linha[$c.Coluna] = $valor
        }
        [void]$linhasUsr.Add([PSCustomObject]$linha)
    }
    Save-CsvRel -Dados $linhasUsr.ToArray() -Arquivo '08_Cobertura_PorUsuario.csv'
    $recs = New-Object System.Collections.ArrayList
    foreach ($w in @($linhasWl.ToArray())) {
        if (-not $w.EvidenciaDisponivel) {
            [void]$recs.Add([PSCustomObject]@{
                Prioridade  = 4
                Workload    = $w.Workload
                Pilar       = $w.Pilar
                Achado      = 'Sem evidencia coletada para este workload'
                Recomendacao = 'Rodar novamente com as permissoes necessarias ou validar manualmente no portal. Detalhe: ' + $w.Detalhe
                Impacto     = 'Baixo'
                Esforco     = 'Baixo'
            })
            continue
        }
        if ([int]$w.Entitled -eq 0) { continue }
        if ([double]$w.PctAdocao -eq 0) {
            [void]$recs.Add([PSCustomObject]@{
                Prioridade  = 1
                Workload    = $w.Workload
                Pilar       = $w.Pilar
                Achado      = ([string]$w.Entitled + ' usuarios licenciados e nenhum coberto - workload pago e nao utilizado')
                Recomendacao = 'Planejar implantacao completa do workload. Evidencia esperada: ' + $w.Evidencia
                Impacto     = 'Alto'
                Esforco     = 'Alto'
            })
        }
        elseif ([double]$w.PctAdocao -lt 80) {
            [void]$recs.Add([PSCustomObject]@{
                Prioridade  = 2
                Workload    = $w.Workload
                Pilar       = $w.Pilar
                Achado      = ('Apenas ' + [string]$w.PctAdocao + '% dos ' + [string]$w.Entitled + ' usuarios licenciados estao cobertos (' + [string]$w.Gap + ' sem cobertura)')
                Recomendacao = 'Ampliar o escopo das politicas ate cobrir todos os usuarios licenciados. Evidencia esperada: ' + $w.Evidencia
                Impacto     = 'Medio'
                Esforco     = 'Medio'
            })
        }
    }
    $ociosas = @($usuarios | Where-Object { $_.TemE5 -and (-not $_.Habilitado) }).Count
    if ($ociosas -gt 0) {
        [void]$recs.Add([PSCustomObject]@{
            Prioridade  = 2
            Workload    = 'Licenciamento'
            Pilar       = 'Licenciamento'
            Achado      = ([string]$ociosas + ' licencas E5 atribuidas a contas desabilitadas')
            Recomendacao = 'Revogar as licencas E5 dessas contas e realocar para usuarios ativos'
            Impacto     = 'Medio'
            Esforco     = 'Baixo'
        })
    }
    if ($script:UsuariosSemMfa -and [int]$script:UsuariosSemMfa -gt 0) {
        [void]$recs.Add([PSCustomObject]@{
            Prioridade  = 1
            Workload    = 'MFA / Autenticacao forte'
            Pilar       = 'Identidade'
            Achado      = ([string]$script:UsuariosSemMfa + ' usuarios sem metodo de MFA registrado')
            Recomendacao = 'Executar campanha de registro de MFA e exigir MFA por Acesso Condicional'
            Impacto     = 'Alto'
            Esforco     = 'Medio'
        })
    }
    Save-CsvRel -Dados @($recs.ToArray() | Sort-Object Prioridade) -Arquivo '09_Recomendacoes.csv'
    $avaliados = @($linhasWl.ToArray() | Where-Object { $_.EvidenciaDisponivel -and [int]$_.Entitled -gt 0 })
    $score = 0
    if ($avaliados.Count -gt 0) {
        $soma = 0
        foreach ($a in $avaliados) { $soma = $soma + [double]$a.PctAdocao }
        $score = [math]::Round($soma / $avaliados.Count, 1)
    }
    $compradosE5 = 0
    $consumidosE5 = 0
    foreach ($s in @($script:SkusE5Linhas)) {
        $compradosE5 = $compradosE5 + [int]$s.Comprados
        $consumidosE5 = $consumidosE5 + [int]$s.Consumidos
    }
    $sumario = [PSCustomObject]@{
        DataExecucao          = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        TenantId              = $script:TenantId
        ExecutadoPor          = $AdminUPN
        PeriodoDias           = $DiasPeriodo
        TotalUsuarios         = $script:Usuarios.Count
        UsuariosComE5         = @($usuarios | Where-Object { $_.TemE5 }).Count
        LicencasE5Compradas   = $compradosE5
        LicencasE5Consumidas  = $consumidosE5
        LicencasE5Disponiveis = ($compradosE5 - $consumidosE5)
        LicencasE5Ociosas     = @($usuarios | Where-Object { $_.TemE5 -and (-not $_.Habilitado) }).Count
        ScoreAdocaoE5Pct      = $score
        WorkloadsAvaliados    = $avaliados.Count
        WorkloadsNaoImplantados = @($linhasWl.ToArray() | Where-Object { $_.Maturidade -eq 'Nao implantado' }).Count
        WorkloadsSemEvidencia = @($linhasWl.ToArray() | Where-Object { -not $_.EvidenciaDisponivel }).Count
        UsuariosSemMfa        = $script:UsuariosSemMfa
        AdminsSemMfa          = $script:AdminsSemMfa
        TotalAdmins           = $script:TotalAdmins
        AdminsElegiveisPim    = $script:AdminsComPim
        DispositivosOnboarded = $script:DispositivosOnboarded
        AlertasNoPeriodo      = $script:TotalAlertas
        DeteccoesDeRisco      = $script:RiskDetections
        AccessReviews         = $script:AccessReviews
        SecureScorePct        = $script:SecureScorePct
        Anonimizado           = $Anonimizar
    }
    Save-CsvRel -Dados @($sumario) -Arquivo '00_Sumario.csv'
    $script:ScoreFinal = $score
    $script:ResumoWorkloads = $linhasWl.ToArray()
}
# =========================================================== ORQUESTRADOR ====
$script:Catalogo = @(
    @{ Nome = 'Entra ID P1'; Pilar = 'Identidade'; Planos = @('AAD_PREMIUM'); Evidencia = 'Usuario no escopo de pelo menos uma politica de Acesso Condicional habilitada' },
    @{ Nome = 'Entra ID P2'; Pilar = 'Identidade'; Planos = @('AAD_PREMIUM_P2'); Evidencia = 'Usuario coberto por CA baseada em risco, elegivel via PIM, avaliado pelo Identity Protection ou no escopo de Access Review ativa' },
    @{ Nome = 'MFA / Autenticacao forte'; Pilar = 'Identidade'; Planos = @('AAD_PREMIUM', 'AAD_PREMIUM_P2'); Evidencia = 'Usuario com metodo de MFA registrado' },
    @{ Nome = 'Defender for Endpoint P2'; Pilar = 'Defender'; Planos = @('WINDEFATP'); Evidencia = 'Usuario com logon em dispositivo onboarded no MDE' },
    @{ Nome = 'Defender for Office 365 P2'; Pilar = 'Defender'; Planos = @('THREAT_INTELLIGENCE'); Evidencia = 'Usuario no escopo de Safe Links / Safe Attachments / Anti-phishing' },
    @{ Nome = 'Defender for Identity'; Pilar = 'Defender'; Planos = @('ATA'); Evidencia = 'Conta com eventos capturados por sensores MDI' },
    @{ Nome = 'Defender for Cloud Apps'; Pilar = 'Defender'; Planos = @('ADALLOM_S_STANDARD', 'ADALLOM_S_O365'); Evidencia = 'Conta com atividade registrada em CloudAppEvents' },
    @{ Nome = 'Sensitivity Labels (MIP)'; Pilar = 'Purview'; Planos = @('MIP_S_CLP1', 'MIP_S_CLP2', 'RMS_S_PREMIUM', 'RMS_S_PREMIUM2'); Evidencia = 'Usuario no escopo de politica de publicacao de rotulos' },
    @{ Nome = 'Auto-labeling (E5)'; Pilar = 'Purview'; Planos = @('MIP_S_CLP2'); Evidencia = 'Tenant com politicas de auto-labeling ativas' },
    @{ Nome = 'DLP'; Pilar = 'Purview'; Planos = @('MIP_S_CLP1', 'MIP_S_CLP2'); Evidencia = 'Usuario no escopo de politica DLP ativa' },
    @{ Nome = 'Endpoint DLP'; Pilar = 'Purview'; Planos = @('MICROSOFTENDPOINTDLP'); Evidencia = 'Usuario no escopo de politica DLP com location Endpoint' },
    @{ Nome = 'Retencao / Records Management'; Pilar = 'Governanca'; Planos = @('RECORDS_MANAGEMENT'); Evidencia = 'Usuario no escopo de politica de retencao habilitada' },
    @{ Nome = 'Insider Risk Management'; Pilar = 'Governanca'; Planos = @('INSIDER_RISK', 'INSIDER_RISK_MANAGEMENT'); Evidencia = 'Tenant com politicas de Insider Risk configuradas' },
    @{ Nome = 'Communication Compliance'; Pilar = 'Governanca'; Planos = @('COMMUNICATIONS_COMPLIANCE'); Evidencia = 'Tenant com politicas de Communication Compliance configuradas' },
    @{ Nome = 'eDiscovery Premium'; Pilar = 'Governanca'; Planos = @('EQUIVIO_ANALYTICS'); Evidencia = 'Tenant com casos de eDiscovery Premium criados' },
    @{ Nome = 'Audit Premium'; Pilar = 'Governanca'; Planos = @('M365_ADVANCED_AUDITING'); Evidencia = 'Mailboxes com eventos de auditoria premium habilitados' },
    @{ Nome = 'Information Barriers'; Pilar = 'Governanca'; Planos = @('INFORMATION_BARRIERS'); Evidencia = 'Tenant com politicas de Information Barriers ativas' },
    @{ Nome = 'Customer Lockbox'; Pilar = 'Governanca'; Planos = @('LOCKBOX_ENTERPRISE'); Evidencia = 'Customer Lockbox habilitado na organizacao' }
)
function Start-E5Assessment {
    $inicio = Get-Date
    Write-Host ''
    Write-Host '  Microsoft 365 E5 - Assessment de adocao de seguranca e Purview' -ForegroundColor Yellow
    Write-Host '  Somente leitura. Nenhuma configuracao do tenant e alterada.' -ForegroundColor Yellow
    Write-Host ''
    $script:DispositivosOnboarded = 0
    $script:TotalAlertas = 0
    $script:RiskDetections = 0
    $script:AccessReviews = 0
    $script:UsuariosSemMfa = 0
    $script:AdminsSemMfa = 0
    $script:TotalAdmins = 0
    $script:AdminsComPim = 0
    $script:SkusE5Linhas = @()
    $preflight = @(Initialize-Ambiente)
    if ($preflight.Count -eq 0 -or -not $preflight[$preflight.Count - 1]) {
        Write-Falha 'Pre-flight falhou. Assessment abortado.'
        return
    }
    try { Get-DadosLicencas } catch { Write-Falha ('Fase 1 falhou: ' + $_.Exception.Message) }
    if ($script:Usuarios.Count -eq 0) {
        Write-Falha 'Nenhum usuario coletado. Verifique as permissoes User.Read.All / Directory.Read.All.'
        return
    }
    if (-not $SkipEntra) {
        try { Get-DadosEntra } catch { Write-Falha ('Fase 2 falhou: ' + $_.Exception.Message) }
    }
    if (-not $SkipDefender) {
        try { Get-DadosDefender } catch { Write-Falha ('Fase 3 falhou: ' + $_.Exception.Message) }
    }
    if (-not $SkipPurview) {
        try { Get-DadosPurview } catch { Write-Falha ('Fase 4 falhou: ' + $_.Exception.Message) }
        try { Get-DadosGovernanca } catch { Write-Falha ('Fase 5 falhou: ' + $_.Exception.Message) }
    }
    try { Get-DadosSecureScore } catch { Write-Falha ('Fase 6 falhou: ' + $_.Exception.Message) }
    try { New-Cobertura } catch { Write-Falha ('Fase 7 falhou: ' + $_.Exception.Message) }
    Write-Etapa 'Resultado'
    foreach ($w in @($script:ResumoWorkloads)) {
        $cor = 'Gray'
        if ($w.Maturidade -eq 'Implantado') { $cor = 'Green' }
        elseif ($w.Maturidade -eq 'Nao implantado') { $cor = 'Red' }
        elseif ($w.Maturidade -eq 'Parcial' -or $w.Maturidade -eq 'Inicial') { $cor = 'Yellow' }
        $txt = '  {0,-32} entitled {1,6}  coberto {2,6}  {3,6}%  {4}' -f $w.Workload, $w.Entitled, $w.Coberto, $w.PctAdocao, $w.Maturidade
        Write-Host $txt -ForegroundColor $cor
    }
    Write-Host ''
    Write-Host ('  Score de adocao E5: ' + $script:ScoreFinal + '%') -ForegroundColor Cyan
    Write-Host ('  Duracao: ' + [math]::Round(((Get-Date) - $inicio).TotalMinutes, 1) + ' min') -ForegroundColor Cyan
    Write-Host ('  Pasta:   ' + $script:PastaSaida) -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  Proximo passo: abra E5-Assessment-Dashboard.html no navegador e selecione a pasta acima.' -ForegroundColor Yellow
    Write-Host ''
}
Start-E5Assessment
