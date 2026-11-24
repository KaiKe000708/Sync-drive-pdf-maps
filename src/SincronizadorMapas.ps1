<#
.SYNOPSIS
    Sincronizador de Mapas PDF, versão 2.1

    Copia para o Google Drive somente os arquivos .pdf da pasta principal do servidor.
    Subpastas da origem são ignoradas e nada é apagado no destino.

.DESCRIPTION
    Sem parâmetros: abre a interface.
    -Auto           executa sem interface (para o Agendador de Tarefas do Windows).
    -Auto -Simular  mostra o que seria feito, sem copiar nada.

    Códigos de saída do modo -Auto:
        0  concluído sem falhas
        1  concluído com falhas ou interrompido
        2  não executou (origem ou destino indisponível, outra execução em andamento, erro grave)

    Configurações: config.json, na mesma pasta deste arquivo.
    Histórico de alterações: historico.csv, na pasta de logs.
#>
[CmdletBinding()]
param(
    [switch]$Auto,
    [switch]$Simular
)

$ErrorActionPreference = 'Stop'
$CaminhoConfig = Join-Path $PSScriptRoot 'config.json'
$CaminhoEstado = Join-Path $PSScriptRoot 'estado.json'

# =====================================================================================
#  MOTOR DE SINCRONIZAÇÃO
#  Fica num texto separado porque também é carregado na thread de segundo plano da
#  interface. Assim a janela nunca trava enquanto copia.
# =====================================================================================
$CodigoMotor = @'
$ErrorActionPreference = 'Stop'

function Get-ConfigPadrao {
    [ordered]@{
        Origem            = '\\servidor\compartilhamento\MAPAS PDF'
        Destino           = 'G:\Meu Drive\MAPAS\MAPAS PDF'
        PastaLogs         = 'G:\Meu Drive\MAPAS\_LOGS'
        PastaArquivados   = 'G:\Meu Drive\MAPAS\_ARQUIVADOS'
        EspelharDrive     = $true
        Tema              = 'Sistema'
        DiasManterLogs    = 90
        Tentativas        = 3
        EsperaSegundos    = 5
    }
}

function Read-SyncConfig {
    param([string]$Caminho)
    $cfg = Get-ConfigPadrao
    if (Test-Path -LiteralPath $Caminho -PathType Leaf) {
        try {
            $json = Get-Content -LiteralPath $Caminho -Raw -Encoding UTF8 | ConvertFrom-Json
        } catch {
            throw "O arquivo config.json está com erro de formatação: $($_.Exception.Message)"
        }
        if ($null -eq $json) { throw 'O arquivo config.json está vazio.' }
        foreach ($chave in @($cfg.Keys)) {
            $prop = $json.PSObject.Properties[$chave]
            if ($null -ne $prop -and $null -ne $prop.Value -and "$($prop.Value)".Trim() -ne '') {
                $cfg[$chave] = $prop.Value
            }
        }
    }
    $cfg.Origem            = "$($cfg.Origem)".Trim()
    $cfg.Destino           = "$($cfg.Destino)".Trim()
    $cfg.PastaLogs         = "$($cfg.PastaLogs)".Trim()
    $cfg.PastaArquivados   = "$($cfg.PastaArquivados)".Trim()
    $cfg.EspelharDrive = ($cfg.EspelharDrive -eq $true -or "$($cfg.EspelharDrive)" -eq 'true')
    if (@('Claro', 'Escuro', 'Sistema') -notcontains "$($cfg.Tema)") { $cfg.Tema = 'Sistema' }
    $cfg.Tema              = (Get-Culture).TextInfo.ToTitleCase("$($cfg.Tema)".ToLower())
    $cfg.DiasManterLogs    = [Math]::Max(0, [int]$cfg.DiasManterLogs)
    $cfg.Tentativas        = [Math]::Max(1, [Math]::Min(10, [int]$cfg.Tentativas))
    $cfg.EsperaSegundos    = [Math]::Max(0, [Math]::Min(120, [int]$cfg.EsperaSegundos))
    return $cfg
}

function Save-SyncConfig {
    param($Config, [string]$Caminho)
    $texto = $Config | ConvertTo-Json
    [IO.File]::WriteAllText($Caminho, $texto, (New-Object Text.UTF8Encoding($true)))
}

function Read-Estado {
    param([string]$Caminho)
    if ([string]::IsNullOrWhiteSpace($Caminho)) { return $null }
    if (-not (Test-Path -LiteralPath $Caminho -PathType Leaf)) { return $null }
    try { return (Get-Content -LiteralPath $Caminho -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}

function Get-ExcecaoBase {
    param($Erro)
    $ex = $Erro
    if ($Erro -is [System.Management.Automation.ErrorRecord]) { $ex = $Erro.Exception }
    while (($ex -is [System.Management.Automation.MethodInvocationException] -or
            $ex -is [System.Management.Automation.MethodException]) -and $null -ne $ex.InnerException) {
        $ex = $ex.InnerException
    }
    return $ex
}

function Get-MensagemErro {
    param($Erro)
    $ex = Get-ExcecaoBase $Erro
    if ($null -eq $ex) { return "$Erro" }
    # Junta as mensagens internas: erros de interface costumam esconder a causa real nelas.
    $partes = New-Object 'System.Collections.Generic.List[string]'
    $atual = $ex
    while ($null -ne $atual -and $partes.Count -lt 4) {
        $msg = "$($atual.Message)".Trim()
        if ($atual.GetType().Name -eq 'XamlParseException' -and $atual.LineNumber -gt 0) { $msg += " (linha $($atual.LineNumber))" }
        if ($msg -and -not $partes.Contains($msg)) { $partes.Add($msg) }
        $atual = $atual.InnerException
    }
    return ($partes -join ' ')
}

function Get-Plural {
    param([int]$N, [string]$Singular, [string]$Plural)
    if ($N -eq 1) { return "1 $Singular" }
    return "$N $Plural"
}

function Format-Tamanho {
    param([long]$Bytes)
    $br = [Globalization.CultureInfo]::GetCultureInfo('pt-BR')
    if ($Bytes -ge 1GB) { return ($Bytes / 1GB).ToString('N2', $br) + ' GB' }
    if ($Bytes -ge 1MB) { return ($Bytes / 1MB).ToString('N1', $br) + ' MB' }
    if ($Bytes -ge 1KB) { return ($Bytes / 1KB).ToString('N0', $br) + ' KB' }
    return "$Bytes B"
}

function Get-TextoResumo {
    param([hashtable]$Resumo)
    $partes = @()
    if ($Resumo.Adicionados -gt 0) { $partes += Get-Plural $Resumo.Adicionados 'adicionado' 'adicionados' }
    if ($Resumo.Atualizados -gt 0) { $partes += Get-Plural $Resumo.Atualizados 'atualizado' 'atualizados' }
    if ($Resumo.Arquivados  -gt 0) { $partes += Get-Plural $Resumo.Arquivados 'arquivado' 'arquivados' }
    if ($Resumo.Falhas      -gt 0) { $partes += Get-Plural $Resumo.Falhas 'falha' 'falhas' }
    if ($partes.Count -eq 0) { return 'nenhuma alteração' }
    return ($partes -join ', ')
}

function Send-SyncEvent {
    param($Sync, [hashtable]$Dados)
    if ($null -ne $Sync -and $null -ne $Sync.Fila) { $Sync.Fila.Enqueue($Dados) }
}

function Write-SyncLog {
    param([hashtable]$Ctx, [string]$Nivel, [string]$Mensagem)
    if ($Nivel -eq 'ERRO' -and -not $Ctx.PrimeiroErro) { $Ctx.PrimeiroErro = $Mensagem }
    $linha = '{0}  {1,-5}  {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Nivel, $Mensagem
    if ($Ctx.LogFile) {
        try { [IO.File]::AppendAllText($Ctx.LogFile, $linha + [Environment]::NewLine, (New-Object Text.UTF8Encoding($true))) } catch { }
    }
    if ($Ctx.Console) {
        $cor = 'Gray'
        if ($Nivel -eq 'OK')    { $cor = 'Green' }
        if ($Nivel -eq 'AVISO') { $cor = 'Yellow' }
        if ($Nivel -eq 'ERRO')  { $cor = 'Red' }
        Write-Host $linha -ForegroundColor $cor
    }
    Send-SyncEvent $Ctx.Sync @{ Tipo = 'log'; Nivel = $Nivel; Mensagem = $Mensagem }
}

# ------------------------------------------------------------------ histórico (planilha)
# Uma linha por alteração feita no Drive. Separador ";" para abrir direto no Excel.
$script:CabecalhoHistorico = '"Data e hora";"Execução";"Ação";"Arquivo";"Bytes";"Detalhe"'

function Initialize-Historico {
    param([string]$Pasta)
    if ([string]::IsNullOrWhiteSpace($Pasta)) { return $null }
    try {
        $arquivo = [IO.Path]::Combine($Pasta, 'historico.csv')
        if (-not [IO.File]::Exists($arquivo)) {
            [IO.File]::WriteAllText($arquivo, $script:CabecalhoHistorico + "`r`n", (New-Object Text.UTF8Encoding($true)))
        }
        return $arquivo
    } catch { return $null }
}

function ConvertTo-LinhaCsv {
    param([object[]]$Campos)
    $saida = foreach ($c in $Campos) {
        $t = ("$c" -replace "[`r`n]+", ' ') -replace '"', '""'
        '"' + $t + '"'
    }
    return ($saida -join ';')
}

function Add-Historico {
    param([hashtable]$Ctx, [string]$Acao, [string]$Arquivo, [long]$Bytes = 0, [string]$Detalhe = '')
    if (-not $Ctx.Historico) { return }
    $tam = ''
    if ($Bytes -gt 0) { $tam = "$Bytes" }
    $linha = ConvertTo-LinhaCsv @((Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Ctx.Execucao, $Acao, $Arquivo, $tam, $Detalhe)
    try { [IO.File]::AppendAllText($Ctx.Historico, $linha + "`r`n", (New-Object Text.UTF8Encoding($true))) } catch { }
}

function Read-Historico {
    param($Config)
    $pastas = @($Config.PastaLogs, [IO.Path]::Combine($env:LOCALAPPDATA, 'SincronizadorMapas', 'Logs'))
    $lista  = New-Object 'System.Collections.Generic.List[object]'
    $lidos  = @{}
    $inv    = [Globalization.CultureInfo]::InvariantCulture
    foreach ($p in $pastas) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $arquivo = [IO.Path]::Combine($p, 'historico.csv')
        if ($lidos.ContainsKey($arquivo.ToLowerInvariant()) -or -not [IO.File]::Exists($arquivo)) { continue }
        $lidos[$arquivo.ToLowerInvariant()] = $true
        $linhas = @([IO.File]::ReadAllLines($arquivo, [Text.Encoding]::UTF8) | Select-Object -Skip 1 | Where-Object { $_.Trim() -ne '' })
        if ($linhas.Count -eq 0) { continue }
        $dados = $linhas | ConvertFrom-Csv -Delimiter ';' -Header 'DataHora', 'Execucao', 'Acao', 'Arquivo', 'Bytes', 'Detalhe'
        foreach ($d in $dados) {
            $quando = [datetime]::MinValue
            if (-not [datetime]::TryParseExact("$($d.DataHora)", 'yyyy-MM-dd HH:mm:ss', $inv, [Globalization.DateTimeStyles]::None, [ref]$quando)) { continue }
            [long]$bytes = 0
            [void][long]::TryParse("$($d.Bytes)", [ref]$bytes)
            $lista.Add([pscustomobject]@{
                Quando   = $quando
                Execucao = "$($d.Execucao)"
                Acao     = "$($d.Acao)"
                Arquivo  = "$($d.Arquivo)"
                Bytes    = $bytes
                Detalhe  = "$($d.Detalhe)"
            })
        }
    }
    $ordenada = @($lista | Sort-Object -Property Quando -Descending)
    return ,$ordenada
}

function Get-ArquivoHistorico {
    param($Config)
    foreach ($p in @($Config.PastaLogs, [IO.Path]::Combine($env:LOCALAPPDATA, 'SincronizadorMapas', 'Logs'))) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $arquivo = [IO.Path]::Combine($p, 'historico.csv')
        if ([IO.File]::Exists($arquivo)) { return $arquivo }
    }
    return [IO.Path]::Combine($Config.PastaLogs, 'historico.csv')
}

# ------------------------------------------------------------------ verificações

function Test-MutexAbandonado {
    param($Erro)
    $ex = $Erro.Exception
    while ($null -ne $ex) {
        if ($ex -is [System.Threading.AbandonedMutexException]) { return $true }
        $ex = $ex.InnerException
    }
    return $false
}

# Confere se origem e destino estão acessíveis antes de qualquer cópia.
function Test-SyncAmbiente {
    param($Config, [switch]$CriarDestino)
    $problemas = New-Object 'System.Collections.Generic.List[string]'
    $origemOk = $false; $destinoOk = $false; $destinoExiste = $false

    if ([string]::IsNullOrWhiteSpace($Config.Origem)) {
        $problemas.Add('A pasta de origem não foi definida.')
    } elseif (Test-Path -LiteralPath $Config.Origem -PathType Container) {
        $origemOk = $true
    } else {
        $problemas.Add("Não foi possível acessar a origem ($($Config.Origem)). Verifique a conexão com a rede e o servidor.")
    }

    if ([string]::IsNullOrWhiteSpace($Config.Destino)) {
        $problemas.Add('A pasta de destino não foi definida.')
    } elseif (Test-Path -LiteralPath $Config.Destino -PathType Container) {
        $destinoOk = $true; $destinoExiste = $true
    } else {
        $pai = [IO.Path]::GetDirectoryName($Config.Destino.TrimEnd('\'))
        if ($pai -and (Test-Path -LiteralPath $pai -PathType Container)) {
            $destinoOk = $true
            if ($CriarDestino) {
                [void][IO.Directory]::CreateDirectory($Config.Destino)
                $destinoExiste = $true
            }
        } else {
            $problemas.Add("A pasta do Google Drive não está disponível ($pai). Verifique se o Google Drive está aberto e conectado.")
        }
    }

    if ($origemOk -and $destinoOk) {
        try {
            $a = [IO.Path]::GetFullPath($Config.Origem).TrimEnd('\')
            $b = [IO.Path]::GetFullPath($Config.Destino).TrimEnd('\')
            if ($a -ieq $b) { $problemas.Add('Origem e destino apontam para a mesma pasta.'); $destinoOk = $false }
        } catch { }
    }

    [pscustomobject]@{
        OrigemOk      = $origemOk
        DestinoOk     = $destinoOk
        DestinoExiste = $destinoExiste
        Problemas     = $problemas.ToArray()
    }
}

# Somente arquivos .pdf do primeiro nível (sem -Recurse). A extensão é conferida
# exatamente, porque o filtro *.pdf do Windows também aceitaria ".pdfx", ".pdf_old" etc.
function Get-PdfsNivelPrincipal {
    param([string]$Pasta)
    if ([string]::IsNullOrWhiteSpace($Pasta) -or -not (Test-Path -LiteralPath $Pasta -PathType Container)) { return @() }
    $ocultos = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System
    return @(Get-ChildItem -LiteralPath $Pasta -File -Force |
        Where-Object {
            $_.Extension -ieq '.pdf' -and
            -not $_.Name.StartsWith('~$') -and
            (($_.Attributes -band $ocultos) -eq 0)
        })
}

# Compara origem e destino e devolve o que precisa ser feito com cada PDF.
#   Novo       não existe no Drive
#   Alterado   tamanho diferente ou versão da origem mais recente
#   EmDia      igual nos dois lados
#   SoNoDrive  existe no Drive, mas não está mais na origem
function Get-SyncPlano {
    param($Config)
    $origem  = @(Get-PdfsNivelPrincipal $Config.Origem)
    $destino = @(Get-PdfsNivelPrincipal $Config.Destino)

    $noDestino = @{}
    foreach ($d in $destino) { $noDestino[$d.Name] = $d }

    $plano = New-Object 'System.Collections.Generic.List[object]'
    foreach ($o in ($origem | Sort-Object Name)) {
        $d = $noDestino[$o.Name]
        if ($null -eq $d) {
            $status = 'Novo'
        } elseif ($d.Length -ne $o.Length) {
            $status = 'Alterado'
        } elseif (($o.LastWriteTimeUtc - $d.LastWriteTimeUtc).TotalSeconds -gt 2) {
            $status = 'Alterado'
        } else {
            $status = 'EmDia'
        }
        if ($null -ne $d) { $noDestino.Remove($o.Name) }
        $plano.Add([pscustomobject]@{
            Nome           = $o.Name
            Bytes          = [long]$o.Length
            Modificado     = $o.LastWriteTime
            Status         = $status
            CaminhoOrigem  = $o.FullName
            CaminhoDestino = [IO.Path]::Combine($Config.Destino, $o.Name)
        })
    }
    foreach ($d in ($noDestino.Values | Sort-Object Name)) {
        $plano.Add([pscustomobject]@{
            Nome           = $d.Name
            Bytes          = [long]$d.Length
            Modificado     = $d.LastWriteTime
            Status         = 'SoNoDrive'
            CaminhoOrigem  = $null
            CaminhoDestino = $d.FullName
        })
    }
    return ,$plano.ToArray()
}

# Cópia em duas etapas: primeiro do servidor para uma pasta temporária local, depois
# para o Google Drive. Se a rede cair no meio, o PDF do Drive não fica pela metade.
# O arquivo do Drive é sobrescrito no lugar, mantendo o mesmo link de compartilhamento.
function Copy-PdfSeguro {
    param([string]$Origem, [string]$Destino, [string]$PastaTemp, [int]$Tentativas, [int]$Espera, [hashtable]$Ctx)
    $ultimoErro = ''
    for ($i = 1; $i -le $Tentativas; $i++) {
        if ($null -ne $Ctx.Sync -and $Ctx.Sync.Cancelar) {
            return [pscustomobject]@{ Ok = $false; Cancelado = $true; Erro = 'Cancelado' }
        }
        $tmp = [IO.Path]::Combine($PastaTemp, [guid]::NewGuid().ToString('N') + '.tmp')
        try {
            $src = New-Object IO.FileInfo -ArgumentList $Origem
            if (-not $src.Exists) { throw 'O arquivo não existe mais na origem.' }
            $tamanho = $src.Length
            $dataMod = $src.LastWriteTimeUtc

            [IO.File]::Copy($Origem, $tmp, $true)
            $t = New-Object IO.FileInfo -ArgumentList $tmp
            if ($t.Length -ne $tamanho) { throw "Cópia incompleta ($($t.Length) de $tamanho bytes)." }
            $src.Refresh()
            if ($src.Length -ne $tamanho -or $src.LastWriteTimeUtc -ne $dataMod) {
                throw 'O arquivo foi modificado no servidor durante a cópia.'
            }

            if ([IO.File]::Exists($Destino)) {
                $attr = [IO.File]::GetAttributes($Destino)
                if (($attr -band [IO.FileAttributes]::ReadOnly) -ne 0) {
                    [IO.File]::SetAttributes($Destino, ($attr -bxor [IO.FileAttributes]::ReadOnly))
                }
            }
            [IO.File]::Copy($tmp, $Destino, $true)
            $d = New-Object IO.FileInfo -ArgumentList $Destino
            if ($d.Length -ne $tamanho) { throw "Tamanho diferente no destino ($($d.Length) de $tamanho bytes)." }
            try { [IO.File]::SetLastWriteTimeUtc($Destino, $dataMod) } catch { }

            return [pscustomobject]@{ Ok = $true; Cancelado = $false; Erro = $null }
        } catch {
            $ultimoErro = Get-MensagemErro $_
            if ($i -lt $Tentativas) {
                Write-SyncLog $Ctx 'AVISO' ("Tentativa {0} de {1} falhou ({2}): {3}" -f $i, $Tentativas, [IO.Path]::GetFileName($Origem), $ultimoErro)
                Start-Sleep -Seconds $Espera
            }
        } finally {
            if ([IO.File]::Exists($tmp)) { try { [IO.File]::Delete($tmp) } catch { } }
        }
    }
    return [pscustomobject]@{ Ok = $false; Cancelado = $false; Erro = $ultimoErro }
}

function New-ArquivoLog {
    param([string]$Pasta, [datetime]$Quando)
    $nome = 'Sincronizacao_' + $Quando.ToString('yyyy-MM-dd_HH-mm-ss') + '.log'
    $alternativa = [IO.Path]::Combine($env:LOCALAPPDATA, 'SincronizadorMapas', 'Logs')
    foreach ($p in @($Pasta, $alternativa)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        try {
            [void][IO.Directory]::CreateDirectory($p)
            $arquivo = [IO.Path]::Combine($p, $nome)
            [IO.File]::WriteAllText($arquivo, '', (New-Object Text.UTF8Encoding($true)))
            return $arquivo
        } catch { }
    }
    return $null
}

function Remove-LogsAntigos {
    param([string]$Pasta, [int]$Dias)
    if ($Dias -le 0 -or [string]::IsNullOrWhiteSpace($Pasta)) { return 0 }
    if (-not (Test-Path -LiteralPath $Pasta -PathType Container)) { return 0 }
    $limite = (Get-Date).AddDays(-$Dias)
    $antigos = @(Get-ChildItem -LiteralPath $Pasta -File -Filter 'Sincronizacao_*.log' | Where-Object { $_.LastWriteTime -lt $limite })
    foreach ($f in $antigos) { try { [IO.File]::Delete($f.FullName) } catch { } }
    return $antigos.Count
}

function Save-Estado {
    param([string]$Caminho, [hashtable]$Resumo)
    if ([string]::IsNullOrWhiteSpace($Caminho)) { return }
    $agora = Get-Date
    $estado = [pscustomobject]@{
        UltimaSincronizacao = $agora.ToString('yyyy-MM-ddTHH:mm:ss')
        Texto               = $agora.ToString('dd/MM/yyyy') + ' às ' + $agora.ToString('HH') + ':' + $agora.ToString('mm')
        Adicionados         = $Resumo.Adicionados
        Atualizados         = $Resumo.Atualizados
        Arquivados          = $Resumo.Arquivados
        Falhas              = $Resumo.Falhas
        Codigo              = $Resumo.Codigo
        SoNoDrive           = @($Resumo.SoNoDrive)
    }
    [IO.File]::WriteAllText($Caminho, ($estado | ConvertTo-Json), (New-Object Text.UTF8Encoding($true)))
}

# ------------------------------------------------------------------ execução

function Invoke-EtapasSincronizacao {
    param($Config, [hashtable]$Ctx, [hashtable]$Resumo, [string]$PastaTemp, [switch]$Simular)
    $Sync = $Ctx.Sync

    if ($Simular) { Write-SyncLog $Ctx 'INFO' 'SIMULAÇÃO: nenhum arquivo será copiado, movido ou apagado.' }
    Write-SyncLog $Ctx 'INFO' ('Origem:  ' + $Config.Origem)
    Write-SyncLog $Ctx 'INFO' ('Destino: ' + $Config.Destino)
    Write-SyncLog $Ctx 'INFO' 'Regra: somente arquivos .pdf da pasta principal. Subpastas são ignoradas.'

    $amb = Test-SyncAmbiente -Config $Config -CriarDestino:(-not $Simular)
    Send-SyncEvent $Sync @{ Tipo = 'ambiente'; Ambiente = $amb }
    if ($amb.Problemas.Count -gt 0) {
        foreach ($p in $amb.Problemas) { Write-SyncLog $Ctx 'ERRO' $p }
        $Resumo.Codigo = 2
        return
    }

    $plano = Get-SyncPlano -Config $Config
    Send-SyncEvent $Sync @{ Tipo = 'plano'; Itens = $plano }

    $copiar   = @($plano | Where-Object { $_.Status -eq 'Novo' -or $_.Status -eq 'Alterado' })
    $extras   = @($plano | Where-Object { $_.Status -eq 'SoNoDrive' })
    $naOrigem = @($plano | Where-Object { $_.Status -ne 'SoNoDrive' }).Count
    $novos    = @($copiar | Where-Object { $_.Status -eq 'Novo' }).Count
    $Resumo.EmDia = @($plano | Where-Object { $_.Status -eq 'EmDia' }).Count
    $Resumo.SoNoDrive = @($extras | ForEach-Object { $_.Nome })

    Write-SyncLog $Ctx 'INFO' ('PDFs na origem: {0}. Novos: {1}. Alterados: {2}. Em dia: {3}. Só no Drive: {4}.' -f `
        $naOrigem, $novos, ($copiar.Count - $novos), $Resumo.EmDia, $extras.Count)

    # Registra no histórico, uma única vez, os PDFs que saíram do servidor.
    $anteriores = @{}
    $estado = Read-Estado $Ctx.CaminhoEstado
    if ($null -ne $estado -and $null -ne $estado.PSObject.Properties['SoNoDrive']) {
        foreach ($n in @($estado.SoNoDrive)) { if ($n) { $anteriores["$n"] = $true } }
    }

    if ($naOrigem -eq 0) {
        Write-SyncLog $Ctx 'AVISO' 'Nenhum PDF foi encontrado na pasta principal da origem.'
        $Resumo.SoNoDrive = @($anteriores.Keys)
    } elseif (-not $Simular) {
        foreach ($x in $extras) {
            if (-not $anteriores.ContainsKey($x.Nome)) {
                Add-Historico $Ctx 'Saiu do servidor' $x.Nome $x.Bytes 'Não existe mais na pasta principal do servidor.'
            }
        }
    }

    $arquivar = ([bool]$Config.EspelharDrive) -and $extras.Count -gt 0
    if ($arquivar -and $naOrigem -eq 0) {
        Write-SyncLog $Ctx 'AVISO' 'Por segurança nada foi arquivado: a origem está sem nenhum PDF (possível falha de acesso ao servidor).'
        $arquivar = $false
    }
    # Trava contra falhas parciais de acesso: não retira mais da metade dos PDFs de uma vez.
    $limite = [Math]::Max(20, [int][Math]::Ceiling($naOrigem * 0.5))
    if ($arquivar -and $extras.Count -gt $limite) {
        Write-SyncLog $Ctx 'AVISO' ("Por segurança nada foi retirado do Drive: {0} PDFs sairiam de uma vez (limite {1}). Confira se a pasta do servidor está completa." -f $extras.Count, $limite)
        $arquivar = $false
    }
    if (-not $arquivar) {
        foreach ($x in $extras) { Write-SyncLog $Ctx 'INFO' ('Mantido no Drive (não existe mais na origem): ' + $x.Nome) }
    }

    $total = $copiar.Count
    if ($arquivar) { $total += $extras.Count }
    $feito = 0
    if ($total -eq 0) { Write-SyncLog $Ctx 'OK' 'Nada a fazer: todos os PDFs já estão atualizados no Drive.' }

    foreach ($item in $copiar) {
        if ($null -ne $Sync -and $Sync.Cancelar) { $Resumo.Cancelado = $true; break }
        Send-SyncEvent $Sync @{ Tipo = 'item'; Nome = $item.Nome; Status = 'Copiando' }
        Send-SyncEvent $Sync @{ Tipo = 'progresso'; Feito = $feito; Total = $total; Atual = $item.Nome }
        $ehNovo = ($item.Status -eq 'Novo')
        $acao = 'Atualizado'
        if ($ehNovo) { $acao = 'Adicionado' }

        if ($Simular) {
            Write-SyncLog $Ctx 'INFO' ('[simulação] {0}: {1} ({2})' -f $acao, $item.Nome, (Format-Tamanho $item.Bytes))
            Send-SyncEvent $Sync @{ Tipo = 'item'; Nome = $item.Nome; Status = $item.Status }
            $feito++
            continue
        }

        $r = Copy-PdfSeguro -Origem $item.CaminhoOrigem -Destino $item.CaminhoDestino -PastaTemp $PastaTemp `
                            -Tentativas $Config.Tentativas -Espera $Config.EsperaSegundos -Ctx $Ctx
        if ($r.Ok) {
            if ($ehNovo) { $Resumo.Adicionados++ } else { $Resumo.Atualizados++ }
            Write-SyncLog $Ctx 'OK' ('{0}: {1} ({2})' -f $acao, $item.Nome, (Format-Tamanho $item.Bytes))
            $detalhe = 'Versão do servidor de ' + ([datetime]$item.Modificado).ToString('dd/MM/yyyy HH:mm') + '.'
            Add-Historico $Ctx $acao $item.Nome $item.Bytes $detalhe
            Send-SyncEvent $Sync @{ Tipo = 'item'; Nome = $item.Nome; Status = 'Copiado' }
        } elseif ($r.Cancelado) {
            $Resumo.Cancelado = $true
            Send-SyncEvent $Sync @{ Tipo = 'item'; Nome = $item.Nome; Status = $item.Status }
            break
        } else {
            $Resumo.Falhas++
            Write-SyncLog $Ctx 'ERRO' ('Falha ao copiar {0}: {1}' -f $item.Nome, $r.Erro)
            Add-Historico $Ctx 'Falha' $item.Nome $item.Bytes ('Não foi copiado. ' + $r.Erro)
            Send-SyncEvent $Sync @{ Tipo = 'item'; Nome = $item.Nome; Status = 'Falhou'; Detalhe = $r.Erro }
        }
        $feito++
        Send-SyncEvent $Sync @{ Tipo = 'progresso'; Feito = $feito; Total = $total; Atual = '' }
    }

    if ($arquivar -and -not $Resumo.Cancelado) {
        $pastaArq = [IO.Path]::Combine($Config.PastaArquivados, (Get-Date -Format 'yyyy-MM-dd'))
        $restantes = New-Object 'System.Collections.Generic.List[string]'
        foreach ($x in $extras) {
            if ($null -ne $Sync -and $Sync.Cancelar) { $Resumo.Cancelado = $true; $restantes.Add($x.Nome); continue }
            Send-SyncEvent $Sync @{ Tipo = 'progresso'; Feito = $feito; Total = $total; Atual = $x.Nome }
            if ($Simular) {
                Write-SyncLog $Ctx 'INFO' ('[simulação] Arquivaria {0} em {1}' -f $x.Nome, $pastaArq)
                $feito++
                continue
            }
            try {
                [void][IO.Directory]::CreateDirectory($pastaArq)
                $alvo = [IO.Path]::Combine($pastaArq, $x.Nome)
                if ([IO.File]::Exists($alvo)) {
                    $alvo = [IO.Path]::Combine($pastaArq, [IO.Path]::GetFileNameWithoutExtension($x.Nome) + '_' + (Get-Date -Format 'HHmmss') + [IO.Path]::GetExtension($x.Nome))
                }
                [IO.File]::Move($x.CaminhoDestino, $alvo)
                $Resumo.Arquivados++
                Write-SyncLog $Ctx 'OK' ('Arquivado: {0}. Movido para {1}' -f $x.Nome, $alvo)
                Add-Historico $Ctx 'Arquivado' $x.Nome $x.Bytes ('Retirado da pasta do Drive e guardado em ' + $alvo)
                Send-SyncEvent $Sync @{ Tipo = 'item'; Nome = $x.Nome; Status = 'Arquivado' }
            } catch {
                $msg = Get-MensagemErro $_
                $Resumo.Falhas++
                $restantes.Add($x.Nome)
                Write-SyncLog $Ctx 'ERRO' ('Falha ao arquivar {0}: {1}' -f $x.Nome, $msg)
                Add-Historico $Ctx 'Falha' $x.Nome $x.Bytes ('Não foi arquivado. ' + $msg)
                Send-SyncEvent $Sync @{ Tipo = 'item'; Nome = $x.Nome; Status = 'Falhou'; Detalhe = $msg }
            }
            $feito++
        }
        if (-not $Simular) { $Resumo.SoNoDrive = $restantes.ToArray() }
    }

    Send-SyncEvent $Sync @{ Tipo = 'progresso'; Feito = $feito; Total = $total; Atual = '' }
    if ($Resumo.Cancelado) { Write-SyncLog $Ctx 'AVISO' 'Sincronização interrompida pelo usuário.' }
    Write-SyncLog $Ctx 'INFO' ('Resumo: {0}. {1} já em dia.' -f (Get-TextoResumo $Resumo), $Resumo.EmDia)

    if ($Resumo.Falhas -gt 0 -or $Resumo.Cancelado) { $Resumo.Codigo = 1 } else { $Resumo.Codigo = 0 }
}

function Invoke-Sincronizacao {
    param($Config, $Sync, [switch]$Console, [switch]$Simular, [string]$CaminhoEstado, [string]$OrigemExecucao = 'Manual')

    $inicio = Get-Date
    $ctx = @{
        LogFile       = $null
        Historico     = $null
        Console       = [bool]$Console
        Sync          = $Sync
        Execucao      = $inicio.ToString('yyyy-MM-dd HH:mm:ss')
        CaminhoEstado = $CaminhoEstado
        PrimeiroErro  = $null
    }
    $resumo = @{
        Adicionados = 0; Atualizados = 0; Arquivados = 0; Falhas = 0; EmDia = 0
        SoNoDrive = @(); Cancelado = $false; Codigo = 0; LogFile = $null; Duracao = 0
    }
    $mutex = $null; $temMutex = $false; $ocupado = $false; $pastaTemp = $null

    try {
        # Impede duas sincronizações ao mesmo tempo (janela e tarefa agendada, por exemplo).
        try {
            $mutex = New-Object System.Threading.Mutex -ArgumentList $false, 'Global\SincronizadorMapasPDF'
            try { $temMutex = $mutex.WaitOne(0) }
            catch { if (Test-MutexAbandonado $_) { $temMutex = $true } else { throw } }
            if (-not $temMutex) { $ocupado = $true }
        } catch {
            if ((Get-ExcecaoBase $_) -is [System.UnauthorizedAccessException]) { $ocupado = $true }
            else { Write-SyncLog $ctx 'AVISO' ('Trava de execução indisponível: ' + (Get-MensagemErro $_)) }
        }

        if ($ocupado) {
            Write-SyncLog $ctx 'ERRO' 'Já existe uma sincronização em andamento (em outra janela ou pela tarefa agendada).'
            $resumo.Codigo = 2
        } else {
            if (-not $Simular) {
                $ctx.LogFile = New-ArquivoLog $Config.PastaLogs $inicio
                if ($ctx.LogFile) { $ctx.Historico = Initialize-Historico ([IO.Path]::GetDirectoryName($ctx.LogFile)) }
            }
            $resumo.LogFile = $ctx.LogFile
            $pastaTemp = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'SincronizadorMapas_' + [guid]::NewGuid().ToString('N'))
            [void][IO.Directory]::CreateDirectory($pastaTemp)

            Invoke-EtapasSincronizacao -Config $Config -Ctx $ctx -Resumo $resumo -PastaTemp $pastaTemp -Simular:$Simular

            if (-not $Simular -and $resumo.Codigo -ne 2) {
                try { Save-Estado -Caminho $CaminhoEstado -Resumo $resumo }
                catch { Write-SyncLog $ctx 'AVISO' ('Não foi possível gravar estado.json: ' + (Get-MensagemErro $_)) }
                try {
                    $apagados = Remove-LogsAntigos -Pasta ([IO.Path]::GetDirectoryName($ctx.LogFile)) -Dias $Config.DiasManterLogs
                    if ($apagados -gt 0) { Write-SyncLog $ctx 'INFO' ("$apagados log(s) com mais de $($Config.DiasManterLogs) dias foram removidos.") }
                } catch { }
            }
        }
    } catch {
        $resumo.Codigo = 2
        Write-SyncLog $ctx 'ERRO' ('Erro inesperado: ' + (Get-MensagemErro $_))
    } finally {
        if ($temMutex) { try { $mutex.ReleaseMutex() } catch { } }
        if ($null -ne $mutex) { $mutex.Dispose() }
        if ($pastaTemp -and [IO.Directory]::Exists($pastaTemp)) { try { [IO.Directory]::Delete($pastaTemp, $true) } catch { } }
        $resumo.Duracao = [int][Math]::Round(((Get-Date) - $inicio).TotalSeconds)

        if ($resumo.Codigo -eq 2) { $detalhe = $OrigemExecucao + '. Não executada: ' + $ctx.PrimeiroErro }
        elseif ($resumo.Cancelado) { $detalhe = $OrigemExecucao + '. Interrompida: ' + (Get-TextoResumo $resumo) + '.' }
        else { $detalhe = $OrigemExecucao + ': ' + (Get-TextoResumo $resumo) + '.' }
        Add-Historico $ctx 'Sincronização' '' 0 $detalhe

        Write-SyncLog $ctx 'INFO' ("Fim. Duração: $($resumo.Duracao) s. Código: $($resumo.Codigo).")
        Send-SyncEvent $Sync @{ Tipo = 'fim'; Resumo = $resumo }
    }
    return $resumo
}
'@

. ([scriptblock]::Create($CodigoMotor))

# =====================================================================================
#  MODO AUTOMÁTICO (sem janela), usado pelo Agendador de Tarefas
# =====================================================================================
if ($Auto) {
    try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
    try {
        $cfg = Read-SyncConfig $CaminhoConfig
        if (-not (Test-Path -LiteralPath $CaminhoConfig)) { Save-SyncConfig $cfg $CaminhoConfig }
    } catch {
        Write-Host ('ERRO  ' + (Get-MensagemErro $_)) -ForegroundColor Red
        exit 2
    }
    $syncAuto = [hashtable]::Synchronized(@{ Fila = $null; Cancelar = $false })
    $res = @(Invoke-Sincronizacao -Config $cfg -Sync $syncAuto -Console -Simular:$Simular -CaminhoEstado $CaminhoEstado -OrigemExecucao 'Automática')[-1]
    exit [int]$res.Codigo
}

# =====================================================================================
#  INTERFACE
# =====================================================================================
$CodigoTarefa = @'
try {
    if ($Modo -eq 'verificar') {
        $amb = Test-SyncAmbiente -Config $Config
        Send-SyncEvent $Sync @{ Tipo = 'ambiente'; Ambiente = $amb }
        if ($amb.OrigemOk) {
            $plano = Get-SyncPlano -Config $Config
            Send-SyncEvent $Sync @{ Tipo = 'plano'; Itens = $plano }
        }
        Send-SyncEvent $Sync @{ Tipo = 'fimVerificacao'; Ambiente = $amb }
    } else {
        [void](Invoke-Sincronizacao -Config $Config -Sync $Sync -CaminhoEstado $CaminhoEstado -OrigemExecucao 'Manual')
    }
} catch {
    Send-SyncEvent $Sync @{ Tipo = 'falha'; Mensagem = (Get-MensagemErro $_) }
}
'@

# Cores dos temas. Toda a interface usa estas chaves, então trocar o tema é só trocar a paleta.
$Paletas = @{
    Claro = @{
        Fundo = '#F3F2EE'; Superficie = '#FFFFFF'; Borda = '#E2DFD8'; BordaSuave = '#F0EEE8'; Divisoria = '#EEEBE5'
        Texto = '#1E2320'; TextoMudo = '#6E716C'; TextoFraco = '#8A8D87'
        Acento = '#2F6A43'; AcentoHover = '#285C3A'; AcentoPressionado = '#1F4A2E'; AcentoBorda = '#173823'; TextoAcento = '#2F6A43'
        DiscretoHover = '#EDF1EC'; DiscretoPressionado = '#E0E8E0'
        SecundarioFundo = '#FFFFFF'; SecundarioBorda = '#D5D2CA'; SecundarioHover = '#F6F5F1'; SecundarioPressionado = '#ECEAE4'
        Campo = '#FAFAF8'; CampoFoco = '#FFFFFF'; CampoBorda = '#DEDBD3'; CampoBordaHover = '#C9C5BC'
        LinhaHover = '#F8F7F4'; LinhaSelecionada = '#EEF3EE'; Trilho = '#E6E3DC'; Barra = '#D3D0C8'; BarraHover = '#ABA89F'
        CaixaFundo = '#FFFFFF'; CaixaBorda = '#B9B6AE'; SegmentoFundo = '#ECEAE4'; SegmentoAtivo = '#FFFFFF'
        DicaFundo = '#FFFFFF'; DicaBorda = '#D5D2CA'
        PontoOk = '#3C8A55'; PontoErro = '#B5473D'; PontoAviso = '#C08A1E'; PontoNeutro = '#C9C6BE'
        StatusErro = '#9A3A33'; StatusSucesso = '#2F6A43'
        NumNovo = '#2F6A43'; NumAlterado = '#85590A'; NumSoNoDrive = '#7A5C4B'; NumFalha = '#9A3A33'
        SeloNovoF = '#E3EEE5'; SeloNovoT = '#2F6A43'; SeloAlteradoF = '#F7EDD6'; SeloAlteradoT = '#85590A'
        SeloEmDiaT = '#8A8D87'; SeloSoNoDriveF = '#EFE8E3'; SeloSoNoDriveT = '#7A5C4B'
        SeloCopiandoF = '#E4EBF1'; SeloCopiandoT = '#3B5873'; SeloCopiadoF = '#2F6A43'; SeloCopiadoT = '#FFFFFF'
        SeloFalhouF = '#F5E1DE'; SeloFalhouT = '#9A3A33'; SeloArquivadoF = '#ECEAE4'; SeloArquivadoT = '#5E615C'
        SeloNeutroF = '#EFEDE8'; SeloNeutroT = '#6E716C'
    }
    Escuro = @{
        Fundo = '#161817'; Superficie = '#1F2221'; Borda = '#2F3331'; BordaSuave = '#272A29'; Divisoria = '#2B2F2D'
        Texto = '#E7E8E5'; TextoMudo = '#A3A7A2'; TextoFraco = '#7F837E'
        Acento = '#3D8655'; AcentoHover = '#47945F'; AcentoPressionado = '#357549'; AcentoBorda = '#8CC9A0'; TextoAcento = '#7CC492'
        DiscretoHover = '#26302A'; DiscretoPressionado = '#2C3A31'
        SecundarioFundo = '#262928'; SecundarioBorda = '#3A3E3C'; SecundarioHover = '#2D3130'; SecundarioPressionado = '#333736'
        Campo = '#1A1C1B'; CampoFoco = '#151716'; CampoBorda = '#353937'; CampoBordaHover = '#474C49'
        LinhaHover = '#252827'; LinhaSelecionada = '#233229'; Trilho = '#2F3331'; Barra = '#3E4240'; BarraHover = '#555A57'
        CaixaFundo = '#1A1C1B'; CaixaBorda = '#5A5F5C'; SegmentoFundo = '#1A1C1B'; SegmentoAtivo = '#2D3130'
        DicaFundo = '#2B2E2D'; DicaBorda = '#3F4341'
        PontoOk = '#4FAE6E'; PontoErro = '#E0675B'; PontoAviso = '#D9A23A'; PontoNeutro = '#5A5F5C'
        StatusErro = '#EE9A90'; StatusSucesso = '#7CC492'
        NumNovo = '#7CC492'; NumAlterado = '#E0B062'; NumSoNoDrive = '#CFA58C'; NumFalha = '#EE9A90'
        SeloNovoF = '#22362A'; SeloNovoT = '#7CC492'; SeloAlteradoF = '#3A3020'; SeloAlteradoT = '#E0B062'
        SeloEmDiaT = '#7F837E'; SeloSoNoDriveF = '#382D27'; SeloSoNoDriveT = '#CFA58C'
        SeloCopiandoF = '#22303D'; SeloCopiandoT = '#8FB4D6'; SeloCopiadoF = '#3D8655'; SeloCopiadoT = '#FFFFFF'
        SeloFalhouF = '#40262A'; SeloFalhouT = '#EE9A90'; SeloArquivadoF = '#2C2F2E'; SeloArquivadoT = '#B0B3AE'
        SeloNeutroF = '#2A2D2C'; SeloNeutroT = '#A3A7A2'
    }
}

# Estilos usados pelas duas janelas (principal e relatório).
$XamlEstilos = @'
    <Style x:Key="Cartao" TargetType="Border">
      <Setter Property="Background" Value="{DynamicResource Superficie}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Borda}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="18,16"/>
    </Style>

    <Style x:Key="BtnPrimario" TargetType="Button">
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="Background" Value="{DynamicResource Acento}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="18,0"/>
      <Setter Property="Height" Value="36"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{DynamicResource Acento}" BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="{DynamicResource AcentoHover}"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="b" Property="Background" Value="{DynamicResource AcentoPressionado}"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource AcentoBorda}"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.4"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnSecundario" TargetType="Button">
      <Setter Property="Foreground" Value="{DynamicResource Texto}"/>
      <Setter Property="Background" Value="{DynamicResource SecundarioFundo}"/>
      <Setter Property="Padding" Value="16,0"/>
      <Setter Property="Height" Value="36"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{DynamicResource SecundarioBorda}" BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Background" Value="{DynamicResource SecundarioHover}"/>
                <Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource CampoBordaHover}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="b" Property="Background" Value="{DynamicResource SecundarioPressionado}"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource Acento}"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnDiscreto" TargetType="Button">
      <Setter Property="Foreground" Value="{DynamicResource TextoAcento}"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Padding" Value="10,0"/>
      <Setter Property="Height" Value="32"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="{DynamicResource DiscretoHover}"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="b" Property="Background" Value="{DynamicResource DiscretoPressionado}"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="Background" Value="{DynamicResource DiscretoHover}"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.4"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="EstiloCampo" TargetType="TextBox">
      <Setter Property="Foreground" Value="{DynamicResource Texto}"/>
      <Setter Property="Background" Value="{DynamicResource Campo}"/>
      <Setter Property="CaretBrush" Value="{DynamicResource Texto}"/>
      <Setter Property="SelectionBrush" Value="{DynamicResource Acento}"/>
      <Setter Property="Padding" Value="9,0"/>
      <Setter Property="Height" Value="34"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{DynamicResource CampoBorda}" BorderThickness="1" CornerRadius="6">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="Center"
                            HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource CampoBordaHover}"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource Acento}"/>
                <Setter TargetName="b" Property="Background" Value="{DynamicResource CampoFoco}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.6"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="CampoCaminho" TargetType="TextBox" BasedOn="{StaticResource EstiloCampo}">
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="12.5"/>
    </Style>

    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="{DynamicResource TextoMudo}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal" Background="Transparent">
              <Border x:Name="caixa" Width="16" Height="16" CornerRadius="4" BorderThickness="1.5" BorderBrush="{DynamicResource CaixaBorda}" Background="{DynamicResource CaixaFundo}" VerticalAlignment="Center">
                <Path x:Name="marca" Data="M2.6,6.6 L5.3,9.2 L10.4,3.6" Stroke="#FFFFFF" StrokeThickness="2"
                      StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Visibility="Collapsed"/>
              </Border>
              <ContentPresenter Margin="8,0,0,0" VerticalAlignment="Center"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="caixa" Property="BorderBrush" Value="{DynamicResource Acento}"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="caixa" Property="BorderBrush" Value="{DynamicResource Acento}"/></Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="caixa" Property="Background" Value="{DynamicResource Acento}"/>
                <Setter TargetName="caixa" Property="BorderBrush" Value="{DynamicResource Acento}"/>
                <Setter TargetName="marca" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.5"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="GrupoSegmentos" TargetType="Border">
      <Setter Property="Background" Value="{DynamicResource SegmentoFundo}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Borda}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="7"/>
      <Setter Property="Padding" Value="2"/>
      <Setter Property="HorizontalAlignment" Value="Left"/>
    </Style>

    <Style x:Key="Segmento" TargetType="RadioButton">
      <Setter Property="Foreground" Value="{DynamicResource TextoMudo}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Padding" Value="11,0"/>
      <Setter Property="Height" Value="26"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="b" Background="Transparent" CornerRadius="5" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="{DynamicResource Texto}"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter Property="Foreground" Value="{DynamicResource Texto}"/></Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="b" Property="Background" Value="{DynamicResource SegmentoAtivo}"/>
                <Setter Property="Foreground" Value="{DynamicResource Texto}"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ProgressBar">
      <Setter Property="Height" Value="4"/>
      <Setter Property="Foreground" Value="{DynamicResource Acento}"/>
      <Setter Property="Background" Value="{DynamicResource Trilho}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ProgressBar">
            <Grid>
              <Border x:Name="PART_Track" Background="{TemplateBinding Background}" CornerRadius="2"/>
              <Border x:Name="PART_Indicator" Background="{TemplateBinding Foreground}" CornerRadius="2" HorizontalAlignment="Left"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ToolTip">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ToolTip">
            <Border Background="{DynamicResource DicaFundo}" BorderBrush="{DynamicResource DicaBorda}" BorderThickness="1" CornerRadius="5" Padding="9,6" MaxWidth="380">
              <ContentPresenter TextElement.Foreground="{DynamicResource Texto}" TextElement.FontSize="12">
                <ContentPresenter.Resources>
                  <Style TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/></Style>
                </ContentPresenter.Resources>
              </ContentPresenter>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <ControlTemplate x:Key="BarraVertical" TargetType="ScrollBar">
      <Grid Background="Transparent" Width="10">
        <Track x:Name="PART_Track" IsDirectionReversed="True">
          <Track.Thumb>
            <Thumb>
              <Thumb.Template>
                <ControlTemplate TargetType="Thumb">
                  <Border x:Name="t" Background="{DynamicResource Barra}" CornerRadius="3" Margin="2,2"/>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="t" Property="Background" Value="{DynamicResource BarraHover}"/></Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </Thumb.Template>
            </Thumb>
          </Track.Thumb>
        </Track>
      </Grid>
    </ControlTemplate>

    <Style x:Key="CabecalhoColuna" TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{DynamicResource TextoFraco}"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Height" Value="34"/>
      <Setter Property="Padding" Value="14,0"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridColumnHeader">
            <Border Background="{TemplateBinding Background}" BorderBrush="{DynamicResource Divisoria}" BorderThickness="0,1,0,1" Padding="{TemplateBinding Padding}">
              <StackPanel Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}">
                <ContentPresenter VerticalAlignment="Center"/>
                <Path x:Name="seta" Width="8" Height="5" Margin="6,0,0,0" VerticalAlignment="Center" Stretch="Fill"
                      Fill="{DynamicResource TextoAcento}" Data="M0,5 L4,0 L8,5 Z" Visibility="Collapsed"/>
              </StackPanel>
            </Border>
            <ControlTemplate.Triggers>
              <MultiTrigger>
                <MultiTrigger.Conditions>
                  <Condition Property="CanUserSort" Value="True"/>
                  <Condition Property="IsMouseOver" Value="True"/>
                </MultiTrigger.Conditions>
                <Setter Property="Foreground" Value="{DynamicResource Texto}"/>
                <Setter Property="Cursor" Value="Hand"/>
              </MultiTrigger>
              <Trigger Property="SortDirection" Value="Ascending">
                <Setter TargetName="seta" Property="Visibility" Value="Visible"/>
                <Setter Property="Foreground" Value="{DynamicResource Texto}"/>
              </Trigger>
              <Trigger Property="SortDirection" Value="Descending">
                <Setter TargetName="seta" Property="Visibility" Value="Visible"/>
                <Setter TargetName="seta" Property="Data" Value="M0,0 L4,5 L8,0 Z"/>
                <Setter Property="Foreground" Value="{DynamicResource Texto}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="CabecalhoDireita" TargetType="DataGridColumnHeader" BasedOn="{StaticResource CabecalhoColuna}">
      <Setter Property="HorizontalContentAlignment" Value="Right"/>
    </Style>

    <Style x:Key="Celula" TargetType="DataGridCell">
      <Setter Property="Padding" Value="14,0"/>
      <Setter Property="Foreground" Value="{DynamicResource Texto}"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridCell">
            <Border Background="Transparent" Padding="{TemplateBinding Padding}">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Linha" TargetType="DataGridRow">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BordaSuave}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{DynamicResource LinhaHover}"/></Trigger>
        <Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="{DynamicResource LinhaSelecionada}"/></Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="Grade" TargetType="DataGrid">
      <Setter Property="AutoGenerateColumns" Value="False"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="CanUserAddRows" Value="False"/>
      <Setter Property="CanUserDeleteRows" Value="False"/>
      <Setter Property="CanUserResizeRows" Value="False"/>
      <Setter Property="CanUserReorderColumns" Value="False"/>
      <Setter Property="CanUserSortColumns" Value="False"/>
      <Setter Property="CanUserResizeColumns" Value="False"/>
      <Setter Property="SelectionMode" Value="Single"/>
      <Setter Property="SelectionUnit" Value="FullRow"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="GridLinesVisibility" Value="None"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="RowBackground" Value="Transparent"/>
      <Setter Property="AlternatingRowBackground" Value="Transparent"/>
      <Setter Property="Foreground" Value="{DynamicResource Texto}"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="RowHeight" Value="38"/>
      <Setter Property="HorizontalScrollBarVisibility" Value="Disabled"/>
      <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
      <Setter Property="EnableRowVirtualization" Value="True"/>
      <Setter Property="VirtualizingPanel.IsVirtualizing" Value="True"/>
      <Setter Property="VirtualizingPanel.VirtualizationMode" Value="Recycling"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="ColumnHeaderStyle" Value="{StaticResource CabecalhoColuna}"/>
      <Setter Property="CellStyle" Value="{StaticResource Celula}"/>
      <Setter Property="RowStyle" Value="{StaticResource Linha}"/>
      <Style.Resources>
        <Style TargetType="ScrollBar">
          <Setter Property="Width" Value="10"/>
          <Setter Property="MinWidth" Value="10"/>
          <Setter Property="Template" Value="{StaticResource BarraVertical}"/>
        </Style>
      </Style.Resources>
    </Style>

    <Style x:Key="TextoCelula" TargetType="TextBlock">
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="TextoCelulaMudo" TargetType="TextBlock" BasedOn="{StaticResource TextoCelula}">
      <Setter Property="Foreground" Value="{DynamicResource TextoMudo}"/>
    </Style>
    <Style x:Key="TextoCelulaDireita" TargetType="TextBlock" BasedOn="{StaticResource TextoCelulaMudo}">
      <Setter Property="HorizontalAlignment" Value="Right"/>
    </Style>

    <Style x:Key="CelulaNumero" TargetType="Border">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Divisoria}"/>
      <Setter Property="Padding" Value="18,12"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{DynamicResource LinhaHover}"/></Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="Numero" TargetType="TextBlock">
      <Setter Property="FontSize" Value="24"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="RotuloNumero" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{DynamicResource TextoMudo}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>

    <DataTemplate x:Key="Selo">
      <Border x:Name="selo" CornerRadius="10" Padding="9,2,9,3" HorizontalAlignment="Left" VerticalAlignment="Center" Background="{DynamicResource SeloNeutroF}">
        <TextBlock x:Name="seloTexto" Text="{Binding StatusTexto}" FontSize="12" FontWeight="SemiBold" Foreground="{DynamicResource SeloNeutroT}" TextTrimming="CharacterEllipsis"/>
      </Border>
      <DataTemplate.Triggers>
        <DataTrigger Binding="{Binding Status}" Value="Novo">
          <Setter TargetName="selo" Property="Background" Value="{DynamicResource SeloNovoF}"/>
          <Setter TargetName="seloTexto" Property="Foreground" Value="{DynamicResource SeloNovoT}"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Alterado">
          <Setter TargetName="selo" Property="Background" Value="{DynamicResource SeloAlteradoF}"/>
          <Setter TargetName="seloTexto" Property="Foreground" Value="{DynamicResource SeloAlteradoT}"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="EmDia">
          <Setter TargetName="selo" Property="Background" Value="Transparent"/>
          <Setter TargetName="seloTexto" Property="Foreground" Value="{DynamicResource SeloEmDiaT}"/>
          <Setter TargetName="seloTexto" Property="FontWeight" Value="Normal"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="SoNoDrive">
          <Setter TargetName="selo" Property="Background" Value="{DynamicResource SeloSoNoDriveF}"/>
          <Setter TargetName="seloTexto" Property="Foreground" Value="{DynamicResource SeloSoNoDriveT}"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Copiando">
          <Setter TargetName="selo" Property="Background" Value="{DynamicResource SeloCopiandoF}"/>
          <Setter TargetName="seloTexto" Property="Foreground" Value="{DynamicResource SeloCopiandoT}"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Copiado">
          <Setter TargetName="selo" Property="Background" Value="{DynamicResource SeloCopiadoF}"/>
          <Setter TargetName="seloTexto" Property="Foreground" Value="{DynamicResource SeloCopiadoT}"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Falhou">
          <Setter TargetName="selo" Property="Background" Value="{DynamicResource SeloFalhouF}"/>
          <Setter TargetName="seloTexto" Property="Foreground" Value="{DynamicResource SeloFalhouT}"/>
          <Setter TargetName="selo" Property="ToolTip" Value="{Binding Detalhe}"/>
          <Setter TargetName="selo" Property="Cursor" Value="Help"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Arquivado">
          <Setter TargetName="selo" Property="Background" Value="{DynamicResource SeloArquivadoF}"/>
          <Setter TargetName="seloTexto" Property="Foreground" Value="{DynamicResource SeloArquivadoT}"/>
        </DataTrigger>
      </DataTemplate.Triggers>
    </DataTemplate>
'@

$XamlPrincipal = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Sincronizador de Mapas PDF"
        Width="1020" Height="700" MinWidth="620" MinHeight="560"
        WindowStartupLocation="CenterScreen"
        Background="{DynamicResource Fundo}" Foreground="{DynamicResource Texto}"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display">

  <Window.Resources>
<!--ESTILOS-->
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Grid Margin="28,22,28,18">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <!-- Cabeçalho -->
      <Grid Margin="0,0,0,18">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel VerticalAlignment="Bottom">
          <TextBlock x:Name="TxtSobretitulo" Text="Mapas PDF" Foreground="{DynamicResource TextoMudo}" FontSize="12" TextTrimming="CharacterEllipsis"/>
          <TextBlock Text="Sincronização com o Google Drive" FontSize="21" FontWeight="SemiBold" Margin="0,2,0,0" TextTrimming="CharacterEllipsis"/>
        </StackPanel>
        <StackPanel Grid.Column="1" VerticalAlignment="Bottom" Margin="16,0,0,0">
          <Border Style="{StaticResource GrupoSegmentos}" HorizontalAlignment="Right" ToolTip="Aparência da janela">
            <StackPanel Orientation="Horizontal">
              <RadioButton x:Name="RbClaro" Tag="Claro" Content="Claro" GroupName="Tema" Style="{StaticResource Segmento}"/>
              <RadioButton x:Name="RbEscuro" Tag="Escuro" Content="Escuro" GroupName="Tema" Style="{StaticResource Segmento}"/>
              <RadioButton x:Name="RbSistema" Tag="Sistema" Content="Igual ao Windows" GroupName="Tema" Style="{StaticResource Segmento}"/>
            </StackPanel>
          </Border>
          <TextBlock x:Name="TxtUltima" Margin="0,8,2,3" Foreground="{DynamicResource TextoMudo}" FontSize="12" TextAlignment="Right"/>
        </StackPanel>
      </Grid>

      <!-- Origem e destino -->
      <UniformGrid x:Name="GradeCaminhos" Grid.Row="1" Columns="2" Margin="0,0,0,14">
        <Border x:Name="CartaoOrigem" Style="{StaticResource Cartao}" Margin="0,0,7,0">
          <StackPanel>
            <DockPanel Margin="0,0,0,10" LastChildFill="False">
              <TextBlock Text="Origem" FontWeight="SemiBold" DockPanel.Dock="Left"/>
              <TextBlock Text="servidor da rede" Foreground="{DynamicResource TextoFraco}" Margin="8,0,0,0" DockPanel.Dock="Left"/>
              <StackPanel Orientation="Horizontal" DockPanel.Dock="Right" VerticalAlignment="Center">
                <Ellipse x:Name="PontoOrigem" Width="8" Height="8" Fill="{DynamicResource PontoNeutro}" Margin="0,0,6,0" VerticalAlignment="Center"/>
                <TextBlock x:Name="StatusOrigem" Text="Verificando" FontSize="12" Foreground="{DynamicResource TextoMudo}"/>
              </StackPanel>
            </DockPanel>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox x:Name="CampoOrigem" Style="{StaticResource CampoCaminho}"/>
              <Button x:Name="BtnAlterarOrigem" Grid.Column="1" Content="Alterar" Style="{StaticResource BtnDiscreto}" Margin="6,0,0,0"/>
              <Button x:Name="BtnAbrirOrigem" Grid.Column="2" Content="Abrir" Style="{StaticResource BtnDiscreto}"/>
            </Grid>
          </StackPanel>
        </Border>

        <Border x:Name="CartaoDestino" Style="{StaticResource Cartao}" Margin="7,0,0,0">
          <StackPanel>
            <DockPanel Margin="0,0,0,10" LastChildFill="False">
              <TextBlock Text="Destino" FontWeight="SemiBold" DockPanel.Dock="Left"/>
              <TextBlock Text="Google Drive" Foreground="{DynamicResource TextoFraco}" Margin="8,0,0,0" DockPanel.Dock="Left"/>
              <StackPanel Orientation="Horizontal" DockPanel.Dock="Right" VerticalAlignment="Center">
                <Ellipse x:Name="PontoDestino" Width="8" Height="8" Fill="{DynamicResource PontoNeutro}" Margin="0,0,6,0" VerticalAlignment="Center"/>
                <TextBlock x:Name="StatusDestino" Text="Verificando" FontSize="12" Foreground="{DynamicResource TextoMudo}"/>
              </StackPanel>
            </DockPanel>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox x:Name="CampoDestino" Style="{StaticResource CampoCaminho}"/>
              <Button x:Name="BtnAlterarDestino" Grid.Column="1" Content="Alterar" Style="{StaticResource BtnDiscreto}" Margin="6,0,0,0"/>
              <Button x:Name="BtnAbrirDestino" Grid.Column="2" Content="Abrir" Style="{StaticResource BtnDiscreto}"/>
            </Grid>
          </StackPanel>
        </Border>
      </UniformGrid>

      <!-- Números -->
      <Border Grid.Row="2" Style="{StaticResource Cartao}" Padding="0" Margin="0,0,0,14">
        <UniformGrid Columns="4">
          <Border x:Name="FiltroNovos" Tag="Novo" Style="{StaticResource CelulaNumero}" BorderThickness="0,0,1,0" CornerRadius="7,0,0,7" ToolTip="Clique para mostrar só os novos">
            <StackPanel>
              <TextBlock x:Name="NumNovos" Text="0" Style="{StaticResource Numero}" Foreground="{DynamicResource NumNovo}"/>
              <TextBlock Text="novos" Style="{StaticResource RotuloNumero}"/>
            </StackPanel>
          </Border>
          <Border x:Name="FiltroAlterados" Tag="Alterado" Style="{StaticResource CelulaNumero}" BorderThickness="0,0,1,0" ToolTip="Clique para mostrar só os alterados">
            <StackPanel>
              <TextBlock x:Name="NumAlterados" Text="0" Style="{StaticResource Numero}" Foreground="{DynamicResource NumAlterado}"/>
              <TextBlock Text="alterados" Style="{StaticResource RotuloNumero}"/>
            </StackPanel>
          </Border>
          <Border x:Name="FiltroEmDia" Tag="EmDia" Style="{StaticResource CelulaNumero}" BorderThickness="0,0,1,0" ToolTip="Clique para mostrar só os que estão em dia">
            <StackPanel>
              <TextBlock x:Name="NumEmDia" Text="0" Style="{StaticResource Numero}"/>
              <TextBlock Text="em dia" Style="{StaticResource RotuloNumero}"/>
            </StackPanel>
          </Border>
          <Border x:Name="FiltroSoNoDrive" Tag="SoNoDrive" Style="{StaticResource CelulaNumero}" BorderThickness="0" CornerRadius="0,7,7,0"
                  ToolTip="PDFs que estão no Drive mas não existem mais no servidor. Clique para mostrar só estes.">
            <StackPanel>
              <TextBlock x:Name="NumSoNoDrive" Text="0" Style="{StaticResource Numero}" Foreground="{DynamicResource NumSoNoDrive}"/>
              <TextBlock Text="só no Drive" Style="{StaticResource RotuloNumero}"/>
            </StackPanel>
          </Border>
        </UniformGrid>
      </Border>

      <!-- Lista -->
      <Border Grid.Row="3" Style="{StaticResource Cartao}" Padding="0">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <DockPanel Margin="18,14,18,12">
            <CheckBox x:Name="ChkArquivar" DockPanel.Dock="Right" VerticalAlignment="Center" Margin="16,0,0,0"
                      Content="Manter o Drive igual ao servidor"
                      ToolTip="PDFs que não existem mais no servidor são retirados da pasta do Drive e guardados em _ARQUIVADOS. O servidor nunca é alterado."/>
            <Grid x:Name="GradeBuscaLista" DockPanel.Dock="Right" Width="220" Height="32" Margin="16,0,0,0" VerticalAlignment="Center">
              <TextBox x:Name="CampoBuscaLista" Style="{StaticResource EstiloCampo}" Height="32"/>
              <TextBlock x:Name="DicaBuscaLista" Text="Buscar arquivo" Foreground="{DynamicResource TextoFraco}" Margin="11,0,0,0"
                         VerticalAlignment="Center" IsHitTestVisible="False"/>
            </Grid>
            <StackPanel VerticalAlignment="Center">
              <TextBlock Text="Arquivos da pasta principal" FontWeight="SemiBold"/>
              <TextBlock x:Name="TxtSubLista" Text="Somente arquivos .pdf. Subpastas são ignoradas. Clique duas vezes para abrir." Foreground="{DynamicResource TextoFraco}" FontSize="12" Margin="0,2,0,0" TextTrimming="CharacterEllipsis"/>
            </StackPanel>
          </DockPanel>

          <DataGrid x:Name="Lista" Grid.Row="1" Style="{StaticResource Grade}" CanUserSortColumns="True">
            <DataGrid.Columns>
              <DataGridTextColumn Header="ARQUIVO" Binding="{Binding Nome}" SortMemberPath="Nome" Width="*" ElementStyle="{StaticResource TextoCelula}"/>
              <DataGridTextColumn Header="TAMANHO" Binding="{Binding TamanhoTexto}" SortMemberPath="Bytes" Width="110" ElementStyle="{StaticResource TextoCelulaDireita}" HeaderStyle="{StaticResource CabecalhoDireita}"/>
              <DataGridTextColumn Header="MODIFICADO" Binding="{Binding ModificadoTexto}" SortMemberPath="Modificado" Width="160" ElementStyle="{StaticResource TextoCelulaMudo}"/>
              <DataGridTemplateColumn Header="SITUAÇÃO" SortMemberPath="Status" Width="150" CellTemplate="{StaticResource Selo}"/>
            </DataGrid.Columns>
          </DataGrid>

          <TextBlock x:Name="ListaVazia" Grid.Row="1" Text="Verificando as pastas..." Foreground="{DynamicResource TextoFraco}"
                     HorizontalAlignment="Center" VerticalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap" Margin="24" IsHitTestVisible="False"/>
        </Grid>
      </Border>
    </Grid>

    <!-- Rodapé -->
    <Border Grid.Row="1" Background="{DynamicResource Superficie}" BorderBrush="{DynamicResource Borda}" BorderThickness="0,1,0,0" Padding="28,14">
      <Grid x:Name="GradeRodape">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel VerticalAlignment="Center" Margin="0,0,20,0">
          <TextBlock x:Name="TxtStatus" Text="Iniciando" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
          <TextBlock x:Name="TxtDetalhe" Foreground="{DynamicResource TextoMudo}" FontSize="12" Margin="0,2,0,0" TextTrimming="CharacterEllipsis"/>
          <ProgressBar x:Name="Barra" Margin="0,8,0,0" Minimum="0" Maximum="1" Value="0" Visibility="Hidden"/>
        </StackPanel>
        <WrapPanel x:Name="PainelBotoes" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <Button x:Name="BtnRelatorio" Content="Relatório" Style="{StaticResource BtnDiscreto}" Margin="0,0,4,0" ToolTip="O que foi adicionado, atualizado ou retirado do Drive"/>
          <Button x:Name="BtnLogs" Content="Logs" Style="{StaticResource BtnDiscreto}" Margin="0,0,8,0" ToolTip="Abre a pasta com o registro técnico de cada execução"/>
          <Button x:Name="BtnVerificar" Content="Verificar" Style="{StaticResource BtnSecundario}" Margin="0,0,8,0" ToolTip="Compara as pastas sem copiar nada (F5)"/>
          <Button x:Name="BtnSincronizar" Content="Sincronizar" Style="{StaticResource BtnPrimario}" MinWidth="130"/>
        </WrapPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

$XamlRelatorio = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Relatório de alterações no Drive"
        Width="1040" Height="660" MinWidth="660" MinHeight="500"
        WindowStartupLocation="CenterOwner"
        Background="{DynamicResource Fundo}" Foreground="{DynamicResource Texto}"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display">

  <Window.Resources>
<!--ESTILOS-->
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Grid Margin="28,22,28,18">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <StackPanel Margin="0,0,0,16">
        <TextBlock Text="Histórico do Google Drive" Foreground="{DynamicResource TextoMudo}" FontSize="12"/>
        <TextBlock Text="Relatório de alterações" FontSize="21" FontWeight="SemiBold" Margin="0,2,0,0"/>
      </StackPanel>

      <!-- Filtros -->
      <WrapPanel Grid.Row="1" Margin="0,0,0,6">
        <Border Style="{StaticResource GrupoSegmentos}" Margin="0,0,10,8">
          <StackPanel Orientation="Horizontal">
            <RadioButton x:Name="RpHoje" Content="Hoje" GroupName="Periodo" Style="{StaticResource Segmento}"/>
            <RadioButton x:Name="Rp7" Content="7 dias" GroupName="Periodo" Style="{StaticResource Segmento}"/>
            <RadioButton x:Name="Rp30" Content="30 dias" GroupName="Periodo" Style="{StaticResource Segmento}" IsChecked="True"/>
            <RadioButton x:Name="RpTudo" Content="Tudo" GroupName="Periodo" Style="{StaticResource Segmento}"/>
          </StackPanel>
        </Border>
        <Border Style="{StaticResource GrupoSegmentos}" Margin="0,0,10,8">
          <StackPanel Orientation="Horizontal">
            <RadioButton x:Name="RaTodas" Content="Todas as alterações" GroupName="Acao" Style="{StaticResource Segmento}" IsChecked="True"/>
            <RadioButton x:Name="RaAdicionados" Content="Adicionados" GroupName="Acao" Style="{StaticResource Segmento}"/>
            <RadioButton x:Name="RaAtualizados" Content="Atualizados" GroupName="Acao" Style="{StaticResource Segmento}"/>
            <RadioButton x:Name="RaSaiu" Content="Saíram do servidor" GroupName="Acao" Style="{StaticResource Segmento}"/>
            <RadioButton x:Name="RaArquivados" Content="Arquivados" GroupName="Acao" Style="{StaticResource Segmento}"/>
            <RadioButton x:Name="RaFalhas" Content="Falhas" GroupName="Acao" Style="{StaticResource Segmento}"/>
            <RadioButton x:Name="RaExecucoes" Content="Execuções" GroupName="Acao" Style="{StaticResource Segmento}"/>
          </StackPanel>
        </Border>
        <Grid Width="230" Margin="0,0,0,8" Height="32">
          <TextBox x:Name="CampoBusca" Style="{StaticResource EstiloCampo}" Height="32"/>
          <TextBlock x:Name="DicaBusca" Text="Buscar arquivo" Foreground="{DynamicResource TextoFraco}" Margin="11,0,0,0"
                     VerticalAlignment="Center" IsHitTestVisible="False"/>
        </Grid>
      </WrapPanel>

      <!-- Totais do período -->
      <Border Grid.Row="2" Style="{StaticResource Cartao}" Padding="0" Margin="0,0,0,14">
        <UniformGrid Columns="5">
          <Border BorderBrush="{DynamicResource Divisoria}" BorderThickness="0,0,1,0" Padding="18,12">
            <StackPanel>
              <TextBlock x:Name="RelAdicionados" Text="0" Style="{StaticResource Numero}" Foreground="{DynamicResource NumNovo}"/>
              <TextBlock Text="adicionados" Style="{StaticResource RotuloNumero}"/>
            </StackPanel>
          </Border>
          <Border BorderBrush="{DynamicResource Divisoria}" BorderThickness="0,0,1,0" Padding="18,12">
            <StackPanel>
              <TextBlock x:Name="RelAtualizados" Text="0" Style="{StaticResource Numero}" Foreground="{DynamicResource NumAlterado}"/>
              <TextBlock Text="atualizados" Style="{StaticResource RotuloNumero}"/>
            </StackPanel>
          </Border>
          <Border BorderBrush="{DynamicResource Divisoria}" BorderThickness="0,0,1,0" Padding="18,12"
                  ToolTip="PDFs que deixaram de existir na pasta principal do servidor.">
            <StackPanel>
              <TextBlock x:Name="RelSaiu" Text="0" Style="{StaticResource Numero}" Foreground="{DynamicResource NumSoNoDrive}"/>
              <TextBlock Text="saíram do servidor" Style="{StaticResource RotuloNumero}"/>
            </StackPanel>
          </Border>
          <Border BorderBrush="{DynamicResource Divisoria}" BorderThickness="0,0,1,0" Padding="18,12"
                  ToolTip="PDFs retirados da pasta do Drive e guardados em _ARQUIVADOS.">
            <StackPanel>
              <TextBlock x:Name="RelArquivados" Text="0" Style="{StaticResource Numero}"/>
              <TextBlock Text="arquivados" Style="{StaticResource RotuloNumero}"/>
            </StackPanel>
          </Border>
          <Border Padding="18,12">
            <StackPanel>
              <TextBlock x:Name="RelFalhas" Text="0" Style="{StaticResource Numero}" Foreground="{DynamicResource NumFalha}"/>
              <TextBlock Text="falhas" Style="{StaticResource RotuloNumero}"/>
            </StackPanel>
          </Border>
        </UniformGrid>
      </Border>

      <!-- Lista -->
      <Border Grid.Row="3" Style="{StaticResource Cartao}" Padding="0">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <StackPanel Margin="18,14,18,12">
            <TextBlock Text="Registros" FontWeight="SemiBold"/>
            <TextBlock x:Name="TxtRelResumo" Foreground="{DynamicResource TextoFraco}" FontSize="12" Margin="0,2,0,0" TextTrimming="CharacterEllipsis"/>
          </StackPanel>
          <DataGrid x:Name="ListaRel" Grid.Row="1" Style="{StaticResource Grade}">
            <DataGrid.Columns>
              <DataGridTextColumn Header="DATA" Binding="{Binding QuandoTexto}" Width="140" ElementStyle="{StaticResource TextoCelulaMudo}"/>
              <DataGridTemplateColumn Header="AÇÃO" Width="160" CellTemplate="{StaticResource Selo}"/>
              <DataGridTextColumn Header="ARQUIVO" Binding="{Binding Arquivo}" Width="1.3*" ElementStyle="{StaticResource TextoCelula}"/>
              <DataGridTextColumn Header="TAMANHO" Binding="{Binding TamanhoTexto}" Width="95" ElementStyle="{StaticResource TextoCelulaDireita}" HeaderStyle="{StaticResource CabecalhoDireita}"/>
              <DataGridTextColumn Header="DETALHE" Binding="{Binding Detalhe}" Width="*" ElementStyle="{StaticResource TextoCelulaMudo}"/>
            </DataGrid.Columns>
          </DataGrid>
          <TextBlock x:Name="RelVazio" Grid.Row="1" Foreground="{DynamicResource TextoFraco}"
                     HorizontalAlignment="Center" VerticalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap" Margin="24" IsHitTestVisible="False"/>
        </Grid>
      </Border>
    </Grid>

    <Border Grid.Row="1" Background="{DynamicResource Superficie}" BorderBrush="{DynamicResource Borda}" BorderThickness="0,1,0,0" Padding="28,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel VerticalAlignment="Center" Margin="0,0,20,0">
          <TextBlock Text="Arquivo do histórico" FontSize="12" Foreground="{DynamicResource TextoMudo}"/>
          <TextBlock x:Name="TxtRelArquivo" FontFamily="Consolas" FontSize="12" Margin="0,2,0,0" TextTrimming="CharacterEllipsis"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <Button x:Name="BtnAbrirHistorico" Content="Abrir pasta" Style="{StaticResource BtnDiscreto}" Margin="0,0,8,0"/>
          <Button x:Name="BtnAtualizarRel" Content="Atualizar" Style="{StaticResource BtnSecundario}" Margin="0,0,8,0" ToolTip="Recarrega o histórico (F5)"/>
          <Button x:Name="BtnExportar" Content="Exportar planilha" Style="{StaticResource BtnPrimario}" ToolTip="Salva os registros filtrados num arquivo que abre no Excel"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

$XamlDialogo = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Sincronizador de Mapas PDF"
        Width="460" SizeToContent="Height" ResizeMode="NoResize" ShowInTaskbar="False"
        WindowStartupLocation="CenterOwner"
        Background="{DynamicResource Superficie}" Foreground="{DynamicResource Texto}"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display">
  <Window.Resources>
<!--ESTILOS-->
  </Window.Resources>
  <StackPanel>
    <Grid Margin="24,22,24,22">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Border x:Name="DlgIcone" Width="34" Height="34" CornerRadius="17" VerticalAlignment="Top" Margin="0,0,16,0"
              Background="{DynamicResource SeloCopiandoF}">
        <TextBlock x:Name="DlgIconeTexto" Text="i" FontSize="17" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center"
                   Foreground="{DynamicResource SeloCopiandoT}"/>
      </Border>
      <StackPanel Grid.Column="1">
        <TextBlock x:Name="DlgTitulo" FontSize="15" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,6,0,8"/>
        <TextBlock x:Name="DlgMensagem" Foreground="{DynamicResource TextoMudo}" TextWrapping="Wrap" LineHeight="19"/>
      </StackPanel>
    </Grid>
    <Border Background="{DynamicResource Fundo}" BorderBrush="{DynamicResource Borda}" BorderThickness="0,1,0,0" Padding="24,14">
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
        <Button x:Name="DlgNao" Style="{StaticResource BtnSecundario}" Margin="0,0,8,0" MinWidth="96"/>
        <Button x:Name="DlgSim" Style="{StaticResource BtnPrimario}" MinWidth="96"/>
      </StackPanel>
    </Border>
  </StackPanel>
</Window>
'@

$XamlRevisao = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Sincronizador de Mapas PDF"
        Width="660" Height="560" MinWidth="480" MinHeight="380" ShowInTaskbar="False"
        WindowStartupLocation="CenterOwner"
        Background="{DynamicResource Superficie}" Foreground="{DynamicResource Texto}"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid Margin="24,22,24,16">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Border Width="34" Height="34" CornerRadius="17" VerticalAlignment="Top" Margin="0,0,16,0" Background="{DynamicResource SeloNovoF}">
        <TextBlock Text="?" FontSize="17" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="{DynamicResource SeloNovoT}"/>
      </Border>
      <StackPanel Grid.Column="1">
        <TextBlock Text="Revisar antes de sincronizar" FontSize="15" FontWeight="SemiBold" Margin="0,6,0,6"/>
        <TextBlock x:Name="RevResumo" Foreground="{DynamicResource TextoMudo}" TextWrapping="Wrap"/>
      </StackPanel>
    </Grid>
    <Border Grid.Row="1" BorderBrush="{DynamicResource Divisoria}" BorderThickness="0,1,0,0">
      <DataGrid x:Name="RevLista" Style="{StaticResource Grade}">
        <DataGrid.Columns>
          <DataGridTemplateColumn Header="AÇÃO" Width="160" CellTemplate="{StaticResource Selo}"/>
          <DataGridTextColumn Header="ARQUIVO" Binding="{Binding Nome}" Width="*" ElementStyle="{StaticResource TextoCelula}"/>
          <DataGridTextColumn Header="TAMANHO" Binding="{Binding TamanhoTexto}" Width="100" ElementStyle="{StaticResource TextoCelulaDireita}" HeaderStyle="{StaticResource CabecalhoDireita}"/>
        </DataGrid.Columns>
      </DataGrid>
    </Border>
    <Border Grid.Row="2" Background="{DynamicResource Fundo}" BorderBrush="{DynamicResource Borda}" BorderThickness="0,1,0,0" Padding="24,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="RevNota" Foreground="{DynamicResource TextoFraco}" FontSize="12" TextWrapping="Wrap" VerticalAlignment="Center" Margin="0,0,16,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal">
          <Button x:Name="RevCancelar" Content="Cancelar" Style="{StaticResource BtnSecundario}" Margin="0,0,8,0" MinWidth="96"/>
          <Button x:Name="RevConfirmar" Style="{StaticResource BtnPrimario}" MinWidth="130"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

function Show-ErroFatal {
    param([string]$Mensagem)
    try {
        Add-Type -AssemblyName PresentationFramework
        [void][System.Windows.MessageBox]::Show($Mensagem, 'Sincronizador de Mapas PDF', 'OK', 'Error')
    } catch {
        Write-Host $Mensagem -ForegroundColor Red
    }
}

try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # Estilos e cores ficam no nível do programa: são lidos uma única vez e valem para
    # todas as janelas. Antes cada caixa de mensagem relia tudo, o que deixava a abertura lenta.
    $wpfApp = [System.Windows.Application]::Current
    if ($null -eq $wpfApp) { $wpfApp = [System.Windows.Application]::new() }
    $wpfApp.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
    $recursos = [Windows.Markup.XamlReader]::Parse(
        '<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">' +
        $XamlEstilos + '</ResourceDictionary>')
    $wpfApp.Resources.MergedDictionaries.Add($recursos)
    $script:TemaAplicado = $null

    # Barra de título escura no Windows 10/11 (se não for possível, segue sem ela).
    $dwmDisponivel = $false
    try {
        if (-not ('SincMapasDwm' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SincMapasDwm {
    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
    public static void Definir(IntPtr hwnd, bool escuro) {
        int v = escuro ? 1 : 0;
        if (DwmSetWindowAttribute(hwnd, 20, ref v, 4) != 0) { DwmSetWindowAttribute(hwnd, 19, ref v, 4); }
    }
}
'@
        }
        $dwmDisponivel = $true
    } catch { }

    $avisoConfig = $null
    try {
        $configInicial = Read-SyncConfig $CaminhoConfig
        if (-not (Test-Path -LiteralPath $CaminhoConfig)) {
            try { Save-SyncConfig $configInicial $CaminhoConfig } catch { }
        }
    } catch {
        $avisoConfig = (Get-MensagemErro $_) + "`n`nAs configurações padrão serão usadas até que você altere algum caminho."
        $configInicial = Get-ConfigPadrao
    }

    # Pincéis prontos de cada tema (criados uma vez e compartilhados pelas janelas).
    $Pinceis = @{}
    foreach ($nomeTema in @($Paletas.Keys)) {
        $mapa = @{}
        foreach ($chave in @($Paletas[$nomeTema].Keys)) {
            $cor = [System.Windows.Media.ColorConverter]::ConvertFromString($Paletas[$nomeTema][$chave])
            # ::new() em vez de New-Object: New-Object embrulha o objeto num PSObject, e o WPF
            # não reconhece esse embrulho como cor quando ele vai para o dicionário de recursos.
            $pincel = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]$cor)
            $pincel.Freeze()
            $mapa[$chave] = $pincel
        }
        $Pinceis[$nomeTema] = $mapa
    }

    function Get-TemaSistema {
        try {
            $v = Get-ItemPropertyValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'AppsUseLightTheme'
            if ([int]$v -eq 0) { return 'Escuro' }
        } catch { }
        return 'Claro'
    }

    function Get-TemaEfetivo {
        param([string]$Escolha)
        if ($Escolha -eq 'Escuro' -or $Escolha -eq 'Claro') { return $Escolha }
        return (Get-TemaSistema)
    }

    function Set-PaletaJanela {
        param($Janela, [string]$Tema)
        if ($script:TemaAplicado -eq $Tema) { return }
        $mapa = $Pinceis[$Tema]
        foreach ($chave in @($mapa.Keys)) {
            $valor = $mapa[$chave]
            if ($valor -is [System.Management.Automation.PSObject]) { $valor = $valor.PSObject.BaseObject }
            $wpfApp.Resources[$chave] = $valor
        }
        $script:TemaAplicado = $Tema
    }

    function Set-BarraTitulo {
        param($Janela, [bool]$Escuro)
        if (-not $dwmDisponivel -or $null -eq $Janela) { return }
        try {
            $h = (New-Object System.Windows.Interop.WindowInteropHelper -ArgumentList $Janela).Handle
            if ($h -ne [IntPtr]::Zero) { [SincMapasDwm]::Definir($h, $Escuro) }
        } catch { }
    }

    # Caixa de mensagem no mesmo visual e tema da ferramenta (substitui a do Windows).
    function Show-Dialogo {
        param($Dono, [string]$Titulo, [string]$Mensagem, [string]$Tipo = 'info', [string]$TextoSim = 'OK', [string]$TextoNao = '')
        try {
            $d = New-Janela $XamlDialogo
            $tema = 'Claro'
            if ($null -ne $app -and $app.TemaEfetivo) { $tema = $app.TemaEfetivo }
            if ($null -ne $Dono -and $Dono.IsVisible) { $d.Owner = $Dono } else { $d.WindowStartupLocation = 'CenterScreen' }
            $el = Get-Elementos $d @('DlgIcone', 'DlgIconeTexto', 'DlgTitulo', 'DlgMensagem', 'DlgSim', 'DlgNao')
            $el.DlgTitulo.Text = $Titulo
            $el.DlgMensagem.Text = $Mensagem
            $cores = @{ info = @('SeloCopiandoF', 'SeloCopiandoT', 'i'); pergunta = @('SeloNovoF', 'SeloNovoT', '?');
                        aviso = @('SeloAlteradoF', 'SeloAlteradoT', '!'); erro = @('SeloFalhouF', 'SeloFalhouT', '!') }
            $c = $cores[$Tipo]
            if ($null -eq $c) { $c = $cores['info'] }
            $el.DlgIcone.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, $c[0])
            $el.DlgIconeTexto.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $c[1])
            $el.DlgIconeTexto.Text = $c[2]
            $el.DlgSim.Content = $TextoSim
            $el.DlgSim.IsDefault = $true
            if ([string]::IsNullOrEmpty($TextoNao)) { $el.DlgNao.Visibility = 'Collapsed' } else { $el.DlgNao.Content = $TextoNao }
            $el.DlgSim.Add_Click({ $d.DialogResult = $true })
            $el.DlgNao.Add_Click({ $d.DialogResult = $false })
            $d.Add_KeyDown({ param($s, $e) if ($e.Key -eq [System.Windows.Input.Key]::Escape) { $e.Handled = $true; $d.DialogResult = $false } })
            $d.Add_SourceInitialized({ Set-BarraTitulo $d ($tema -eq 'Escuro') })
            $d.Add_ContentRendered({ [void]$el.DlgSim.Focus() })
            return ($d.ShowDialog() -eq $true)
        } catch {
            $botoes = 'OK'
            if ($TextoNao) { $botoes = 'YesNo' }
            $res = [System.Windows.MessageBox]::Show(($Titulo + "`n`n" + $Mensagem), 'Sincronizador de Mapas PDF', $botoes)
            return ($res -eq [System.Windows.MessageBoxResult]::Yes -or $res -eq [System.Windows.MessageBoxResult]::OK)
        }
    }

    function New-Janela {
        param([string]$Xaml)
        $texto = $Xaml.Replace('<!--ESTILOS-->', '')
        $w = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$texto)))
        return $w
    }

    function Get-Elementos {
        param($Janela, [string[]]$Nomes)
        $h = @{}
        foreach ($n in $Nomes) {
            $h[$n] = $Janela.FindName($n)
            if ($null -eq $h[$n]) { throw "Elemento da interface não encontrado: $n" }
        }
        return $h
    }

    $janela = New-Janela $XamlPrincipal
    $ui = Get-Elementos $janela @(
        'TxtSobretitulo','TxtUltima','RbClaro','RbEscuro','RbSistema','GradeCaminhos','CartaoOrigem','CartaoDestino',
        'PontoOrigem','StatusOrigem','CampoOrigem','BtnAlterarOrigem','BtnAbrirOrigem',
        'PontoDestino','StatusDestino','CampoDestino','BtnAlterarDestino','BtnAbrirDestino',
        'NumNovos','NumAlterados','NumEmDia','NumSoNoDrive',
        'ChkArquivar','Lista','ListaVazia','FiltroNovos','FiltroAlterados','FiltroEmDia','FiltroSoNoDrive',
        'GradeBuscaLista','CampoBuscaLista','DicaBuscaLista','TxtSubLista',
        'GradeRodape','TxtStatus','TxtDetalhe','Barra','PainelBotoes',
        'BtnRelatorio','BtnLogs','BtnVerificar','BtnSincronizar')

    $app = @{
        Config           = $configInicial
        TemaEfetivo      = 'Claro'
        Carregando       = $true
        Executando       = $false
        Modo             = $null
        Sync             = $null
        PS               = $null
        RS               = $null
        Handle           = $null
        Itens            = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
        Indice           = @{}
        Visiveis         = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
        IndiceVisivel    = @{}
        Filtro           = $null
        OrdemCampo       = 'Status'
        OrdemDesc        = $false
        TimerBusca       = New-Object System.Windows.Threading.DispatcherTimer
        OrigemOk         = $false
        DestinoOk        = $false
        TemDados         = $false
        UltimoErro       = $null
        UltimoLog        = $null
        FecharAoTerminar = $false
        Estreito         = $null
        Timer            = New-Object System.Windows.Threading.DispatcherTimer
    }
    $app.Timer.Interval = [TimeSpan]::FromMilliseconds(120)
    $ui.Lista.ItemsSource = $app.Visiveis
    $app.TimerBusca.Interval = [TimeSpan]::FromMilliseconds(220)

    # Ordem da coluna Situação: o que precisa de atenção vem primeiro.
    $OrdemSituacao = @{ Falhou = 0; Copiando = 1; Novo = 2; Alterado = 3; SoNoDrive = 4; Copiado = 5; Arquivado = 6; EmDia = 7 }

    $rel = @{ Janela = $null; Ui = @{}; Dados = @(); Itens = $null; Estreito = $null }

    # ---------------------------------------------------------------- tema

    function Set-Tema {
        param([string]$Escolha)
        $efetivo = Get-TemaEfetivo $Escolha
        $app.TemaEfetivo = $efetivo
        foreach ($w in @($janela, $rel.Janela)) {
            if ($null -ne $w) {
                Set-PaletaJanela $w $efetivo
                Set-BarraTitulo $w ($efetivo -eq 'Escuro')
            }
        }
    }

    # ---------------------------------------------------------------- utilidades da janela

    function Invoke-Seguro {
        param([scriptblock]$Acao)
        try { & $Acao }
        catch {
            [void](Show-Dialogo $janela 'Ocorreu um erro inesperado' (Get-MensagemErro $_) 'erro')
        }
    }

    function Save-Config {
        try { Save-SyncConfig $app.Config $CaminhoConfig; return $true }
        catch { Set-Status 'Configuração não salva' ('Não foi possível gravar config.json: ' + (Get-MensagemErro $_)) 'erro'; return $false }
    }

    function Get-StatusTexto {
        param([string]$Status)
        switch ($Status) {
            'Novo'      { return 'Novo' }
            'Alterado'  { return 'Alterado' }
            'EmDia'     { return 'Em dia' }
            'SoNoDrive' { return 'Só no Drive' }
            'Copiando'  { return 'Copiando...' }
            'Copiado'   { return 'Copiado' }
            'Falhou'    { return 'Falhou' }
            'Arquivado' { return 'Arquivado' }
        }
        return $Status
    }

    function New-LinhaLista {
        param($Item, [string]$Status, [string]$Detalhe, [string]$StatusOriginal)
        $mod = ''
        if ($null -ne $Item.Modificado) { $mod = ([datetime]$Item.Modificado).ToString('dd/MM/yyyy  HH:mm') }
        [pscustomobject]@{
            Nome            = $Item.Nome
            Bytes           = $Item.Bytes
            Modificado      = $Item.Modificado
            TamanhoTexto    = Format-Tamanho $Item.Bytes
            ModificadoTexto = $mod
            Status          = $Status
            StatusTexto     = Get-StatusTexto $Status
            StatusOriginal  = $StatusOriginal
            Detalhe         = $Detalhe
            CaminhoOrigem   = $Item.CaminhoOrigem
            CaminhoDestino  = $Item.CaminhoDestino
        }
    }

    function Set-Status {
        param([string]$Titulo, [string]$Detalhe, [string]$Tom = 'normal')
        $ui.TxtStatus.Text = $Titulo
        $ui.TxtDetalhe.Text = $Detalhe
        if ([string]::IsNullOrEmpty($Detalhe)) { $ui.TxtDetalhe.ToolTip = $null } else { $ui.TxtDetalhe.ToolTip = $Detalhe }
        $chave = 'Texto'
        if ($Tom -eq 'erro') { $chave = 'StatusErro' }
        if ($Tom -eq 'sucesso') { $chave = 'StatusSucesso' }
        $ui.TxtStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $chave)
    }

    function Set-Indicador {
        param($Ponto, $Texto, [string]$Estado, [string]$Rotulo)
        $Texto.Text = $Rotulo
        $chave = 'PontoNeutro'
        if ($Estado -eq 'ok')    { $chave = 'PontoOk' }
        if ($Estado -eq 'erro')  { $chave = 'PontoErro' }
        if ($Estado -eq 'aviso') { $chave = 'PontoAviso' }
        $Ponto.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, $chave)
    }

    function Update-UltimaSincronizacao {
        $texto = 'Nenhuma sincronização registrada'
        $estado = Read-Estado $CaminhoEstado
        if ($null -ne $estado -and $estado.Texto) {
            $texto = 'Última sincronização: ' + $estado.Texto
            if ([int]$estado.Falhas -gt 0) { $texto += ' (' + (Get-Plural ([int]$estado.Falhas) 'falha' 'falhas') + ')' }
        }
        $ui.TxtUltima.Text = $texto
    }

    function Update-Sobretitulo {
        $nome = ''
        try { $nome = [IO.Path]::GetFileName($app.Config.Origem.TrimEnd('\')) } catch { }
        if ([string]::IsNullOrWhiteSpace($nome)) { $nome = 'Mapas PDF' }
        $ui.TxtSobretitulo.Text = $nome
    }

    function Update-Numeros {
        $novos = 0; $alterados = 0; $emDia = 0; $soNoDrive = 0
        foreach ($l in $app.Itens) {
            $s = $l.Status
            if ($s -eq 'Copiando' -or $s -eq 'Falhou') { $s = $l.StatusOriginal }
            switch ($s) {
                'Novo'      { $novos++ }
                'Alterado'  { $alterados++ }
                'EmDia'     { $emDia++ }
                'Copiado'   { $emDia++ }
                'SoNoDrive' { $soNoDrive++ }
            }
        }
        $opacidade = 1.0
        if (-not $app.TemDados) { $opacidade = 0.35 }
        foreach ($par in @(@('NumNovos', $novos), @('NumAlterados', $alterados), @('NumEmDia', $emDia), @('NumSoNoDrive', $soNoDrive))) {
            $ui[$par[0]].Text = "$($par[1])"
            $ui[$par[0]].Opacity = $opacidade
        }
    }

    function Get-Contagem {
        $c = @{ Novo = 0; Alterado = 0; EmDia = 0; SoNoDrive = 0; Falhou = 0 }
        foreach ($l in $app.Itens) { if ($c.ContainsKey($l.Status)) { $c[$l.Status]++ } }
        return $c
    }

    function Set-Plano {
        param($Itens)
        $nova = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
        $indice = @{}
        foreach ($i in @($Itens)) {
            $indice[$i.Nome] = $nova.Count
            $nova.Add((New-LinhaLista -Item $i -Status $i.Status -StatusOriginal $i.Status))
        }
        $app.Itens = $nova
        $app.Indice = $indice
        $app.TemDados = $true
        if ($nova.Count -eq 0) { $ui.ListaVazia.Text = 'Nenhum PDF encontrado na pasta principal da origem.' }
        Update-Lista
        Update-Numeros
    }

    # Monta o que aparece na tabela: aplica filtro dos números, busca e ordenação.
    function Update-Lista {
        $busca = $ui.CampoBuscaLista.Text.Trim()
        $filtro = $app.Filtro
        $selecionados = foreach ($l in $app.Itens) {
            if ($filtro) {
                if ($filtro -eq 'EmDia') { if ($l.StatusOriginal -ne 'EmDia') { continue } }
                elseif ($l.StatusOriginal -ne $filtro) { continue }
            }
            if ($busca -and $l.Nome.IndexOf($busca, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
            $l
        }
        switch ($app.OrdemCampo) {
            'Nome'       { $chave = { $_.Nome } }
            'Bytes'      { $chave = { [long]$_.Bytes } }
            'Modificado' { $chave = { if ($null -ne $_.Modificado) { [datetime]$_.Modificado } else { [datetime]::MinValue } } }
            default      { $chave = { $o = $OrdemSituacao[$_.Status]; if ($null -eq $o) { 9 } else { $o } } }
        }
        $ordenados = @($selecionados | Sort-Object -Property @{ Expression = $chave; Descending = [bool]$app.OrdemDesc }, @{ Expression = { $_.Nome }; Descending = $false })

        $col = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
        $indice = @{}
        foreach ($l in $ordenados) { $indice[$l.Nome] = $col.Count; $col.Add($l) }
        $app.Visiveis = $col
        $app.IndiceVisivel = $indice
        $ui.Lista.ItemsSource = $col

        foreach ($c in $ui.Lista.Columns) {
            if ($c.SortMemberPath -eq $app.OrdemCampo) {
                if ($app.OrdemDesc) { $c.SortDirection = [System.ComponentModel.ListSortDirection]::Descending }
                else { $c.SortDirection = [System.ComponentModel.ListSortDirection]::Ascending }
            } else { $c.SortDirection = $null }
        }

        foreach ($n in 'FiltroNovos', 'FiltroAlterados', 'FiltroEmDia', 'FiltroSoNoDrive') {
            if ($filtro -and "$($ui[$n].Tag)" -eq $filtro) {
                $ui[$n].SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'LinhaSelecionada')
            } else {
                $ui[$n].ClearValue([System.Windows.Controls.Border]::BackgroundProperty)
            }
        }

        $padrao = 'Somente arquivos .pdf. Subpastas são ignoradas. Clique duas vezes para abrir.'
        if ($filtro -or $busca) {
            $nomes = @{ Novo = 'novos'; Alterado = 'alterados'; EmDia = 'em dia'; SoNoDrive = 'só no Drive' }
            $txt = "Mostrando $($col.Count) de $($app.Itens.Count)"
            if ($filtro) { $txt += ' (' + $nomes[$filtro] + ')' }
            if ($filtro) { $txt += '. Clique no número de novo para ver todos.' } else { $txt += '.' }
            $ui.TxtSubLista.Text = $txt
        } else {
            $ui.TxtSubLista.Text = $padrao
        }

        if ($app.Itens.Count -eq 0) {
            $ui.ListaVazia.Visibility = 'Visible'
        } elseif ($col.Count -eq 0) {
            $ui.ListaVazia.Text = 'Nenhum arquivo com este filtro.'
            $ui.ListaVazia.Visibility = 'Visible'
        } else {
            $ui.ListaVazia.Visibility = 'Collapsed'
        }
    }

    function Set-Ordenacao {
        param([string]$Campo)
        if ([string]::IsNullOrEmpty($Campo)) { return }
        if ($app.OrdemCampo -eq $Campo) { $app.OrdemDesc = -not $app.OrdemDesc }
        else {
            $app.OrdemCampo = $Campo
            $app.OrdemDesc = ($Campo -eq 'Bytes' -or $Campo -eq 'Modificado')
        }
        Update-Lista
    }

    function Set-Filtro {
        param([string]$Filtro)
        if ($app.Filtro -eq $Filtro) { $app.Filtro = $null } else { $app.Filtro = $Filtro }
        Update-Lista
    }

    # Lista do que a sincronização vai fazer, para a revisão antes de confirmar.
    function Get-Pendentes {
        $acoes = @{ Novo = @(0, 'Adicionar'); Alterado = @(1, 'Atualizar'); SoNoDrive = @(2, 'Retirar do Drive') }
        $lista = New-Object 'System.Collections.Generic.List[object]'
        foreach ($l in $app.Itens) {
            $s = $l.Status
            if ($s -eq 'Falhou' -or $s -eq 'Copiando') { $s = $l.StatusOriginal }
            if ($s -eq 'SoNoDrive' -and -not $app.Config.EspelharDrive) { continue }
            if (-not $acoes.ContainsKey($s)) { continue }
            $lista.Add([pscustomobject]@{
                Ordem        = $acoes[$s][0]
                Status       = $s
                StatusTexto  = $acoes[$s][1]
                Nome         = $l.Nome
                TamanhoTexto = $l.TamanhoTexto
            })
        }
        $ordenada = @($lista | Sort-Object Ordem, Nome)
        return ,$ordenada
    }

    function Show-Revisao {
        param($Pendentes)
        try {
            $w = New-Janela $XamlRevisao
            $w.Owner = $janela
            $el = Get-Elementos $w @('RevResumo', 'RevLista', 'RevNota', 'RevCancelar', 'RevConfirmar')
            $novos = @($Pendentes | Where-Object { $_.Status -eq 'Novo' }).Count
            $alterados = @($Pendentes | Where-Object { $_.Status -eq 'Alterado' }).Count
            $retirar = @($Pendentes | Where-Object { $_.Status -eq 'SoNoDrive' }).Count
            $partes = @()
            if ($novos -gt 0) { $partes += "$novos para adicionar" }
            if ($alterados -gt 0) { $partes += "$alterados para atualizar" }
            if ($retirar -gt 0) { $partes += "$retirar para retirar do Drive" }
            $el.RevResumo.Text = ($partes -join ', ') + '. Confira a lista abaixo. O servidor não é alterado.'

            $col = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
            foreach ($p in $Pendentes) { $col.Add($p) }
            $el.RevLista.ItemsSource = $col

            $nota = 'A lista é conferida de novo no momento da cópia.'
            if ($retirar -gt 0) {
                $naOrigem = 0
                foreach ($l in $app.Itens) { if ($l.StatusOriginal -ne 'SoNoDrive') { $naOrigem++ } }
                $limite = [Math]::Max(20, [int][Math]::Ceiling($naOrigem * 0.5))
                if ($retirar -gt $limite) {
                    $nota = "Por segurança, nada será retirado do Drive: são mais de $limite PDFs de uma vez. Confira se a pasta do servidor está completa."
                } else {
                    $nota = 'Os retirados são guardados em ' + [IO.Path]::Combine($app.Config.PastaArquivados, (Get-Date -Format 'yyyy-MM-dd')) + '.'
                }
            }
            $el.RevNota.Text = $nota
            $el.RevConfirmar.Content = 'Sincronizar ' + (Get-Plural $Pendentes.Count 'arquivo' 'arquivos')
            $el.RevConfirmar.IsDefault = $true
            $el.RevConfirmar.Add_Click({ $w.DialogResult = $true })
            $el.RevCancelar.Add_Click({ $w.DialogResult = $false })
            $w.Add_KeyDown({ param($s, $e) if ($e.Key -eq [System.Windows.Input.Key]::Escape) { $e.Handled = $true; $w.DialogResult = $false } })
            $w.Add_SourceInitialized({ Set-BarraTitulo $w ($app.TemaEfetivo -eq 'Escuro') })
            $w.Add_ContentRendered({ [void]$el.RevConfirmar.Focus() })
            return ($w.ShowDialog() -eq $true)
        } catch {
            $txt = (Get-Plural $Pendentes.Count 'arquivo será alterado' 'arquivos serão alterados') + ' no Drive. O servidor não é alterado.'
            return (Show-Dialogo $janela 'Confirmar sincronização' $txt 'pergunta' 'Sincronizar' 'Cancelar')
        }
    }

    function Set-Linha {
        param([string]$Nome, [string]$Status, [string]$Detalhe)
        if (-not $app.Indice.ContainsKey($Nome)) { return }
        $idx = $app.Indice[$Nome]
        $atual = $app.Itens[$idx]
        $nova = New-LinhaLista -Item $atual -Status $Status -Detalhe $Detalhe -StatusOriginal $atual.StatusOriginal
        $app.Itens[$idx] = $nova
        if ($app.IndiceVisivel.ContainsKey($Nome)) {
            $app.Visiveis[$app.IndiceVisivel[$Nome]] = $nova
            if ($Status -eq 'Copiando') { $ui.Lista.ScrollIntoView($nova) }
        }
    }

    function Set-Ambiente {
        param($Amb)
        $app.OrigemOk = [bool]$Amb.OrigemOk
        $app.DestinoOk = [bool]$Amb.DestinoOk
        if ($Amb.OrigemOk) { Set-Indicador $ui.PontoOrigem $ui.StatusOrigem 'ok' 'Conectado' }
        else { Set-Indicador $ui.PontoOrigem $ui.StatusOrigem 'erro' 'Indisponível' }
        if ($Amb.DestinoOk -and $Amb.DestinoExiste) { Set-Indicador $ui.PontoDestino $ui.StatusDestino 'ok' 'Conectado' }
        elseif ($Amb.DestinoOk) { Set-Indicador $ui.PontoDestino $ui.StatusDestino 'aviso' 'Será criada' }
        else { Set-Indicador $ui.PontoDestino $ui.StatusDestino 'erro' 'Indisponível' }
    }

    function Update-Controles {
        $rodando = [bool]$app.Executando
        foreach ($n in 'CampoOrigem','CampoDestino','BtnAlterarOrigem','BtnAlterarDestino','ChkArquivar','BtnVerificar') {
            $ui[$n].IsEnabled = -not $rodando
        }
        if ($rodando -and $app.Modo -eq 'sincronizar') {
            $ui.BtnSincronizar.Style = $janela.FindResource('BtnSecundario')
            $ui.BtnSincronizar.Content = 'Cancelar'
            $ui.BtnSincronizar.IsEnabled = -not $app.Sync.Cancelar
            $ui.Barra.Visibility = 'Visible'
        } else {
            $ui.BtnSincronizar.Style = $janela.FindResource('BtnPrimario')
            $ui.BtnSincronizar.Content = 'Sincronizar'
            $ui.BtnSincronizar.IsEnabled = (-not $rodando) -and $app.OrigemOk -and $app.DestinoOk
            $ui.Barra.Visibility = 'Hidden'
        }
    }

    function Set-Layout {
        param([double]$Largura)
        $estreito = $Largura -lt 860
        if ($estreito -eq $app.Estreito) { return }
        $app.Estreito = $estreito
        if ($estreito) {
            $ui.GradeCaminhos.Columns = 1
            $ui.CartaoOrigem.Margin  = New-Object System.Windows.Thickness -ArgumentList 0, 0, 0, 10
            $ui.CartaoDestino.Margin = New-Object System.Windows.Thickness -ArgumentList 0
            $ui.Lista.Columns[2].Visibility = 'Collapsed'
            [System.Windows.Controls.DockPanel]::SetDock($ui.ChkArquivar, 'Bottom')
            $ui.ChkArquivar.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 12, 0, 0
            $ui.ChkArquivar.HorizontalAlignment = 'Left'
            [System.Windows.Controls.DockPanel]::SetDock($ui.GradeBuscaLista, 'Bottom')
            $ui.GradeBuscaLista.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 12, 0, 0
            $ui.GradeBuscaLista.HorizontalAlignment = 'Left'
            $ui.GradeBuscaLista.Width = 260
            [System.Windows.Controls.Grid]::SetRow($ui.PainelBotoes, 1)
            [System.Windows.Controls.Grid]::SetColumn($ui.PainelBotoes, 0)
            [System.Windows.Controls.Grid]::SetColumnSpan($ui.PainelBotoes, 2)
            $ui.PainelBotoes.HorizontalAlignment = 'Right'
            $ui.PainelBotoes.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 12, 0, 0
            $ui.RbSistema.Content = 'Windows'
        } else {
            $ui.GradeCaminhos.Columns = 2
            $ui.CartaoOrigem.Margin  = New-Object System.Windows.Thickness -ArgumentList 0, 0, 7, 0
            $ui.CartaoDestino.Margin = New-Object System.Windows.Thickness -ArgumentList 7, 0, 0, 0
            $ui.Lista.Columns[2].Visibility = 'Visible'
            [System.Windows.Controls.DockPanel]::SetDock($ui.ChkArquivar, 'Right')
            $ui.ChkArquivar.Margin = New-Object System.Windows.Thickness -ArgumentList 16, 0, 0, 0
            $ui.ChkArquivar.HorizontalAlignment = 'Stretch'
            [System.Windows.Controls.DockPanel]::SetDock($ui.GradeBuscaLista, 'Right')
            $ui.GradeBuscaLista.Margin = New-Object System.Windows.Thickness -ArgumentList 16, 0, 0, 0
            $ui.GradeBuscaLista.HorizontalAlignment = 'Stretch'
            $ui.GradeBuscaLista.Width = 220
            [System.Windows.Controls.Grid]::SetRow($ui.PainelBotoes, 0)
            [System.Windows.Controls.Grid]::SetColumn($ui.PainelBotoes, 1)
            [System.Windows.Controls.Grid]::SetColumnSpan($ui.PainelBotoes, 1)
            $ui.PainelBotoes.Margin = New-Object System.Windows.Thickness -ArgumentList 0
            $ui.RbSistema.Content = 'Igual ao Windows'
        }
    }

    function Open-Pasta {
        param([string]$Caminho, [string]$Selecionar, $Dono = $janela)
        if ($Selecionar -and [IO.File]::Exists($Selecionar)) {
            Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"{0}"' -f $Selecionar)
            return
        }
        if ([string]::IsNullOrWhiteSpace($Caminho) -or -not (Test-Path -LiteralPath $Caminho -PathType Container)) {
            [void](Show-Dialogo $Dono 'Pasta indisponível' "Esta pasta não está acessível no momento:`n`n$Caminho" 'aviso')
            return
        }
        Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $Caminho)
    }

    function Select-Pasta {
        param([string]$Atual, [string]$Descricao)
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = $Descricao
        $dlg.ShowNewFolderButton = $false
        if ($Atual -and (Test-Path -LiteralPath $Atual -PathType Container)) { $dlg.SelectedPath = $Atual }
        try {
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
        } finally { $dlg.Dispose() }
        return $null
    }

    function Save-ConfigDaTela {
        $origem  = $ui.CampoOrigem.Text.Trim().Trim('"').Trim()
        $destino = $ui.CampoDestino.Text.Trim().Trim('"').Trim()
        $arquivar = ($ui.ChkArquivar.IsChecked -eq $true)
        $mudou = ($origem -ne $app.Config.Origem) -or ($destino -ne $app.Config.Destino) -or ($arquivar -ne [bool]$app.Config.EspelharDrive)
        $ui.CampoOrigem.Text = $origem
        $ui.CampoDestino.Text = $destino
        if (-not $mudou) { return $false }
        $app.Config.Origem = $origem
        $app.Config.Destino = $destino
        $app.Config.EspelharDrive = $arquivar
        [void](Save-Config)
        Update-Sobretitulo
        return $true
    }

    # ---------------------------------------------------------------- relatório

    function Get-CodigoSelo {
        param([string]$Acao)
        switch ($Acao) {
            'Adicionado'       { return 'Novo' }
            'Atualizado'       { return 'Alterado' }
            'Arquivado'        { return 'Arquivado' }
            'Saiu do servidor' { return 'SoNoDrive' }
            'Falha'            { return 'Falhou' }
        }
        return 'Neutro'
    }

    function Import-Relatorio {
        if ($null -eq $rel.Janela) { return }
        try {
            $rel.Dados = Read-Historico $app.Config
            $rel.Erro = $null
        } catch {
            $rel.Dados = @()
            $rel.Erro = Get-MensagemErro $_
        }
        $rel.Ui.TxtRelArquivo.Text = Get-ArquivoHistorico $app.Config
        $rel.Ui.TxtRelArquivo.ToolTip = $rel.Ui.TxtRelArquivo.Text
        Update-Relatorio
    }

    function Update-Relatorio {
        if ($null -eq $rel.Janela) { return }
        $r = $rel.Ui
        $hoje = (Get-Date).Date
        $desde = [datetime]::MinValue
        if ($r.RpHoje.IsChecked -eq $true) { $desde = $hoje }
        elseif ($r.Rp7.IsChecked -eq $true) { $desde = $hoje.AddDays(-6) }
        elseif ($r.Rp30.IsChecked -eq $true) { $desde = $hoje.AddDays(-29) }

        $filtro = '*'
        if ($r.RaAdicionados.IsChecked -eq $true) { $filtro = 'Adicionado' }
        elseif ($r.RaAtualizados.IsChecked -eq $true) { $filtro = 'Atualizado' }
        elseif ($r.RaSaiu.IsChecked -eq $true) { $filtro = 'Saiu do servidor' }
        elseif ($r.RaArquivados.IsChecked -eq $true) { $filtro = 'Arquivado' }
        elseif ($r.RaFalhas.IsChecked -eq $true) { $filtro = 'Falha' }
        elseif ($r.RaExecucoes.IsChecked -eq $true) { $filtro = 'Sincronização' }

        $busca = $r.CampoBusca.Text.Trim()
        $cont = @{ Adicionado = 0; Atualizado = 0; 'Saiu do servidor' = 0; Arquivado = 0; Falha = 0; 'Sincronização' = 0 }
        $linhas = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'

        foreach ($d in @($rel.Dados)) {
            if ($d.Quando -lt $desde) { continue }
            if ($cont.ContainsKey($d.Acao)) { $cont[$d.Acao]++ }
            if ($filtro -eq '*') { if ($d.Acao -eq 'Sincronização') { continue } }
            elseif ($d.Acao -ne $filtro) { continue }
            if ($busca -and $d.Arquivo.IndexOf($busca, [StringComparison]::OrdinalIgnoreCase) -lt 0 -and
                            $d.Detalhe.IndexOf($busca, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
            $tam = ''
            if ($d.Bytes -gt 0) { $tam = Format-Tamanho $d.Bytes }
            $linhas.Add([pscustomobject]@{
                Quando       = $d.Quando
                QuandoTexto  = $d.Quando.ToString('dd/MM/yyyy  HH:mm')
                Status       = Get-CodigoSelo $d.Acao
                StatusTexto  = $d.Acao
                Arquivo      = $d.Arquivo
                TamanhoTexto = $tam
                Detalhe      = $d.Detalhe
            })
        }

        $r.RelAdicionados.Text = "$($cont['Adicionado'])"
        $r.RelAtualizados.Text = "$($cont['Atualizado'])"
        $r.RelSaiu.Text        = "$($cont['Saiu do servidor'])"
        $r.RelArquivados.Text  = "$($cont['Arquivado'])"
        $r.RelFalhas.Text      = "$($cont['Falha'])"

        $rel.Itens = $linhas
        $r.ListaRel.ItemsSource = $linhas
        $r.TxtRelResumo.Text = (Get-Plural $linhas.Count 'registro' 'registros') + ' com os filtros atuais. ' +
                               (Get-Plural $cont['Sincronização'] 'sincronização' 'sincronizações') + ' no período.'

        if ($rel.Erro) {
            $r.RelVazio.Text = "Não foi possível ler o histórico.`n" + $rel.Erro
            $r.RelVazio.Visibility = 'Visible'
        } elseif (@($rel.Dados).Count -eq 0) {
            $r.RelVazio.Text = "Ainda não há registros.`nO histórico passa a ser gravado a partir desta versão, a cada sincronização."
            $r.RelVazio.Visibility = 'Visible'
        } elseif ($linhas.Count -eq 0) {
            $r.RelVazio.Text = 'Nada encontrado com esses filtros.'
            $r.RelVazio.Visibility = 'Visible'
        } else {
            $r.RelVazio.Visibility = 'Collapsed'
        }
    }

    function Export-Relatorio {
        $w = $rel.Janela
        if ($null -eq $rel.Itens -or $rel.Itens.Count -eq 0) {
            [void](Show-Dialogo $w 'Nada para exportar' 'Não há registros com os filtros atuais.' 'info')
            return
        }
        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Title = 'Exportar relatório'
        $dlg.FileName = 'Relatorio_Mapas_' + (Get-Date -Format 'yyyy-MM-dd') + '.csv'
        $dlg.Filter = 'Planilha CSV (*.csv)|*.csv'
        $dlg.DefaultExt = '.csv'
        if ($dlg.ShowDialog($w) -ne $true) { return }

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append((ConvertTo-LinhaCsv @('Data', 'Hora', 'Ação', 'Arquivo', 'Tamanho', 'Detalhe')) + "`r`n")
        foreach ($l in $rel.Itens) {
            [void]$sb.Append((ConvertTo-LinhaCsv @($l.Quando.ToString('dd/MM/yyyy'), $l.Quando.ToString('HH:mm:ss'), $l.StatusTexto, $l.Arquivo, $l.TamanhoTexto, $l.Detalhe)) + "`r`n")
        }
        [IO.File]::WriteAllText($dlg.FileName, $sb.ToString(), (New-Object Text.UTF8Encoding($true)))

        $abrir = Show-Dialogo $w 'Planilha exportada' ((Get-Plural $rel.Itens.Count 'registro exportado' 'registros exportados') + '. Deseja abrir a planilha agora?') 'info' 'Abrir planilha' 'Agora não'
        if ($abrir) { [void][System.Diagnostics.Process]::Start($dlg.FileName) }
    }

    function Set-LayoutRelatorio {
        param([double]$Largura)
        $estreito = $Largura -lt 900
        if ($estreito -eq $rel.Estreito) { return }
        $rel.Estreito = $estreito
        if ($estreito) { $rel.Ui.ListaRel.Columns[4].Visibility = 'Collapsed' }
        else { $rel.Ui.ListaRel.Columns[4].Visibility = 'Visible' }
    }

    function Open-Relatorio {
        if ($null -ne $rel.Janela) {
            if ($rel.Janela.WindowState -eq [System.Windows.WindowState]::Minimized) { $rel.Janela.WindowState = 'Normal' }
            [void]$rel.Janela.Activate()
            return
        }
        $w = New-Janela $XamlRelatorio
        $w.Owner = $janela
        $rel.Janela = $w
        $rel.Estreito = $null
        $rel.Ui = Get-Elementos $w @(
            'RpHoje','Rp7','Rp30','RpTudo',
            'RaTodas','RaAdicionados','RaAtualizados','RaSaiu','RaArquivados','RaFalhas','RaExecucoes',
            'CampoBusca','DicaBusca',
            'RelAdicionados','RelAtualizados','RelSaiu','RelArquivados','RelFalhas',
            'TxtRelResumo','ListaRel','RelVazio','TxtRelArquivo',
            'BtnAbrirHistorico','BtnAtualizarRel','BtnExportar')

        foreach ($n in 'RpHoje','Rp7','Rp30','RpTudo','RaTodas','RaAdicionados','RaAtualizados','RaSaiu','RaArquivados','RaFalhas','RaExecucoes') {
            $rel.Ui[$n].Add_Checked({ Invoke-Seguro { Update-Relatorio } })
        }
        $rel.Ui.CampoBusca.Add_TextChanged({
            Invoke-Seguro {
                if ($rel.Ui.CampoBusca.Text.Length -gt 0) { $rel.Ui.DicaBusca.Visibility = 'Collapsed' } else { $rel.Ui.DicaBusca.Visibility = 'Visible' }
                Update-Relatorio
            }
        })
        $rel.Ui.BtnAtualizarRel.Add_Click({ Invoke-Seguro { Import-Relatorio } })
        $rel.Ui.BtnExportar.Add_Click({ Invoke-Seguro { Export-Relatorio } })
        $rel.Ui.BtnAbrirHistorico.Add_Click({
            Invoke-Seguro {
                $arq = Get-ArquivoHistorico $app.Config
                Open-Pasta -Caminho ([IO.Path]::GetDirectoryName($arq)) -Selecionar $arq -Dono $rel.Janela
            }
        })
        $rel.Ui.ListaRel.Add_MouseDoubleClick({
            Invoke-Seguro {
                $item = $rel.Ui.ListaRel.SelectedItem
                if ($null -eq $item -or [string]::IsNullOrEmpty($item.Arquivo)) { return }
                $alvo = [IO.Path]::Combine($app.Config.Destino, $item.Arquivo)
                if ([IO.File]::Exists($alvo)) { [void][System.Diagnostics.Process]::Start($alvo) }
            }
        })
        $w.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq [System.Windows.Input.Key]::Escape) { $e.Handled = $true; $rel.Janela.Close() }
            elseif ($e.Key -eq [System.Windows.Input.Key]::F5) { $e.Handled = $true; Invoke-Seguro { Import-Relatorio } }
        })
        $w.Add_SourceInitialized({ Set-BarraTitulo $rel.Janela ($app.TemaEfetivo -eq 'Escuro') })
        $w.Add_SizeChanged({ Invoke-Seguro { Set-LayoutRelatorio $rel.Janela.ActualWidth } })
        $w.Add_Closed({ $rel.Janela = $null; $rel.Ui = @{}; $rel.Itens = $null })

        Import-Relatorio
        $w.Show()
    }

    # ---------------------------------------------------------------- execução em segundo plano

    function Start-Tarefa {
        param([string]$Modo)
        if ($app.Executando) { return }
        [void](Save-ConfigDaTela)

        $copiaConfig = [ordered]@{}
        foreach ($k in @($app.Config.Keys)) { $copiaConfig[$k] = $app.Config[$k] }

        $sync = [hashtable]::Synchronized(@{
            Fila     = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
            Cancelar = $false
        })
        $rs = [runspacefactory]::CreateRunspace()
        $rs.ThreadOptions = 'ReuseThread'
        $rs.Open()
        $rs.SessionStateProxy.SetVariable('Sync', $sync)
        $rs.SessionStateProxy.SetVariable('Config', $copiaConfig)
        $rs.SessionStateProxy.SetVariable('Modo', $Modo)
        $rs.SessionStateProxy.SetVariable('CaminhoEstado', $CaminhoEstado)
        $ps = [powershell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($CodigoMotor + [Environment]::NewLine + $CodigoTarefa)

        $app.Sync = $sync; $app.PS = $ps; $app.RS = $rs
        $app.Modo = $Modo; $app.Executando = $true; $app.UltimoErro = $null
        $ui.Barra.Value = 0

        if ($Modo -eq 'verificar') {
            Set-Indicador $ui.PontoOrigem $ui.StatusOrigem 'neutro' 'Verificando'
            Set-Indicador $ui.PontoDestino $ui.StatusDestino 'neutro' 'Verificando'
            Set-Status 'Verificando as pastas...' 'Lendo a pasta principal do servidor e do Google Drive.'
            if ($app.Itens.Count -eq 0) { $ui.ListaVazia.Text = 'Verificando as pastas...'; $ui.ListaVazia.Visibility = 'Visible' }
        } else {
            Set-Status 'Sincronizando...' 'Preparando a cópia.'
        }
        Update-Controles
        $app.Handle = $ps.BeginInvoke()
        $app.Timer.Start()
    }

    function Complete-Tarefa {
        $app.Timer.Stop()
        try { [void]$app.PS.EndInvoke($app.Handle) }
        catch { Set-Status 'Falha inesperada' (Get-MensagemErro $_) 'erro' }
        try { $app.PS.Dispose() } catch { }
        try { $app.RS.Close(); $app.RS.Dispose() } catch { }
        $modo = $app.Modo
        $app.PS = $null; $app.RS = $null; $app.Handle = $null
        $app.Executando = $false
        Update-Controles
        if ($modo -eq 'sincronizar' -and $null -ne $rel.Janela) { Import-Relatorio }
        if ($app.FecharAoTerminar) { $janela.Close() }
    }

    function Receive-Eventos {
        if ($null -eq $app.Sync) { return }
        $evt = $null
        $n = 0
        while ($n -lt 400 -and $app.Sync.Fila.TryDequeue([ref]$evt)) {
            Invoke-Evento $evt
            $n++
        }
        if ($null -ne $app.Handle -and $app.Handle.IsCompleted -and $app.Sync.Fila.IsEmpty) { Complete-Tarefa }
    }

    function Invoke-Evento {
        param([hashtable]$E)
        switch ($E.Tipo) {
            'ambiente' {
                Set-Ambiente $E.Ambiente
            }
            'plano' {
                Set-Plano $E.Itens
            }
            'item' {
                Set-Linha -Nome $E.Nome -Status $E.Status -Detalhe $E.Detalhe
                Update-Numeros
                if ($E.Status -eq 'Copiando') { $ui.TxtDetalhe.Text = 'Copiando ' + $E.Nome }
            }
            'progresso' {
                if ([int]$E.Total -gt 0) {
                    $ui.Barra.Value = [double]$E.Feito / [double]$E.Total
                    $ui.TxtStatus.Text = 'Sincronizando {0} de {1}' -f ([Math]::Min([int]$E.Feito + 1, [int]$E.Total)), $E.Total
                }
            }
            'log' {
                if ($E.Nivel -eq 'ERRO') { $app.UltimoErro = $E.Mensagem }
            }
            'fimVerificacao' {
                $amb = $E.Ambiente
                if (-not $amb.OrigemOk -or -not $amb.DestinoOk) {
                    $titulo = 'Destino indisponível'
                    if (-not $amb.OrigemOk) { $titulo = 'Origem indisponível' }
                    Set-Status $titulo (@($amb.Problemas) -join ' ') 'erro'
                    if (-not $amb.OrigemOk) {
                        $app.Itens.Clear(); $app.Visiveis.Clear(); $app.IndiceVisivel = @{}; $app.Indice = @{}; $app.TemDados = $false
                        $ui.ListaVazia.Text = "Não foi possível ler a pasta de origem.`nConfira a conexão com o servidor e clique em Verificar."
                        $ui.ListaVazia.Visibility = 'Visible'
                        Update-Numeros
                    }
                } else {
                    $c = Get-Contagem
                    $retirar = 0
                    if ($app.Config.EspelharDrive) { $retirar = $c.SoNoDrive }
                    $pendentes = $c.Novo + $c.Alterado + $retirar
                    if ($pendentes -gt 0) {
                        $partes = @()
                        if ($c.Novo -gt 0) { $partes += Get-Plural $c.Novo 'novo' 'novos' }
                        if ($c.Alterado -gt 0) { $partes += Get-Plural $c.Alterado 'alterado' 'alterados' }
                        if ($retirar -gt 0) { $partes += Get-Plural $retirar 'a retirar do Drive' 'a retirar do Drive' }
                        Set-Status (Get-Plural $pendentes 'arquivo para sincronizar' 'arquivos para sincronizar') (($partes -join ', ') + '. Clique em Sincronizar para deixar o Drive igual ao servidor.')
                    } else {
                        $detalhe = 'Os PDFs do servidor e do Drive são iguais.'
                        if ($c.EmDia -gt 0) { $detalhe = (Get-Plural $c.EmDia 'PDF igual' 'PDFs iguais') + ' no servidor e no Drive.' }
                        Set-Status 'Tudo em dia' $detalhe 'sucesso'
                    }
                }
            }
            'fim' {
                $r = $E.Resumo
                $app.UltimoLog = $r.LogFile
                $copiados = $r.Adicionados + $r.Atualizados
                if ($r.Codigo -eq 2) {
                    $msg = $app.UltimoErro
                    if (-not $msg) { $msg = 'Veja o log para mais detalhes.' }
                    Set-Status 'A sincronização não foi executada' $msg 'erro'
                } elseif ($r.Cancelado) {
                    Set-Status 'Sincronização interrompida' ((Get-Plural $copiados 'arquivo copiado' 'arquivos copiados') + ' antes de parar.')
                } elseif ($r.Falhas -gt 0) {
                    Set-Status ('Concluída com ' + (Get-Plural $r.Falhas 'falha' 'falhas')) ((Get-TextoResumo $r) + '. Passe o mouse sobre "Falhou" para ver o motivo.') 'erro'
                } elseif ($copiados -eq 0 -and $r.Arquivados -eq 0) {
                    Set-Status 'Sincronização concluída' 'Nenhuma alteração: o Drive já estava em dia.' 'sucesso'
                } else {
                    Set-Status 'Sincronização concluída' ((Get-TextoResumo $r) + " em $($r.Duracao) s. Veja os detalhes em Relatório.") 'sucesso'
                }
                Update-UltimaSincronizacao
            }
            'falha' {
                Set-Status 'Falha inesperada' $E.Mensagem 'erro'
            }
        }
    }

    # Se algo falhar ao atualizar a tela, o timer para antes de avisar (evita várias janelas de erro).
    $app.Timer.Add_Tick({
        try { Receive-Eventos }
        catch {
            $app.Timer.Stop()
            $msg = Get-MensagemErro $_
            Set-Status 'Falha ao atualizar a tela' $msg 'erro'
            [void](Show-Dialogo $janela 'Ocorreu um erro inesperado' $msg 'erro')
            if ($null -ne $app.Handle -and $app.Handle.IsCompleted) { try { Complete-Tarefa } catch { } }
            else { $app.Timer.Start() }
        }
    })

    # ---------------------------------------------------------------- eventos da janela principal

    $aoEscolherTema = {
        param($s, $e)
        if ($app.Carregando) { return }
        Invoke-Seguro {
            $app.Config.Tema = "$($s.Tag)"
            Set-Tema $app.Config.Tema
            [void](Save-Config)
        }
    }
    $ui.RbClaro.Add_Checked($aoEscolherTema)
    $ui.RbEscuro.Add_Checked($aoEscolherTema)
    $ui.RbSistema.Add_Checked($aoEscolherTema)

    $ui.BtnVerificar.Add_Click({ Invoke-Seguro { Start-Tarefa 'verificar' } })
    $ui.BtnRelatorio.Add_Click({ Invoke-Seguro { Open-Relatorio } })

    $ui.BtnSincronizar.Add_Click({
        Invoke-Seguro {
            if ($app.Executando) {
                if ($app.Modo -eq 'sincronizar' -and -not $app.Sync.Cancelar) {
                    $app.Sync.Cancelar = $true
                    Set-Status 'Cancelando...' 'A cópia para assim que o arquivo atual terminar.'
                    Update-Controles
                }
                return
            }
            [void](Save-ConfigDaTela)
            $pendentes = Get-Pendentes
            if ($pendentes.Count -gt 0) {
                if (-not (Show-Revisao $pendentes)) { return }
            }
            Start-Tarefa 'sincronizar'
        }
    })

    $ui.BtnLogs.Add_Click({
        Invoke-Seguro {
            if ($app.UltimoLog -and [IO.File]::Exists($app.UltimoLog)) { Open-Pasta -Caminho ([IO.Path]::GetDirectoryName($app.UltimoLog)) -Selecionar $app.UltimoLog }
            else { Open-Pasta -Caminho $app.Config.PastaLogs }
        }
    })

    $ui.BtnAbrirOrigem.Add_Click({ Invoke-Seguro { Open-Pasta -Caminho $ui.CampoOrigem.Text.Trim() } })
    $ui.BtnAbrirDestino.Add_Click({ Invoke-Seguro { Open-Pasta -Caminho $ui.CampoDestino.Text.Trim() } })

    $ui.BtnAlterarOrigem.Add_Click({
        Invoke-Seguro {
            $p = Select-Pasta -Atual $ui.CampoOrigem.Text.Trim() -Descricao 'Escolha a pasta de origem (servidor). Somente os PDFs da pasta principal serão copiados.'
            if ($p) { $ui.CampoOrigem.Text = $p; if (Save-ConfigDaTela) { Start-Tarefa 'verificar' } }
        }
    })
    $ui.BtnAlterarDestino.Add_Click({
        Invoke-Seguro {
            $p = Select-Pasta -Atual $ui.CampoDestino.Text.Trim() -Descricao 'Escolha a pasta de destino no Google Drive.'
            if ($p) { $ui.CampoDestino.Text = $p; if (Save-ConfigDaTela) { Start-Tarefa 'verificar' } }
        }
    })

    $aoSairDoCampo = { Invoke-Seguro { if (-not $app.Executando -and (Save-ConfigDaTela)) { Start-Tarefa 'verificar' } } }
    $ui.CampoOrigem.Add_LostKeyboardFocus($aoSairDoCampo)
    $ui.CampoDestino.Add_LostKeyboardFocus($aoSairDoCampo)
    $aoTeclarNoCampo = {
        param($s, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
            $e.Handled = $true
            Invoke-Seguro { if (-not $app.Executando -and (Save-ConfigDaTela)) { Start-Tarefa 'verificar' } }
        }
    }
    $ui.CampoOrigem.Add_KeyDown($aoTeclarNoCampo)
    $ui.CampoDestino.Add_KeyDown($aoTeclarNoCampo)

    $ui.Lista.Add_Sorting({
        param($s, $e)
        $e.Handled = $true
        $campo = $e.Column.SortMemberPath
        Invoke-Seguro { Set-Ordenacao $campo }
    })
    foreach ($n in 'FiltroNovos', 'FiltroAlterados', 'FiltroEmDia', 'FiltroSoNoDrive') {
        $ui[$n].Add_MouseLeftButtonUp({ param($s, $e) $tag = "$($s.Tag)"; Invoke-Seguro { Set-Filtro $tag } })
    }
    # A busca espera o usuário parar de digitar antes de filtrar.
    $ui.CampoBuscaLista.Add_TextChanged({
        if ($ui.CampoBuscaLista.Text.Length -gt 0) { $ui.DicaBuscaLista.Visibility = 'Collapsed' } else { $ui.DicaBuscaLista.Visibility = 'Visible' }
        $app.TimerBusca.Stop(); $app.TimerBusca.Start()
    })
    $app.TimerBusca.Add_Tick({ $app.TimerBusca.Stop(); Invoke-Seguro { Update-Lista } })
    $ui.CampoBuscaLista.Add_KeyDown({
        param($s, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Escape) { $e.Handled = $true; $ui.CampoBuscaLista.Text = '' }
    })

    $ui.ChkArquivar.Add_Click({ Invoke-Seguro { [void](Save-ConfigDaTela) } })

    $ui.Lista.Add_MouseDoubleClick({
        Invoke-Seguro {
            $item = $ui.Lista.SelectedItem
            if ($null -eq $item) { return }
            $alvo = $null
            if ($item.CaminhoDestino -and [IO.File]::Exists($item.CaminhoDestino)) { $alvo = $item.CaminhoDestino }
            elseif ($item.CaminhoOrigem -and [IO.File]::Exists($item.CaminhoOrigem)) { $alvo = $item.CaminhoOrigem }
            if ($alvo) { [void][System.Diagnostics.Process]::Start($alvo) }
        }
    })

    $janela.Add_KeyDown({
        param($s, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::F5 -and -not $app.Executando) {
            $e.Handled = $true
            Invoke-Seguro { Start-Tarefa 'verificar' }
        }
    })

    $janela.Add_SizeChanged({ Invoke-Seguro { Set-Layout $janela.ActualWidth } })

    # Com "Igual ao Windows", acompanha a troca de tema feita no Windows.
    $janela.Add_Activated({
        if ($app.Carregando -or $app.Config.Tema -ne 'Sistema') { return }
        Invoke-Seguro { if ((Get-TemaSistema) -ne $app.TemaEfetivo) { Set-Tema 'Sistema' } }
    })

    $janela.Add_SourceInitialized({ Set-BarraTitulo $janela ($app.TemaEfetivo -eq 'Escuro') })

    $janela.Add_Closing({
        param($s, $e)
        if (-not $app.Executando) { return }
        if ($app.Modo -eq 'sincronizar') {
            if ($app.FecharAoTerminar) { $e.Cancel = $true; return }
            $e.Cancel = $true
            $sair = Show-Dialogo $janela 'Sincronização em andamento' 'Deseja interromper e fechar? O arquivo que está sendo copiado agora termina antes de fechar.' 'aviso' 'Interromper e fechar' 'Continuar'
            if ($sair) {
                $app.FecharAoTerminar = $true
                $app.Sync.Cancelar = $true
                Set-Status 'Encerrando...' 'Aguardando o arquivo atual terminar.'
                Update-Controles
            }
        } else {
            try { $app.Timer.Stop(); [void]$app.PS.BeginStop($null, $null) } catch { }
        }
    })

    $janela.Dispatcher.Add_UnhandledException({
        param($s, $e)
        $e.Handled = $true
        [void][System.Windows.MessageBox]::Show(('Ocorreu um erro inesperado:' + "`n`n" + $e.Exception.Message), 'Sincronizador de Mapas PDF', 'OK', 'Error')
    })

    $janela.Add_Loaded({
        Invoke-Seguro {
            $ui.CampoOrigem.Text = $app.Config.Origem
            $ui.CampoDestino.Text = $app.Config.Destino
            $ui.ChkArquivar.IsChecked = [bool]$app.Config.EspelharDrive
            switch ($app.Config.Tema) {
                'Claro'  { $ui.RbClaro.IsChecked = $true }
                'Escuro' { $ui.RbEscuro.IsChecked = $true }
                default  { $ui.RbSistema.IsChecked = $true }
            }
            $app.Carregando = $false
            Update-Sobretitulo
            Update-UltimaSincronizacao
            Update-Numeros
            Set-Layout $janela.ActualWidth
            if ($avisoConfig) {
                [void](Show-Dialogo $janela 'Configuração' $avisoConfig 'aviso')
            }
            Start-Tarefa 'verificar'
        }
    })

    # Aplica o tema antes de mostrar a janela, para não piscar com as cores erradas.
    Set-Tema $app.Config.Tema
    [void]$janela.ShowDialog()
}
catch {
    Show-ErroFatal ("Não foi possível abrir o Sincronizador de Mapas PDF.`n`n" + (Get-MensagemErro $_))
    exit 1
}
