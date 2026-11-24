<#
.SYNOPSIS
    Cria um ambiente de demonstração do Sincronizador de Mapas PDF, só com dados fictícios.

.DESCRIPTION
    Serve para testar a ferramenta e tirar capturas de tela sem expor caminhos,
    nomes de servidores ou arquivos reais. Cria, dentro da pasta escolhida:

        Servidor\MAPAS PDF\          "servidor" com PDFs fictícios (e uma subpasta que deve ser ignorada)
        Meu Drive\MAPAS\MAPAS PDF\   "Google Drive" com PDFs em todas as situações
        Meu Drive\MAPAS\_LOGS\       histórico de exemplo dos últimos dias
        App\                         cópia da ferramenta já configurada para estas pastas

    Nada fora desta pasta é lido ou alterado.

.PARAMETER Pasta
    Onde criar a demonstração. Padrão: C:\SincronizadorDemo

.PARAMETER Recriar
    Apaga a demonstração existente e cria de novo. Só apaga pastas criadas por este script.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\ferramentas\preparar-demo.ps1
#>
[CmdletBinding()]
param(
    [string]$Pasta = 'C:\SincronizadorDemo',
    [switch]$Recriar
)

$ErrorActionPreference = 'Stop'
$inv = [Globalization.CultureInfo]::InvariantCulture
$utf8Bom = New-Object Text.UTF8Encoding($true)

$origemApp = Join-Path (Split-Path -Parent $PSScriptRoot) 'src'
if (-not (Test-Path -LiteralPath (Join-Path $origemApp 'SincronizadorMapas.ps1'))) {
    throw 'Não encontrei src\SincronizadorMapas.ps1. Rode este script de dentro da pasta "ferramentas" do repositório.'
}

$marcador = Join-Path $Pasta '.demo-sincronizador'
if (Test-Path -LiteralPath $Pasta) {
    if (-not $Recriar) {
        Write-Host "A pasta $Pasta já existe. Use -Recriar para apagar e criar de novo." -ForegroundColor Yellow
        exit 1
    }
    if (-not (Test-Path -LiteralPath $marcador)) {
        throw "A pasta $Pasta não foi criada por este script. Por segurança, nada foi apagado."
    }
    Remove-Item -LiteralPath $Pasta -Recurse -Force
}

$servidor  = Join-Path $Pasta 'Servidor\MAPAS PDF'
$drive     = Join-Path $Pasta 'Meu Drive\MAPAS'
$destino   = Join-Path $drive 'MAPAS PDF'
$logs      = Join-Path $drive '_LOGS'
$arquivo   = Join-Path $drive '_ARQUIVADOS'
$app       = Join-Path $Pasta 'App'

foreach ($p in @($servidor, (Join-Path $servidor 'Antigos'), $destino, $logs, $app)) {
    [void][IO.Directory]::CreateDirectory($p)
}
[IO.File]::WriteAllText($marcador, 'Pasta criada por ferramentas\preparar-demo.ps1', $utf8Bom)

# ------------------------------------------------------------------ PDFs fictícios

function New-PdfFicticio {
    param([string]$Caminho, [string]$Titulo, [int]$Kb, [datetime]$Data)
    # Texto do PDF só em ASCII, para não depender de codificação de fonte.
    $texto = $Titulo.Normalize([Text.NormalizationForm]::FormD) -replace '[^\x20-\x7E]', '' -replace '[()\\]', ''
    $conteudo = "BT /F1 20 Tf 72 720 Td (Mapa ficticio para demonstracao) Tj 0 -30 Td /F1 14 Tf ($texto) Tj ET"
    $objetos = @(
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
        "<< /Length $($conteudo.Length) >>`nstream`n$conteudo`nendstream",
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
    )
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("%PDF-1.4`n")
    # Comentários para o arquivo ter um tamanho realista.
    $linha = '%' + ('0' * 99) + "`n"
    for ($i = 0; $i -lt ($Kb * 10); $i++) { [void]$sb.Append($linha) }
    $posicoes = @()
    for ($i = 0; $i -lt $objetos.Count; $i++) {
        $posicoes += $sb.Length
        [void]$sb.Append("$($i + 1) 0 obj`n$($objetos[$i])`nendobj`n")
    }
    $xref = $sb.Length
    [void]$sb.Append("xref`n0 $($objetos.Count + 1)`n0000000000 65535 f `n")
    foreach ($p in $posicoes) { [void]$sb.Append(('{0:D10} 00000 n ' -f $p) + "`n") }
    [void]$sb.Append("trailer`n<< /Size $($objetos.Count + 1) /Root 1 0 R >>`nstartxref`n$xref`n%%EOF`n")
    [IO.File]::WriteAllText($Caminho, $sb.ToString(), [Text.Encoding]::ASCII)
    [IO.File]::SetLastWriteTime($Caminho, $Data)
}

$agora = Get-Date
function Get-Dia { param([int]$Dias, [int]$Hora, [int]$Minuto) return $agora.Date.AddDays(-$Dias).AddHours($Hora).AddMinutes($Minuto) }

# Em dia: iguais no servidor e no Drive
$emDia = @(
    @('Mapa Boa Vista Talhão 01.pdf', 420, 40), @('Mapa Boa Vista Talhão 02.pdf', 388, 40),
    @('Mapa Boa Vista Talhão 03.pdf', 512, 38), @('Mapa Santa Clara Talhão 01.pdf', 296, 35),
    @('Mapa Santa Clara Talhão 02.pdf', 344, 35), @('Mapa Três Lagoas Talhão 01.pdf', 610, 30),
    @('Mapa Três Lagoas Talhão 02.pdf', 575, 30), @('Mapa Geral Safra Demonstração.pdf', 1240, 20)
)
foreach ($a in $emDia) {
    $data = Get-Dia $a[2] 9 15
    New-PdfFicticio (Join-Path $servidor $a[0]) $a[0] $a[1] $data
    Copy-Item -LiteralPath (Join-Path $servidor $a[0]) -Destination (Join-Path $destino $a[0])
}

# Alterados: versão nova no servidor, versão antiga no Drive
$alterados = @(@('Mapa Santa Clara Talhão 03.pdf', 402, 455), @('Mapa Três Lagoas Talhão 03.pdf', 530, 548))
foreach ($a in $alterados) {
    New-PdfFicticio (Join-Path $destino $a[0]) $a[0] $a[1] (Get-Dia 12 10 0)
    New-PdfFicticio (Join-Path $servidor $a[0]) $a[0] $a[2] (Get-Dia 0 7 40)
}

# Novos: só no servidor
foreach ($a in @(@('Mapa Boa Vista Talhão 04.pdf', 466), @('Mapa Boa Vista Talhão 05.pdf', 433), @('Mapa Santa Clara Talhão 04.pdf', 389))) {
    New-PdfFicticio (Join-Path $servidor $a[0]) $a[0] $a[1] (Get-Dia 0 8 5)
}

# Só no Drive: saíram do servidor
New-PdfFicticio (Join-Path $destino 'Mapa Safra Anterior Talhão 07.pdf') 'Mapa Safra Anterior Talhão 07' 350 (Get-Dia 40 11 0)
New-PdfFicticio (Join-Path $destino 'Mapa Safra Anterior Talhão 08.pdf') 'Mapa Safra Anterior Talhão 08' 362 (Get-Dia 40 11 0)

# Devem ser ignorados: subpasta e arquivo que não é PDF
New-PdfFicticio (Join-Path $servidor 'Antigos\Mapa Antigo 2019.pdf') 'Mapa Antigo 2019' 300 (Get-Dia 900 9 0)
New-PdfFicticio (Join-Path $servidor 'Antigos\Mapa Antigo 2020.pdf') 'Mapa Antigo 2020' 300 (Get-Dia 700 9 0)
[IO.File]::WriteAllText((Join-Path $servidor 'Observações.txt'), 'Arquivo de texto fictício. Não deve ser copiado.', $utf8Bom)

# ------------------------------------------------------------------ histórico de exemplo

function ConvertTo-LinhaCsv {
    param([object[]]$Campos)
    $saida = foreach ($c in $Campos) { '"' + ("$c" -replace '"', '""') + '"' }
    return ($saida -join ';')
}

$linhas = New-Object 'System.Collections.Generic.List[string]'
$linhas.Add('"Data e hora";"Execução";"Ação";"Arquivo";"Bytes";"Detalhe"')
function Add-Registro {
    param([datetime]$Execucao, [int]$Segundos, [string]$Acao, [string]$Arquivo, [long]$Bytes, [string]$Detalhe)
    $quando = $Execucao.AddSeconds($Segundos).ToString('yyyy-MM-dd HH:mm:ss', $inv)
    $tam = ''
    if ($Bytes -gt 0) { $tam = "$Bytes" }
    $linhas.Add((ConvertTo-LinhaCsv @($quando, $Execucao.ToString('yyyy-MM-dd HH:mm:ss', $inv), $Acao, $Arquivo, $tam, $Detalhe)))
}
function Get-Versao { param([datetime]$D) return 'Versão do servidor de ' + $D.ToString('dd/MM/yyyy HH:mm', $inv) + '.' }

$e1 = Get-Dia 6 8 10
Add-Registro $e1 2 'Adicionado' 'Mapa Santa Clara Talhão 01.pdf' 303104 (Get-Versao (Get-Dia 7 16 20))
Add-Registro $e1 3 'Adicionado' 'Mapa Santa Clara Talhão 02.pdf' 352256 (Get-Versao (Get-Dia 7 16 22))
Add-Registro $e1 5 'Adicionado' 'Mapa Três Lagoas Talhão 01.pdf' 624640 (Get-Versao (Get-Dia 7 17 5))
Add-Registro $e1 6 'Sincronização' '' 0 'Automática: 3 adicionados.'

$e2 = Get-Dia 4 14 32
Add-Registro $e2 1 'Saiu do servidor' 'Mapa Safra Anterior Talhão 07.pdf' 358400 'Não existe mais na pasta principal do servidor.'
Add-Registro $e2 3 'Atualizado' 'Mapa Boa Vista Talhão 02.pdf' 397312 (Get-Versao (Get-Dia 4 13 50))
Add-Registro $e2 4 'Atualizado' 'Mapa Geral Safra Demonstração.pdf' 1269760 (Get-Versao (Get-Dia 4 11 12))
Add-Registro $e2 5 'Sincronização' '' 0 'Manual: 2 atualizados.'

$e3 = Get-Dia 2 8 10
Add-Registro $e3 2 'Adicionado' 'Mapa Três Lagoas Talhão 02.pdf' 588800 (Get-Versao (Get-Dia 3 15 40))
Add-Registro $e3 19 'Falha' 'Mapa Boa Vista Talhão 03.pdf' 524288 'Não foi copiado. O arquivo está sendo usado por outro processo.'
Add-Registro $e3 20 'Sincronização' '' 0 'Automática: 1 adicionado, 1 falha.'

$e4 = Get-Dia 1 8 10
Add-Registro $e4 1 'Atualizado' 'Mapa Boa Vista Talhão 03.pdf' 524288 (Get-Versao (Get-Dia 2 17 30))
Add-Registro $e4 2 'Sincronização' '' 0 'Automática: 1 atualizado.'

[IO.File]::WriteAllText((Join-Path $logs 'historico.csv'), (($linhas -join "`r`n") + "`r`n"), $utf8Bom)

# ------------------------------------------------------------------ cópia da ferramenta configurada

Get-ChildItem -LiteralPath $origemApp -File | Where-Object { $_.Name -notin @('config.json', 'estado.json') } |
    ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $app $_.Name) }

$config = [ordered]@{
    Origem            = $servidor
    Destino           = $destino
    PastaLogs         = $logs
    PastaArquivados   = $arquivo
    EspelharDrive     = $true
    Tema              = 'Sistema'
    DiasManterLogs    = 90
    Tentativas        = 3
    EsperaSegundos    = 5
}
[IO.File]::WriteAllText((Join-Path $app 'config.json'), ($config | ConvertTo-Json), $utf8Bom)

$ultima = $e4.AddSeconds(2)
$estado = [ordered]@{
    UltimaSincronizacao = $ultima.ToString('yyyy-MM-ddTHH:mm:ss', $inv)
    Texto               = $ultima.ToString('dd/MM/yyyy', $inv) + ' às ' + $ultima.ToString('HH:mm', $inv)
    Adicionados         = 0
    Atualizados         = 1
    Arquivados          = 0
    Falhas              = 0
    Codigo              = 0
    SoNoDrive           = @('Mapa Safra Anterior Talhão 07.pdf')
}
[IO.File]::WriteAllText((Join-Path $app 'estado.json'), ($estado | ConvertTo-Json), $utf8Bom)

Write-Host ''
Write-Host 'Demonstração criada em ' -NoNewline; Write-Host $Pasta -ForegroundColor Green
Write-Host ''
Write-Host '  Servidor fictício : ' $servidor
Write-Host '  Drive fictício    : ' $destino
Write-Host '  Ferramenta        : ' $app
Write-Host ''
Write-Host 'Situação inicial: 8 em dia, 2 alterados, 3 novos e 2 só no Drive.'
Write-Host 'A subpasta "Antigos" e o arquivo .txt devem ser ignorados.'
Write-Host ''
Write-Host 'Para abrir:  ' -NoNewline; Write-Host (Join-Path $app 'Abrir Sincronizador.bat') -ForegroundColor Cyan
Write-Host 'Para recomeçar do zero, rode este script de novo com -Recriar.'
