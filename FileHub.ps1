#Requires -Version 5.1
<#
.SYNOPSIS
    FileHub - Gerenciador de Downloads por Categoria

.DESCRIPTION
    Execute com o comando abaixo no PowerShell:

    irm "https://raw.githubusercontent.com/SEU_USUARIO/SEU_REPOSITORIO/main/FileHub.ps1" | iex

    ─────────────────────────────────────────────────────────────
    INSTRUCOES PARA O DESENVOLVEDOR:
    1. Preencha as 4 variáveis na secao CONFIGURACAO abaixo
    2. Suba este arquivo para o seu repositório no GitHub
    3. Compartilhe o comando acima com seus usuarios
    ─────────────────────────────────────────────────────────────

    ESTRUTURA DO REPOSITORIO:
    SEU_REPO/
    ├── FileHub.ps1           <- este arquivo
    ├── catalog.json          <- opcional: icones e descricoes
    └── plugins/
          ├── Ferramentas/    <- vira uma categoria
          │     ├── app.zip
          │     └── tool.exe
          └── Scripts/        <- vira outra categoria
                └── script.ps1

    CATALOG.JSON (opcional):
    {
      "categories": [
        { "id": "Ferramentas", "name": "Ferramentas", "icon": "🔧", "description": "Utilitarios" }
      ],
      "files": [
        { "filename": "app.zip", "title": "Meu App", "description": "Faz X e Y" }
      ]
    }
#>

# ============================================================
# CONFIGURACAO - Preencha antes de subir ao GitHub
# ============================================================
$GitHubUser    = "Clowrdl"        # Seu usuario do GitHub
$GitHubRepo    = "Hubcriativo"    # Nome do repositorio
$GitHubBranch  = "main"               # Branch (main ou master)
$PluginsFolder = "plugins"            # Pasta dos plugins no repo
# ============================================================
# Apos preencher, o comando para seus usuarios sera:
# irm "https://raw.githubusercontent.com/$GitHubUser/$GitHubRepo/$GitHubBranch/FileHub.ps1" | iex
# ============================================================

# URLs derivadas automaticamente
$RawBase    = "https://raw.githubusercontent.com/$GitHubUser/$GitHubRepo/$GitHubBranch"
$ApiBase    = "https://api.github.com/repos/$GitHubUser/$GitHubRepo/contents"
$CatalogUrl = "$RawBase/catalog.json"
$ScriptUrl  = "$RawBase/FileHub.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ── Paleta de cores ──────────────────────────────────────────
$Colors = @{
    Background    = [System.Drawing.Color]::FromArgb(15, 15, 20)
    Surface       = [System.Drawing.Color]::FromArgb(25, 25, 35)
    SurfaceHover  = [System.Drawing.Color]::FromArgb(35, 35, 50)
    Accent        = [System.Drawing.Color]::FromArgb(99, 102, 241)
    AccentHover   = [System.Drawing.Color]::FromArgb(129, 140, 248)
    AccentLight   = [System.Drawing.Color]::FromArgb(30, 30, 60)
    TextPrimary   = [System.Drawing.Color]::FromArgb(240, 240, 250)
    TextSecondary = [System.Drawing.Color]::FromArgb(148, 148, 180)
    TextMuted     = [System.Drawing.Color]::FromArgb(90, 90, 120)
    Success       = [System.Drawing.Color]::FromArgb(52, 211, 153)
    Error         = [System.Drawing.Color]::FromArgb(248, 113, 113)
    Warning       = [System.Drawing.Color]::FromArgb(251, 191, 36)
    Border        = [System.Drawing.Color]::FromArgb(45, 45, 65)
    CategoryBg    = [System.Drawing.Color]::FromArgb(20, 20, 30)
    CommandBg     = [System.Drawing.Color]::FromArgb(10, 10, 18)
}

# ── Estado global ─────────────────────────────────────────────
$Script:Catalog     = $null
$Script:SelectedCat = $null
$Script:CheckBoxes  = @{}
$Script:SavePath    = [Environment]::GetFolderPath("Desktop")
$Script:Downloading = $false

# ── Helpers ──────────────────────────────────────────────────
function Get-GitHubContents($path) {
    try {
        $req = [System.Net.WebRequest]::Create("$ApiBase/$path")
        $req.Method    = "GET"
        $req.Timeout   = 12000
        $req.UserAgent = "FileHub-PowerShell/1.0"
        $resp   = $req.GetResponse()
        $stream = $resp.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $json   = $reader.ReadToEnd()
        $reader.Close(); $resp.Close()
        return $json | ConvertFrom-Json
    } catch { return $null }
}

function Get-CatalogMeta {
    try {
        $client = New-Object System.Net.WebClient
        $client.Encoding = [System.Text.Encoding]::UTF8
        return $client.DownloadString($CatalogUrl) | ConvertFrom-Json
    } catch { return $null }
}

function Get-Catalog {
    $meta         = Get-CatalogMeta
    $rootContents = Get-GitHubContents $PluginsFolder
    if (-not $rootContents) { return $null }

    $dirs = $rootContents | Where-Object { $_.type -eq "dir" }
    if (-not $dirs) { return $null }

    $categories = @()
    foreach ($dir in $dirs) {
        $catMeta     = if ($meta -and $meta.categories) { $meta.categories | Where-Object { $_.id -eq $dir.name } } else { $null }
        $icon        = if ($catMeta -and $catMeta.icon)        { $catMeta.icon }        else { "📁" }
        $displayName = if ($catMeta -and $catMeta.name)        { $catMeta.name }        else { $dir.name }
        $desc        = if ($catMeta -and $catMeta.description) { $catMeta.description } else { "/$PluginsFolder/$($dir.name)" }

        $folderContents = Get-GitHubContents $dir.path
        $files = @()
        if ($folderContents) {
            foreach ($item in ($folderContents | Where-Object { $_.type -eq "file" })) {
                $fileMeta = if ($meta -and $meta.files) { $meta.files | Where-Object { $_.filename -eq $item.name } } else { $null }
                $files += [PSCustomObject]@{
                    title       = if ($fileMeta -and $fileMeta.title)       { $fileMeta.title }       else { [System.IO.Path]::GetFileNameWithoutExtension($item.name) }
                    filename    = $item.name
                    description = if ($fileMeta -and $fileMeta.description) { $fileMeta.description } else { "" }
                    url         = $item.download_url
                    size        = $item.size
                }
            }
        }

        $categories += [PSCustomObject]@{
            id = $dir.name; name = $displayName; icon = $icon; description = $desc; files = $files
        }
    }
    return [PSCustomObject]@{ categories = $categories }
}

function Format-FileSize($bytes) {
    if ($bytes -le 0)   { return "?" }
    if ($bytes -ge 1GB) { return "{0:N1} GB" -f ($bytes / 1GB) }
    if ($bytes -ge 1MB) { return "{0:N1} MB" -f ($bytes / 1MB) }
    if ($bytes -ge 1KB) { return "{0:N1} KB" -f ($bytes / 1KB) }
    return "$bytes B"
}

# ── Janela principal ─────────────────────────────────────────
$Form = New-Object System.Windows.Forms.Form
$Form.Text            = "FileHub"
$Form.Size            = New-Object System.Drawing.Size(1000, 720)
$Form.StartPosition   = "CenterScreen"
$Form.BackColor       = $Colors.Background
$Form.ForeColor       = $Colors.TextPrimary
$Form.FormBorderStyle = "FixedSingle"
$Form.MaximizeBox     = $false
$Form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

# ── Header ───────────────────────────────────────────────────
$PanelHeader = New-Object System.Windows.Forms.Panel
$PanelHeader.Size      = New-Object System.Drawing.Size(1000, 70)
$PanelHeader.Location  = New-Object System.Drawing.Point(0, 0)
$PanelHeader.BackColor = $Colors.Surface
$Form.Controls.Add($PanelHeader)

$LblTitle = New-Object System.Windows.Forms.Label
$LblTitle.Text      = "FileHub"
$LblTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
$LblTitle.ForeColor = $Colors.Accent
$LblTitle.Location  = New-Object System.Drawing.Point(25, 12)
$LblTitle.Size      = New-Object System.Drawing.Size(200, 45)
$PanelHeader.Controls.Add($LblTitle)

$LblRepo = New-Object System.Windows.Forms.Label
$LblRepo.Text      = "github.com/$GitHubUser/$GitHubRepo  →  /$PluginsFolder"
$LblRepo.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$LblRepo.ForeColor = $Colors.TextMuted
$LblRepo.Location  = New-Object System.Drawing.Point(27, 50)
$LblRepo.Size      = New-Object System.Drawing.Size(500, 16)
$PanelHeader.Controls.Add($LblRepo)

$BtnRefresh = New-Object System.Windows.Forms.Button
$BtnRefresh.Text      = "↻  Atualizar"
$BtnRefresh.Size      = New-Object System.Drawing.Size(130, 34)
$BtnRefresh.Location  = New-Object System.Drawing.Point(850, 18)
$BtnRefresh.BackColor = $Colors.AccentLight
$BtnRefresh.ForeColor = $Colors.AccentHover
$BtnRefresh.FlatStyle = "Flat"
$BtnRefresh.FlatAppearance.BorderColor = $Colors.Accent
$BtnRefresh.FlatAppearance.BorderSize  = 1
$BtnRefresh.Cursor    = [System.Windows.Forms.Cursors]::Hand
$PanelHeader.Controls.Add($BtnRefresh)

# ── Separador header ─────────────────────────────────────────
$Sep = New-Object System.Windows.Forms.Panel
$Sep.Size      = New-Object System.Drawing.Size(1000, 1)
$Sep.Location  = New-Object System.Drawing.Point(0, 70)
$Sep.BackColor = $Colors.Border
$Form.Controls.Add($Sep)

# ── Barra do comando (como usar) ──────────────────────────────
$PanelCommand = New-Object System.Windows.Forms.Panel
$PanelCommand.Size      = New-Object System.Drawing.Size(1000, 42)
$PanelCommand.Location  = New-Object System.Drawing.Point(0, 71)
$PanelCommand.BackColor = $Colors.CommandBg
$Form.Controls.Add($PanelCommand)

$LblPS = New-Object System.Windows.Forms.Label
$LblPS.Text      = "PS>"
$LblPS.Font      = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
$LblPS.ForeColor = $Colors.Success
$LblPS.Location  = New-Object System.Drawing.Point(18, 13)
$LblPS.Size      = New-Object System.Drawing.Size(35, 18)
$LblPS.BackColor = $Colors.CommandBg
$PanelCommand.Controls.Add($LblPS)

$CommandText = "irm `"$ScriptUrl`" | iex"
$TxtCommand = New-Object System.Windows.Forms.TextBox
$TxtCommand.Text        = $CommandText
$TxtCommand.Font        = New-Object System.Drawing.Font("Consolas", 9)
$TxtCommand.ForeColor   = $Colors.AccentHover
$TxtCommand.BackColor   = $Colors.CommandBg
$TxtCommand.Location    = New-Object System.Drawing.Point(55, 11)
$TxtCommand.Size        = New-Object System.Drawing.Size(780, 20)
$TxtCommand.BorderStyle = "None"
$TxtCommand.ReadOnly    = $true
$PanelCommand.Controls.Add($TxtCommand)

$BtnCopy = New-Object System.Windows.Forms.Button
$BtnCopy.Text      = "📋 Copiar"
$BtnCopy.Size      = New-Object System.Drawing.Size(90, 26)
$BtnCopy.Location  = New-Object System.Drawing.Point(888, 8)
$BtnCopy.BackColor = $Colors.AccentLight
$BtnCopy.ForeColor = $Colors.TextSecondary
$BtnCopy.FlatStyle = "Flat"
$BtnCopy.FlatAppearance.BorderSize = 0
$BtnCopy.Cursor    = [System.Windows.Forms.Cursors]::Hand
$BtnCopy.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText($CommandText)
    $BtnCopy.Text      = "✔ Copiado!"
    $BtnCopy.ForeColor = $Colors.Success
    $t = New-Object System.Windows.Forms.Timer
    $t.Interval = 2000
    $t.Add_Tick({ $BtnCopy.Text = "📋 Copiar"; $BtnCopy.ForeColor = $Colors.TextSecondary; $t.Stop() })
    $t.Start()
})
$PanelCommand.Controls.Add($BtnCopy)

$SepCmd = New-Object System.Windows.Forms.Panel
$SepCmd.Size      = New-Object System.Drawing.Size(1000, 1)
$SepCmd.Location  = New-Object System.Drawing.Point(0, 41)
$SepCmd.BackColor = $Colors.Border
$PanelCommand.Controls.Add($SepCmd)

# ── Painel esquerdo - Categorias ──────────────────────────────
$PanelLeft = New-Object System.Windows.Forms.Panel
$PanelLeft.Size      = New-Object System.Drawing.Size(220, 528)
$PanelLeft.Location  = New-Object System.Drawing.Point(0, 113)
$PanelLeft.BackColor = $Colors.CategoryBg
$Form.Controls.Add($PanelLeft)

$LblCatTitle = New-Object System.Windows.Forms.Label
$LblCatTitle.Text      = "CATEGORIAS"
$LblCatTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5, [System.Drawing.FontStyle]::Bold)
$LblCatTitle.ForeColor = $Colors.TextMuted
$LblCatTitle.Location  = New-Object System.Drawing.Point(18, 16)
$LblCatTitle.Size      = New-Object System.Drawing.Size(180, 18)
$PanelLeft.Controls.Add($LblCatTitle)

$FlowCategories = New-Object System.Windows.Forms.FlowLayoutPanel
$FlowCategories.Location      = New-Object System.Drawing.Point(10, 40)
$FlowCategories.Size          = New-Object System.Drawing.Size(200, 478)
$FlowCategories.FlowDirection = "TopDown"
$FlowCategories.WrapContents  = $false
$FlowCategories.BackColor     = $Colors.CategoryBg
$PanelLeft.Controls.Add($FlowCategories)

# ── Separador vertical ────────────────────────────────────────
$SepV = New-Object System.Windows.Forms.Panel
$SepV.Size      = New-Object System.Drawing.Size(1, 528)
$SepV.Location  = New-Object System.Drawing.Point(220, 113)
$SepV.BackColor = $Colors.Border
$Form.Controls.Add($SepV)

# ── Painel direito - Arquivos ─────────────────────────────────
$PanelRight = New-Object System.Windows.Forms.Panel
$PanelRight.Size      = New-Object System.Drawing.Size(779, 528)
$PanelRight.Location  = New-Object System.Drawing.Point(221, 113)
$PanelRight.BackColor = $Colors.Background
$Form.Controls.Add($PanelRight)

$LblCatName = New-Object System.Windows.Forms.Label
$LblCatName.Text      = "Selecione uma categoria"
$LblCatName.Font      = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$LblCatName.ForeColor = $Colors.TextPrimary
$LblCatName.Location  = New-Object System.Drawing.Point(20, 18)
$LblCatName.Size      = New-Object System.Drawing.Size(500, 30)
$PanelRight.Controls.Add($LblCatName)

$LblCatDesc = New-Object System.Windows.Forms.Label
$LblCatDesc.Text      = ""
$LblCatDesc.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$LblCatDesc.ForeColor = $Colors.TextSecondary
$LblCatDesc.Location  = New-Object System.Drawing.Point(22, 50)
$LblCatDesc.Size      = New-Object System.Drawing.Size(730, 18)
$PanelRight.Controls.Add($LblCatDesc)

$BtnSelectAll = New-Object System.Windows.Forms.Button
$BtnSelectAll.Text      = "Selecionar Todos"
$BtnSelectAll.Size      = New-Object System.Drawing.Size(140, 28)
$BtnSelectAll.Location  = New-Object System.Drawing.Point(490, 18)
$BtnSelectAll.BackColor = $Colors.Surface
$BtnSelectAll.ForeColor = $Colors.TextSecondary
$BtnSelectAll.FlatStyle = "Flat"
$BtnSelectAll.FlatAppearance.BorderColor = $Colors.Border
$BtnSelectAll.FlatAppearance.BorderSize  = 1
$BtnSelectAll.Cursor    = [System.Windows.Forms.Cursors]::Hand
$BtnSelectAll.Visible   = $false
$PanelRight.Controls.Add($BtnSelectAll)

$BtnDeselectAll = New-Object System.Windows.Forms.Button
$BtnDeselectAll.Text      = "Desmarcar Todos"
$BtnDeselectAll.Size      = New-Object System.Drawing.Size(140, 28)
$BtnDeselectAll.Location  = New-Object System.Drawing.Point(635, 18)
$BtnDeselectAll.BackColor = $Colors.Surface
$BtnDeselectAll.ForeColor = $Colors.TextSecondary
$BtnDeselectAll.FlatStyle = "Flat"
$BtnDeselectAll.FlatAppearance.BorderColor = $Colors.Border
$BtnDeselectAll.FlatAppearance.BorderSize  = 1
$BtnDeselectAll.Cursor    = [System.Windows.Forms.Cursors]::Hand
$BtnDeselectAll.Visible   = $false
$PanelRight.Controls.Add($BtnDeselectAll)

$PanelFiles = New-Object System.Windows.Forms.Panel
$PanelFiles.Location   = New-Object System.Drawing.Point(0, 75)
$PanelFiles.Size       = New-Object System.Drawing.Size(779, 453)
$PanelFiles.BackColor  = $Colors.Background
$PanelFiles.AutoScroll = $true
$PanelRight.Controls.Add($PanelFiles)

# ── Footer ────────────────────────────────────────────────────
$PanelFooter = New-Object System.Windows.Forms.Panel
$PanelFooter.Size      = New-Object System.Drawing.Size(1000, 79)
$PanelFooter.Location  = New-Object System.Drawing.Point(0, 641)
$PanelFooter.BackColor = $Colors.Surface
$Form.Controls.Add($PanelFooter)

$SepFooter = New-Object System.Windows.Forms.Panel
$SepFooter.Size      = New-Object System.Drawing.Size(1000, 1)
$SepFooter.Location  = New-Object System.Drawing.Point(0, 0)
$SepFooter.BackColor = $Colors.Border
$PanelFooter.Controls.Add($SepFooter)

$LblSavePath = New-Object System.Windows.Forms.Label
$LblSavePath.Text      = "Salvar em:  $($Script:SavePath)"
$LblSavePath.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5)
$LblSavePath.ForeColor = $Colors.TextSecondary
$LblSavePath.Location  = New-Object System.Drawing.Point(18, 14)
$LblSavePath.Size      = New-Object System.Drawing.Size(500, 18)
$PanelFooter.Controls.Add($LblSavePath)

$BtnChooseDir = New-Object System.Windows.Forms.Button
$BtnChooseDir.Text      = "Escolher Pasta"
$BtnChooseDir.Size      = New-Object System.Drawing.Size(130, 30)
$BtnChooseDir.Location  = New-Object System.Drawing.Point(18, 36)
$BtnChooseDir.BackColor = $Colors.Surface
$BtnChooseDir.ForeColor = $Colors.TextSecondary
$BtnChooseDir.FlatStyle = "Flat"
$BtnChooseDir.FlatAppearance.BorderColor = $Colors.Border
$BtnChooseDir.FlatAppearance.BorderSize  = 1
$BtnChooseDir.Cursor    = [System.Windows.Forms.Cursors]::Hand
$PanelFooter.Controls.Add($BtnChooseDir)

$LblStatus = New-Object System.Windows.Forms.Label
$LblStatus.Text      = ""
$LblStatus.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5)
$LblStatus.ForeColor = $Colors.TextMuted
$LblStatus.Location  = New-Object System.Drawing.Point(160, 46)
$LblStatus.Size      = New-Object System.Drawing.Size(460, 18)
$PanelFooter.Controls.Add($LblStatus)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Size      = New-Object System.Drawing.Size(460, 6)
$ProgressBar.Location  = New-Object System.Drawing.Point(160, 36)
$ProgressBar.Style     = "Continuous"
$ProgressBar.BackColor = $Colors.Border
$ProgressBar.ForeColor = $Colors.Accent
$ProgressBar.Visible   = $false
$PanelFooter.Controls.Add($ProgressBar)

$BtnDownload = New-Object System.Windows.Forms.Button
$BtnDownload.Text      = "▼  Baixar Selecionados"
$BtnDownload.Size      = New-Object System.Drawing.Size(180, 52)
$BtnDownload.Location  = New-Object System.Drawing.Point(800, 14)
$BtnDownload.BackColor = $Colors.Accent
$BtnDownload.ForeColor = [System.Drawing.Color]::White
$BtnDownload.FlatStyle = "Flat"
$BtnDownload.FlatAppearance.BorderSize = 0
$BtnDownload.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$BtnDownload.Cursor    = [System.Windows.Forms.Cursors]::Hand
$PanelFooter.Controls.Add($BtnDownload)

# ── Funcoes de UI ─────────────────────────────────────────────
function Show-Status($msg, $color = $null) {
    $LblStatus.Text     = $msg
    $LblStatus.ForeColor = if ($color) { $color } else { $Colors.TextMuted }
    $Form.Refresh()
}

function New-CategoryButton($cat) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text      = "  $($cat.icon)  $($cat.name)"
    $btn.Size      = New-Object System.Drawing.Size(190, 40)
    $btn.Margin    = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
    $btn.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $btn.BackColor = $Colors.CategoryBg
    $btn.ForeColor = $Colors.TextSecondary
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btn.Tag       = $cat.id
    $btn.Add_Click({ param($s, $e) Load-Category $s.Tag })
    $btn.Add_MouseEnter({ param($s, $e) if ($s.Tag -ne $Script:SelectedCat) { $s.BackColor = $Colors.SurfaceHover; $s.ForeColor = $Colors.TextPrimary } })
    $btn.Add_MouseLeave({ param($s, $e) if ($s.Tag -ne $Script:SelectedCat) { $s.BackColor = $Colors.CategoryBg;  $s.ForeColor = $Colors.TextSecondary } })
    return $btn
}

function Load-Categories {
    $FlowCategories.Controls.Clear()
    if (-not $Script:Catalog) { return }
    foreach ($cat in $Script:Catalog.categories) {
        $FlowCategories.Controls.Add((New-CategoryButton $cat))
    }
}

function Load-Category($catId) {
    $Script:SelectedCat = $catId
    $Script:CheckBoxes  = @{}

    foreach ($btn in $FlowCategories.Controls) {
        if ($btn.Tag -eq $catId) { $btn.BackColor = $Colors.AccentLight; $btn.ForeColor = $Colors.AccentHover }
        else                     { $btn.BackColor = $Colors.CategoryBg;  $btn.ForeColor = $Colors.TextSecondary }
    }

    $cat = $Script:Catalog.categories | Where-Object { $_.id -eq $catId }
    if (-not $cat) { return }

    $LblCatName.Text        = "$($cat.icon)  $($cat.name)"
    $LblCatDesc.Text        = $cat.description
    $BtnSelectAll.Visible   = $true
    $BtnDeselectAll.Visible = $true
    $PanelFiles.Controls.Clear()
    $y = 10

    if (-not $cat.files -or $cat.files.Count -eq 0) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text      = "Nenhum arquivo encontrado nesta categoria."
        $lbl.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
        $lbl.ForeColor = $Colors.TextMuted
        $lbl.Location  = New-Object System.Drawing.Point(20, 20)
        $lbl.Size      = New-Object System.Drawing.Size(600, 24)
        $PanelFiles.Controls.Add($lbl)
        return
    }

    foreach ($file in $cat.files) {
        $card = New-Object System.Windows.Forms.Panel
        $card.Size      = New-Object System.Drawing.Size(748, 72)
        $card.Location  = New-Object System.Drawing.Point(15, $y)
        $card.BackColor = $Colors.Surface
        $card.Cursor    = [System.Windows.Forms.Cursors]::Hand

        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Size      = New-Object System.Drawing.Size(20, 20)
        $cb.Location  = New-Object System.Drawing.Point(15, 26)
        $cb.BackColor = $Colors.Surface
        $cb.ForeColor = $Colors.Accent
        $card.Controls.Add($cb)
        $Script:CheckBoxes[$file.filename] = $cb

        $lblName = New-Object System.Windows.Forms.Label
        $lblName.Text      = $file.title
        $lblName.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $lblName.ForeColor = $Colors.TextPrimary
        $lblName.Location  = New-Object System.Drawing.Point(50, 12)
        $lblName.Size      = New-Object System.Drawing.Size(500, 22)
        $lblName.BackColor = [System.Drawing.Color]::Transparent
        $card.Controls.Add($lblName)

        $lblFilename = New-Object System.Windows.Forms.Label
        $lblFilename.Text      = $file.filename
        $lblFilename.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5)
        $lblFilename.ForeColor = $Colors.TextMuted
        $lblFilename.Location  = New-Object System.Drawing.Point(51, 33)
        $lblFilename.Size      = New-Object System.Drawing.Size(300, 16)
        $lblFilename.BackColor = [System.Drawing.Color]::Transparent
        $card.Controls.Add($lblFilename)

        $lblDesc = New-Object System.Windows.Forms.Label
        $lblDesc.Text      = $file.description
        $lblDesc.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5)
        $lblDesc.ForeColor = $Colors.TextSecondary
        $lblDesc.Location  = New-Object System.Drawing.Point(51, 50)
        $lblDesc.Size      = New-Object System.Drawing.Size(500, 16)
        $lblDesc.BackColor = [System.Drawing.Color]::Transparent
        $card.Controls.Add($lblDesc)

        $ext = [System.IO.Path]::GetExtension($file.filename).TrimStart('.').ToUpper()
        $lblExt = New-Object System.Windows.Forms.Label
        $lblExt.Text      = $ext
        $lblExt.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5, [System.Drawing.FontStyle]::Bold)
        $lblExt.ForeColor = $Colors.AccentHover
        $lblExt.Location  = New-Object System.Drawing.Point(630, 12)
        $lblExt.Size      = New-Object System.Drawing.Size(100, 16)
        $lblExt.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
        $lblExt.BackColor = [System.Drawing.Color]::Transparent
        $card.Controls.Add($lblExt)

        $lblSize = New-Object System.Windows.Forms.Label
        $lblSize.Text      = Format-FileSize $file.size
        $lblSize.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5)
        $lblSize.ForeColor = $Colors.TextMuted
        $lblSize.Location  = New-Object System.Drawing.Point(630, 30)
        $lblSize.Size      = New-Object System.Drawing.Size(100, 16)
        $lblSize.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
        $lblSize.BackColor = [System.Drawing.Color]::Transparent
        $card.Controls.Add($lblSize)

        $cbRef = $cb
        $card.Add_Click({ param($s, $e) $cbRef.Checked = -not $cbRef.Checked })
        foreach ($ctrl in $card.Controls) {
            if ($ctrl -isnot [System.Windows.Forms.CheckBox]) {
                $ctrl.Add_Click({ param($s, $e) $cbRef.Checked = -not $cbRef.Checked })
            }
        }
        $card.Add_MouseEnter({ param($s, $e) $s.BackColor = $Colors.SurfaceHover })
        $card.Add_MouseLeave({ param($s, $e) $s.BackColor = $Colors.Surface })

        $PanelFiles.Controls.Add($card)
        $y += 82
    }
    $PanelFiles.AutoScrollMinSize = New-Object System.Drawing.Size(0, $y)
}

function Start-Downloads {
    if ($Script:Downloading) { return }

    $cat = $Script:Catalog.categories | Where-Object { $_.id -eq $Script:SelectedCat }
    if (-not $cat) { Show-Status "Selecione uma categoria primeiro." $Colors.Warning; return }

    $toDownload = @()
    foreach ($file in $cat.files) {
        if ($Script:CheckBoxes.ContainsKey($file.filename) -and $Script:CheckBoxes[$file.filename].Checked) {
            $toDownload += $file
        }
    }

    if ($toDownload.Count -eq 0) { Show-Status "Nenhum arquivo selecionado." $Colors.Warning; return }

    $Script:Downloading   = $true
    $BtnDownload.Enabled  = $false
    $BtnDownload.Text     = "Baixando..."
    $ProgressBar.Visible  = $true
    $ProgressBar.Maximum  = $toDownload.Count
    $ProgressBar.Value    = 0
    $success = 0; $failed = 0

    for ($i = 0; $i -lt $toDownload.Count; $i++) {
        $file = $toDownload[$i]
        Show-Status ("Baixando {0}/{1}: {2}" -f ($i + 1), $toDownload.Count, $file.filename) $Colors.TextSecondary
        $ProgressBar.Value = $i
        $destPath = Join-Path $Script:SavePath $file.filename
        try {
            $client = New-Object System.Net.WebClient
            $client.Headers.Add("User-Agent", "FileHub-PowerShell/1.0")
            $client.DownloadFile($file.url, $destPath)
            $success++
        } catch {
            $failed++
            [System.Windows.Forms.MessageBox]::Show(
                "Erro ao baixar '$($file.filename)':`n$($_.Exception.Message)",
                "Erro de Download",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
        $ProgressBar.Value = $i + 1
        $Form.Refresh()
    }

    $Script:Downloading   = $false
    $BtnDownload.Enabled  = $true
    $BtnDownload.Text     = "▼  Baixar Selecionados"

    $msg = "Concluido: $success arquivo(s) baixado(s)"
    if ($failed -gt 0) { $msg += " | $failed falhou(ram)" }
    Show-Status $msg $Colors.Success

    if ($success -gt 0) {
        $open = [System.Windows.Forms.MessageBox]::Show(
            "$success arquivo(s) baixado(s) com sucesso!`nDeseja abrir a pasta de destino?",
            "Download Concluido",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        if ($open -eq [System.Windows.Forms.DialogResult]::Yes) { Start-Process explorer.exe $Script:SavePath }
    }

    Start-Sleep -Seconds 3
    $ProgressBar.Visible = $false
    Show-Status ""
}

# ── Eventos ───────────────────────────────────────────────────
$BtnRefresh.Add_Click({
    Show-Status "Detectando plugins em /$PluginsFolder ..." $Colors.TextSecondary
    $FlowCategories.Controls.Clear()
    $PanelFiles.Controls.Clear()
    $LblCatName.Text = "Aguarde..."
    $Script:Catalog  = Get-Catalog
    if ($Script:Catalog) {
        Load-Categories
        $LblCatName.Text        = "Selecione uma categoria"
        $LblCatDesc.Text        = ""
        $BtnSelectAll.Visible   = $false
        $BtnDeselectAll.Visible = $false
        $total = ($Script:Catalog.categories | ForEach-Object { $_.files.Count } | Measure-Object -Sum).Sum
        Show-Status "$($Script:Catalog.categories.Count) categoria(s)  |  $total arquivo(s) detectado(s)" $Colors.Success
    } else {
        $LblCatName.Text = "Erro ao carregar"
        Show-Status "Falha ao acessar o repositorio. Verifique usuario/repo e conexao." $Colors.Error
    }
})

$BtnChooseDir.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description  = "Escolha a pasta para salvar os arquivos"
    $dlg.SelectedPath = $Script:SavePath
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Script:SavePath  = $dlg.SelectedPath
        $LblSavePath.Text = "Salvar em:  $($Script:SavePath)"
    }
})

$BtnSelectAll.Add_Click({   foreach ($cb in $Script:CheckBoxes.Values) { $cb.Checked = $true  } })
$BtnDeselectAll.Add_Click({ foreach ($cb in $Script:CheckBoxes.Values) { $cb.Checked = $false } })
$BtnDownload.Add_Click({ Start-Downloads })

# ── Carregamento inicial ──────────────────────────────────────
$Form.Add_Shown({
    Show-Status "Conectando ao GitHub e detectando plugins..." $Colors.TextSecondary
    $Script:Catalog = Get-Catalog
    if ($Script:Catalog) {
        Load-Categories
        $total = ($Script:Catalog.categories | ForEach-Object { $_.files.Count } | Measure-Object -Sum).Sum
        Show-Status "$($Script:Catalog.categories.Count) categoria(s)  |  $total arquivo(s) detectado(s)" $Colors.Success
    } else {
        Show-Status "ERRO: Nao foi possivel conectar ao repositorio." $Colors.Error
        [System.Windows.Forms.MessageBox]::Show(
            "Nao foi possivel carregar os plugins do GitHub.`n`nVerifique:`n- $GitHubUser/$GitHubRepo existe e e publico`n- A pasta '$PluginsFolder' existe no repositorio`n- Sua conexao com a internet",
            "Erro ao Carregar",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

[System.Windows.Forms.Application]::Run($Form)