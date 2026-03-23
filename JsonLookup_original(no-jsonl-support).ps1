#!/usr/bin/env pwsh
#Requires -Version 7.0
<#
.SYNOPSIS
    Generic JSON Lookup Tool - GUI front-end for searching and editing JSON files.
.DESCRIPTION
    Provides a searchable, editable GUI for any JSON file. Sections are discovered
    automatically. Checkbox state, window geometry, splitter ratio all persist per-file.
.PARAMETER Path
    Path to a JSON file. If omitted, loads default.json from the working directory,
    or opens a file picker 1 second after the form appears.
.EXAMPLE
    # From a shortcut (no console at all — the script relaunches itself hidden):
    pwsh -ExecutionPolicy Bypass -File "E:\JsonLookup\JsonLookup.ps1"
    pwsh -ExecutionPolicy Bypass -File "E:\JsonLookup\JsonLookup.ps1" -Path "E:\data.json"
    # Or from a console (also works):
    ./JsonLookup.ps1 -Path ./data.json
.AUTHOR
    (c) 2026 @drgfragkos
#>
param(
    [string]$Path,
    [switch]$_Hidden   # Internal flag: set by the self-relaunch; do not use manually
)

# ══════════════════════════════════════════════════════════════════════════════
#   SELF-RELAUNCH HIDDEN — zero console flash
#   If not already relaunched, start a new hidden pwsh process and exit this one.
#   The shortcut target is simply:
#       pwsh -ExecutionPolicy Bypass -File "E:\JsonLookup\JsonLookup.ps1"
#   No -WindowStyle Hidden needed on the shortcut; the script handles it.
# ══════════════════════════════════════════════════════════════════════════════
if (-not $_Hidden) {
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden',
                 '-File', "`"$PSCommandPath`"", '-_Hidden')
    if ($Path) { $argList += '-Path'; $argList += "`"$Path`"" }
    Start-Process -FilePath 'pwsh' -ArgumentList $argList -WindowStyle Hidden -WorkingDirectory (Get-Location).Path
    exit
}

# ── Native helpers (P/Invoke) ─────────────────────────────────────────────────
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class WinHelper {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
    public static extern int SetWindowTheme(IntPtr hwnd, string appName, string idList);
    public static void HideConsole() {
        IntPtr h = GetConsoleWindow();
        if (h != IntPtr.Zero) ShowWindow(h, 0);
    }
    public static void EnableDarkTitleBar(IntPtr hwnd) {
        int val = 1;
        DwmSetWindowAttribute(hwnd, 20, ref val, sizeof(int));
    }
    public static void ApplyDarkScrollbars(IntPtr hwnd) {
        SetWindowTheme(hwnd, "DarkMode_Explorer", null);
    }
}
"@ -ErrorAction SilentlyContinue

# Belt-and-suspenders: hide any residual console window in the relaunched process
try { [WinHelper]::HideConsole() } catch {}

# ── Assemblies ────────────────────────────────────────────────────────────────
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ── Theme ─────────────────────────────────────────────────────────────────────
$script:FontFamily     = "Calibri"
$script:BaseFontSize   = 11
$script:Font           = [System.Drawing.Font]::new($script:FontFamily, $script:BaseFontSize)
$script:FontBold       = [System.Drawing.Font]::new($script:FontFamily, $script:BaseFontSize, [System.Drawing.FontStyle]::Bold)
$script:TextBoxFontSz  = $script:BaseFontSize + 1   # text boxes are +1
$script:DetailFontSize = $script:TextBoxFontSz
$script:MinDetailFont  = 6
$script:MaxDetailFont  = 30

$script:BgDark     = [System.Drawing.Color]::FromArgb(32, 32, 32)
$script:BgMid      = [System.Drawing.Color]::FromArgb(43, 43, 43)
$script:BgInput    = [System.Drawing.Color]::FromArgb(51, 51, 51)
$script:BgHover    = [System.Drawing.Color]::FromArgb(62, 62, 62)
$script:FgPrimary  = [System.Drawing.Color]::FromArgb(230, 230, 230)
$script:FgSecondary= [System.Drawing.Color]::FromArgb(160, 160, 160)
$script:FgDim      = [System.Drawing.Color]::FromArgb(110, 110, 115)
$script:Accent     = [System.Drawing.Color]::FromArgb(76, 194, 255)
$script:AccentDim  = [System.Drawing.Color]::FromArgb(0, 103, 163)
$script:Border     = [System.Drawing.Color]::FromArgb(60, 60, 60)

# Checkbox adaptive sizing
$script:ChkPreferredWidth = 150
$script:ChkMinWidth       = 50

# ── State ─────────────────────────────────────────────────────────────────────
$script:JsonPath        = $null
$script:JsonData        = $null
$script:Sections        = $null
$script:ScalarKeys      = @()
$script:Checkboxes      = [ordered]@{}
$script:CurrentResults  = @()
$script:SelectedSection = $null
$script:SelectedIndex   = -1
$script:DetailDirty     = $false
$script:SavedState      = @{}
$script:StatePath       = $null
$script:IsRootArray     = $false  # true if JSON root is an array, not an object
$script:SelectAllActive = $false
$script:PreSelectAllState = @{}   # stores checkbox state before Select All
$script:UserSplitterRatio = $null  # tracks user's manual splitter position
$script:RevealActive = $false
$script:PreRevealDetailText = ""
$script:PreRevealDetailLabel = ""
$script:PreRevealReadOnly = $true

# ── JSON helpers ──────────────────────────────────────────────────────────────
function Load-JsonData {
    $raw = Get-Content -Path $script:JsonPath -Raw -Encoding utf8
    return ($raw | ConvertFrom-Json -AsHashtable -Depth 100)
}

function Save-JsonData {
    param($Data)
    $json = $Data | ConvertTo-Json -Depth 100 -EnumsAsStrings
    [System.IO.File]::WriteAllText($script:JsonPath, $json, [System.Text.Encoding]::UTF8)
}

function Load-ProjectState {
    if ($script:StatePath -and (Test-Path $script:StatePath)) {
        # Ensure hidden attribute on load
        try {
            $fi = [System.IO.FileInfo]::new($script:StatePath)
            if (-not ($fi.Attributes -band [System.IO.FileAttributes]::Hidden)) {
                $fi.Attributes = $fi.Attributes -bor [System.IO.FileAttributes]::Hidden
            }
        } catch {}
        try {
            return (Get-Content $script:StatePath -Raw | ConvertFrom-Json -AsHashtable)
        } catch { return @{} }
    }
    return @{}
}

function Save-ProjectState {
    param($State)
    if ($script:StatePath) {
        ($State | ConvertTo-Json -Depth 5) | Set-Content -Path $script:StatePath -Encoding utf8
        # Ensure hidden attribute on save
        try {
            $fi = [System.IO.FileInfo]::new($script:StatePath)
            if (-not ($fi.Attributes -band [System.IO.FileAttributes]::Hidden)) {
                $fi.Attributes = $fi.Attributes -bor [System.IO.FileAttributes]::Hidden
            }
        } catch {}
    }
}

function Get-EntryLabel {
    param($Entry, $Section, $Index)
    if ($Entry -is [System.Collections.IDictionary]) {
        foreach ($f in @('name','title','id','label','key','description','Name','Title','Id','hostname','host','subject')) {
            if ($Entry.Contains($f) -and $null -ne $Entry[$f]) {
                $v = "$($Entry[$f])"
                if ($v.Length -gt 120) { $v = $v.Substring(0,117) + "..." }
                return $v
            }
        }
        foreach ($k in $Entry.Keys) {
            $v = $Entry[$k]
            if ($v -is [string] -and $v.Length -gt 0) {
                $lab = "${k}: $v"
                if ($lab.Length -gt 120) { $lab = $lab.Substring(0,117) + "..." }
                return $lab
            }
        }
        return "$Section [$Index]"
    }
    return "$Entry"
}

function Search-Entries {
    param([string]$Query, [string[]]$ActiveSections)
    $results = [System.Collections.ArrayList]::new()
    $q = $Query.Trim().ToLowerInvariant()

    foreach ($sec in $ActiveSections) {
        if ($null -eq $script:Sections -or -not ($script:Sections.Contains($sec))) { continue }
        $entries = $script:Sections[$sec]
        for ($i = 0; $i -lt $entries.Count; $i++) {
            $entry = $entries[$i]
            $match = $false
            if ([string]::IsNullOrEmpty($q)) {
                $match = $true
            }
            elseif ($entry -is [System.Collections.IDictionary]) {
                foreach ($k in $entry.Keys) {
                    $val = $entry[$k]
                    if ($null -ne $val -and "$val".ToLowerInvariant().Contains($q)) {
                        $match = $true; break
                    }
                }
            }
            else {
                if ("$entry".ToLowerInvariant().Contains($q)) { $match = $true }
            }
            if ($match) {
                [void]$results.Add(@{
                    Section = $sec
                    Index   = $i
                    Entry   = $entry
                    Label   = Get-EntryLabel $entry $sec $i
                })
            }
        }
    }
    return $results
}

# Helper: collect all state to save
function Build-FullState {
    $st = @{}
    # Checkboxes - always save the "real" selection, not the Select All override
    if ($script:SelectAllActive -and $script:PreSelectAllState.Count -gt 0) {
        # Save the pre-Select-All state (the user's custom selection)
        foreach ($sec in $script:PreSelectAllState.Keys) {
            $st[$sec] = $script:PreSelectAllState[$sec]
        }
    }
    else {
        foreach ($sec in $script:Checkboxes.Keys) {
            $st[$sec] = $script:Checkboxes[$sec].Checked
        }
    }
    # Splitter ratio
    $totalW = $split.ClientSize.Width - $split.SplitterWidth
    if ($totalW -gt 0) {
        $st["__splitterRatio"] = [Math]::Round($split.SplitterDistance / $totalW, 4)
    }
    # Window geometry (use RestoreBounds for maximized/normal)
    $bounds = if ($form.WindowState -eq "Normal") { $form.Bounds } else { $form.RestoreBounds }
    $st["__windowX"]      = $bounds.X
    $st["__windowY"]      = $bounds.Y
    $st["__windowWidth"]  = $bounds.Width
    $st["__windowHeight"] = $bounds.Height
    $st["__windowState"]  = "$($form.WindowState)"
    return $st
}

# ══════════════════════════════════════════════════════════════════════════════
#   BUILD FORM
# ══════════════════════════════════════════════════════════════════════════════
$form = [System.Windows.Forms.Form]::new()
$form.Text          = "JSON Lookup"
$form.StartPosition = "CenterScreen"
$form.BackColor     = $script:BgDark
$form.ForeColor     = $script:FgPrimary
$form.Font          = $script:Font
$form.KeyPreview    = $true

# Default size: 80% of primary screen
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Size        = [System.Drawing.Size]::new([int]($screen.Width * 0.8), [int]($screen.Height * 0.8))
$form.MinimumSize = [System.Drawing.Size]::new(640, 400)

# ── Application icon (generated programmatically) ────────────────────────────
# Draws a magnifying glass search icon with accent colors on a dark background
function New-AppIcon {
    $sz  = 32
    $bmp = [System.Drawing.Bitmap]::new($sz, $sz, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = "AntiAlias"
    $g.Clear([System.Drawing.Color]::FromArgb(32, 32, 32))

    # Lens circle
    $penLens = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(76, 194, 255), 2.6)
    $g.DrawEllipse($penLens, 4, 4, 17, 17)
    $penLens.Dispose()

    # Fill the lens with a subtle tint
    $fillBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(30, 76, 194, 255))
    $g.FillEllipse($fillBrush, 5, 5, 15, 15)
    $fillBrush.Dispose()

    # Handle
    $penHandle = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(200, 200, 200), 3)
    $penHandle.StartCap = "Round"
    $penHandle.EndCap   = "Round"
    $g.DrawLine($penHandle, 19, 19, 27, 27)
    $penHandle.Dispose()

    # JSON curly brace inside the lens
    $braceFont  = [System.Drawing.Font]::new("Consolas", 9, [System.Drawing.FontStyle]::Bold)
    $braceBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(200, 76, 194, 255))
    $sf = [System.Drawing.StringFormat]::new()
    $sf.Alignment     = "Center"
    $sf.LineAlignment = "Center"
    $lensRect = [System.Drawing.RectangleF]::new(4, 3, 17, 19)
    $g.DrawString("{}", $braceFont, $braceBrush, $lensRect, $sf)
    $braceBrush.Dispose()
    $braceFont.Dispose()
    $sf.Dispose()

    $g.Dispose()

    # Convert to icon
    $hIcon = $bmp.GetHicon()
    $icon  = [System.Drawing.Icon]::FromHandle($hIcon)
    return $icon
}

$form.Icon = New-AppIcon

$form.Add_HandleCreated({
    try { [WinHelper]::EnableDarkTitleBar($form.Handle) } catch {}
})

# ── Top panel (search + checkboxes) ──────────────────────────────────────────
$topPanel = [System.Windows.Forms.Panel]::new()
$topPanel.Dock      = "Top"
$topPanel.BackColor = $script:BgMid
$topPanel.Height    = 82

$txtSearch = [System.Windows.Forms.TextBox]::new()
$txtSearch.Location    = [System.Drawing.Point]::new(12, 10)
$txtSearch.Height      = 28
$txtSearch.Font        = [System.Drawing.Font]::new($script:FontFamily, $script:TextBoxFontSz)
$txtSearch.BackColor   = $script:BgInput
$txtSearch.ForeColor   = $script:FgPrimary
$txtSearch.BorderStyle = "FixedSingle"
$txtSearch.Enabled     = $false
$topPanel.Controls.Add($txtSearch)

# Panel for checkboxes + Select button
$chkRow = [System.Windows.Forms.Panel]::new()
$chkRow.Location  = [System.Drawing.Point]::new(12, 44)
$chkRow.Height    = 28
$chkRow.BackColor = $script:BgMid
$topPanel.Controls.Add($chkRow)

# "Select" toggle button - placed to the LEFT of checkboxes
$btnSelectAll = [System.Windows.Forms.Button]::new()
$btnSelectAll.Text      = "Select"
$btnSelectAll.Size      = [System.Drawing.Size]::new(60, 22)
$btnSelectAll.Location  = [System.Drawing.Point]::new(0, 0)
$btnSelectAll.FlatStyle = "Flat"
$btnSelectAll.FlatAppearance.BorderColor = $script:Border
$btnSelectAll.FlatAppearance.BorderSize  = 1
$btnSelectAll.BackColor = $script:BgInput
$btnSelectAll.ForeColor = $script:FgPrimary
$btnSelectAll.Font      = [System.Drawing.Font]::new($script:FontFamily, $script:BaseFontSize - 1)
$btnSelectAll.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnSelectAll.Visible   = $false
$chkRow.Controls.Add($btnSelectAll)

# The actual checkbox container (inside chkRow, RIGHT of the Select button + 12px gap)
$chkPanel = [System.Windows.Forms.Panel]::new()
$chkPanel.Location  = [System.Drawing.Point]::new(78, 0)
$chkPanel.Height    = 28
$chkPanel.BackColor = $script:BgMid
$chkRow.Controls.Add($chkPanel)

$form.Controls.Add($topPanel)

# ── Status bar ────────────────────────────────────────────────────────────────
$statusBar = [System.Windows.Forms.StatusStrip]::new()
$statusBar.BackColor  = $script:BgMid
$statusBar.SizingGrip = $false
$statusLabel = [System.Windows.Forms.ToolStripStatusLabel]::new()
$statusLabel.ForeColor = $script:FgSecondary
$statusLabel.Font      = $script:Font
$statusLabel.Text      = "No file loaded"
$statusBar.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusBar)

# ── Split container ──────────────────────────────────────────────────────────
$split = [System.Windows.Forms.SplitContainer]::new()
$split.Dock          = "Fill"
$split.Orientation   = "Vertical"
$split.BackColor     = $script:Border
$split.SplitterWidth = 3
$split.Panel1MinSize = 100
$split.Panel2MinSize = 100
$split.Panel1.BackColor = $script:BgDark
$split.Panel2.BackColor = $script:BgDark
$form.Controls.Add($split)
$split.BringToFront()

# Track user's manual splitter position
$split.Add_SplitterMoved({
    $totalW = $split.ClientSize.Width - $split.SplitterWidth
    if ($totalW -gt 0) {
        $script:UserSplitterRatio = [Math]::Round($split.SplitterDistance / $totalW, 4)
    }
})

# Set initial splitter after form is sized
$form.Add_Load({
    $totalW = $split.ClientSize.Width - $split.SplitterWidth
    $target = [int]($totalW * 0.5)
    $target = [Math]::Max($split.Panel1MinSize, [Math]::Min($target, $totalW - $split.Panel2MinSize))
    $split.SplitterDistance = $target
})

# ── Left panel: results list ─────────────────────────────────────────────────
$lblResults = [System.Windows.Forms.Label]::new()
$lblResults.Text      = " Results"
$lblResults.Dock      = "Top"
$lblResults.Font      = $script:FontBold
$lblResults.ForeColor = $script:Accent
$lblResults.BackColor = $script:BgMid
$lblResults.Height    = 26
$lblResults.TextAlign = "MiddleLeft"
$split.Panel1.Controls.Add($lblResults)

$listBox = [System.Windows.Forms.ListBox]::new()
$listBox.Dock           = "Fill"
$listBox.Font           = [System.Drawing.Font]::new($script:FontFamily, $script:TextBoxFontSz)
$listBox.BackColor      = $script:BgDark
$listBox.ForeColor      = $script:FgPrimary
$listBox.BorderStyle    = "None"
$listBox.ItemHeight     = 26
$listBox.DrawMode       = "OwnerDrawFixed"
$listBox.IntegralHeight = $false
$split.Panel1.Controls.Add($listBox)

# Reveal button bar at bottom of Panel1 (same height as detail button bar)
$revealBar = [System.Windows.Forms.Panel]::new()
$revealBar.Dock      = "Bottom"
$revealBar.Height    = 36
$revealBar.BackColor = $script:BgMid
$split.Panel1.Controls.Add($revealBar)

$btnReveal = [System.Windows.Forms.Button]::new()
$btnReveal.Text      = "Reveal"
$btnReveal.Size      = [System.Drawing.Size]::new(90, 28)
$btnReveal.Location  = [System.Drawing.Point]::new(4, 4)
$btnReveal.FlatStyle = "Flat"
$btnReveal.FlatAppearance.BorderColor = $script:Border
$btnReveal.FlatAppearance.BorderSize  = 1
$btnReveal.BackColor = $script:BgInput
$btnReveal.ForeColor = $script:FgPrimary
$btnReveal.Font      = $script:Font
$btnReveal.Cursor    = [System.Windows.Forms.Cursors]::Hand
$revealBar.Controls.Add($btnReveal)

$listBox.BringToFront()

# Dark scrollbars on the listbox
$listBox.Add_HandleCreated({
    try { [WinHelper]::ApplyDarkScrollbars($listBox.Handle) } catch {}
})

$listBox.Add_DrawItem({
    param($s, $e)
    if ($e.Index -lt 0) { return }
    $item = $s.Items[$e.Index]
    $sel  = (($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -ne 0)
    $bg   = if ($sel) { $script:AccentDim } else { $script:BgDark }
    $fg   = if ($sel) { [System.Drawing.Color]::White } else { $script:FgPrimary }
    $brush = [System.Drawing.SolidBrush]::new($bg)
    $e.Graphics.FillRectangle($brush, $e.Bounds)
    $brush.Dispose()

    $text = "$item"
    $secTag = ""
    if ($text.StartsWith("[")) {
        $end = $text.IndexOf("]")
        if ($end -gt 0) {
            $secTag = $text.Substring(0, $end + 1)
            $text   = $text.Substring($end + 2)
        }
    }
    $x = $e.Bounds.X + 6
    $y = $e.Bounds.Y + 3
    if ($secTag) {
        $tagFont  = [System.Drawing.Font]::new($script:FontFamily, 8)
        $tagColor = if ($sel) { [System.Drawing.Color]::FromArgb(180,210,255) } else { $script:FgDim }
        $tagBrush = [System.Drawing.SolidBrush]::new($tagColor)
        $e.Graphics.DrawString($secTag, $tagFont, $tagBrush, $x, $y + 1)
        $x += [System.Windows.Forms.TextRenderer]::MeasureText($secTag, $tagFont).Width
        $tagBrush.Dispose()
        $tagFont.Dispose()
    }
    $fgBrush = [System.Drawing.SolidBrush]::new($fg)
    $e.Graphics.DrawString($text, $s.Font, $fgBrush, $x, $y)
    $fgBrush.Dispose()
})

# ── Right-click context menu on listbox ──────────────────────────────────────
$ctxMenu = [System.Windows.Forms.ContextMenuStrip]::new()
$ctxMenu.BackColor = $script:BgMid
$ctxMenu.Font      = $script:Font
$ctxMenu.ShowImageMargin = $false
# Custom dark renderer for context menus and submenus — draws text manually
# so per-item ForeColor cannot override our hover logic
Add-Type -TypeDefinition @"
using System.Drawing;
using System.Windows.Forms;
public class DarkMenuRenderer : ToolStripProfessionalRenderer {
    private Color _bg;
    private Color _bgHover;
    private Color _border;
    private Color _fgNormal;
    private Color _fgHover;
    public DarkMenuRenderer(Color bg, Color bgHover, Color border, Color fgNormal, Color fgHover)
        : base() { _bg = bg; _bgHover = bgHover; _border = border; _fgNormal = fgNormal; _fgHover = fgHover; }
    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) {
        using (var b = new SolidBrush(_bg)) e.Graphics.FillRectangle(b, e.AffectedBounds);
    }
    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) {
        using (var p = new Pen(_border)) e.Graphics.DrawRectangle(p, 0, 0,
            e.AffectedBounds.Width - 1, e.AffectedBounds.Height - 1);
    }
    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) {
        var rc = new Rectangle(Point.Empty, e.Item.Size);
        var c = (e.Item.Selected || e.Item.Pressed) ? _bgHover : _bg;
        using (var b = new SolidBrush(c)) e.Graphics.FillRectangle(b, rc);
    }
    protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e) {
        var c = (e.Item.Selected || e.Item.Pressed) ? _fgHover : _fgNormal;
        using (var b = new SolidBrush(c))
            e.Graphics.DrawString(e.Text, e.TextFont, b, e.TextRectangle,
                new StringFormat { LineAlignment = StringAlignment.Center });
    }
    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e) {
        using (var b = new SolidBrush(_bg)) e.Graphics.FillRectangle(b, e.Item.Bounds);
        int y = e.Item.ContentRectangle.Height / 2;
        using (var p = new Pen(_border)) e.Graphics.DrawLine(p, 4, y, e.Item.Width - 4, y);
    }
    protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e) {
        e.ArrowColor = (e.Item.Selected || e.Item.Pressed) ? _fgHover : _fgNormal;
        base.OnRenderArrow(e);
    }
    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) {
        using (var b = new SolidBrush(_bg)) e.Graphics.FillRectangle(b, e.AffectedBounds);
    }
}
"@ -ReferencedAssemblies System.Windows.Forms, System.Drawing -ErrorAction SilentlyContinue

$darkRenderer = [DarkMenuRenderer]::new(
    $script:BgMid,                                   # background
    $script:AccentDim,                                # hover background
    $script:Border,                                   # border
    $script:FgPrimary,                                # normal text (light)
    [System.Drawing.Color]::FromArgb(20, 20, 20)      # hover text (dark)
)
$ctxMenu.Renderer = $darkRenderer

# Clone submenu
$mnuClone       = [System.Windows.Forms.ToolStripMenuItem]::new("Clone")
$mnuCloneFirst  = [System.Windows.Forms.ToolStripMenuItem]::new("Place at Position First")
$mnuCloneBefore = [System.Windows.Forms.ToolStripMenuItem]::new("Place at Position -1")
$mnuCloneAfter  = [System.Windows.Forms.ToolStripMenuItem]::new("Place at Position +1")
$mnuCloneLast   = [System.Windows.Forms.ToolStripMenuItem]::new("Place at Position Last")

$mnuClone.DropDownItems.AddRange(@($mnuCloneFirst, $mnuCloneBefore, $mnuCloneAfter, $mnuCloneLast))
$mnuClone.DropDown.BackColor = $script:BgMid
$mnuClone.DropDown.Renderer  = $darkRenderer
$mnuClone.DropDown.ShowImageMargin = $false

# Remove submenu
$mnuRemove = [System.Windows.Forms.ToolStripMenuItem]::new("Remove")
$mnuDelete = [System.Windows.Forms.ToolStripMenuItem]::new("Delete Object")

$mnuRemove.DropDownItems.Add($mnuDelete)
$mnuRemove.DropDown.BackColor = $script:BgMid
$mnuRemove.DropDown.Renderer  = $darkRenderer
$mnuRemove.DropDown.ShowImageMargin = $false

$ctxMenu.Items.AddRange(@($mnuClone, $mnuRemove))

# Right-click: select the item under cursor, show menu
$listBox.Add_MouseDown({
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $idx = $listBox.IndexFromPoint($e.Location)
        if ($idx -ge 0 -and $idx -lt $listBox.Items.Count) {
            $listBox.SelectedIndex = $idx
            $ctxMenu.Show($listBox, $e.Location)
        }
    }
})

# ── Clone / Delete helper functions ──────────────────────────────────────────
function Deep-CloneEntry {
    param($Entry)
    # Round-trip through JSON for a true deep copy
    $json = $Entry | ConvertTo-Json -Depth 100 -EnumsAsStrings
    return ($json | ConvertFrom-Json -AsHashtable -Depth 100)
}

function Resequence-Ids {
    param($SectionEntries)
    # Only resequence if entries have an 'id' field that is numeric (int/long/double)
    if ($SectionEntries.Count -eq 0) { return }
    $first = $SectionEntries[0]
    if (-not ($first -is [System.Collections.IDictionary])) { return }
    if (-not $first.Contains('id')) { return }
    $firstId = $first['id']
    if (-not ($firstId -is [int] -or $firstId -is [long] -or $firstId -is [double] -or $firstId -is [decimal])) { return }

    for ($i = 0; $i -lt $SectionEntries.Count; $i++) {
        $entry = $SectionEntries[$i]
        if ($entry -is [System.Collections.IDictionary] -and $entry.Contains('id')) {
            $entry['id'] = $i + 1
        }
    }
}

function Rebuild-JsonFromSections {
    # Rebuild $script:JsonData from $script:Sections, then save to disk
    if ($script:IsRootArray) {
        # Root array: the single "_entries" section IS the entire file
        $script:JsonData = [System.Collections.ArrayList]@($script:Sections["_entries"])
    }
    else {
        foreach ($sec in $script:Sections.Keys) {
            if ($sec -eq '_scalars') {
                $scalarEntry = $script:Sections['_scalars'][0]
                foreach ($sk in $scalarEntry.Keys) { $script:JsonData[$sk] = $scalarEntry[$sk] }
            }
            else {
                $origVal = $script:JsonData[$sec]
                if ($origVal -is [System.Collections.IList]) {
                    $script:JsonData[$sec] = [System.Collections.ArrayList]@($script:Sections[$sec])
                }
                else {
                    if ($script:Sections[$sec].Count -gt 0) {
                        $script:JsonData[$sec] = $script:Sections[$sec][0]
                    }
                }
            }
        }
    }
    Save-JsonData $script:JsonData
}

function Save-AfterStructuralChange {
    Rebuild-JsonFromSections
}

function Do-CloneEntry {
    param([string]$Position)  # "first", "before", "after", "last"
    if ($null -eq $script:SelectedSection -or $script:SelectedIndex -lt 0) { return }

    $sec     = $script:SelectedSection
    $idx     = $script:SelectedIndex
    $entries = $script:Sections[$sec]
    $clone   = Deep-CloneEntry $entries[$idx]

    switch ($Position) {
        "first" {
            $entries.Insert(0, $clone)
            # Original moved to idx+1
        }
        "before" {
            $entries.Insert($idx, $clone)
            # Original moved to idx+1; keep viewing original
        }
        "after" {
            $entries.Insert($idx + 1, $clone)
        }
        "last" {
            $entries.Add($clone) | Out-Null
        }
    }

    Resequence-Ids $entries
    Save-AfterStructuralChange

    # Update checkbox label with new count
    if ($script:Checkboxes.Contains($sec)) {
        $displayName = if ($sec -eq '_scalars') { '(scalars)' } elseif ($sec -eq '_entries') { '(all entries)' } else { $sec }
        $script:Checkboxes[$sec].Text = "$displayName ($($entries.Count))"
    }

    # Refresh and re-select
    $script:RunSearch.Invoke()
    $newIdx = switch ($Position) {
        "first"  { $idx + 1 }  # original shifted down
        "before" { $idx + 1 }  # original shifted down
        "after"  { $idx }
        "last"   { $idx }
    }
    # Find the entry in the current results by section + index
    for ($ri = 0; $ri -lt $script:CurrentResults.Count; $ri++) {
        $r = $script:CurrentResults[$ri]
        if ($r.Section -eq $sec -and $r.Index -eq $newIdx) {
            $listBox.SelectedIndex = $ri
            break
        }
    }
    $statusLabel.Text = "Cloned entry ($Position) in [$sec]"
}

function Do-DeleteEntry {
    if ($null -eq $script:SelectedSection -or $script:SelectedIndex -lt 0) { return }

    $sec     = $script:SelectedSection
    $idx     = $script:SelectedIndex
    $entries = $script:Sections[$sec]

    if ($entries.Count -le 0) { return }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Delete this entry from [$sec]?",
        "Confirm Delete",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $entries.RemoveAt($idx)
    Resequence-Ids $entries
    Save-AfterStructuralChange

    # Update checkbox label
    if ($script:Checkboxes.Contains($sec)) {
        $displayName = if ($sec -eq '_scalars') { '(scalars)' } elseif ($sec -eq '_entries') { '(all entries)' } else { $sec }
        $script:Checkboxes[$sec].Text = "$displayName ($($entries.Count))"
    }

    # Clear detail and refresh
    $script:SelectedSection = $null
    $script:SelectedIndex   = -1
    $script:DetailDirty     = $false
    $txtDetail.Text     = ""
    $txtDetail.ReadOnly = $true
    $lblDetail.Text     = " Detail"
    $btnSave.Enabled    = $false
    $btnRevert.Enabled  = $false

    $script:RunSearch.Invoke()

    # Select nearest item
    if ($listBox.Items.Count -gt 0) {
        $newIdx = [Math]::Min($idx, $listBox.Items.Count - 1)
        # Find an item in the same section near the old index
        for ($ri = 0; $ri -lt $script:CurrentResults.Count; $ri++) {
            if ($script:CurrentResults[$ri].Section -eq $sec) {
                $listBox.SelectedIndex = $ri
                break
            }
        }
    }
    $statusLabel.Text = "Deleted entry from [$sec]"
}

# Wire up menu items
$mnuCloneFirst.Add_Click({ Do-CloneEntry "first" })
$mnuCloneBefore.Add_Click({ Do-CloneEntry "before" })
$mnuCloneAfter.Add_Click({ Do-CloneEntry "after" })
$mnuCloneLast.Add_Click({ Do-CloneEntry "last" })
$mnuDelete.Add_Click({ Do-DeleteEntry })

# ── Right panel: detail view ─────────────────────────────────────────────────
$detailOuter = [System.Windows.Forms.Panel]::new()
$detailOuter.Dock      = "Fill"
$detailOuter.BackColor = $script:BgDark
$split.Panel2.Controls.Add($detailOuter)

# Header bar
$detailHeader = [System.Windows.Forms.Panel]::new()
$detailHeader.Dock      = "Top"
$detailHeader.Height    = 26
$detailHeader.BackColor = $script:BgMid
$detailOuter.Controls.Add($detailHeader)

$lblDetail = [System.Windows.Forms.Label]::new()
$lblDetail.Text      = " Detail"
$lblDetail.Dock      = "Fill"
$lblDetail.Font      = $script:FontBold
$lblDetail.ForeColor = $script:Accent
$lblDetail.BackColor = $script:BgMid
$lblDetail.TextAlign = "MiddleLeft"
$detailHeader.Controls.Add($lblDetail)

# Zoom controls docked right
$zoomPanel = [System.Windows.Forms.Panel]::new()
$zoomPanel.Dock      = "Right"
$zoomPanel.Width     = 110
$zoomPanel.BackColor = $script:BgMid
$detailHeader.Controls.Add($zoomPanel)

$btnZoomOut = [System.Windows.Forms.Button]::new()
$btnZoomOut.Text      = "-"
$btnZoomOut.Location  = [System.Drawing.Point]::new(2, 2)
$btnZoomOut.Size      = [System.Drawing.Size]::new(28, 22)
$btnZoomOut.FlatStyle = "Flat"
$btnZoomOut.FlatAppearance.BorderSize = 0
$btnZoomOut.BackColor = $script:BgInput
$btnZoomOut.ForeColor = $script:FgPrimary
$btnZoomOut.Font      = [System.Drawing.Font]::new($script:FontFamily, 10, [System.Drawing.FontStyle]::Bold)
$btnZoomOut.Cursor    = [System.Windows.Forms.Cursors]::Hand
$zoomPanel.Controls.Add($btnZoomOut)

$lblZoom = [System.Windows.Forms.Label]::new()
$lblZoom.Text      = "$($script:DetailFontSize)pt"
$lblZoom.Location  = [System.Drawing.Point]::new(32, 2)
$lblZoom.Size      = [System.Drawing.Size]::new(42, 22)
$lblZoom.ForeColor = $script:FgSecondary
$lblZoom.BackColor = $script:BgMid
$lblZoom.TextAlign = "MiddleCenter"
$lblZoom.Font      = [System.Drawing.Font]::new($script:FontFamily, 8)
$zoomPanel.Controls.Add($lblZoom)

$btnZoomIn = [System.Windows.Forms.Button]::new()
$btnZoomIn.Text      = "+"
$btnZoomIn.Location  = [System.Drawing.Point]::new(76, 2)
$btnZoomIn.Size      = [System.Drawing.Size]::new(28, 22)
$btnZoomIn.FlatStyle = "Flat"
$btnZoomIn.FlatAppearance.BorderSize = 0
$btnZoomIn.BackColor = $script:BgInput
$btnZoomIn.ForeColor = $script:FgPrimary
$btnZoomIn.Font      = [System.Drawing.Font]::new($script:FontFamily, 10, [System.Drawing.FontStyle]::Bold)
$btnZoomIn.Cursor    = [System.Windows.Forms.Cursors]::Hand
$zoomPanel.Controls.Add($btnZoomIn)

# Button bar at bottom of detail panel
$btnBar = [System.Windows.Forms.Panel]::new()
$btnBar.Dock      = "Bottom"
$btnBar.Height    = 36
$btnBar.BackColor = $script:BgMid
$detailOuter.Controls.Add($btnBar)

# Open button (left-aligned in button bar)
$btnOpen = [System.Windows.Forms.Button]::new()
$btnOpen.Text      = "Open  (Ctrl+O)"
$btnOpen.Size      = [System.Drawing.Size]::new(120, 28)
$btnOpen.Location  = [System.Drawing.Point]::new(4, 4)
$btnOpen.FlatStyle = "Flat"
$btnOpen.FlatAppearance.BorderColor = $script:Border
$btnOpen.BackColor = $script:BgInput
$btnOpen.ForeColor = $script:FgPrimary
$btnOpen.Font      = $script:Font
$btnOpen.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnBar.Controls.Add($btnOpen)

# Save button (right-aligned)
$btnSave = [System.Windows.Forms.Button]::new()
$btnSave.Text      = "Save  (Ctrl+S)"
$btnSave.Size      = [System.Drawing.Size]::new(120, 28)
$btnSave.FlatStyle = "Flat"
$btnSave.FlatAppearance.BorderColor = $script:AccentDim
$btnSave.BackColor = $script:BgInput
$btnSave.ForeColor = $script:FgPrimary
$btnSave.Font      = $script:Font
$btnSave.Enabled   = $false
$btnSave.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnBar.Controls.Add($btnSave)

# Revert button (right-aligned)
$btnRevert = [System.Windows.Forms.Button]::new()
$btnRevert.Text      = "Revert"
$btnRevert.Size      = [System.Drawing.Size]::new(80, 28)
$btnRevert.FlatStyle = "Flat"
$btnRevert.FlatAppearance.BorderColor = $script:Border
$btnRevert.BackColor = $script:BgInput
$btnRevert.ForeColor = $script:FgPrimary
$btnRevert.Font      = $script:Font
$btnRevert.Enabled   = $false
$btnRevert.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnBar.Controls.Add($btnRevert)

# Detail editor - RichTextBox
$txtDetail = [System.Windows.Forms.RichTextBox]::new()
$txtDetail.Dock       = "Fill"
$txtDetail.Font       = [System.Drawing.Font]::new($script:FontFamily, $script:DetailFontSize)
$txtDetail.BackColor  = [System.Drawing.Color]::FromArgb(38, 38, 42)
$txtDetail.ForeColor  = $script:FgPrimary
$txtDetail.BorderStyle = "None"
$txtDetail.WordWrap   = $false
$txtDetail.ReadOnly   = $true
$txtDetail.DetectUrls = $false
$txtDetail.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
$txtDetail.AcceptsTab = $true
$detailOuter.Controls.Add($txtDetail)
$txtDetail.BringToFront()

# Dark scrollbars on the RichTextBox
$txtDetail.Add_HandleCreated({
    try { [WinHelper]::ApplyDarkScrollbars($txtDetail.Handle) } catch {}
})

# ══════════════════════════════════════════════════════════════════════════════
#   ZOOM
# ══════════════════════════════════════════════════════════════════════════════
function Update-DetailFont {
    $txtDetail.Font = [System.Drawing.Font]::new($script:FontFamily, $script:DetailFontSize)
    $lblZoom.Text   = "$($script:DetailFontSize)pt"
}

$btnZoomIn.Add_Click({
    if ($script:DetailFontSize -lt $script:MaxDetailFont) {
        $script:DetailFontSize++
        Update-DetailFont
    }
})
$btnZoomOut.Add_Click({
    if ($script:DetailFontSize -gt $script:MinDetailFont) {
        $script:DetailFontSize--
        Update-DetailFont
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#   SELECT ALL TOGGLE
# ══════════════════════════════════════════════════════════════════════════════
$btnSelectAll.Add_Click({
    if (-not $script:SelectAllActive) {
        # Activate: save current checkbox state, then check all
        $script:PreSelectAllState = @{}
        foreach ($sec in $script:Checkboxes.Keys) {
            $script:PreSelectAllState[$sec] = $script:Checkboxes[$sec].Checked
        }
        foreach ($sec in $script:Checkboxes.Keys) {
            $script:Checkboxes[$sec].Checked = $true
        }
        $script:SelectAllActive = $true
        $btnSelectAll.BackColor = $script:Accent
        $btnSelectAll.ForeColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
        $btnSelectAll.FlatAppearance.BorderColor = $script:Accent
    }
    else {
        # Deactivate: restore previous checkbox state
        foreach ($sec in $script:PreSelectAllState.Keys) {
            if ($script:Checkboxes.Contains($sec)) {
                $script:Checkboxes[$sec].Checked = $script:PreSelectAllState[$sec]
            }
        }
        $script:SelectAllActive = $false
        $btnSelectAll.BackColor = $script:BgInput
        $btnSelectAll.ForeColor = $script:FgPrimary
        $btnSelectAll.FlatAppearance.BorderColor = $script:Border
    }
    $script:RunSearch.Invoke()
})

# ══════════════════════════════════════════════════════════════════════════════
#   REVEAL TOGGLE - show full JSON, disable list/checkboxes
# ══════════════════════════════════════════════════════════════════════════════
$script:DisabledFg = [System.Drawing.Color]::FromArgb(90, 90, 90)

$btnReveal.Add_Click({
    if (-not $script:RevealActive) {
        # Activate Reveal: save current detail state, show full JSON
        $script:PreRevealDetailText  = $txtDetail.Text
        $script:PreRevealDetailLabel = $lblDetail.Text
        $script:PreRevealReadOnly    = $txtDetail.ReadOnly

        # Show full JSON in the viewer (read-only)
        if ($null -ne $script:JsonData) {
            $fullJson = $script:JsonData | ConvertTo-Json -Depth 100 -EnumsAsStrings
            $txtDetail.ReadOnly = $true
            $txtDetail.Text     = $fullJson
            $txtDetail.SelectionStart = 0
            $lblDetail.Text     = " Full JSON  -  $([System.IO.Path]::GetFileName($script:JsonPath))"
        }

        # Disable controls
        $listBox.Enabled      = $false
        $listBox.BackColor    = [System.Drawing.Color]::FromArgb(40, 40, 40)
        $btnSelectAll.Enabled = $false
        $btnSelectAll.ForeColor = $script:DisabledFg
        $btnSave.Enabled      = $false
        $btnRevert.Enabled    = $false
        foreach ($sec in $script:Checkboxes.Keys) {
            $script:Checkboxes[$sec].Enabled  = $false
            $script:Checkboxes[$sec].ForeColor = $script:DisabledFg
        }

        $script:RevealActive = $true
        $btnReveal.BackColor = $script:Accent
        $btnReveal.ForeColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
        $btnReveal.FlatAppearance.BorderColor = $script:Accent
        $statusLabel.Text = "Reveal mode  -  showing full JSON (read-only)"
    }
    else {
        # Deactivate Reveal: restore previous state
        $txtDetail.ReadOnly = $script:PreRevealReadOnly
        $txtDetail.Text     = $script:PreRevealDetailText
        $lblDetail.Text     = $script:PreRevealDetailLabel
        if ($script:PreRevealDetailText.Length -gt 0) {
            $txtDetail.SelectionStart = 0
        }

        # Re-enable controls
        $listBox.Enabled      = $true
        $listBox.BackColor    = $script:BgDark
        $btnSelectAll.Enabled = $true
        if ($script:SelectAllActive) {
            $btnSelectAll.ForeColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
        } else {
            $btnSelectAll.ForeColor = $script:FgPrimary
        }
        foreach ($sec in $script:Checkboxes.Keys) {
            $script:Checkboxes[$sec].Enabled  = $true
            $script:Checkboxes[$sec].ForeColor = $script:FgPrimary
        }
        # Restore save/revert state based on dirty flag
        $btnSave.Enabled   = $script:DetailDirty
        $btnRevert.Enabled = $script:DetailDirty

        $script:RevealActive = $false
        $btnReveal.BackColor = $script:BgInput
        $btnReveal.ForeColor = $script:FgPrimary
        $btnReveal.FlatAppearance.BorderColor = $script:Border
        $statusLabel.Text = "$($script:CurrentResults.Count) match(es)"
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#   LAYOUT
# ══════════════════════════════════════════════════════════════════════════════
function Do-Layout {
    $w = $topPanel.ClientSize.Width
    $txtSearch.Width = [Math]::Max(100, $w - 24)
    $chkRow.Width    = [Math]::Max(100, $w - 24)

    # Select button at LEFT edge of chkRow
    $selectBtnW = 64
    $btnSelectAll.Location = [System.Drawing.Point]::new(0, 0)
    $btnSelectAll.Size     = [System.Drawing.Size]::new($selectBtnW, 22)

    # Checkbox panel starts after Select button + 12px breathing space
    $chkPanelX = $selectBtnW + 14
    $chkPanel.Location = [System.Drawing.Point]::new($chkPanelX, 0)
    $chkPanel.Width = [Math]::Max(50, $chkRow.Width - $chkPanelX - 2)

    # Adaptive checkbox layout
    if ($script:Checkboxes.Count -gt 0) {
        $panelW = $chkPanel.Width
        $count  = $script:Checkboxes.Count

        $fitsAtPref = [Math]::Max(1, [Math]::Floor($panelW / $script:ChkPreferredWidth))

        if ($count -le $fitsAtPref) {
            $cellW = $script:ChkPreferredWidth
        }
        else {
            $targetRows = [Math]::Min(3, [Math]::Ceiling($count / $fitsAtPref))
            $perRow = [Math]::Ceiling($count / $targetRows)
            $cellW  = [Math]::Max($script:ChkMinWidth, [Math]::Floor($panelW / $perRow))
        }

        $cols = [Math]::Max(1, [Math]::Floor($panelW / $cellW))
        $idx  = 0
        foreach ($sec in $script:Checkboxes.Keys) {
            $chk = $script:Checkboxes[$sec]
            $col = $idx % $cols
            $row = [Math]::Floor($idx / $cols)
            $chk.Location = [System.Drawing.Point]::new($col * $cellW, $row * 24)
            $chk.Width    = $cellW - 4
            $idx++
        }
        $totalRows = [Math]::Ceiling($count / $cols)
        $chkPanel.Height = [Math]::Max(24, $totalRows * 24 + 2)
    }
    else {
        $chkPanel.Height = 4
    }

    $chkRow.Height   = [Math]::Max($chkPanel.Height, 24)
    $topPanel.Height = 44 + $chkRow.Height + 6

    # Center the Select button vertically in the chkRow
    $btnSelectAll.Location = [System.Drawing.Point]::new(
        0,
        [Math]::Max(0, [int](($chkRow.Height - 22) / 2))
    )

    # Button bar: Open left-aligned, Save/Revert right-aligned
    $bw = $btnBar.ClientSize.Width
    $btnOpen.Location   = [System.Drawing.Point]::new(4, 4)
    $btnSave.Location   = [System.Drawing.Point]::new($bw - 210, 4)
    $btnRevert.Location = [System.Drawing.Point]::new($bw - 86, 4)
}

$topPanel.Add_Resize({ Do-Layout })
$btnBar.Add_Resize({ Do-Layout })

# ══════════════════════════════════════════════════════════════════════════════
#   SEARCH
# ══════════════════════════════════════════════════════════════════════════════
$script:RunSearch = {
    if ($null -eq $script:Sections) { return }
    $active = [System.Collections.ArrayList]::new()
    foreach ($sec in $script:Checkboxes.Keys) {
        if ($script:Checkboxes[$sec].Checked) { [void]$active.Add($sec) }
    }
    $script:CurrentResults = Search-Entries -Query $txtSearch.Text -ActiveSections $active.ToArray()

    $listBox.BeginUpdate()
    $listBox.Items.Clear()
    foreach ($r in $script:CurrentResults) {
        $listBox.Items.Add("[$($r.Section)] $($r.Label)") | Out-Null
    }
    $listBox.EndUpdate()

    $statusLabel.Text = "$($script:CurrentResults.Count) match(es) in $($active.Count) section(s)"
}

$searchTimer = [System.Windows.Forms.Timer]::new()
$searchTimer.Interval = 180
$searchTimer.Add_Tick({ $searchTimer.Stop(); $script:RunSearch.Invoke() })
$txtSearch.Add_TextChanged({ $searchTimer.Stop(); $searchTimer.Start() })

# ══════════════════════════════════════════════════════════════════════════════
#   SELECTION / DETAIL DISPLAY
# ══════════════════════════════════════════════════════════════════════════════
$listBox.Add_SelectedIndexChanged({
    if ($script:RevealActive) { return }  # ignore during Reveal mode
    $idx = $listBox.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:CurrentResults.Count) { return }
    $r = $script:CurrentResults[$idx]
    $json = $r.Entry | ConvertTo-Json -Depth 50 -EnumsAsStrings
    $txtDetail.ReadOnly = $false
    $txtDetail.Text     = $json
    $txtDetail.SelectionStart = 0
    $lblDetail.Text     = " [$($r.Section)] $($r.Label)"
    $script:SelectedSection = $r.Section
    $script:SelectedIndex   = $r.Index
    $script:DetailDirty     = $false
    $btnSave.Enabled   = $false
    $btnRevert.Enabled = $false

    # Only snap to 20% on first selection (when no user position has been set yet)
    if ($null -eq $script:UserSplitterRatio) {
        $totalW  = $split.ClientSize.Width - $split.SplitterWidth
        $target  = [int]($totalW * 0.20)
        $maxDist = $totalW - $split.Panel2MinSize
        $target  = [Math]::Max($split.Panel1MinSize, [Math]::Min($target, $maxDist))
        try { $split.SplitterDistance = $target } catch {}
        $script:UserSplitterRatio = 0.20
    }
})

$txtDetail.Add_TextChanged({
    if (-not $txtDetail.ReadOnly -and $null -ne $script:SelectedSection) {
        $script:DetailDirty = $true
        $btnSave.Enabled    = $true
        $btnRevert.Enabled  = $true
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#   SAVE / REVERT
# ══════════════════════════════════════════════════════════════════════════════
function Do-SaveDetail {
    if (-not $script:DetailDirty -or $null -eq $script:SelectedSection) { return $false }
    try {
        $newEntry = $txtDetail.Text | ConvertFrom-Json -AsHashtable -Depth 100
        $script:Sections[$script:SelectedSection][$script:SelectedIndex] = $newEntry

        Rebuild-JsonFromSections

        $script:DetailDirty = $false
        $btnSave.Enabled    = $false
        $btnRevert.Enabled  = $false
        $statusLabel.Text   = "Saved"

        $savedIdx = $listBox.SelectedIndex
        $script:RunSearch.Invoke()
        if ($savedIdx -ge 0 -and $savedIdx -lt $listBox.Items.Count) {
            $listBox.SelectedIndex = $savedIdx
        }
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Invalid JSON:`n$($_.Exception.Message)", "Parse Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
}

$btnSave.Add_Click({ Do-SaveDetail })

$btnRevert.Add_Click({
    if ($null -ne $script:SelectedSection -and $script:SelectedIndex -ge 0) {
        $entry = $script:Sections[$script:SelectedSection][$script:SelectedIndex]
        $txtDetail.Text = $entry | ConvertTo-Json -Depth 50 -EnumsAsStrings
        $script:DetailDirty = $false
        $btnSave.Enabled    = $false
        $btnRevert.Enabled  = $false
        $statusLabel.Text   = "Reverted"
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#   OPEN FILE (shared logic for button + shortcut)
# ══════════════════════════════════════════════════════════════════════════════
function Show-OpenDialog {
    $ofd = [System.Windows.Forms.OpenFileDialog]::new()
    $ofd.InitialDirectory = if ($script:JsonPath) { [System.IO.Path]::GetDirectoryName($script:JsonPath) } else { (Get-Location).Path }
    $ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $ofd.Title  = "Select a JSON file"
    if ($ofd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        Open-JsonFile -FilePath $ofd.FileName
    }
}

$btnOpen.Add_Click({ Show-OpenDialog })

# ══════════════════════════════════════════════════════════════════════════════
#   KEYBOARD SHORTCUTS
# ══════════════════════════════════════════════════════════════════════════════
$form.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq 'S') {
        $e.SuppressKeyPress = $true
        if ($btnSave.Enabled) { Do-SaveDetail }
    }
    if ($e.Control -and $e.KeyCode -eq 'F') {
        $e.SuppressKeyPress = $true
        $txtSearch.Focus(); $txtSearch.SelectAll()
    }
    if ($e.Control -and $e.KeyCode -eq 'O') {
        $e.SuppressKeyPress = $true
        Show-OpenDialog
    }
    if ($e.Control -and ($e.KeyCode -eq 'Oemplus' -or $e.KeyCode -eq 'Add')) {
        $e.SuppressKeyPress = $true
        if ($script:DetailFontSize -lt $script:MaxDetailFont) { $script:DetailFontSize++; Update-DetailFont }
    }
    if ($e.Control -and ($e.KeyCode -eq 'OemMinus' -or $e.KeyCode -eq 'Subtract')) {
        $e.SuppressKeyPress = $true
        if ($script:DetailFontSize -gt $script:MinDetailFont) { $script:DetailFontSize--; Update-DetailFont }
    }
    if ($e.KeyCode -eq 'Escape') {
        if ($txtSearch.Focused -and $txtSearch.Text.Length -gt 0) {
            $txtSearch.Clear()
        }
        elseif ($listBox.SelectedIndex -ge 0) {
            $listBox.ClearSelected()
            $txtDetail.Text = ""
            $txtDetail.ReadOnly = $true
            $btnSave.Enabled = $false; $btnRevert.Enabled = $false
            $lblDetail.Text = " Detail"
            $script:SelectedSection = $null
            $script:UserSplitterRatio = $null  # reset so next selection snaps to 20% again
            # Restore saved splitter ratio or default 50%
            $totalW  = $split.ClientSize.Width - $split.SplitterWidth
            $ratio   = 0.5
            if ($script:SavedState -and $script:SavedState -is [hashtable] -and $script:SavedState.ContainsKey("__splitterRatio")) {
                $ratio = [double]$script:SavedState["__splitterRatio"]
            }
            $target  = [int]($totalW * $ratio)
            $maxDist = $totalW - $split.Panel2MinSize
            $target  = [Math]::Max($split.Panel1MinSize, [Math]::Min($target, $maxDist))
            try { $split.SplitterDistance = $target } catch {}
        }
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#   FILE LOADING
# ══════════════════════════════════════════════════════════════════════════════
function Open-JsonFile {
    param([string]$FilePath)

    # Auto-save current project state
    if ($script:DetailDirty -and $null -ne $script:SelectedSection) {
        Do-SaveDetail | Out-Null
    }
    if ($script:Checkboxes.Count -gt 0 -and $script:StatePath) {
        Save-ProjectState (Build-FullState)
    }

    $script:JsonPath = [System.IO.Path]::GetFullPath($FilePath)
    if (-not (Test-Path $script:JsonPath)) {
        [System.Windows.Forms.MessageBox]::Show("File not found: $($script:JsonPath)", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $jsonDir      = [System.IO.Path]::GetDirectoryName($script:JsonPath)
    $jsonBaseName = [System.IO.Path]::GetFileNameWithoutExtension($script:JsonPath)
    $script:StatePath = [System.IO.Path]::Combine($jsonDir, ".${jsonBaseName}_lookup_state.jlconf")

    try { $script:JsonData = Load-JsonData }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to parse JSON:`n$($_.Exception.Message)", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Discover sections (use ArrayList for mutability: Clone/Delete support)
    $script:Sections   = [ordered]@{}
    $script:ScalarKeys = @()
    $script:IsRootArray = $false

    if ($script:JsonData -is [System.Collections.IList]) {
        # Root is an array — wrap in a single virtual section
        $script:IsRootArray = $true
        $script:Sections["_entries"] = [System.Collections.ArrayList]@($script:JsonData)
    }
    else {
        # Root is an object — discover named sections
        foreach ($key in ($script:JsonData.Keys | Sort-Object)) {
            $val = $script:JsonData[$key]
            if ($val -is [System.Collections.IList]) {
                $script:Sections[$key] = [System.Collections.ArrayList]@($val)
            }
            elseif ($val -is [System.Collections.IDictionary]) {
                $script:Sections[$key] = [System.Collections.ArrayList]@(,$val)
            }
            else {
                $script:ScalarKeys += $key
            }
        }
        if ($script:ScalarKeys.Count -gt 0) {
            $scalarObj = [ordered]@{}
            foreach ($sk in $script:ScalarKeys) { $scalarObj[$sk] = $script:JsonData[$sk] }
            $script:Sections["_scalars"] = [System.Collections.ArrayList]@(,$scalarObj)
        }
    }

    # Load saved project state
    $script:SavedState = Load-ProjectState

    # Reset Select All state
    $script:SelectAllActive = $false
    $script:PreSelectAllState = @{}
    $btnSelectAll.BackColor = $script:BgInput
    $btnSelectAll.ForeColor = $script:FgPrimary
    $btnSelectAll.FlatAppearance.BorderColor = $script:Border

    # Reset Reveal state
    $script:RevealActive = $false
    $btnReveal.BackColor = $script:BgInput
    $btnReveal.ForeColor = $script:FgPrimary
    $btnReveal.FlatAppearance.BorderColor = $script:Border
    $listBox.Enabled   = $true
    $listBox.BackColor = $script:BgDark
    $btnSelectAll.Enabled = $true

    # Reset splitter tracking
    $script:UserSplitterRatio = $null

    # Build checkboxes
    $chkPanel.Controls.Clear()
    $script:Checkboxes = [ordered]@{}
    foreach ($sec in $script:Sections.Keys) {
        $chk = [System.Windows.Forms.CheckBox]::new()
        $displayName = if ($sec -eq '_scalars') { '(scalars)' } elseif ($sec -eq '_entries') { '(all entries)' } else { $sec }
        $count = $script:Sections[$sec].Count
        $chk.Text      = "$displayName ($count)"
        $chk.Tag       = $sec
        $chk.ForeColor = $script:FgPrimary
        $chk.BackColor = $script:BgMid
        $chk.Font      = $script:Font
        $chk.AutoSize  = $false
        $chk.Height    = 22
        $checked = $true
        if ($script:SavedState -and $script:SavedState -is [hashtable] -and $script:SavedState.ContainsKey($sec)) {
            $checked = [bool]$script:SavedState[$sec]
        }
        $chk.Checked = $checked
        $chk.Add_CheckedChanged({
            # If user manually changes a checkbox while Select All is active, deactivate Select All
            if ($script:SelectAllActive) {
                $script:SelectAllActive = $false
                $btnSelectAll.BackColor = $script:BgInput
                $btnSelectAll.ForeColor = $script:FgPrimary
                $btnSelectAll.FlatAppearance.BorderColor = $script:Border
            }
            $script:RunSearch.Invoke()
        })
        $chkPanel.Controls.Add($chk)
        $script:Checkboxes[$sec] = $chk
    }
    $btnSelectAll.Visible = ($script:Checkboxes.Count -gt 0)

    # Restore window geometry from this project's state
    $restoredGeometry = $false
    if ($script:SavedState -is [hashtable]) {
        $hasX = $script:SavedState.ContainsKey("__windowX")
        $hasW = $script:SavedState.ContainsKey("__windowWidth")
        if ($hasX -and $hasW) {
            $wx = [int]$script:SavedState["__windowX"]
            $wy = [int]$script:SavedState["__windowY"]
            $ww = [int]$script:SavedState["__windowWidth"]
            $wh = [int]$script:SavedState["__windowHeight"]

            # Validate the position is on a visible screen
            $testRect = [System.Drawing.Rectangle]::new($wx, $wy, $ww, $wh)
            $onScreen = $false
            foreach ($scr in [System.Windows.Forms.Screen]::AllScreens) {
                if ($scr.WorkingArea.IntersectsWith($testRect)) { $onScreen = $true; break }
            }
            if ($onScreen -and $ww -ge 640 -and $wh -ge 400) {
                $form.StartPosition = "Manual"
                $form.Location = [System.Drawing.Point]::new($wx, $wy)
                $form.Size     = [System.Drawing.Size]::new($ww, $wh)
                $restoredGeometry = $true
            }
        }
        # Restore window state (Normal, Minimized, Maximized)
        if ($script:SavedState.ContainsKey("__windowState")) {
            $ws = "$($script:SavedState['__windowState'])"
            switch ($ws) {
                "Maximized" { $form.WindowState = "Maximized" }
                "Minimized" { $form.WindowState = "Normal" }  # don't restore minimized
                default      { $form.WindowState = "Normal" }
            }
        }
    }

    $form.Text = "JSON Lookup  -  $([System.IO.Path]::GetFileName($script:JsonPath))"
    $txtSearch.Enabled = $true
    $txtSearch.Text    = ""
    $txtDetail.Text    = ""
    $txtDetail.ReadOnly = $true
    $lblDetail.Text    = " Detail"
    $script:SelectedSection = $null
    $script:SelectedIndex   = -1
    $script:DetailDirty     = $false
    $btnSave.Enabled   = $false
    $btnRevert.Enabled = $false

    Do-Layout

    # Restore splitter ratio
    if ($script:SavedState -is [hashtable] -and $script:SavedState.ContainsKey("__splitterRatio")) {
        $ratio  = [double]$script:SavedState["__splitterRatio"]
        $totalW = $split.ClientSize.Width - $split.SplitterWidth
        if ($totalW -gt 0) {
            $target  = [int]($totalW * $ratio)
            $maxDist = $totalW - $split.Panel2MinSize
            $target  = [Math]::Max($split.Panel1MinSize, [Math]::Min($target, $maxDist))
            try { $split.SplitterDistance = $target } catch {}
            $script:UserSplitterRatio = $ratio
        }
    }

    $script:RunSearch.Invoke()
    $txtSearch.Focus()
    $statusLabel.Text = "Loaded $([System.IO.Path]::GetFileName($script:JsonPath))"
}

# ══════════════════════════════════════════════════════════════════════════════
#   ON CLOSE - always save everything
# ══════════════════════════════════════════════════════════════════════════════
$form.Add_FormClosing({
    param($sender, $e)
    # Auto-save pending detail edits
    if ($script:DetailDirty -and $null -ne $script:SelectedSection) {
        $ok = Do-SaveDetail
        if (-not $ok) {
            $answer = [System.Windows.Forms.MessageBox]::Show(
                "Current edits contain invalid JSON and cannot be saved.`nClose anyway and discard?",
                "Invalid Edits",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($answer -eq [System.Windows.Forms.DialogResult]::No) {
                $e.Cancel = $true
                return
            }
        }
    }
    # Save full project state (checkboxes, splitter, window geometry)
    if ($script:StatePath) {
        Save-ProjectState (Build-FullState)
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#   STARTUP SEQUENCE
# ══════════════════════════════════════════════════════════════════════════════
$startTimer = [System.Windows.Forms.Timer]::new()
$startTimer.Interval = 1000
$startTimer.Add_Tick({
    $startTimer.Stop()
    $startTimer.Dispose()

    if ($script:JsonPath -and (Test-Path $script:JsonPath)) { return }

    # Try default.json
    $defaultPath = [System.IO.Path]::Combine((Get-Location).Path, "default.json")
    if (Test-Path $defaultPath) {
        Open-JsonFile -FilePath $defaultPath
        return
    }

    # File picker
    $ofd = [System.Windows.Forms.OpenFileDialog]::new()
    $ofd.InitialDirectory = (Get-Location).Path
    $ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $ofd.Title  = "Select a JSON file to work with"
    if ($ofd.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        Open-JsonFile -FilePath $ofd.FileName
    }
    else {
        $statusLabel.Text = "No file loaded  -  Ctrl+O to open"
    }
})

$form.Add_Shown({
    if ($Path -and $Path.Length -gt 0) {
        $resolved = if ([System.IO.Path]::IsPathRooted($Path)) { $Path }
                    else { [System.IO.Path]::Combine((Get-Location).Path, $Path) }
        if (Test-Path $resolved) {
            Open-JsonFile -FilePath $resolved
        }
        else {
            $statusLabel.Text = "File not found: $Path"
            $startTimer.Start()
        }
    }
    else {
        $startTimer.Start()
    }
    Do-Layout
})

# ── Run ───────────────────────────────────────────────────────────────────────
[System.Windows.Forms.Application]::Run($form)
