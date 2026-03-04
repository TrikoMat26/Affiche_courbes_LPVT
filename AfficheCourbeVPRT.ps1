[CmdletBinding()]
param()

# --- Configuration et Styles ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# --- Intégration C# pour FolderPicker Moderne ---
# Permet d'avoir la fenêtre de sélection "style Explorateur" au lieu de l'arbre WinForms classique
$folderPickerSource = @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class FolderPicker
{
    public virtual string ResultPath { get; protected set; }
    public virtual string Title { get; set; }
    public virtual string InputPath { get; set; }

    public bool ShowDialog(IntPtr ownerHandle)
    {
        IFileDialog dialog = null;
        try
        {
            dialog = (IFileDialog)new FileOpenDialog();
            dialog.SetOptions(FOS.FOS_PICKFOLDERS | FOS.FOS_FORCEFILESYSTEM);
            
            if (!string.IsNullOrEmpty(InputPath))
            {
                IShellItem item;
                if (SHCreateItemFromParsingName(InputPath, IntPtr.Zero, typeof(IShellItem).GUID, out item) == 0)
                {
                    dialog.SetFolder(item);
                }
            }

            if (!string.IsNullOrEmpty(Title))
            {
                dialog.SetTitle(Title);
            }

            if (dialog.Show(ownerHandle) == 0)
            {
                IShellItem result;
                dialog.GetResult(out result);
                string path;
                result.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out path);
                ResultPath = path;
                return true;
            }
        }
        catch { }
        finally
        {
            if (dialog != null) Marshal.ReleaseComObject(dialog);
        }
        return false;
    }

    [DllImport("shell32.dll")]
    private static extern int SHCreateItemFromParsingName([MarshalAs(UnmanagedType.LPWStr)] string pszPath, IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem ppv);

    [ComImport]
    [Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    [ClassInterface(ClassInterfaceType.None)]
    private class FileOpenDialog { }

    [ComImport]
    [Guid("42f85136-db7e-439c-85f1-e4075d135fc8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileDialog
    {
        [PreserveSig] int Show(IntPtr parent);
        void SetFileTypes();
        void SetFileTypeIndex();
        void GetFileTypeIndex();
        void Advise();
        void Unadvise();
        void SetOptions(FOS fos);
        void GetOptions();
        void SetDefaultFolder();
        void SetFolder(IShellItem psi);
        void GetFolder();
        void GetCurrentSelection();
        void SetFileName();
        void GetFileName();
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel();
        void SetFileNameLabel();
        void GetResult(out IShellItem ppsi);
        void AddPlace();
        void SetDefaultExtension();
        void Close();
        void SetClientGuid();
        void ClearClientData();
        void SetFilter();
    }

    [ComImport]
    [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        void BindToHandler();
        void GetParent();
        void GetDisplayName(SIGDN sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes();
        void Compare();
    }

    private enum SIGDN : uint
    {
        SIGDN_FILESYSPATH = 0x80058000
    }

    [Flags]
    private enum FOS : uint
    {
        FOS_PICKFOLDERS = 0x00000020,
        FOS_FORCEFILESYSTEM = 0x00000040
    }
}
'@

try {
    Add-Type -TypeDefinition $folderPickerSource -Language CSharp -ReferencedAssemblies 'System.Windows.Forms'
}
catch {
    # Ignore error if type already exists (when running script multiple times in same session)
}

# --- Constantes et Regex (Inchangés) ---
$RX_OPTS_I = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
$RX_OPTS_S = [System.Text.RegularExpressions.RegexOptions]::Singleline

$patVprtBegin = 'Begin\s*Sequence:\s*V.{1}RIFICATION\s+PR.{1}CISION\s+RAPPORT\s+DE\s+TRANSFERT'
$patVprtEnd = 'End\s*Sequence:\s*V.{1}RIFICATION\s+PR.{1}CISION\s+RAPPORT\s+DE\s+TRANSFERT'
$rxBegin = [regex]::new($patVprtBegin, $RX_OPTS_I -bor $RX_OPTS_S)
$rxEnd = [regex]::new($patVprtEnd, $RX_OPTS_I -bor $RX_OPTS_S)
$rxRoueLine = [regex]::new('ROUE\s+CODEUSE\s*=\s*([0-9A-Z])\s*\.\.\s*GAMME\s*=\s*(\d+)\s*\.\.\s*Voie\s+([UVW])', $RX_OPTS_I)
$rxPrecLine = [regex]::new('Pr.{1}cision[^<]*%[^<]*:</td>\s*<td[^>]*>([^<]+)</td>', $RX_OPTS_I -bor $RX_OPTS_S)
$rxReportStamp = [regex]::new('\[(\d{2})\s+(\d{2})\s+(\d{2})\]\[(\d{2})\s+(\d{2})\s+(\d{4})\]', $RX_OPTS_I)

# --- Fonctions Utilitaires ---

function Get-ReportTimestamp {
    param([string]$FileName)
    if (-not $FileName) { return $null }
    $match = $rxReportStamp.Match($FileName)
    if (-not $match.Success) { return $null }
    $hour = $match.Groups[1].Value
    $minute = $match.Groups[2].Value
    $second = $match.Groups[3].Value
    $day = $match.Groups[4].Value
    $month = $match.Groups[5].Value
    $year = $match.Groups[6].Value
    $stamp = "$day/$month/$year ${hour}:${minute}:${second}"
    try {
        return [datetime]::ParseExact($stamp, 'dd/MM/yyyy HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        return $null
    }
}

function Convert-PSObjectToHashtable {
    param([Parameter(ValueFromPipeline = $true)]$InputObject)
    
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $hash = @{}
        foreach ($prop in $InputObject.psobject.Properties) {
            $hash[$prop.Name] = Convert-PSObjectToHashtable $prop.Value
        }
        return $hash
    }
    elseif ($InputObject -is [object[]]) {
        $arr = @()
        foreach ($item in $InputObject) {
            $arr += Convert-PSObjectToHashtable $item
        }
        return $arr
    }
    else {
        return $InputObject
    }
}

function Get-Resistances {
    if (Test-Path $script:resistancePath) {
        try {
            $json = Get-Content $script:resistancePath -Raw -ErrorAction SilentlyContinue
            if ($json) {
                # Version modernisée (-AsHashtable existe)
                if ($PSVersionTable.PSVersion.Major -ge 6) {
                    return $json | ConvertFrom-Json -AsHashtable
                }
                else {
                    # Version compatible PowerShell 5.1
                    $obj = $json | ConvertFrom-Json
                    if ($obj) {
                        return Convert-PSObjectToHashtable $obj
                    }
                }
            }
        }
        catch {}
    }
    return @{}
}

function Save-Resistances {
    param($Dict)
    try {
        $Dict | ConvertTo-Json -Depth 10 | Set-Content $script:resistancePath -Force
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la sauvegarde des résistances : $_")
    }
}

# --- Fonctions de gestion des marquages CMS ---
$script:smdPath = Join-Path $PSScriptRoot "smd_markings.json"

function Get-SMDMarkings {
    if (Test-Path $script:smdPath) {
        try {
            $content = Get-Content $script:smdPath | Out-String
            return ($content | ConvertFrom-Json | Sort-Object R)
        }
        catch {
            Write-Warning "Erreur lors de la lecture du fichier de marquages CMS : $($_.Exception.Message)"
        }
    }
    # Valeurs par défaut si le fichier n'existe pas ou est corrompu
    return @(
        @{ R = 221; M = "-" }
        @{ R = 274; M = "-" }
        @{ R = 332; M = "51A" }
        @{ R = 383; M = "57A" }
        @{ R = 432; M = "62A" }
        @{ R = 475; M = "66A" }
        @{ R = 536; M = "71A" }
        @{ R = 576; M = "74A" }
        @{ R = 634; M = "78A" }
        @{ R = 681; M = "81A" }
        @{ R = 845; M = "-" }
        @{ R = 931; M = "-" }
        @{ R = 1000; M = "-" }
    )
}

function Save-SMDMarkings {
    param(
        [Parameter(Mandatory = $true)]
        $Data
    )
    try {
        $Data | ConvertTo-Json -Depth 100 | Set-Content $script:smdPath -Force
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la sauvegarde des marquages CMS : $($_.Exception.Message)", "Erreur de sauvegarde", 'OK', 'Error')
    }
}

function Get-FileText {
    param([string]$Path)
    $encodings = @([System.Text.Encoding]::UTF8,
        [System.Text.Encoding]::Default,
        [System.Text.Encoding]::GetEncoding(28591))
    foreach ($enc in $encodings) {
        try { return [System.IO.File]::ReadAllText($Path, $enc) } catch {}
    }
    try { return (Get-Content -Path $Path -Raw) } catch { return $null }
}

function Parse-VprtMeasurements {
    param([string]$Path)
    $text = Get-FileText -Path $Path
    if (-not $text) { return @() }

    $block = $null
    $mb = $rxBegin.Match($text)
    if ($mb.Success) {
        $start = $mb.Index + $mb.Length
        $me = $rxEnd.Match($text, $start)
        if ($me.Success) {
            $block = $text.Substring($start, $me.Index - $start)
        }
        else {
            $block = $text.Substring($start)
        }
    }
    else {
        $block = $text
    }

    $data = @()
    $match = $rxRoueLine.Match($block)
    while ($match.Success) {
        $roue = $match.Groups[1].Value
        $gamme = [int]$match.Groups[2].Value
        $voie = $match.Groups[3].Value.ToUpper()

        $precMatch = $rxPrecLine.Match($block, $match.Index)
        if ($precMatch.Success) {
            $values = $precMatch.Groups[1].Value -split '/' | ForEach-Object { ($_ -replace '\s+', '') }
            if ($values.Count -ge 4) {
                $precision = [double]($values[-1] -replace ',', '.')
                $data += [pscustomobject]@{
                    Roue      = $roue
                    Gamme     = $gamme
                    Voie      = $voie
                    Precision = $precision
                    Position  = "R=$roue G=$gamme"
                }
            }
        }
        $match = $match.NextMatch()
    }
    return $data
}

function Load-VprtReportsFromFolder {
    param([string]$FolderPath, [double]$LimitHigh, [double]$LimitLow)
    if (-not (Test-Path $FolderPath -PathType Container)) { return @() }
    $reports = @()
    $files = Get-ChildItem -Path $FolderPath -Filter *.htm* -File | Sort-Object @{
        Expression = {
            $ts = Get-ReportTimestamp $_.BaseName
            if ($ts) { $ts } else { $_.LastWriteTime }
        }
    }
    
    foreach ($file in $files) {
        $data = Parse-VprtMeasurements -Path $file.FullName
        if ($data.Count -gt 0) {
            $timestamp = Get-ReportTimestamp $file.BaseName
            if (-not $timestamp) { $timestamp = $file.LastWriteTime }
            
            # Calcul des stats immédiat pour le tableau de bord
            $maxVal = ($data | Measure-Object -Property Precision -Maximum).Maximum
            $minVal = ($data | Measure-Object -Property Precision -Minimum).Minimum
            $maxAbs = [Math]::Max([Math]::Abs($maxVal), [Math]::Abs($minVal))
            
            # Statut OK/NOK
            $isOk = ($maxVal -le $LimitHigh) -and ($minVal -ge $LimitLow)

            $reports += [pscustomobject]@{
                Name         = $file.Name
                Path         = $file.FullName
                Data         = $data
                Timestamp    = $timestamp
                MaxPrecision = $maxVal
                MinPrecision = $minVal
                MaxAbsError  = $maxAbs
                IsOk         = $isOk
            }
        }
    }
    return $reports
}

# --- Variables Globales ---
$script:limitHigh = 0.22
$script:limitLow = -0.22
$script:reportCache = @()
$script:reportColors = @{}
$script:resistancePath = Join-Path $PSScriptRoot "resistances.json"
$script:resistanceDict = Get-Resistances
$script:smdPath = Join-Path $PSScriptRoot "smd_markings.json"
$script:smdDict = Get-SMDMarkings

# Liste fixe des 15 positions pour l'axe X
$script:StandardPositions = @(
    "R=F G=1", "R=E G=2", "R=D G=3", "R=C G=4", "R=B G=5",
    "R=A G=6", "R=9 G=7", "R=8 G=8", "R=7 G=9", "R=6 G=10",
    "R=5 G=11", "R=4 G=12", "R=3 G=13", "R=2 G=14", "R=1 G=15"
)

# --- Fonctions Graphiques ---

function Set-FixedYAxis {
    param([System.Windows.Forms.DataVisualization.Charting.ChartArea]$Area)
    if (-not $Area) { return }
    $Area.AxisY.Minimum = -0.25
    $Area.AxisY.Maximum = 0.25
    $Area.AxisY.Interval = 0.05
    $Area.AxisY.MajorGrid.LineDashStyle = 'Dot'
    $Area.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::LightGray
    $Area.AxisY.StripLines.Clear()
    
    # Ligne Zéro plus visible
    $stripZero = New-Object System.Windows.Forms.DataVisualization.Charting.StripLine
    $stripZero.IntervalOffset = 0
    $stripZero.StripWidth = 0.002
    $stripZero.BackColor = 'Black'
    $Area.AxisY.StripLines.Add($stripZero)
}

function New-BaseChart {
    param([string]$Title)
    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.BackColor = [System.Drawing.Color]::White
    $chart.BorderlineColor = [System.Drawing.Color]::LightGray
    $chart.BorderlineDashStyle = 'Solid'
    $chart.BorderlineWidth = 1
    
    $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $area.Name = 'main'
    $area.BackColor = [System.Drawing.Color]::White
    $area.AxisX.LabelStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $area.AxisY.LabelStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $area.AxisX.MajorGrid.LineColor = [System.Drawing.Color]::LightGray
    $area.AxisX.MajorGrid.LineDashStyle = 'Dot'
    $chart.ChartAreas.Add($area) | Out-Null
    
    # Ajout Légende
    $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
    $legend.Name = "Legend1"
    $legend.Docking = "Top"
    $legend.Alignment = "Center"
    $legend.LegendStyle = "Row"
    $chart.Legends.Add($legend) | Out-Null
    
    # --- Interaction : CTRL + Roulette pour changer la hauteur ---
    $chart.Add_MouseEnter({ $this.Focus() })
    $chart.Add_MouseWheel({
            if ([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Control) {
                $step = 50
                $newHeight = $this.Height + ([Math]::Sign($_.Delta) * $step)
            
                # Limites de hauteur (Min 150px, Max 2000px)
                if ($newHeight -lt 150) { $newHeight = 150 }
                if ($newHeight -gt 2000) { $newHeight = 2000 }
            
                $this.Height = $newHeight
            }
        })

    if (-not [string]::IsNullOrWhiteSpace($Title)) {
        $t = New-Object System.Windows.Forms.DataVisualization.Charting.Title
        $t.Text = $Title
        $t.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        $chart.Titles.Add($t) | Out-Null
    }
    
    return $chart
}

# --- IHM ---

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Analyse VPRT 2.0 (SEQ-02)'
$form.Size = New-Object System.Drawing.Size(1400, 850)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Split Principal
$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'
$split.FixedPanel = 'Panel1'
$split.SplitterDistance = 520
$split.SplitterWidth = 5
$split.Panel1MinSize = 480
$form.Controls.Add($split)

# --- Panneau Gauche (Contrôles + Liste) ---
$panelLeft = New-Object System.Windows.Forms.Panel
$panelLeft.Dock = 'Fill'
$panelLeft.Padding = New-Object System.Windows.Forms.Padding(10)
$split.Panel1.Controls.Add($panelLeft)

# Zone Dossier
$grpFolder = New-Object System.Windows.Forms.GroupBox
$grpFolder.Text = "Source des données"
$grpFolder.Height = 100
$grpFolder.Dock = 'Top'
$grpFolder.Padding = New-Object System.Windows.Forms.Padding(10)
$panelLeft.Controls.Add($grpFolder)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = '📂 Choisir dossier...'
$btnBrowse.Location = New-Object System.Drawing.Point(10, 25)
$btnBrowse.Size = New-Object System.Drawing.Size(120, 30)
$btnBrowse.FlatStyle = 'System'
$grpFolder.Controls.Add($btnBrowse)

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Location = New-Object System.Drawing.Point(140, 27)
$txtFolder.Size = New-Object System.Drawing.Size(370, 25)
$txtFolder.Anchor = 'Top, Left, Right'
$txtFolder.ReadOnly = $false
$grpFolder.Controls.Add($txtFolder)

$btnNext = New-Object System.Windows.Forms.Button
$btnNext.Text = '▶ Dossier Suivant'
$btnNext.Height = 30
$btnNext.Width = 120
$btnNext.Dock = 'Right'
$btnNext.FlatStyle = 'Standard'
$btnNext.Cursor = [System.Windows.Forms.Cursors]::Hand

$lblCurrentFolder = New-Object System.Windows.Forms.Label
$lblCurrentFolder.Text = "N° de série : -"
$lblCurrentFolder.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblCurrentFolder.ForeColor = [System.Drawing.Color]::DarkBlue
$lblCurrentFolder.AutoSize = $false
$lblCurrentFolder.Dock = 'Fill'
$lblCurrentFolder.TextAlign = 'MiddleLeft'

$pnlBottomFolder = New-Object System.Windows.Forms.Panel
$pnlBottomFolder.Height = 30
$pnlBottomFolder.Dock = 'Bottom'
$pnlBottomFolder.Controls.Add($lblCurrentFolder)
$pnlBottomFolder.Controls.Add($btnNext)

$grpFolder.Controls.Add($pnlBottomFolder)

# Zone Boutons des Sous-Dossiers (N° Série)
$pnlSeries = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlSeries.Dock = 'Top'
$pnlSeries.Height = 150
$pnlSeries.AutoScroll = $true
$pnlSeries.Padding = New-Object System.Windows.Forms.Padding(5)
$pnlSeries.BackColor = [System.Drawing.Color]::WhiteSmoke
$panelLeft.Controls.Add($pnlSeries)
$pnlSeries.BringToFront()

# Zone Limites
$grpLimits = New-Object System.Windows.Forms.GroupBox
$grpLimits.Text = "Critères d'acceptation (%)"
$grpLimits.Height = 60
$grpLimits.Dock = 'Top'
$panelLeft.Controls.Add($grpLimits)
$grpLimits.BringToFront() # Pour l'ordre d'affichage (dessous le pnlSeries)

$lblH = New-Object System.Windows.Forms.Label
$lblH.Text = "Max :"
$lblH.AutoSize = $true
$lblH.Location = New-Object System.Drawing.Point(10, 25)
$grpLimits.Controls.Add($lblH)

$nudHigh = New-Object System.Windows.Forms.NumericUpDown
$nudHigh.Location = New-Object System.Drawing.Point(50, 22)
$nudHigh.DecimalPlaces = 2
$nudHigh.Increment = 0.01
$nudHigh.Minimum = -10
$nudHigh.Maximum = 10
$nudHigh.Value = 0.22
$grpLimits.Controls.Add($nudHigh)

$lblL = New-Object System.Windows.Forms.Label
$lblL.Text = "Min :"
$lblL.AutoSize = $true
$lblL.Location = New-Object System.Drawing.Point(140, 25)
$grpLimits.Controls.Add($lblL)

$nudLow = New-Object System.Windows.Forms.NumericUpDown
$nudLow.Location = New-Object System.Drawing.Point(180, 22)
$nudLow.DecimalPlaces = 2
$nudLow.Increment = 0.01
$nudLow.Minimum = -10
$nudLow.Maximum = 10
$nudLow.Value = -0.22
$grpLimits.Controls.Add($nudLow)

$pnlOptions = New-Object System.Windows.Forms.Panel
$pnlOptions.Height = 40
$pnlOptions.Dock = 'Top'
$pnlOptions.Padding = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
$panelLeft.Controls.Add($pnlOptions)
$pnlOptions.BringToFront()

$btnSMDMarking = New-Object System.Windows.Forms.Button
$btnSMDMarking.Text = "🔗 Editer Codes CMS"
$btnSMDMarking.Location = New-Object System.Drawing.Point(10, 5)
$btnSMDMarking.Size = New-Object System.Drawing.Size(150, 30)
$btnSMDMarking.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnSMDMarking.FlatStyle = 'Standard'
$btnSMDMarking.Add_Click({ Show-SMDMarkingWindow })
$pnlOptions.Controls.Add($btnSMDMarking)

# Zone Liste (DataGridView)
$dgvReports = New-Object System.Windows.Forms.DataGridView
$dgvReports.Dock = 'Fill'
$dgvReports.AllowUserToAddRows = $false
$dgvReports.AllowUserToDeleteRows = $false
$dgvReports.RowHeadersVisible = $false
$dgvReports.SelectionMode = 'FullRowSelect'
$dgvReports.MultiSelect = $true
$dgvReports.ReadOnly = $true
$dgvReports.AutoSizeColumnsMode = 'Fill'
$dgvReports.BackgroundColor = [System.Drawing.Color]::WhiteSmoke
$dgvReports.BorderStyle = 'None'

# Colonnes
$colIcon = New-Object System.Windows.Forms.DataGridViewImageColumn
$colIcon.Name = "Status"
$colIcon.HeaderText = ""
$colIcon.Width = 30
$colIcon.AutoSizeMode = 'None'
$dgvReports.Columns.Add($colIcon) | Out-Null

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = "Rapport"
$colName.HeaderText = "Rapport"
$colName.FillWeight = 50
$dgvReports.Columns.Add($colName) | Out-Null

$colDate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDate.Name = "Date"
$colDate.HeaderText = "Date"
$colDate.FillWeight = 30
$dgvReports.Columns.Add($colDate) | Out-Null

$colRes = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colRes.Name = "Resistances"
$colRes.HeaderText = "Résistances (U,V,W)"
$colRes.FillWeight = 30
$colRes.DefaultCellStyle.Alignment = 'MiddleLeft'
$dgvReports.Columns.Add($colRes) | Out-Null

$pnlGrid = New-Object System.Windows.Forms.Panel
$pnlGrid.Dock = 'Fill'
$pnlGrid.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
$pnlGrid.Controls.Add($dgvReports)
$panelLeft.Controls.Add($pnlGrid)
$pnlGrid.BringToFront()

# --- Panneau Droit (Onglets) ---
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$split.Panel2.Controls.Add($tabControl)

# Onglet 1 : Profils (Graphiques détaillés)
$tabProfiles = New-Object System.Windows.Forms.TabPage "📊 Profils Détaillés"
$tabControl.TabPages.Add($tabProfiles) | Out-Null

$panelCharts = New-Object System.Windows.Forms.FlowLayoutPanel
$panelCharts.Dock = 'Fill'
$panelCharts.AutoScroll = $true
$panelCharts.FlowDirection = 'TopDown'
$panelCharts.WrapContents = $false
$panelCharts.BackColor = [System.Drawing.Color]::White
$tabProfiles.Controls.Add($panelCharts)

# Barre d'outils onglet Profils
$pnlTools = New-Object System.Windows.Forms.Panel
$pnlTools.Dock = 'Top'
$pnlTools.Height = 40
$pnlTools.BackColor = [System.Drawing.Color]::WhiteSmoke
$tabProfiles.Controls.Add($pnlTools)

$cbDisplayMode = New-Object System.Windows.Forms.ComboBox
$cbDisplayMode.Items.AddRange(@('Vue Multi-Rapports', 'Comparaison par Voie'))
$cbDisplayMode.SelectedIndex = 0
$cbDisplayMode.Location = New-Object System.Drawing.Point(10, 8)
$cbDisplayMode.Width = 200
$cbDisplayMode.DropDownStyle = 'DropDownList'
$pnlTools.Controls.Add($cbDisplayMode)

$cbVoie = New-Object System.Windows.Forms.ComboBox
$cbVoie.Items.AddRange(@('U', 'V', 'W'))
$cbVoie.SelectedIndex = 0
$cbVoie.Location = New-Object System.Drawing.Point(220, 8)
$cbVoie.Width = 50
$cbVoie.DropDownStyle = 'DropDownList'
$cbVoie.Enabled = $false
$pnlTools.Controls.Add($cbVoie)

# Onglet 2 : Tendance (Evolution temporelle)
$tabTrend = New-Object System.Windows.Forms.TabPage "📈 Tendance / Évolution"
$tabControl.TabPages.Add($tabTrend) | Out-Null

$chartTrend = New-BaseChart "Évolution de l'Erreur Max Absolue"
$chartTrend.Dock = 'Fill'
$tabTrend.Controls.Add($chartTrend)

# Onglet 3 : Données Brutes
$tabData = New-Object System.Windows.Forms.TabPage "📋 Données Brutes"
$tabControl.TabPages.Add($tabData) | Out-Null

$dgvData = New-Object System.Windows.Forms.DataGridView
$dgvData.Dock = 'Fill'
$dgvData.ReadOnly = $true
$dgvData.AutoSizeColumnsMode = 'Fill'
$tabData.Controls.Add($dgvData)

# --- Logique Applicative ---

# Images pour le statut
$bmpOK = New-Object System.Drawing.Bitmap(16, 16)
$gOK = [System.Drawing.Graphics]::FromImage($bmpOK)
$gOK.FillEllipse([System.Drawing.Brushes]::Green, 2, 2, 12, 12)
$gOK.Dispose()

$bmpNOK = New-Object System.Drawing.Bitmap(16, 16)
$gNOK = [System.Drawing.Graphics]::FromImage($bmpNOK)
$gNOK.FillEllipse([System.Drawing.Brushes]::Red, 2, 2, 12, 12)
$gNOK.Dispose()

$UpdateTrend = {
    $chartTrend.Series.Clear()
    $chartTrend.ChartAreas['main'].AxisX.LabelStyle.Format = "dd/MM"
    $chartTrend.ChartAreas['main'].AxisX.IntervalType = 'Days'
    
    # Séries pour U, V, W
    foreach ($v in 'U', 'V', 'W') {
        $s = New-Object System.Windows.Forms.DataVisualization.Charting.Series "Voie $v (Max Abs)"
        $s.ChartType = 'Line'
        $s.BorderWidth = 3
        $s.MarkerStyle = 'Circle'
        $s.MarkerSize = 8
        $s.ToolTip = "Date: #VALX{dd/MM/yy}`nMax Err: #VALY{N3}%"
        $chartTrend.Series.Add($s) | Out-Null
    }
    
    # Limites
    $sHigh = New-Object System.Windows.Forms.DataVisualization.Charting.Series "Limite +"
    $sHigh.ChartType = 'Line'
    $sHigh.Color = 'Red'
    $sHigh.BorderDashStyle = 'Dash'
    $chartTrend.Series.Add($sHigh) | Out-Null
    
    $sLow = New-Object System.Windows.Forms.DataVisualization.Charting.Series "Limite -"
    $sLow.ChartType = 'Line'
    $sLow.Color = 'Red'
    $sLow.BorderDashStyle = 'Dash'
    $chartTrend.Series.Add($sLow) | Out-Null

    # Données triées par date
    $sortedReports = $script:reportCache | Sort-Object Timestamp
    
    foreach ($r in $sortedReports) {
        if (-not $r.Timestamp) { continue }
        
        # Trouver le Max Abs par voie pour ce rapport
        $grp = $r.Data | Group-Object Voie
        foreach ($g in $grp) {
            $maxAbs = ($g.Group | Select-Object @{N = 'A'; E = { [Math]::Abs($_.Precision) } } | Measure-Object A -Maximum).Maximum
            $chartTrend.Series["Voie $($g.Name) (Max Abs)"].Points.AddXY($r.Timestamp, $maxAbs) | Out-Null
        }
        
        $chartTrend.Series["Limite +"].Points.AddXY($r.Timestamp, [Math]::Abs($script:limitHigh)) | Out-Null
        $chartTrend.Series["Limite -"].Points.AddXY($r.Timestamp, 0) | Out-Null # On affiche l'absolu donc 0 à LimitHigh
    }
    
    $chartTrend.ChartAreas['main'].AxisY.Title = "Erreur Absolue Max (%)"
}

$UpdateCharts = {
    if ($script:disableChartUpdate) { return }
    
    $panelCharts.Controls.Clear()
    $panelCharts.SuspendLayout()
    
    $selectedRows = $dgvReports.SelectedRows
    if ($selectedRows.Count -eq 0) { 
        $panelCharts.ResumeLayout()
        return 
    }
    
    $reportsToShow = @()
    foreach ($row in $selectedRows) {
        $rName = $row.Cells["Rapport"].Value
        $r = $script:reportCache | Where-Object Name -eq $rName | Select-Object -First 1
        if ($r) { $reportsToShow += $r }
    }
    
    # Tri chronologique (Ancien -> Récent)
    $reportsToShow = $reportsToShow | Sort-Object Timestamp
    
    # Utilisation des positions fixes
    $allPos = $script:StandardPositions

    # Mode Comparaison
    if ($cbDisplayMode.SelectedIndex -eq 1) {
        $voie = $cbVoie.SelectedItem
        $chart = New-BaseChart "Comparaison Voie $voie"
        
        $targetHeight = [math]::Floor(($panelCharts.ClientSize.Height - 20) / 2)
        if ($targetHeight -lt 200) { $targetHeight = 200 }
        
        $chart.Height = $targetHeight
        $chart.Width = $panelCharts.ClientSize.Width - 30
        Set-FixedYAxis $chart.ChartAreas['main']
        
        # Limites
        $slH = New-Object System.Windows.Forms.DataVisualization.Charting.Series "Lim+"
        $slH.ChartType = 'Line'; $slH.Color = 'Red'; $slH.BorderDashStyle = 'Dash'
        $slH.IsVisibleInLegend = $false
        $chart.Series.Add($slH) | Out-Null
        
        $slL = New-Object System.Windows.Forms.DataVisualization.Charting.Series "Lim-"
        $slL.ChartType = 'Line'; $slL.Color = 'Red'; $slL.BorderDashStyle = 'Dash'
        $slL.IsVisibleInLegend = $false
        $chart.Series.Add($slL) | Out-Null
        
        # Tracer Limites et Labels X
        for ($i = 0; $i -lt $allPos.Count; $i++) {
            $slH.Points.AddXY($i + 1, $script:limitHigh) | Out-Null
            $slL.Points.AddXY($i + 1, $script:limitLow) | Out-Null
            $lbl = New-Object System.Windows.Forms.DataVisualization.Charting.CustomLabel
            $lbl.FromPosition = $i + 0.5; $lbl.ToPosition = $i + 1.5; $lbl.Text = $allPos[$i]
            $chart.ChartAreas['main'].AxisX.CustomLabels.Add($lbl) | Out-Null
        }
        
        # Tracer Courbes
        foreach ($rep in $reportsToShow) {
            $s = New-Object System.Windows.Forms.DataVisualization.Charting.Series $rep.Name
            $s.ChartType = 'Line'
            $s.BorderWidth = 2
            $chart.Series.Add($s) | Out-Null
            
            for ($i = 0; $i -lt $allPos.Count; $i++) {
                $p = $rep.Data | Where-Object { $_.Position -eq $allPos[$i] -and $_.Voie -eq $voie } | Select-Object -First 1
                if ($p) { $s.Points.AddXY($i + 1, $p.Precision) | Out-Null }
                else { $s.Points.AddXY($i + 1, [double]::NaN) | Out-Null }
            }
        }
        $panelCharts.Controls.Add($chart)
    } 
    # Mode Multi-Rapports
    else {
        $prevRes = @{ U = $null; V = $null; W = $null }
        $isFirstReport = $true
        
        foreach ($rep in $reportsToShow) {
            $chart = New-BaseChart $rep.Name
            
            $targetHeight = [math]::Floor(($panelCharts.ClientSize.Height - 20) / 2)
            if ($targetHeight -lt 200) { $targetHeight = 200 }
            
            $chart.Height = $targetHeight
            $chart.Width = $panelCharts.ClientSize.Width - 30
            Set-FixedYAxis $chart.ChartAreas['main']
            
            # Limites
            $sH = New-Object System.Windows.Forms.DataVisualization.Charting.Series "L+"; $sH.ChartType = 'Line'; $sH.Color = 'Red'; $sH.BorderDashStyle = 'Dash'; $sH.IsVisibleInLegend = $false; $chart.Series.Add($sH) | Out-Null
            $sL = New-Object System.Windows.Forms.DataVisualization.Charting.Series "L-"; $sL.ChartType = 'Line'; $sL.Color = 'Red'; $sL.BorderDashStyle = 'Dash'; $sL.IsVisibleInLegend = $false; $chart.Series.Add($sL) | Out-Null
            
            $res = $script:resistanceDict[$rep.Name]
            foreach ($v in 'U', 'V', 'W') {
                $seriesName = "Voie $v"
                if ($res -and $res.$v) {
                    if (-not $isFirstReport -and $prevRes[$v] -ne $null -and $prevRes[$v] -ne $res.$v) {
                        $seriesName += " ($($prevRes[$v]) -> $($res.$v)Ω)"
                    }
                    else {
                        $seriesName += " ($($res.$v)Ω)"
                    }
                    $prevRes[$v] = $res.$v
                }
                
                $s = New-Object System.Windows.Forms.DataVisualization.Charting.Series $seriesName
                $s.ChartType = 'Line'; $s.BorderWidth = 2
                $s.IsVisibleInLegend = $true
                if ($v -eq 'U') { $s.Color = 'SteelBlue' }
                elseif ($v -eq 'V') { $s.Color = 'DarkOrange' }
                else { $s.Color = 'ForestGreen' }
                $chart.Series.Add($s) | Out-Null
                
                for ($i = 0; $i -lt $allPos.Count; $i++) {
                    $pt = $rep.Data | Where-Object { $_.Position -eq $allPos[$i] -and $_.Voie -eq $v } | Select-Object -First 1
                    if ($pt) { 
                        $idx = $s.Points.AddXY($i + 1, $pt.Precision)
                        $s.Points[$idx].ToolTip = "$($allPos[$i]) : $($pt.Precision)%"
                    }
                    else { $s.Points.AddXY($i + 1, [double]::NaN) | Out-Null }
                }
            }
            
            # Labels X et Limites Points
            for ($i = 0; $i -lt $allPos.Count; $i++) {
                $sH.Points.AddXY($i + 1, $script:limitHigh) | Out-Null
                $sL.Points.AddXY($i + 1, $script:limitLow) | Out-Null
                $lbl = New-Object System.Windows.Forms.DataVisualization.Charting.CustomLabel
                $lbl.FromPosition = $i + 0.5; $lbl.ToPosition = $i + 1.5; $lbl.Text = $allPos[$i]
                $chart.ChartAreas['main'].AxisX.CustomLabels.Add($lbl) | Out-Null
            }
            
            $panelCharts.Controls.Add($chart)
            $isFirstReport = $false
        }
    }
    
    $panelCharts.ResumeLayout()
}

# --- Mise à jour de la grille de boutons de série ---
$UpdateSeriesButtons = {
    param([string]$currentPath)
    if ([string]::IsNullOrWhiteSpace($currentPath) -or -not (Test-Path $currentPath)) { return }
    
    try {
        $parent = [System.IO.Directory]::GetParent($currentPath)
        if (-not $parent) { return }
        
        $siblings = Get-ChildItem -Path $parent.FullName -Directory | Sort-Object Name
        
        # Suspendre l'affichage pour éviter les clignotements
        $pnlSeries.SuspendLayout()
        
        # Nettoyer les anciens boutons
        foreach ($ctrl in $pnlSeries.Controls) { $ctrl.Dispose() }
        $pnlSeries.Controls.Clear()
        
        # Créer les nouveaux boutons
        foreach ($folder in $siblings) {
            $btn = New-Object System.Windows.Forms.Button
            $btn.Text = $folder.Name
            # Largeur calculée pour afficher environ 8 boutons sur une largeur de 480px avec marges
            $btn.Size = New-Object System.Drawing.Size(55, 30)
            $btn.FlatStyle = 'Flat'
            $btn.Margin = New-Object System.Windows.Forms.Padding(2)
            $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
            $btn.Tag = $folder.FullName # Stocker le chemin complet dans le tag
            
            # Mettre en évidence le dossier actuellement sélectionné
            if ($folder.FullName -eq $currentPath) {
                $btn.BackColor = [System.Drawing.Color]::SteelBlue
                $btn.ForeColor = [System.Drawing.Color]::White
                $btn.Font = New-Object System.Drawing.Font($btn.Font, [System.Drawing.FontStyle]::Bold)
            }
            else {
                $btn.BackColor = [System.Drawing.Color]::White
                $btn.ForeColor = [System.Drawing.Color]::Black
            }
            
            # Événement de clic pour charger ce dossier spécifique
            $btn.Add_Click({
                    $txtFolder.Text = $this.Tag
                    & $LoadAction
                })
            
            $pnlSeries.Controls.Add($btn)
        }
    }
    catch {
        # Ignorer silencieusement les erreurs de lecture de dossier pour ne pas bloquer l'UI
    }
    finally {
        $pnlSeries.ResumeLayout()
    }
}

function Show-ResistanceDialog {
    param([string]$ReportName, $CurrentValues)
    $diag = New-Object System.Windows.Forms.Form
    $diag.Text = "Saisie des résistances - $ReportName"
    $diag.Size = New-Object System.Drawing.Size(300, 240)
    $diag.StartPosition = 'CenterParent'
    $diag.FormBorderStyle = 'FixedDialog'
    $diag.MaximizeBox = $false
    $diag.MinimizeBox = $false
    $diag.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Valeurs de résistance (Ohms) :"
    $lblInfo.Location = New-Object System.Drawing.Point(20, 10)
    $lblInfo.Size = New-Object System.Drawing.Size(250, 20)
    $diag.Controls.Add($lblInfo)

    # Extraction des valeurs normalisées depuis smdDict (compatible hashtable et PSCustomObject)
    $normValues = @($script:smdDict | ForEach-Object { $_.R } | Sort-Object)

    # U
    $lblU = New-Object System.Windows.Forms.Label
    $lblU.Text = "Voie U :"
    $lblU.Location = New-Object System.Drawing.Point(30, 45)
    $lblU.AutoSize = $true
    $diag.Controls.Add($lblU)
    $cbU = New-Object System.Windows.Forms.ComboBox
    $cbU.Location = New-Object System.Drawing.Point(100, 42)
    $cbU.DropDownStyle = 'DropDownList'
    $cbU.Items.AddRange($normValues)
    if ($CurrentValues.U -in $normValues) { $cbU.SelectedItem = [int]$CurrentValues.U }
    $diag.Controls.Add($cbU)

    # V
    $lblV = New-Object System.Windows.Forms.Label
    $lblV.Text = "Voie V :"
    $lblV.Location = New-Object System.Drawing.Point(30, 80)
    $lblV.AutoSize = $true
    $diag.Controls.Add($lblV)
    $cbV = New-Object System.Windows.Forms.ComboBox
    $cbV.Location = New-Object System.Drawing.Point(100, 77)
    $cbV.DropDownStyle = 'DropDownList'
    $cbV.Items.AddRange($normValues)
    if ($CurrentValues.V -in $normValues) { $cbV.SelectedItem = [int]$CurrentValues.V }
    $diag.Controls.Add($cbV)

    # W
    $lblW = New-Object System.Windows.Forms.Label
    $lblW.Text = "Voie W :"
    $lblW.Location = New-Object System.Drawing.Point(30, 115)
    $lblW.AutoSize = $true
    $diag.Controls.Add($lblW)
    $cbW = New-Object System.Windows.Forms.ComboBox
    $cbW.Location = New-Object System.Drawing.Point(100, 112)
    $cbW.DropDownStyle = 'DropDownList'
    $cbW.Items.AddRange($normValues)
    if ($CurrentValues.W -in $normValues) { $cbW.SelectedItem = [int]$CurrentValues.W }
    $diag.Controls.Add($cbW)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Valider"
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOk.Location = New-Object System.Drawing.Point(50, 160)
    $diag.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.Location = New-Object System.Drawing.Point(150, 160)
    $diag.Controls.Add($btnCancel)

    $diag.AcceptButton = $btnOk
    $diag.CancelButton = $btnCancel

    if ($diag.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return @{ U = $cbU.SelectedItem; V = $cbV.SelectedItem; W = $cbW.SelectedItem }
    }
    return $null
}

function Show-SMDMarkingWindow {
    $script:smdForm = New-Object System.Windows.Forms.Form
    $script:smdForm.Text = "Marquage CMS des Résistances"
    $script:smdForm.Size = New-Object System.Drawing.Size(350, 450)
    $script:smdForm.StartPosition = 'CenterScreen'
    $script:smdForm.FormBorderStyle = 'FixedDialog'
    $script:smdForm.MaximizeBox = $false
    $script:smdForm.MinimizeBox = $false
    
    $pnlBottom = New-Object System.Windows.Forms.Panel
    $pnlBottom.Height = 40
    $pnlBottom.Dock = 'Bottom'
    $script:smdForm.Controls.Add($pnlBottom)
    
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Enregistrer"
    $btnSave.Location = New-Object System.Drawing.Point(120, 5)
    $btnSave.Size = New-Object System.Drawing.Size(100, 30)
    $pnlBottom.Controls.Add($btnSave)
    
    $script:smdDgv = New-Object System.Windows.Forms.DataGridView
    $script:smdDgv.Dock = 'Fill'
    $script:smdDgv.AllowUserToAddRows = $true
    $script:smdDgv.AllowUserToDeleteRows = $true
    $script:smdDgv.ReadOnly = $false
    $script:smdDgv.RowHeadersVisible = $true
    $script:smdDgv.AutoSizeColumnsMode = 'Fill'
    $script:smdDgv.BackgroundColor = [System.Drawing.Color]::White
    $script:smdDgv.SelectionMode = 'FullRowSelect'
    
    $col1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col1.Name = "Resistance"
    $col1.HeaderText = "Résistance (Ω)"
    $col1.ValueType = [int]
    $script:smdDgv.Columns.Add($col1) | Out-Null
    
    $col2 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col2.Name = "Marking"
    $col2.HeaderText = "Marquage CMS (séparés par ,)"
    $script:smdDgv.Columns.Add($col2) | Out-Null
    
    # Remplir avec les données actuelles
    foreach ($item in $script:smdDict) {
        $idx = $script:smdDgv.Rows.Add()
        $script:smdDgv.Rows[$idx].Cells["Resistance"].Value = $item.R
        $script:smdDgv.Rows[$idx].Cells["Marking"].Value = $item.M
    }
    
    $script:smdForm.Controls.Add($script:smdDgv)
    
    # Événement de sauvegarde - utilise uniquement des variables $script: (pas de closure)
    $btnSave.Add_Click({
            $script:smdDgv.EndEdit()
            $newData = @()
            foreach ($row in $script:smdDgv.Rows) {
                if (-not $row.IsNewRow -and $null -ne $row.Cells["Resistance"].Value) {
                    $rVal = [int]$row.Cells["Resistance"].Value
                    $mVal = $row.Cells["Marking"].Value
                    if (-not $mVal) { $mVal = "-" }
                    $newData += @{ "R" = $rVal; "M" = [string]$mVal }
                }
            }
            $newData = @($newData | Sort-Object { $_.R })
            if ($newData.Count -gt 0) {
                $script:smdDict = $newData
                try {
                    $jsonStr = $newData | ConvertTo-Json -Depth 10
                    [System.IO.File]::WriteAllText($script:smdPath, $jsonStr, [System.Text.Encoding]::UTF8)
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Erreur lors de la sauvegarde : $_", "Erreur", 'OK', 'Error')
                }
            }
            $script:smdForm.Close()
        })
    
    # Rend la fenêtre "volante" pour ne pas la perdre derrière
    $script:smdForm.TopMost = $true
    $script:smdForm.Show() | Out-Null
}

$LoadAction = {
    $path = $txtFolder.Text
    if (-not (Test-Path $path)) { return }
    
    $script:limitHigh = [double]$nudHigh.Value
    $script:limitLow = [double]$nudLow.Value
    
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $script:disableChartUpdate = $true
    
    $reports = Load-VprtReportsFromFolder -FolderPath $path -LimitHigh $script:limitHigh -LimitLow $script:limitLow
    
    # Si le dossier sélectionné ne contient pas de rapports HTML mais contient des sous-dossiers (cas du dossier racine)
    if ($reports.Count -eq 0) {
        $subfolders = Get-ChildItem -Path $path -Directory | Sort-Object Name
        if ($subfolders.Count -gt 0) {
            # On passe automatiquement au premier sous-dossier
            $txtFolder.Text = $subfolders[0].FullName
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            & $LoadAction
            return
        }
    }
    
    # Mise à jour du label avec le nom du dossier courant (N° de série)
    $folderName = Split-Path $txtFolder.Text -Leaf
    $lblCurrentFolder.Text = "N° de série : $folderName"
    
    # Mettre à jour les boutons de série basés sur le vrai chemin courant
    & $UpdateSeriesButtons $txtFolder.Text
    
    $script:reportCache = $reports
    
    $dgvReports.Rows.Clear()
    foreach ($r in $reports) {
        $idx = $dgvReports.Rows.Add()
        $row = $dgvReports.Rows[$idx]
        $row.Cells["Rapport"].Value = $r.Name
        $row.Cells["Date"].Value = if ($r.Timestamp) { $r.Timestamp.ToString("dd/MM/yyyy HH:mm") } else { "-" }
        
        # Afficher les résistances si elles existent
        $res = $script:resistanceDict[$r.Name]
        if ($res) {
            $row.Cells["Resistances"].Value = "U:$($res.U) V:$($res.V) W:$($res.W)"
        }
        else {
            $row.Cells["Resistances"].Value = "Cliq. pour saisir..."
        }
        
        $row.Cells["Status"].Value = if ($r.IsOk) { $bmpOK } else { $bmpNOK }
        
        if (-not $r.IsOk) { $row.DefaultCellStyle.ForeColor = 'DarkRed' }
    }
    
    # Update Data Grid (Raw)
    $allData = $reports | Select-Object -ExpandProperty Data
    $dgvData.DataSource = [System.Collections.ArrayList]$allData
    
    # Update Trend
    & $UpdateTrend
    
    # --- 
    # Libérer le verrou de rafraîchissement des graphiques AVANT de sélectionner
    # pour que SelectionChanged enregistre bien les nouvelles sélections !
    $script:disableChartUpdate = $false
    
    # Sélectionner toutes les lignes par défaut
    $dgvReports.SelectAll()
    
    # Forcer la mise à jour (au cas où SelectAll() ne modifierait rien si 1 seule ligne déjà auto-sélectionnée)
    & $UpdateCharts
    
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}

# --- Events ---
$script:firstFolderPickerOpen = $true
$script:currentScriptDir = $PSScriptRoot
if (-not $script:currentScriptDir -and $MyInvocation.MyCommand.Path) {
    $script:currentScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
}
if (-not $script:currentScriptDir) {
    $script:currentScriptDir = [Environment]::CurrentDirectory
}

$btnBrowse.Add_Click({
        $picker = New-Object FolderPicker
        $picker.Title = "Sélectionner un dossier"
        
        if ($script:firstFolderPickerOpen -or [string]::IsNullOrWhiteSpace($txtFolder.Text)) {
            $picker.InputPath = $script:currentScriptDir
            $script:firstFolderPickerOpen = $false
        }
        else {
            try {
                $parent = [System.IO.Directory]::GetParent($txtFolder.Text)
                if ($parent) {
                    $picker.InputPath = $parent.FullName
                }
                else {
                    $picker.InputPath = $txtFolder.Text
                }
            }
            catch {
                $picker.InputPath = $txtFolder.Text
            }
        }
    
        if ($picker.ShowDialog($form.Handle)) {
            $txtFolder.Text = $picker.ResultPath
            & $LoadAction
        }
    })

$btnNext.Add_Click({
        $currentPath = $txtFolder.Text
        if ([string]::IsNullOrWhiteSpace($currentPath) -or -not (Test-Path $currentPath)) { return }
        
        try {
            $parent = [System.IO.Directory]::GetParent($currentPath)
            if ($parent) {
                # Récupérer tous les sous-dossiers du parent, triés par nom
                $siblings = Get-ChildItem -Path $parent.FullName -Directory | Sort-Object Name
                
                $currentIndex = -1
                for ($i = 0; $i -lt $siblings.Count; $i++) {
                    if ($siblings[$i].FullName -eq $currentPath) {
                        $currentIndex = $i
                        break
                    }
                }
                
                # Si on a trouvé le dossier courant
                if ($currentIndex -ge 0) {
                    if ($currentIndex -lt ($siblings.Count - 1)) {
                        # Dossier suivant
                        $nextFolder = $siblings[$currentIndex + 1].FullName
                    }
                    else {
                        # Revenir au premier dossier de la liste
                        $nextFolder = $siblings[0].FullName
                    }
                    $txtFolder.Text = $nextFolder
                    & $LoadAction
                }
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erreur lors de la recherche du dossier suivant : $_", "Erreur", 'OK', 'Error')
        }
    })

$script:lastValidSelection = @()
$script:disableChartUpdate = $false
$script:isRestoringSelection = $false

$dgvReports.Add_SelectionChanged({
        if ($script:isRestoringSelection) { return }

        if ($script:disableChartUpdate) {
            $script:isRestoringSelection = $true
            $dgvReports.ClearSelection()
            foreach ($idx in $script:lastValidSelection) {
                if ($idx -lt $dgvReports.Rows.Count) {
                    $dgvReports.Rows[$idx].Selected = $true
                }
            }
            $script:isRestoringSelection = $false
        }
        else {
            $script:lastValidSelection = @()
            foreach ($r in $dgvReports.SelectedRows) {
                $script:lastValidSelection += $r.Index
            }
            & $UpdateCharts
        }
    })

$dgvReports.Add_CellMouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left -and $e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0) {
            if ($dgvReports.Columns[$e.ColumnIndex].Name -eq "Resistances") {
                $script:disableChartUpdate = $true
            }
            else {
                $script:disableChartUpdate = $false
            }
        }
    })

$dgvReports.Add_CellClick({
        param($s, $e)
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0 -and $dgvReports.Columns[$e.ColumnIndex].Name -eq "Resistances") {
            $row = $dgvReports.Rows[$e.RowIndex]
            $rName = $row.Cells["Rapport"].Value
            $current = $script:resistanceDict[$rName]
            if (-not $current) { $current = @{ U = ""; V = ""; W = "" } }
        
            $newVals = Show-ResistanceDialog -ReportName $rName -CurrentValues $current
        
            $script:disableChartUpdate = $false
        
            if ($newVals) {
                $script:resistanceDict[$rName] = $newVals
                Save-Resistances -Dict $script:resistanceDict
                $row.Cells["Resistances"].Value = "U:$($newVals.U) V:$($newVals.V) W:$($newVals.W)"
            
                & $UpdateCharts
            }
        }
    })

$cbDisplayMode.Add_SelectedIndexChanged({
        $cbVoie.Enabled = ($cbDisplayMode.SelectedIndex -eq 1)
        & $UpdateCharts
    })

$cbVoie.Add_SelectedIndexChanged({ & $UpdateCharts })

# Resize Events
$panelCharts.Add_SizeChanged({
        $targetHeight = [math]::Floor(($panelCharts.ClientSize.Height - 20) / 2)
        if ($targetHeight -lt 200) { $targetHeight = 200 }
    
        foreach ($c in $panelCharts.Controls) {
            $c.Width = $panelCharts.ClientSize.Width - 30
            $c.Height = $targetHeight
        }
    })

# --- Init ---
$form.Add_Shown({
        $split.SplitterDistance = 520
        $form.Activate()
    })

[void]$form.ShowDialog()
