<#
.SYNOPSIS
Exports the book open in Logos Bible Software to HTML, one file per chunk of pages.

.DESCRIPTION
Logos's Print/Export panel lays out at most 100 printed pages at a time, so a whole
book has to be exported in pieces. This script repeats the manual routine for
consecutive page ranges: it puts a range such as "Pages 101-200" in the panel's range
box, presses Enter, waits until Logos has laid the pages out, clicks "Web Page (HTML)"
and saves the file as 101-200.html.

For each volume it first exports the front matter with roman page numbers (i, ii, ...),
then the numbered pages. Files of a multi-volume book get a volume prefix:
v1-i-xxiii.html, v1-1-100.html, ..., v2-1-100.html. Volumes and front matter that
don't exist are recognised because Logos ignores ranges for them.

Logos is driven through Windows UI Automation (the UIAutomationClient assembly that
ships with Windows), plus a few Win32 calls for keystrokes and the Save As dialog.
Logos's library catalog, an SQLite database, is read with the SQLite that ships with
Windows (winsqlite3.dll); it tells whether a book is numbered by page, by volume and
page, or not at all. Nothing needs to be installed. Works in Windows PowerShell 5.1
and PowerShell 7.

Existing chunk files in the output folder are skipped, so an interrupted export can
be resumed by running the same command again.

.EXAMPLE
powershell -NoProfile -File export-book.ps1 -OutDir C:\Books\BDAG

.EXAMPLE
powershell -NoProfile -File export-book.ps1 -OutDir C:\Books\LouwNida -Volume 2 -NoFrontMatter

.EXAMPLE
powershell -NoProfile -File export-book.ps1 -OutDir C:\Books\BDAG -VerifyOnly

.EXAMPLE
powershell -NoProfile -File export-book.ps1 -ListBooks "Louw"

.EXAMPLE
powershell -NoProfile -File export-book.ps1 -Book "Louw-Nida" -OutDir C:\Books\LouwNida

.EXAMPLE
powershell -NoProfile -File export-book.ps1 -ListSections

.EXAMPLE
powershell -NoProfile -File export-book.ps1 -OutDir C:\Books\NA28 -Sections 8-34

Section mode exports top-level sections of the book (as listed under "Use these
sections" in Print/Export), one file per section, named after its number and title:
"08 KATA MATHTHAION.html" (in the book's own script). A section that fills Logos's
100 printed pages is split into halves of its subsections, named after the first and
last subsection: "10 KATA LOUKAN, Chapter 1 to Chapter 6.html".
#>
[CmdletBinding()]
param(
    # Folder for the HTML files. Required, except with -ListBooks, -ListSections and -Probe.
    [string] $OutDir = '',
    # The book to export: its title, abbreviated title or resource ID, or part of it, as
    # long as only one book in the library matches. The script opens it in Logos.
    # Without -Book, the script uses the book in the open Print/Export window, or the
    # one visible book.
    [string] $Book = '',
    # Search the library for books whose title, abbreviated title or resource ID
    # contains this text, list them, and stop.
    [string] $ListBooks = '',
    # List the book's top-level sections, numbered, and stop.
    [switch] $ListSections,
    # Export these top-level sections (numbers from -ListSections, e.g. "8-34" or
    # "1,3,8-10"), one file per section, instead of page ranges.
    [string] $Sections = '',
    # First numbered page to export in each volume. Front matter is only exported from 1.
    [int] $Start = 1,
    # Last numbered page to export in each volume; 0 means up to the end.
    [int] $End = 0,
    # Export only this volume of a multi-volume book; 0 means all volumes.
    [int] $Volume = 0,
    # Pages per file.
    [int] $Chunk = 100,
    # Stop unless the open book's title contains this text.
    [string] $ExpectTitle = '',
    # Skip the front matter with roman page numbers.
    [switch] $NoFrontMatter,
    # Export chunks again even if their files exist.
    [switch] $Overwrite,
    # Only check the files in OutDir; don't touch Logos.
    [switch] $VerifyOnly,
    # Process name of the Logos application (Verbum uses "Verbum").
    [string] $ProcessName = 'Logos',
    # Logos's limit on printed pages per export. Lower it only to test how the script
    # splits ranges and sections.
    [int] $PrintedPageLimit = 100,
    # Stop after this many new files ("Status: paused", exit code 2); run the same
    # command again for the next batch. 0 means no limit.
    [int] $MaxChunks = 0,
    # The number of volumes, when known; the script then doesn't look for more.
    [int] $Volumes = 0,
    # When Logos ignores a range of numbered pages, take it as the end of the volume
    # without first looking ten chunks further ahead for more pages.
    [switch] $NoLookAhead,
    # Try this range in the range box (e.g. "Volume 2, Pages 4-100") and report how
    # Logos takes it, without saving anything.
    [string] $Probe = ''
)

$ErrorActionPreference = 'Stop'

$PanelTitle = 'Print/Export'
$RangeBoxId = 'NavigatorTextBox'   # automation id of the "Pages a-b" edit box
$NextPageId = 'NextPage'           # preview button; gone from the UI tree while Logos lays out pages
$HtmlLinkName = 'Web Page (HTML)'
$PrintedPageCap = $PrintedPageLimit   # Logos never lays out more printed pages than this
$SettleSeconds = 3                 # pause between typing a range and pressing Enter
$LayoutStartSeconds = 10           # how long to wait for a layout to start after Enter
$EnDash = [char]0x2013
$PageLabel = '([0-9]+|[ivxlcdm]+)'
# Range text as Logos rewrites it: "Pages 1101-1108", "Vol. 1, pp 801-841", "Volume 2, Pages i-c".
$RangePattern = "^(?:(?:Volume|Vol\.?)\s*(\d+)\s*,\s*)?(?:Pages|Page|pp\.?|p\.?)\s*$PageLabel(?:\s*[\u2013\-]\s*$PageLabel)?$"
$MultiVolumePattern = '^(?:Volume|Vol\.?)\s*\d'
$ChunkFilePattern = "^(?:v(\d+)-)?$PageLabel-$PageLabel\.html$"
# Page markers in exported text, shown when page numbers are switched on in Logos:
# <span style="color:...; font-family:#Libronix User Interface; ...">  p 1102  </span>,
# or "V 1, p 802" in a book with volumes. A page's marker sits where the page begins.
$PageMarkerPattern = "<span style=""[^""]*Libronix User Interface[^""]*"">\s*(?:V\s*(\d+),\s*)?p\s+$PageLabel\s*</span>"

if (-not ('LogosExportNative' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class LogosExportNative
{
    [StructLayout(LayoutKind.Sequential)]
    struct RECT { public int Left, Top, Right, Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    struct GUITHREADINFO
    {
        public int cbSize;
        public int flags;
        public IntPtr hwndActive, hwndFocus, hwndCapture, hwndMenuOwner, hwndMoveSize, hwndCaret;
        public RECT rcCaret;
    }

    delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc proc, IntPtr lParam);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr hWnd, uint cmd);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hWnd, StringBuilder name, int max);
    [DllImport("user32.dll")] static extern bool GetGUIThreadInfo(uint thread, ref GUITHREADINFO info);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, StringBuilder lParam);
    [DllImport("user32.dll")] static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);

    const uint GW_OWNER = 4;
    const uint WM_NULL = 0x0000, WM_SETTEXT = 0x000C, WM_GETTEXT = 0x000D, WM_GETTEXTLENGTH = 0x000E, WM_CLOSE = 0x0010;
    const uint WM_KEYDOWN = 0x0100, WM_KEYUP = 0x0101;
    const uint SMTO_ABORTIFHUNG = 0x0002;
    const int VK_RETURN = 0x0D;

    // Visible top-level windows of a process, including owned ones such as dialogs.
    public static IntPtr[] TopLevelWindows(int processId)
    {
        List<IntPtr> found = new List<IntPtr>();
        EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
        {
            uint owner;
            GetWindowThreadProcessId(hWnd, out owner);
            if (owner == (uint)processId && IsWindowVisible(hWnd)) found.Add(hWnd);
            return true;
        }, IntPtr.Zero);
        return found.ToArray();
    }

    public static IntPtr Owner(IntPtr hWnd) { return GetWindow(hWnd, GW_OWNER); }

    public static string Title(IntPtr hWnd)
    {
        StringBuilder text = new StringBuilder(512);
        GetWindowText(hWnd, text, text.Capacity);
        return text.ToString();
    }

    public static string ClassName(IntPtr hWnd)
    {
        StringBuilder name = new StringBuilder(256);
        GetClassName(hWnd, name, name.Capacity);
        return name.ToString();
    }

    // True when the window's UI thread answers within the timeout; false while it is busy.
    public static bool IsResponsive(IntPtr hWnd, uint timeoutMs)
    {
        IntPtr result;
        return SendMessageTimeout(hWnd, WM_NULL, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, timeoutMs, out result) != IntPtr.Zero;
    }

    // The window that has keyboard focus on the UI thread that owns hWnd.
    public static IntPtr FocusedWindow(IntPtr hWnd)
    {
        uint processId;
        uint thread = GetWindowThreadProcessId(hWnd, out processId);
        GUITHREADINFO info = new GUITHREADINFO();
        info.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
        return GetGUIThreadInfo(thread, ref info) ? info.hwndFocus : IntPtr.Zero;
    }

    // Posts Enter to a window; it reaches whichever control has keyboard focus there.
    // Posted keys go only to that window, so keys the user holds down don't interfere.
    public static void PostEnter(IntPtr hWnd)
    {
        PostMessage(hWnd, WM_KEYDOWN, new IntPtr(VK_RETURN), new IntPtr(0x001C0001));
        PostMessage(hWnd, WM_KEYUP, new IntPtr(VK_RETURN), new IntPtr(unchecked((int)0xC01C0001)));
    }

    public static void Close(IntPtr hWnd) { PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero); }

    public static string GetText(IntPtr hWnd)
    {
        int length = SendMessage(hWnd, WM_GETTEXTLENGTH, IntPtr.Zero, (string)null).ToInt32();
        StringBuilder text = new StringBuilder(length + 1);
        SendMessage(hWnd, WM_GETTEXT, new IntPtr(text.Capacity), text);
        return text.ToString();
    }

    public static void SetText(IntPtr hWnd, string text) { SendMessage(hWnd, WM_SETTEXT, IntPtr.Zero, text); }

    // SQLite as shipped with Windows 10 and later (winsqlite3.dll), for reading Logos's
    // library catalog.
    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.StdCall)] static extern int sqlite3_open_v2(byte[] filename, out IntPtr db, int flags, IntPtr vfs);
    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.StdCall)] static extern int sqlite3_prepare_v2(IntPtr db, byte[] sql, int length, out IntPtr statement, IntPtr tail);
    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.StdCall)] static extern int sqlite3_bind_text(IntPtr statement, int index, byte[] value, int length, IntPtr destructor);
    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.StdCall)] static extern int sqlite3_step(IntPtr statement);
    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.StdCall)] static extern int sqlite3_column_count(IntPtr statement);
    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.StdCall)] static extern IntPtr sqlite3_column_text(IntPtr statement, int column);
    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.StdCall)] static extern int sqlite3_column_bytes(IntPtr statement, int column);
    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.StdCall)] static extern int sqlite3_finalize(IntPtr statement);
    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.StdCall)] static extern int sqlite3_close(IntPtr db);
    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.StdCall)] static extern IntPtr sqlite3_errmsg(IntPtr db);

    const int SQLITE_OPEN_READWRITE = 2, SQLITE_ROW = 100;

    static byte[] Utf8(string text) { return Encoding.UTF8.GetBytes(text + "\0"); }

    // Runs a query with an optional text parameter (?1) and returns the rows as strings.
    public static List<string[]> Query(string path, string sql, string parameter)
    {
        IntPtr db;
        if (sqlite3_open_v2(Utf8(path), out db, SQLITE_OPEN_READWRITE, IntPtr.Zero) != 0)
        {
            sqlite3_close(db);
            throw new Exception("Could not open the Logos library catalog.");
        }
        try
        {
            IntPtr statement;
            if (sqlite3_prepare_v2(db, Utf8(sql), -1, out statement, IntPtr.Zero) != 0)
                throw new Exception("Library catalog query failed: " + Marshal.PtrToStringAnsi(sqlite3_errmsg(db)));
            try
            {
                if (parameter != null) sqlite3_bind_text(statement, 1, Utf8(parameter), -1, new IntPtr(-1));
                List<string[]> rows = new List<string[]>();
                int columns = sqlite3_column_count(statement);
                while (sqlite3_step(statement) == SQLITE_ROW)
                {
                    string[] row = new string[columns];
                    for (int i = 0; i < columns; i++)
                    {
                        IntPtr text = sqlite3_column_text(statement, i);
                        if (text == IntPtr.Zero) continue;
                        byte[] bytes = new byte[sqlite3_column_bytes(statement, i)];
                        Marshal.Copy(text, bytes, 0, bytes.Length);
                        row[i] = Encoding.UTF8.GetString(bytes);
                    }
                    rows.Add(row);
                }
                return rows;
            }
            finally { sqlite3_finalize(statement); }
        }
        finally { sqlite3_close(db); }
    }
}
'@
}

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$Walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker

# --- Logging and small helpers --------------------------------------------------------

$script:LogFile = $null

function Write-Log([string] $Message) {
    $line = '{0:HH:mm:ss} {1}' -f (Get-Date), $Message
    Write-Host $line
    if ($script:LogFile) { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 }
}

# Stops the export for a reason that running the same command again won't fix.
function Stop-Export([string] $Message) {
    $script:Final = $true
    throw $Message
}

function New-Condition($Property, $Value) {
    New-Object System.Windows.Automation.PropertyCondition($Property, $Value)
}

function Wait-Until([scriptblock] $Condition, [double] $Seconds, [int] $IntervalMs = 300) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ($true) {
        $value = & $Condition
        if ($value -or (Get-Date) -ge $deadline) { return $value }
        Start-Sleep -Milliseconds $IntervalMs
    }
}

# --- Page labels and ranges -----------------------------------------------------------

function ConvertTo-Roman([int] $Number) {
    $numerals = @(@(1000, 'm'), @(900, 'cm'), @(500, 'd'), @(400, 'cd'), @(100, 'c'), @(90, 'xc'),
        @(50, 'l'), @(40, 'xl'), @(10, 'x'), @(9, 'ix'), @(5, 'v'), @(4, 'iv'), @(1, 'i'))
    $text = ''
    foreach ($numeral in $numerals) {
        while ($Number -ge $numeral[0]) { $text += $numeral[1]; $Number -= $numeral[0] }
    }
    $text
}

function ConvertFrom-PageLabel([string] $Label) {
    if ($Label -match '^\d+$') { return [int]$Label }
    $values = @{ i = 1; v = 5; x = 10; l = 50; c = 100; d = 500; m = 1000 }
    $letters = $Label.ToLowerInvariant().ToCharArray()
    $total = 0
    for ($index = 0; $index -lt $letters.Length; $index++) {
        $value = $values[[string]$letters[$index]]
        $isSubtracted = $index + 1 -lt $letters.Length -and $values[[string]$letters[$index + 1]] -gt $value
        if ($isSubtracted) { $total -= $value } else { $total += $value }
    }
    $total
}

function Format-PageLabel([int] $Number, [bool] $Roman) {
    if ($Roman) { ConvertTo-Roman $Number } else { [string]$Number }
}

# The range as typed into Logos: "Pages 1-100" or "Volume 2, Pages i-c" (with an en dash).
function Format-Range([int] $Vol, [bool] $Roman, [int] $First, [int] $Last) {
    $pages = "Pages $(Format-PageLabel $First $Roman)$EnDash$(Format-PageLabel $Last $Roman)"
    if ($Vol -gt 0) { "Volume $Vol, $pages" } else { $pages }
}

# The range in words for the log: "volume 1, pages i to c".
function Get-RangeDescription([int] $Vol, [bool] $Roman, [int] $First, [int] $Last) {
    $pages = "pages $(Format-PageLabel $First $Roman) to $(Format-PageLabel $Last $Roman)"
    if ($Vol -gt 0) { "volume $Vol, $pages" } else { $pages }
}

function Get-ChunkName([int] $Vol, [bool] $Roman, [int] $First, [int] $Last) {
    $prefix = if ($Vol -gt 0) { "v$Vol-" } else { '' }
    "$prefix$(Format-PageLabel $First $Roman)-$(Format-PageLabel $Last $Roman).html"
}

# --- Logos windows --------------------------------------------------------------------

function Get-LogosProcessId {
    $process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $process) { throw "Logos is not running. Start Logos, open the book, and run again." }
    $process.Id
}

function Find-Panel {
    foreach ($hwnd in [LogosExportNative]::TopLevelWindows($script:LogosId)) {
        if ([LogosExportNative]::Title($hwnd) -eq $PanelTitle) { return $hwnd }
    }
    [IntPtr]::Zero
}

function Find-SaveDialog {
    foreach ($hwnd in [LogosExportNative]::TopLevelWindows($script:LogosId)) {
        if ([LogosExportNative]::Owner($hwnd) -eq $script:Panel -and [LogosExportNative]::ClassName($hwnd) -eq '#32770') {
            return $hwnd
        }
    }
    [IntPtr]::Zero
}

function Get-PanelElement { $AE::FromHandle($script:Panel) }

function Find-InPanel($Condition) {
    try { (Get-PanelElement).FindFirst([System.Windows.Automation.TreeScope]::Descendants, $Condition) }
    catch [System.Windows.Automation.ElementNotAvailableException] { $null }
}

function Get-RangeBox { Find-InPanel (New-Condition $AE::AutomationIdProperty $RangeBoxId) }

function Get-RangeText {
    $box = Get-RangeBox
    if (-not $box) { return $null }
    try { $box.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value }
    catch { $null }
}

function Get-MainWindow {
    $main = [LogosExportNative]::TopLevelWindows($script:LogosId) |
        Where-Object { [LogosExportNative]::Owner($_) -eq [IntPtr]::Zero -and [LogosExportNative]::Title($_) -and [LogosExportNative]::Title($_) -ne $PanelTitle } |
        Select-Object -First 1
    if (-not $main) { throw "Logos has no open window. Open the book in Logos and run again." }
    $mainElement = $AE::FromHandle($main)
    $window = $mainElement.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
    if ($window.Current.WindowVisualState -eq [System.Windows.Automation.WindowVisualState]::Minimized) {
        Write-Log 'Logos is minimized; restoring its window.'
        $window.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
        Start-Sleep -Seconds 1
    }
    $mainElement
}

# The books visible in the main window, as objects with the book's Title (from its
# tab) and its Document (the text element).
function Get-VisibleBooks($MainElement) {
    $documents = $MainElement.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.AndCondition(
            (New-Condition $AE::ControlTypeProperty ([System.Windows.Automation.ControlType]::Edit)),
            (New-Condition $AE::NameProperty 'Document'))))
    foreach ($document in $documents) {
        if ($document.Current.IsOffscreen) { continue }
        $tab = $document
        while ($tab -and $tab.Current.ControlType -ne [System.Windows.Automation.ControlType]::TabItem) { $tab = $Walker.GetParent($tab) }
        [pscustomobject]@{ Title = $(if ($tab) { $tab.Current.Name } else { '' }); Document = $document }
    }
}

# Presses Ctrl+P in Logos, which opens Print/Export for the active book.
function Send-PrintShortcut($MainElement, $Document) {
    # Ctrl+P is a WPF keyboard shortcut and only fires on real key events, which
    # combine with keys held down physically (such as a screen reader key); hence
    # the second attempt.
    Add-Type -AssemblyName System.Windows.Forms
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        $MainElement.SetFocus()
        if ($Document) { try { $Document.SetFocus() } catch { } }
        Start-Sleep -Milliseconds 300
        [System.Windows.Forms.SendKeys]::SendWait('^p')
        $panel = Wait-Until { $p = Find-Panel; if ($p -ne [IntPtr]::Zero) { $p } } 15
        if ($panel) { return $panel }
    }
    throw "Ctrl+P did not open the Print/Export window. Press Ctrl+P on the book yourself, then run again."
}

# Returns the Print/Export window, opening it with Ctrl+P on the one visible book if
# needed, and notes the book's title when it can tell.
function Open-Panel {
    $panel = Find-Panel
    $books = @(Get-VisibleBooks (Get-MainWindow))
    if ($books.Count -eq 1) { $script:BookTitle = $books[0].Title }
    if ($panel -ne [IntPtr]::Zero) {
        Write-Log 'The Print/Export window is already open; using it.'
        return $panel
    }
    if ($books.Count -ne 1) {
        throw ("Found $($books.Count) visible books in Logos, so I can't tell which one to export. " +
            "Name the book with -Book, or open Print/Export yourself with Ctrl+P on the book, then run again.")
    }
    Write-Log "Opening Print/Export with Ctrl+P on: $($books[0].Title)"
    Send-PrintShortcut (Get-MainWindow) $books[0].Document
}

# Opens the book in Logos through its logosres: link and then Print/Export for it.
# A Print/Export window that is already open belongs to some other book, so it is
# closed first.
function Open-Book($Record) {
    $panel = Find-Panel
    if ($panel -ne [IntPtr]::Zero) {
        Write-Log 'Closing the Print/Export window that is open, which may belong to another book.'
        $AE::FromHandle($panel).GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close()
        [void](Wait-Until { (Find-Panel) -eq [IntPtr]::Zero } 10)
    }
    Write-Log "Opening $($Record.Title) in Logos."
    Start-Process "logosres:$($Record.LinkId)"
    $main = Get-MainWindow
    $opened = Wait-Until { Get-VisibleBooks $main | Where-Object { $_.Title -eq $Record.Title } | Select-Object -First 1 } 60 500
    if (-not $opened) { throw "Logos did not open $($Record.Title) within a minute." }
    $script:BookTitle = $Record.Title
    Start-Sleep -Seconds 2   # let Logos finish showing the book before Ctrl+P
    # The book Logos just opened is the active one, which is where Ctrl+P goes.
    Send-PrintShortcut $main $null
}

# --- Library catalog ------------------------------------------------------------------

# Logos keeps its library catalog in an SQLite database, LibraryCatalog\catalog.db in
# its data folder. For each book it lists, among other things, how the book is indexed
# (MilestoneIndexes): "Reference;page;..." for page numbers, "Reference;vp;..." for
# volume and page, and neither for books without page numbers (Bibles, for example).
# Books in the user's library have Availability 2. Logos keeps the database open, so
# the script reads a copy.

function Get-CatalogPath {
    $exe = (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object -First 1).Path
    $root = if ($exe) { Split-Path (Split-Path $exe) } else { Join-Path $env:LOCALAPPDATA $ProcessName }
    Get-ChildItem -Path (Join-Path $root 'Data\*\LibraryCatalog\catalog.db') -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
}

function Invoke-CatalogQuery([string] $Sql, [string] $Parameter) {
    $catalog = Get-CatalogPath
    if (-not $catalog) { throw "Could not find the Logos library catalog." }
    $copy = Join-Path ([System.IO.Path]::GetTempPath()) "logos-export-catalog-$PID"
    New-Item -ItemType Directory -Force -Path $copy | Out-Null
    try {
        foreach ($suffix in '', '-wal', '-shm') {
            if (Test-Path -LiteralPath "$catalog$suffix") {
                Copy-Item -LiteralPath "$catalog$suffix" -Destination (Join-Path $copy "catalog.db$suffix") -Force
            }
        }
        , [LogosExportNative]::Query((Join-Path $copy 'catalog.db'), $Sql, $Parameter)
    }
    finally { Remove-Item -LiteralPath $copy -Recurse -Force -ErrorAction SilentlyContinue }
}

# Books whose title, abbreviated title or resource ID contains the text, books in the
# library first. LinkId is the short ID for logosres: links (such as "bdag").
function Find-CatalogBooks([string] $Text) {
    $sql = "select r.ResourceId, r.Title, r.AbbreviatedTitle, r.MilestoneIndexes, r.Availability, " +
        "(select min(a.AlternateResourceId) from AlternateResourceIds a where a.RecordId = r.RecordId and a.AlternateResourceId not like 'lls:%') " +
        "from Records r where r.Title like '%' || ?1 || '%' or r.AbbreviatedTitle like '%' || ?1 || '%' or r.ResourceId like '%' || ?1 || '%' " +
        "order by r.Availability desc, r.Title"
    foreach ($row in (Invoke-CatalogQuery $sql $Text)) {
        $linkId = if ($row[5]) { $row[5] } else { $row[0] -replace '^LLS:', '' }
        [pscustomobject]@{
            ResourceId = $row[0]; Title = $row[1]; ShortTitle = $row[2]; Indexes = "$($row[3])"
            InLibrary = $row[4] -eq '2'; LinkId = $linkId
        }
    }
}

# How a book is numbered, from its catalog record: 'volumes' (volume and page),
# 'pages', or 'none'.
function Get-Numbering($Record) {
    if ($Record.Indexes -match '(^|\t)Reference;vp;') { 'volumes' }
    elseif ($Record.Indexes -match '(^|\t)Reference;page;') { 'pages' }
    else { 'none' }
}

function Get-NumberingDescription($Record) {
    switch (Get-Numbering $Record) {
        'volumes' { 'numbered by volume and page' }
        'pages' { 'numbered by page' }
        default { 'without page numbers (export it by section)' }
    }
}

# The catalog record of the book being exported, or $null when the catalog can't be
# read or the title isn't known.
function Get-BookRecord {
    if (-not $script:BookTitle) { return $null }
    try {
        Find-CatalogBooks $script:BookTitle | Where-Object { $_.Title -eq $script:BookTitle } | Select-Object -First 1
    }
    catch {
        Write-Log "  Could not read the Logos library catalog ($($_.Exception.Message)); working without it."
        $null
    }
}

# Finds the one library book matching -Book.
function Resolve-Book([string] $Text) {
    $found = @(Find-CatalogBooks $Text | Where-Object InLibrary)
    $exact = @($found | Where-Object { $_.Title -eq $Text -or $_.ShortTitle -eq $Text -or $_.ResourceId -eq $Text -or $_.ResourceId -eq "LLS:$Text" -or $_.LinkId -eq $Text })
    if ($exact.Count -eq 1) { return $exact[0] }
    if ($found.Count -eq 1) { return $found[0] }
    if ($found.Count -eq 0) { Stop-Export "No book in your Logos library matches '$Text'. Search with -ListBooks." }
    $sample = ($found | Select-Object -First 5 | ForEach-Object { $_.Title }) -join '; '
    Stop-Export "$($found.Count) books in your library match '$Text', for example: $sample. Use the exact title (see -ListBooks)."
}

function Show-Books([string] $Text) {
    $found = @(Find-CatalogBooks $Text)
    $inLibrary = @($found | Where-Object InLibrary)
    Write-Log "$($inLibrary.Count) books in your library match '$Text':"
    foreach ($record in $inLibrary) {
        $short = if ($record.ShortTitle -and $record.ShortTitle -ne $record.Title) { " ($($record.ShortTitle))" } else { '' }
        Write-Log "  $($record.Title)$short; $(Get-NumberingDescription $record); resource ID $($record.ResourceId)"
    }
    $others = $found.Count - $inLibrary.Count
    if ($others -gt 0) { Write-Log "$others more matching books in the Logos catalog are not in your library." }
}

# Connects to the book to export: opens the one named with -Book, or else uses the
# open Print/Export window or the one visible book.
function Connect-Book {
    $script:LogosId = Get-LogosProcessId
    if ($Book) {
        $record = Resolve-Book $Book
        return Open-Book $record
    }
    Open-Panel
}

function Select-RangeMode {
    $box = Get-RangeBox
    if (-not $box) { throw "Could not find the page range box in the Print/Export window." }
    $radio = $Walker.GetParent($box)
    if ($radio.Current.ControlType -ne [System.Windows.Automation.ControlType]::RadioButton) {
        throw "Could not find the page range option in the Print/Export window."
    }
    $selection = $radio.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    if (-not $selection.Current.IsSelected) {
        # Selecting the option makes Logos lay out the range the box holds. Logos may
        # put back an earlier range, possibly the whole book (minutes of work), so a
        # single page goes in first; usually Logos then shows the reading position.
        $box.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue("Page $Start")
        $selection.Select()
    }
    [void](Wait-Settled)
}

# --- Layout ---------------------------------------------------------------------------

# The preview's "1 of N" counter, as an object with the page count N and the runtime id
# of the Next Page button. Returns $null while Logos is laying out pages, because the
# counter is then missing from the UI tree.
function Get-LayoutState {
    $next = Find-InPanel (New-Condition $AE::AutomationIdProperty $NextPageId)
    if (-not $next) { return $null }
    try {
        $countLabel = $Walker.GetPreviousSibling($next)
        if (-not $countLabel -or $countLabel.Current.Name -notmatch '^\d+$') { return $null }
        [pscustomobject]@{ Count = [int]$countLabel.Current.Name; Id = ($next.GetRuntimeId() -join '.') }
    }
    catch [System.Windows.Automation.ElementNotAvailableException] { $null }
}

function Test-LayoutChanged($Before, $After) {
    if (-not $After) { return $false }
    if (-not $Before) { return $true }
    $After.Count -ne $Before.Count -or $After.Id -ne $Before.Id
}

# Waits until Logos is idle: responsive, with the page counter present and unchanged
# for a few seconds, and returns that layout state. Opening the panel or selecting the
# range option can start a layout of its own; waiting it out keeps it from being
# mistaken for the layout of the next range.
function Wait-Settled([int] $QuietSeconds = 3) {
    $deadline = (Get-Date).AddMinutes(30)
    $last = $null
    $since = Get-Date
    while ((Get-Date) -lt $deadline) {
        $state = $null
        if ([LogosExportNative]::IsResponsive($script:Panel, 500)) { $state = Get-LayoutState }
        if ($state -and $last -and -not (Test-LayoutChanged $last $state)) {
            if (((Get-Date) - $since).TotalSeconds -ge $QuietSeconds) { return $state }
        }
        else {
            $last = $state
            $since = Get-Date
        }
        Start-Sleep -Milliseconds 500
    }
    throw "Logos was still busy after 30 minutes."
}

# Reads the range as Logos rewrote it on Enter, as an object with Volume (0 if none),
# First and Last. Logos narrows a range to the pages that exist: a start on a missing
# page moves forward to the next page, and the range stops at the next missing page
# or the end of the volume ("Pages 1101-1200" becomes "Pages 1101-1108", "Volume 2,
# Pages 4-100" becomes "Volume 2, Pages 73-100", "Volume 2, Pages 1-100" can become
# "Volume 2, Page 3"). For a range without any existing page, Logos either leaves the
# text alone or puts the previous range back.
function Read-ConfirmedRange([int] $First, [int] $Last) {
    Start-Sleep -Milliseconds 1000   # let Logos handle the posted Enter first
    $deadline = (Get-Date).AddSeconds(3)
    do {
        $text = Get-RangeText
        if ($text -and $text.Trim() -match $RangePattern) {
            $lastLabel = if ($Matches[3]) { $Matches[3] } else { $Matches[2] }
            $vol = if ($Matches[1]) { [int]$Matches[1] } else { 0 }
            return [pscustomobject]@{
                Volume = $vol
                First  = ConvertFrom-PageLabel $Matches[2]
                Last   = ConvertFrom-PageLabel $lastLabel
            }
        }
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $deadline)
    $null
}

# Waits for the layout that Enter should start and returns the printed page count, or
# $null when no layout was seen.
function Wait-Layout($Before) {
    $started = $false
    $deadline = (Get-Date).AddSeconds($LayoutStartSeconds)
    while ((Get-Date) -lt $deadline) {
        $state = Get-LayoutState
        if (-not $state) { $started = $true; break }
        # A short range can be laid out completely between two checks.
        if (Test-LayoutChanged $Before $state) { return $state.Count }
        Start-Sleep -Milliseconds 200
    }
    if (-not $started) { return $null }

    $deadline = (Get-Date).AddMinutes(30)
    while ((Get-Date) -lt $deadline) {
        if ([LogosExportNative]::IsResponsive($script:Panel, 500)) {
            $state = Get-LayoutState
            if ($state -and (Get-RangeBox)) { return $state.Count }
        }
        Start-Sleep -Milliseconds 500
    }
    throw "Logos was still laying out the pages after 30 minutes."
}

# Enters the range and waits until Logos has laid the pages out. Returns the range as
# Logos confirmed it plus the printed page count, or $null when the pages don't exist:
# Logos either put another range back or ignored the range twice.
function Submit-Range([int] $Vol, [bool] $Roman, [int] $First, [int] $Last) {
    $description = Get-RangeDescription $Vol $Roman $First $Last
    $before = Wait-Settled
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        $box = Get-RangeBox
        if (-not $box) { throw "Could not find the page range box in the Print/Export window." }
        $box.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue((Format-Range $Vol $Roman $First $Last))
        Start-Sleep -Seconds $SettleSeconds
        (Get-RangeBox).SetFocus()
        [LogosExportNative]::PostEnter($script:Panel)

        $confirmed = Read-ConfirmedRange $First $Last
        if ($confirmed) {
            $otherVolume = $Vol -gt 0 -and $confirmed.Volume -gt 0 -and $confirmed.Volume -ne $Vol
            if ($otherVolume -or $confirmed.First -lt $First -or $confirmed.Last -gt $Last -or $confirmed.Last -lt $confirmed.First) {
                Write-Log "  Logos put back a different range, so $description doesn't exist."
                return $null
            }
        }
        $count = Wait-Layout $before
        if ($null -eq $count -and $confirmed -and ($confirmed.First -ne $First -or $confirmed.Last -ne $Last)) {
            # Logos rewrote the range, so it accepted it; the layout was just too quick
            # to notice (a few pages with the same printed page count as before).
            $count = (Wait-Settled).Count
        }
        if ($null -ne $count) {
            $assumed = $false
            if (-not $confirmed) {
                # During a heavy layout the range box can vanish from the UI tree. Once
                # the layout is done it usually shows the range Logos settled on.
                $text = Get-RangeText
                if ($text -and $text.Trim() -match $RangePattern -and $Matches[3]) {
                    $after = [pscustomobject]@{ First = ConvertFrom-PageLabel $Matches[2]; Last = ConvertFrom-PageLabel $Matches[3] }
                    if ($after.First -ge $First -and $after.Last -le $Last -and $after.First -le $after.Last) { $confirmed = $after }
                }
            }
            if (-not $confirmed) {
                Write-Log "  Could not read the range back from Logos; assuming $description for now."
                $confirmed = [pscustomobject]@{ Volume = $Vol; First = $First; Last = $Last }
                $assumed = $true
            }
            return [pscustomobject]@{ First = $confirmed.First; Last = $confirmed.Last; Printed = $count; Assumed = $assumed }
        }
        if ($attempt -eq 1) { Write-Log "  Logos did not respond to $description; trying once more." }
    }
    Write-Log "  Logos ignored $description twice, so I take it these pages don't exist."
    $null
}

# --- Saving ---------------------------------------------------------------------------

function Test-HtmlComplete([string] $Path) {
    try {
        $stream = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
        try {
            $length = [Math]::Min(256, $stream.Length)
            [void]$stream.Seek(-$length, 'End')
            $buffer = New-Object byte[] $length
            [void]$stream.Read($buffer, 0, $length)
            [System.Text.Encoding]::ASCII.GetString($buffer).ToLowerInvariant().Contains('</html>')
        }
        finally { $stream.Dispose() }
    }
    catch { $false }
}

# Waits until Logos has finished writing the file: complete HTML and a size that no
# longer changes. Returns the size in bytes.
function Wait-File([string] $Path, [int] $Seconds) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    $lastSize = -1
    $stableSince = Get-Date
    while ((Get-Date) -lt $deadline) {
        $size = if (Test-Path -LiteralPath $Path) { (Get-Item -LiteralPath $Path).Length } else { -1 }
        if ($size -ne $lastSize) {
            $lastSize = $size
            $stableSince = Get-Date
        }
        elseif ($size -gt 0 -and ((Get-Date) - $stableSince).TotalSeconds -ge 2 -and (Test-HtmlComplete $Path)) {
            return $size
        }
        Start-Sleep -Milliseconds 500
    }
    throw "Logos did not finish writing $(Split-Path -Leaf $Path) within $Seconds seconds."
}

# Clicks "Web Page (HTML)" and saves the result as $Path. Returns the file size.
function Save-Html([string] $Path) {
    $link = Find-InPanel (New-Object System.Windows.Automation.AndCondition(
        (New-Condition $AE::ControlTypeProperty ([System.Windows.Automation.ControlType]::Hyperlink)),
        (New-Condition $AE::NameProperty $HtmlLinkName)))
    if (-not $link) { throw "Could not find the '$HtmlLinkName' link in the Print/Export window." }

    # Invoking the link opens a modal Save As dialog, and the UI Automation call does
    # not return until that dialog closes, so it runs on a separate thread.
    $invoker = [PowerShell]::Create()
    [void]$invoker.AddScript({
        param($Element)
        try { $Element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() } catch { }
    }).AddArgument($link)
    [void]$invoker.BeginInvoke()
    $script:Invokers += $invoker

    $dialog = Wait-Until { $d = Find-SaveDialog; if ($d -ne [IntPtr]::Zero) { $d } } 30
    if (-not $dialog) { throw "Clicking '$HtmlLinkName' did not open the Save As dialog." }

    # The file name box has keyboard focus once the dialog is ready.
    $nameBox = Wait-Until {
        $f = [LogosExportNative]::FocusedWindow($dialog)
        if ($f -ne [IntPtr]::Zero -and [LogosExportNative]::ClassName($f) -eq 'Edit') { $f }
    } 10
    if (-not $nameBox) {
        [LogosExportNative]::Close($dialog)
        throw "Could not find the file name box in the Save As dialog. Cancelled the save."
    }

    if ($script:FirstExport) {
        $suggested = [LogosExportNative]::GetText($nameBox)
        $title = $suggested -replace '\.html$', ''
        Write-Log "Logos suggests the file name '$suggested', so the book is: $title"
        # The suggested name is the book's title with characters that can't be in file
        # names replaced by "_". A long title may be cut short, so only its start is
        # compared. This catches a Print/Export window left open for another book.
        $expected = $null
        $wanted = $null
        if ($ExpectTitle) {
            $wanted = $ExpectTitle
            $expected = $ExpectTitle -replace '[\\/:*?"<>|]', '_'
        }
        elseif ($script:BookTitle) {
            $wanted = $script:BookTitle
            $full = $script:BookTitle -replace '[\\/:*?"<>|]', '_'
            $expected = $full.Substring(0, [Math]::Min(60, $full.Length))
        }
        if ($expected -and $suggested.IndexOf($expected, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
            [LogosExportNative]::Close($dialog)
            throw ("The Print/Export window is not for '$wanted'. Cancelled the save; nothing was written. " +
                "Close Print/Export and run again, or name the book with -Book.")
        }
        $script:FirstExport = $false
    }

    [LogosExportNative]::SetText($nameBox, $Path)
    if ([LogosExportNative]::GetText($nameBox) -ne $Path) {
        [LogosExportNative]::Close($dialog)
        throw "Could not put the file name in the Save As dialog. Cancelled the save."
    }
    [LogosExportNative]::PostEnter($nameBox)

    $closed = Wait-Until { -not [LogosExportNative]::IsWindow($dialog) } 60
    if (-not $closed) { throw "The Save As dialog did not close after pressing Enter. It may be asking a question." }
    Wait-File $Path 600
}

# --- Output folder --------------------------------------------------------------------

function Get-ChunkFiles {
    Get-ChildItem -LiteralPath $script:Out -Filter '*.html' -File |
        Where-Object { $_.Name -match $ChunkFilePattern } |
        ForEach-Object {
            [void]($_.Name -match $ChunkFilePattern)
            $vol, $firstLabel, $lastLabel = $Matches[1], $Matches[2], $Matches[3]
            [pscustomobject]@{
                Volume = if ($vol) { [int]$vol } else { 0 }
                Roman  = $firstLabel -notmatch '^\d+$'
                First  = ConvertFrom-PageLabel $firstLabel
                Last   = ConvertFrom-PageLabel $lastLabel
                File   = $_
            }
        } |
        Sort-Object Volume, @{ Expression = { -not $_.Roman } }, First
}

function Find-ExistingChunk([int] $Vol, [bool] $Roman, [int] $First) {
    Get-ChunkFiles |
        Where-Object { $_.Volume -eq $Vol -and $_.Roman -eq $Roman -and $_.First -eq $First -and (Test-HtmlComplete $_.File.FullName) } |
        Select-Object -First 1
}

# The highest page marker of the given volume and kind in a file, or 0 when none.
function Get-HighestPageMarker([string] $Path, [int] $Vol, [bool] $Roman) {
    $highest = 0
    foreach ($marker in [regex]::Matches([System.IO.File]::ReadAllText($Path), $PageMarkerPattern, 'IgnoreCase')) {
        $label = $marker.Groups[2].Value
        $markerVolume = if ($marker.Groups[1].Value) { [int]$marker.Groups[1].Value } else { 0 }
        if (($label -notmatch '^\d+$') -ne $Roman) { continue }
        if ($Vol -gt 0 -and $markerVolume -gt 0 -and $markerVolume -ne $Vol) { continue }
        $page = ConvertFrom-PageLabel $label
        if ($page -gt $highest) { $highest = $page }
    }
    $highest
}

# Compares the page markers in a chunk file with the pages its name promises. Returns
# $null when the file has no markers, otherwise a list of problems (empty when fine).
# Not every page has a marker (pages without text don't), so gaps are allowed, but
# every marker must lie within the range and the last page must be there: a chunk cut
# short at Logos's printed page limit ends too early.
function Test-PageMarkers($Chunk) {
    $text = [System.IO.File]::ReadAllText($Chunk.File.FullName)
    $found = [regex]::Matches($text, $PageMarkerPattern, 'IgnoreCase')
    if ($found.Count -eq 0) { return $null }
    $problems = @()
    $outside = @()
    $highest = 0
    foreach ($marker in $found) {
        $label = $marker.Groups[2].Value
        $page = ConvertFrom-PageLabel $label
        $markerVolume = if ($marker.Groups[1].Value) { [int]$marker.Groups[1].Value } else { 0 }
        $isRoman = $label -notmatch '^\d+$'
        $wrongVolume = $Chunk.Volume -gt 0 -and $markerVolume -gt 0 -and $markerVolume -ne $Chunk.Volume
        if ($wrongVolume -or $isRoman -ne $Chunk.Roman -or $page -lt $Chunk.First -or $page -gt $Chunk.Last) {
            $outside += $marker.Value -replace '<[^>]+>', '' -replace '\s+', ' '
        }
        elseif ($page -gt $highest) { $highest = $page }
    }
    if ($outside.Count -gt 0) {
        $problems += "$($Chunk.File.Name) contains pages outside its range, such as '$($outside[0].Trim())'."
    }
    if ($Chunk.Last -gt $Chunk.First -and $highest -lt $Chunk.Last) {
        $problems += ("$($Chunk.File.Name) ends at page $(Format-PageLabel $highest $Chunk.Roman) instead of " +
            "$(Format-PageLabel $Chunk.Last $Chunk.Roman): the export may have been cut short.")
    }
    , $problems
}

# Checks that within each volume the front matter and the numbered pages each form one
# unbroken range, that every file is complete, and, when page numbers are shown in the
# text, that each file holds the pages its name promises.
function Test-Output {
    $files = @(Get-ChunkFiles)
    $sectionFiles = @(Get-ChildItem -LiteralPath $script:Out -Filter '*.html' -File |
        Where-Object { $_.Name -match '^\d+ .+\.html$' } | Sort-Object Name)
    if ($files.Count -eq 0 -and $sectionFiles.Count -gt 0) { return Test-SectionFiles $sectionFiles }
    if ($files.Count -eq 0) {
        Write-Log "No chunk files in $script:Out."
        return $false
    }
    $problems = @()
    $gaps = @(Get-Gaps)
    foreach ($group in ($files | Group-Object Volume, Roman)) {
        $chunks = @($group.Group)
        $first = $chunks[0]
        $vol = if ($first.Volume -gt 0) { "Volume $($first.Volume), " } else { '' }
        $part = if ($first.Roman) { 'front matter' } else { 'pages' }
        $fileCount = if ($chunks.Count -eq 1) { '1 file' } else { "$($chunks.Count) files" }
        Write-Log ("Check: {0}{1} {2} to {3} in {4}." -f $vol, $part,
            (Format-PageLabel $first.First $first.Roman), (Format-PageLabel $chunks[-1].Last $first.Roman), $fileCount)
        for ($index = 1; $index -lt $chunks.Count; $index++) {
            $previous = $chunks[$index - 1]
            $next = $chunks[$index]
            $between = "$($previous.File.Name) and $($next.File.Name)"
            if ($next.First -le $previous.Last) {
                $problems += "$between overlap."
            }
            elseif ($next.First -gt $previous.Last + 1) {
                $from = $previous.Last + 1
                $to = $next.First - 1
                $known = $gaps | Where-Object {
                    $_.Volume -eq $next.Volume -and $_.Roman -eq $next.Roman -and $_.First -le $from -and $_.Last -ge $to
                }
                $pages = "$(Format-PageLabel $from $next.Roman) to $(Format-PageLabel $to $next.Roman)"
                if ($known) { Write-Log "Check: no file has pages $pages (between $between); Logos reported the book has no such pages." }
                else { $problems += "Pages $pages are missing between $between." }
            }
        }
    }
    $withoutMarkers = 0
    foreach ($chunk in $files) {
        if (-not (Test-HtmlComplete $chunk.File.FullName)) {
            $problems += "$($chunk.File.Name) looks incomplete: it has no closing </html> tag."
            continue
        }
        $markerProblems = Test-PageMarkers $chunk
        if ($null -eq $markerProblems) { $withoutMarkers++ } else { $problems += $markerProblems }
    }
    foreach ($part in $script:TooBig) {
        $problems += "$part fills Logos's printed page limit on its own and was not exported; try exporting its section instead."
    }
    if ($withoutMarkers -eq $files.Count) {
        Write-Log ("Check: the files have no page numbers in the text, so their content could not be compared " +
            "with their page ranges. Showing page numbers in Logos makes this check possible.")
    }
    else {
        $note = if ($withoutMarkers -gt 0) { "; $withoutMarkers files have no page numbers to compare" } else { '' }
        Write-Log "Check: compared the page numbers in the text with the file names$note."
    }
    foreach ($problem in $problems) { Write-Log "Problem: $problem" }
    if ($problems.Count -eq 0) { Write-Log 'Check passed: pages are continuous and every file is complete.' }
    $problems.Count -eq 0
}

# --- Sections -------------------------------------------------------------------------

# The tree under "Use these sections:" lists the book's sections. Every section has a
# check box; ticking one includes only that section's own text, not its subsections,
# which have check boxes of their own. Subsections only appear in the UI tree while
# their parent is expanded. Logos always keeps at least one section ticked: unticking
# the last one is refused, so a new selection is ticked before the old one is cleared.

function Get-SectionChildren($Item) {
    $children = @()
    for ($child = $Walker.GetFirstChild($Item); $child; $child = $Walker.GetNextSibling($child)) {
        if ($child.Current.ControlType -eq [System.Windows.Automation.ControlType]::TreeItem) { $children += $child }
    }
    , $children
}

function Get-TopSections {
    $tree = Find-InPanel (New-Condition $AE::AutomationIdProperty 'OurTreeView')
    if (-not $tree) { throw "Could not find the section list in the Print/Export window." }
    Get-SectionChildren $tree
}

function Get-Expander($Item) {
    try { $Item.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern) } catch { $null }
}

function Test-HasSubsections($Item) {
    $expander = Get-Expander $Item
    $expander -and $expander.Current.ExpandCollapseState -ne [System.Windows.Automation.ExpandCollapseState]::LeafNode
}

# Expands a section and returns its subsections, plus whether it was collapsed before.
function Open-Section($Item) {
    $expander = Get-Expander $Item
    $wasCollapsed = $expander.Current.ExpandCollapseState -eq [System.Windows.Automation.ExpandCollapseState]::Collapsed
    if ($wasCollapsed) { $expander.Expand() }
    [void](Wait-Until { (Get-SectionChildren $Item).Count -gt 0 } 3 100)
    [pscustomobject]@{ Children = (Get-SectionChildren $Item); WasCollapsed = $wasCollapsed }
}

function Close-Section($Item, $Opened) {
    if ($Opened.WasCollapsed) { (Get-Expander $Item).Collapse() }
}

function Set-SectionCheck($Item, [bool] $On) {
    $box = $Item.FindFirst([System.Windows.Automation.TreeScope]::Children,
        (New-Condition $AE::ControlTypeProperty ([System.Windows.Automation.ControlType]::CheckBox)))
    if (-not $box) { return }
    $toggle = $box.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
    for ($i = 0; $i -lt 3; $i++) {
        if (($toggle.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On) -eq $On) { return }
        $toggle.Toggle()
    }
}

# Ticks or unticks a section and everything below it.
function Set-SectionTree($Item, [bool] $On) {
    Set-SectionCheck $Item $On
    if (-not (Test-HasSubsections $Item)) { return }
    $opened = Open-Section $Item
    foreach ($child in $opened.Children) { Set-SectionTree $child $On }
    Close-Section $Item $opened
}

# A unit is what goes into one file: top-level section number Index, its own text when
# Self is set, and its subsections From to To (1-based; To = 0 when there are none).
function New-SectionUnit([int] $Index, [bool] $Self, [int] $From, [int] $To) {
    [pscustomobject]@{ Index = $Index; Self = $Self; From = $From; To = $To }
}

function Set-UnitChecks($Unit, [bool] $On, [int] $SkipFrom = 0, [int] $SkipTo = -1) {
    $top = (Get-TopSections)[$Unit.Index - 1]
    if ($Unit.Self) { Set-SectionCheck $top $On }
    if ($Unit.To -lt $Unit.From) { return }
    $opened = Open-Section $top
    for ($n = $Unit.From; $n -le $Unit.To; $n++) {
        if ($n -ge $SkipFrom -and $n -le $SkipTo) { continue }
        Set-SectionTree $opened.Children[$n - 1] $On
    }
    Close-Section $top $opened
}

# Unticks every section outside the unit, for the first selection of a run.
function Clear-OtherSections($Unit) {
    $tops = Get-TopSections
    for ($index = 1; $index -le $tops.Count; $index++) {
        if ($index -ne $Unit.Index) { Set-SectionTree $tops[$index - 1] $false; continue }
        if (-not $Unit.Self) { Set-SectionCheck $tops[$index - 1] $false }
        if (-not (Test-HasSubsections $tops[$index - 1])) { continue }
        $opened = Open-Section $tops[$index - 1]
        for ($n = 1; $n -le $opened.Children.Count; $n++) {
            if ($n -lt $Unit.From -or $n -gt $Unit.To) { Set-SectionTree $opened.Children[$n - 1] $false }
        }
        Close-Section $tops[$index - 1] $opened
    }
}

# Makes the unit the only ticked selection and returns the printed page count once
# Logos has laid it out.
function Select-SectionUnit($Unit) {
    $before = Wait-Settled
    $old = $script:CurrentUnit
    if (-not $old) {
        Set-UnitChecks $Unit $true
        Clear-OtherSections $Unit
    }
    elseif ($old.Index -eq $Unit.Index -and $Unit.From -ge $old.From -and $Unit.To -le $old.To -and ($old.Self -or -not $Unit.Self)) {
        # The unit is part of the previous one (after halving): untick the rest.
        if ($old.Self -and -not $Unit.Self) { Set-SectionCheck (Get-TopSections)[$Unit.Index - 1] $false }
        Set-UnitChecks (New-SectionUnit $old.Index $false $old.From $old.To) $false $Unit.From $Unit.To
    }
    else {
        Set-UnitChecks $Unit $true
        Set-UnitChecks $old $false
    }
    $script:CurrentUnit = $Unit

    # Wait for the layout to start, or to be over already, then for Logos to settle.
    [void](Wait-Until { $s = Get-LayoutState; (-not $s) -or (Test-LayoutChanged $before $s) } 10 200)
    (Wait-Settled).Count
}

function ConvertTo-FileName([string] $Text) {
    ($Text -replace '[\\/:*?"<>|]', '_').Trim().TrimEnd('.')
}

function ConvertFrom-SectionList([string] $Text, [int] $Count) {
    $numbers = @()
    foreach ($part in ($Text -split ',')) {
        if ($part.Trim() -match '^(\d+)(?:\s*-\s*(\d+))?$') {
            $from = [int]$Matches[1]
            $to = if ($Matches[2]) { [int]$Matches[2] } else { $from }
            $numbers += $from..$to
        }
        elseif ($part.Trim()) { throw "Can't read '$part' in -Sections; use numbers and ranges like 8-34 or 1,3,8-10." }
    }
    $outside = @($numbers | Where-Object { $_ -lt 1 -or $_ -gt $Count })
    if ($outside.Count -gt 0) { throw "The book has $Count sections; there is no section $($outside[0])." }
    $numbers | Select-Object -Unique
}

function Show-Sections {
    $tops = Get-TopSections
    Write-Log "The book has $($tops.Count) top-level sections:"
    for ($index = 1; $index -le $tops.Count; $index++) {
        $note = if (Test-HasSubsections $tops[$index - 1]) { ' (has subsections)' } else { '' }
        Write-Log ("{0,4}  {1}{2}" -f $index, $tops[$index - 1].Current.Name, $note)
    }
}

# Page markers in a file, for the log: "printed pages 2 to 101".
function Get-MarkerSpan([string] $Path) {
    $found = [regex]::Matches([System.IO.File]::ReadAllText($Path), $PageMarkerPattern, 'IgnoreCase')
    if ($found.Count -eq 0) { return '' }
    "; page markers $($found[0].Groups[2].Value) to $($found[$found.Count - 1].Groups[2].Value)"
}

# Exports a unit to one file, or, when it fills Logos's printed page limit, splits its
# subsections in two and exports each half.
function Export-SectionUnit($Unit, [string] $Title, [string[]] $SubTitles, [int] $Width) {
    $whole = $Unit.Self -and $Unit.From -eq 1 -and $Unit.To -eq $SubTitles.Count
    $name = "{0} {1}" -f $Unit.Index.ToString("D$Width"), $Title
    if (-not $whole) {
        $name += ", $($SubTitles[$Unit.From - 1])"
        if ($Unit.To -gt $Unit.From) { $name += " to $($SubTitles[$Unit.To - 1])" }
    }
    $name = ConvertTo-FileName $name
    $path = Join-Path $script:Out "$name.html"
    if (-not $Overwrite -and (Test-Path -LiteralPath $path) -and (Test-HtmlComplete $path)) {
        Write-Log "Already exported; skipping $name.html."
        return
    }
    if (Test-Paused) { return }

    $started = Get-Date
    $printed = Select-SectionUnit $Unit
    if ($printed -ge $PrintedPageCap) {
        $parts = $Unit.To - $Unit.From + 1
        if ($parts -lt 2) {
            Write-Log "  $name fills $printed printed pages, Logos's limit, and can't be split further; skipping it."
            $script:TooBig += $name
            return
        }
        $middle = $Unit.From + [int][Math]::Floor($parts / 2) - 1
        Write-Log "  $name fills $printed printed pages, Logos's limit; splitting it in two."
        Export-SectionUnit (New-SectionUnit $Unit.Index $Unit.Self $Unit.From $middle) $Title $SubTitles $Width
        Export-SectionUnit (New-SectionUnit $Unit.Index $false ($middle + 1) $Unit.To) $Title $SubTitles $Width
        return
    }
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path }
    $bytes = Save-Html $path
    $script:Saved++
    Write-Log ("Saved {0}.html: {1} printed pages, {2:N1} MB, {3:N0} seconds{4}." -f $name, $printed,
        ($bytes / 1MB), ((Get-Date) - $started).TotalSeconds, (Get-MarkerSpan $path))
}

function Export-Sections {
    $script:Panel = Connect-Book
    $tops = Get-TopSections
    $numbers = @(ConvertFrom-SectionList $Sections $tops.Count)
    $width = ([string]$tops.Count).Length

    $option = Find-InPanel (New-Object System.Windows.Automation.AndCondition(
        (New-Condition $AE::ControlTypeProperty ([System.Windows.Automation.ControlType]::RadioButton)),
        (New-Condition $AE::NameProperty 'Use these sections:')))
    if (-not $option) { throw "Could not find the 'Use these sections' option in the Print/Export window." }
    $selection = $option.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    if (-not $selection.Current.IsSelected) { $selection.Select() }

    $script:CurrentUnit = $null
    $script:TooBig = @()
    foreach ($index in $numbers) {
        if ($script:Paused) { break }
        $top = (Get-TopSections)[$index - 1]
        $title = $top.Current.Name
        $subTitles = @()
        if (Test-HasSubsections $top) {
            $opened = Open-Section $top
            $subTitles = @($opened.Children | ForEach-Object { $_.Current.Name })
            Close-Section $top $opened
        }
        Write-Log "Section $index of $($tops.Count): $title"
        Export-SectionUnit (New-SectionUnit $index $true 1 $subTitles.Count) $title $subTitles $width
    }
}

# Checks the files of a section export: every file must be complete. The page markers
# of each file are listed so gaps between sections can be spotted.
function Test-SectionFiles($Files) {
    $problems = @()
    foreach ($file in $Files) {
        if (-not (Test-HtmlComplete $file.FullName)) {
            $problems += "$($file.Name) looks incomplete: it has no closing </html> tag."
        }
        else {
            Write-Log "Check: $($file.Name)$(Get-MarkerSpan $file.FullName)."
        }
    }
    foreach ($part in $script:TooBig) {
        $problems += "$part fills Logos's printed page limit on its own and was not exported; export it by page range instead."
    }
    foreach ($problem in $problems) { Write-Log "Problem: $problem" }
    if ($problems.Count -eq 0) { Write-Log "Check passed: $(@($Files).Count) section files, every file is complete." }
    $problems.Count -eq 0
}

# --- Main -----------------------------------------------------------------------------

# Page numbers the book doesn't have, as reported by Logos during the export, are kept
# in export-gaps.txt ("<volume> <front|pages> <first> <last>" per line), so the final
# check can tell them apart from pages that are missing from the output.
function Add-Gap([int] $Vol, [bool] $Roman, [int] $First, [int] $Last) {
    $kind = if ($Roman) { 'front' } else { 'pages' }
    Add-Content -LiteralPath (Join-Path $script:Out 'export-gaps.txt') -Value "$Vol $kind $First $Last" -Encoding UTF8
    Write-Log "  The book has no $(Get-RangeDescription $Vol $Roman $First $Last); skipping those page numbers."
}

function Get-Gaps {
    $path = Join-Path $script:Out 'export-gaps.txt'
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    Get-Content -LiteralPath $path -Encoding UTF8 | Where-Object { $_ -match '^(\d+) (front|pages) (\d+) (\d+)$' } | ForEach-Object {
        [pscustomobject]@{ Volume = [int]$Matches[1]; Roman = $Matches[2] -eq 'front'; First = [int]$Matches[3]; Last = [int]$Matches[4] }
    }
}

# What an export has established is kept in export-progress.txt, so a later run (such
# as the next batch with -MaxChunks) doesn't ask Logos again. One line each:
# "<volume> front done <last>" or "<volume> pages done <last>" when that part of a
# volume is complete (last is its last page, 0 when it has no pages), and
# "<volume> missing" when a volume doesn't exist.
function Add-Progress([string] $Line) {
    Add-Content -LiteralPath (Join-Path $script:Out 'export-progress.txt') -Value $Line -Encoding UTF8
}

function Get-Progress {
    $path = Join-Path $script:Out 'export-progress.txt'
    if ($Overwrite -or -not (Test-Path -LiteralPath $path)) { return @() }
    @(Get-Content -LiteralPath $path -Encoding UTF8)
}

function Find-CompletePart([int] $Vol, [bool] $Roman) {
    $kind = if ($Roman) { 'front' } else { 'pages' }
    foreach ($line in Get-Progress) {
        if ($line -match "^$Vol $kind done (\d+)$") { return [int]$Matches[1] }
    }
    $null
}

function Test-Paused {
    if ($MaxChunks -gt 0 -and $script:Saved -ge $MaxChunks) { $script:Paused = $true }
    $script:Paused
}

# Exports consecutive chunks of one kind of pages (front matter or numbered pages) of
# one volume, until Logos has no more pages. Logos skips page numbers the book doesn't
# have, so a chunk that ends early is followed by the next one; the pages end when
# Logos ignores a range. Returns $true when any pages were found.
function Export-Pages([int] $Vol, [bool] $Roman, [int] $From, [int] $To) {
    $kind = if ($Roman) { 'front' } else { 'pages' }
    $part = if ($Roman) { 'front matter' } else { 'numbered pages' }
    if ($Vol -gt 0) { $part = "volume $Vol, $part" }
    # Only a run without -End can establish where the pages end.
    $openEnded = $To -eq 0
    if ($openEnded) {
        $complete = Find-CompletePart $Vol $Roman
        if ($null -ne $complete) {
            $extent = if ($complete -gt 0) { "up to page $(Format-PageLabel $complete $Roman)" } else { 'none in this book' }
            Write-Log "Already complete: $part ($extent)."
            return $complete -gt 0
        }
    }
    $first = $From
    $size = $Chunk
    $found = $false
    $lastFound = 0
    $lookingAhead = $false
    while ($To -eq 0 -or $first -le $To) {
        if (-not $Overwrite) {
            $gap = Get-Gaps | Where-Object { $_.Volume -eq $Vol -and $_.Roman -eq $Roman -and $_.First -eq $first } | Select-Object -First 1
            if ($gap) {
                $first = $gap.Last + 1
                continue
            }
            $done = Find-ExistingChunk $Vol $Roman $first
            if ($done) {
                Write-Log "Already exported: $(Get-RangeDescription $Vol $Roman $done.First $done.Last); skipping $($done.File.Name)."
                $found = $true
                $lastFound = $done.Last
                $first = $done.Last + 1
                continue
            }
        }
        if (Test-Paused) { return $found }

        $last = $first + $size - 1
        if ($To -ne 0 -and $last -gt $To) { $last = $To }
        $chunkStarted = Get-Date
        Write-Log "Asking Logos for $(Get-RangeDescription $Vol $Roman $first $last)."
        $range = Submit-Range $Vol $Roman $first $last
        if (-not $range) {
            # A gap longer than a chunk looks the same as the end, so numbered pages get
            # one more, much wider look before giving up.
            if ($Roman -or $lookingAhead -or $NoLookAhead -or -not $openEnded) {
                if ($openEnded) { Add-Progress "$Vol $kind done $lastFound" }
                break
            }
            $lookingAhead = $true
            $size = $Chunk * 10
            Write-Log "  Looking further ahead, in case the book skips more page numbers."
            continue
        }
        if ($range.First -gt $first) { Add-Gap $Vol $Roman $first ($range.First - 1) }
        $found = $true
        if ($lookingAhead) {
            $lookingAhead = $false
            $size = $Chunk
            $first = $range.First
            Write-Log "  Found more pages from $(Format-PageLabel $range.First $Roman) on."
            continue
        }
        if ($range.Printed -ge $PrintedPageCap) {
            if ($range.Last -gt $range.First) {
                $first = $range.First
                $size = [Math]::Max(1, [int][Math]::Floor(($range.Last - $range.First + 1) * 3 / 4))
                Write-Log ("  This range fills $($range.Printed) printed pages, Logos's limit, so the export could be " +
                    "cut short. Retrying with $size pages per chunk.")
                continue
            }
            $description = Get-RangeDescription $Vol $Roman $range.First $range.Last
            Write-Log "  $description alone fills Logos's limit of $PrintedPageCap printed pages, so it can't be exported completely by page; skipping it."
            $script:TooBig += $description
            $lastFound = $range.Last
            $first = $range.Last + 1
            continue
        }

        $name = Get-ChunkName $Vol $Roman $range.First $range.Last
        $path = Join-Path $script:Out $name
        if (-not $Overwrite -and (Test-Path -LiteralPath $path) -and (Test-HtmlComplete $path)) {
            Write-Log "Already exported: $name; skipping it."
            $lastFound = $range.Last
            $first = $range.Last + 1
            continue
        }
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path }
        $bytes = Save-Html $path
        if ($range.Assumed) {
            # The page numbers in the text show where the export really ends.
            $highest = Get-HighestPageMarker $path $Vol $Roman
            if ($highest -ge $range.First -and $highest -lt $range.Last) {
                $range.Last = $highest
                $newName = Get-ChunkName $Vol $Roman $range.First $range.Last
                $newPath = Join-Path $script:Out $newName
                if (Test-Path -LiteralPath $newPath) { Remove-Item -LiteralPath $newPath }
                Move-Item -LiteralPath $path -Destination $newPath
                Write-Log "  The page numbers in the file end at $(Format-PageLabel $highest $Roman); renamed it to $newName."
                $name = $newName
            }
        }
        $script:Saved++
        Write-Log ("Saved {0}: {1}, {2} printed pages, {3:N1} MB, {4:N0} seconds." -f $name,
            (Get-RangeDescription $Vol $Roman $range.First $range.Last), $range.Printed, ($bytes / 1MB),
            ((Get-Date) - $chunkStarted).TotalSeconds)
        $lastFound = $range.Last
        $first = $range.Last + 1
    }
    $found
}

# Exports the front matter and the numbered pages of one volume (0 for a book without
# volumes). Returns $true when the volume has any pages.
function Export-Volume([int] $Vol) {
    $found = $false
    if (-not $NoFrontMatter -and $Start -eq 1) {
        if (Export-Pages $Vol $true 1 0) { $found = $true }
        if ($script:Paused) { return $found }
    }
    if (Export-Pages $Vol $false $Start $End) { $found = $true }
    $found
}

function Export-Book {
    $script:Panel = Connect-Book
    Select-RangeMode

    # The library catalog says whether the book is numbered by volume and page, by page,
    # or not at all. Without the catalog, the range box is a hint: it shows "Volume 1,
    # Page i" for a book with volumes, but may still hold a range typed earlier for
    # another book, so then the other style is tried as well.
    $record = Get-BookRecord
    $numbering = $null
    if ($record) {
        $numbering = Get-Numbering $record
        Write-Log "The library catalog says this book is $(Get-NumberingDescription $record)."
        if ($numbering -eq 'none') {
            Stop-Export "This book has no page numbers in Logos, so it can't be exported by page range. Export it by section instead (-ListSections, then -Sections)."
        }
    }
    $progress = Get-Progress
    $looksMultiVolume = (Get-RangeText) -match $MultiVolumePattern
    $styles = if ($Volume -gt 0 -or $Volumes -gt 0 -or ($progress -match '^[1-9]\d* ')) { @($true) }
              elseif ($progress -match '^0 ') { @($false) }
              elseif ($numbering -eq 'volumes') { @($true) }
              elseif ($numbering -eq 'pages') { @($false) }
              elseif ($looksMultiVolume) { @($true, $false) }
              else { @($false, $true) }
    foreach ($withVolumes in $styles) {
        if (-not $withVolumes) {
            Write-Log 'Trying page ranges without volume numbers.'
            $found = Export-Volume 0
            if ($found -or $script:Paused) { return }
        }
        else {
            Write-Log 'Trying page ranges with volume numbers.'
            $list = if ($Volume -gt 0) { @($Volume) } elseif ($Volumes -gt 0) { 1..$Volumes } else { 1..999 }
            $anyVolume = $false
            foreach ($vol in $list) {
                if ($progress -contains "$vol missing") {
                    Write-Log "Already established: volume $vol does not exist."
                    break
                }
                $found = Export-Volume $vol
                if ($script:Paused) { return }
                if ($found) { $anyVolume = $true; continue }
                if ($Volume -gt 0 -or $Volumes -gt 0) { Write-Log "Logos has no pages for volume $vol."; continue }
                if ($anyVolume) {
                    Write-Log "Volume $vol does not exist, so the book has $($vol - 1) volumes."
                    Add-Progress "$vol missing"
                }
                break
            }
            if ($anyVolume) { return }
        }
        Write-Log 'Logos ignored those ranges.'
    }
    throw "Logos ignored every page range, with and without volume numbers. Maybe this book has no page numbers."
}

# Tries one range in the range box and reports how Logos takes it, without saving
# anything: for exploring how a book is numbered before exporting it.
function Invoke-Probe([string] $Text) {
    $script:Panel = Connect-Book
    Select-RangeMode
    $typed = $Text -replace '(?<=[0-9ivxlcdm])\s*-\s*(?=[0-9ivxlcdm])', $EnDash
    $before = Wait-Settled
    (Get-RangeBox).GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($typed)
    Start-Sleep -Seconds $SettleSeconds
    (Get-RangeBox).SetFocus()
    [LogosExportNative]::PostEnter($script:Panel)
    Start-Sleep -Milliseconds 1000
    $rewritten = Get-RangeText
    $count = Wait-Layout $before
    if (-not $rewritten) { $rewritten = Get-RangeText }
    if ($null -eq $count -and $rewritten -ne $typed) { $count = (Wait-Settled).Count }
    Write-Log "Typed: $typed"
    if ($rewritten -eq $typed) { Write-Log 'Logos left the range as typed.' }
    else { Write-Log "Logos rewrote it to: $rewritten" }
    if ($null -eq $count) {
        Write-Log 'Logos laid out nothing new: it ignored the range, or the Enter came too early (probe again to be sure).'
    }
    else { Write-Log "Printed pages: $count" }
}

if ($OutDir) { $script:Out = [System.IO.Path]::GetFullPath($OutDir) }
$script:Invokers = @()
$script:Saved = 0
$script:FirstExport = $true
$script:TooBig = @()
$script:Paused = $false
$script:BookTitle = $null
$script:Final = $false

# Dot-sourcing the script (". .\export-book.ps1 -OutDir <folder>") only loads its
# functions, so single steps can be tried by hand.
if ($MyInvocation.InvocationName -eq '.') { return }

# Section titles can be Greek, Hebrew, etc.; without this, Windows PowerShell writes
# them as question marks when its output is captured.
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

if ($ListBooks -or $ListSections -or $Probe) {
    try {
        if ($ListBooks) { Show-Books $ListBooks }
        elseif ($ListSections) {
            $script:Panel = Connect-Book
            Show-Sections
        }
        else { Invoke-Probe $Probe }
        exit 0
    }
    catch {
        Write-Log "Stopped: $($_.Exception.Message)"
        exit 1
    }
}

if (-not $OutDir) {
    Write-Log 'Stopped: -OutDir is required: the folder for the HTML files.'
    exit 1
}

if ($VerifyOnly) {
    if (-not (Test-Path -LiteralPath $script:Out)) { Write-Log "Folder $script:Out does not exist."; exit 1 }
    if (Test-Output) { exit 0 } else { exit 1 }
}

New-Item -ItemType Directory -Force -Path $script:Out | Out-Null
$script:LogFile = Join-Path $script:Out 'export-log.txt'

Write-Log "Exporting to $script:Out"
$started = Get-Date
try {
    if ($Sections) { Export-Sections } else { Export-Book }
}
catch {
    Write-Log "Stopped: $($_.Exception.Message)"
    if (-not $script:Final) {
        Write-Log 'Run the same command again to continue; finished chunks are kept.'
        [void](Test-Output)
    }
    Write-Log 'Status: stopped'
    exit 1
}
$elapsed = (Get-Date) - $started
$newFiles = if ($script:Saved -eq 1) { '1 new file' } else { "$($script:Saved) new files" }
$duration = '{0} min {1} s' -f [int][Math]::Floor($elapsed.TotalMinutes), $elapsed.Seconds
if ($script:Paused) {
    Write-Log "Paused after $newFiles in $duration, as asked with -MaxChunks. Run the same command again for the next batch."
    [void](Test-Output)
    Write-Log 'Status: paused'
    exit 2
}
Write-Log "Finished: $newFiles in $duration."
if (Test-Output) {
    Write-Log 'Status: done'
    exit 0
}
Write-Log 'Status: done, with problems'
exit 1
