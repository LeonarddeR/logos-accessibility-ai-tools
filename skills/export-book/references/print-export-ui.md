# The Logos Print/Export window: how it behaves

Notes from inspecting Logos (desktop app, WPF, October 2026) with UI Automation. Read this when `export-book.ps1` stops because a control can't be found or Logos behaves differently than below. Check what changed before you edit the script.

## Windows

All windows belong to the Logos process (`Logos.exe`; Verbum is `Verbum.exe`).

- **Main window**: title like "Logos Legacy Edition". Each open book is a `TabItem` whose name is the book title. The book text is an `Edit` element named "Document".
- **Print/Export**: a separate top-level window titled "Print/Export". Ctrl+P on a book opens it, for the book that was active at that moment. Window class `HwndWrapper[Logos;;...]`.
- **Save As**: a standard Windows file dialog (class `#32770`), owned by the Print/Export window. Its title is localized, so find it by class and owner.

## Controls in Print/Export

Find controls by automation id or name. winapp's generated slugs (such as `lnk-thehyperlink-bd69`) change between calls.

- **Range option**: an unnamed `RadioButton` whose only child is the range box. Select it with SelectionItemPattern. Its siblings are "Use selected text", "Use filtered text" and "Use these sections:". The last one has a tree of section check boxes and is often selected by default.
- **Range box**: `Edit` with automation id `NavigatorTextBox`. It takes text such as `Pages 101–200` (en dash; a hyphen also parses). ValuePattern.SetValue works. Oddly, it reports IsPassword = true.
- **Page counter**: under the preview are `PreviousPage` (button), Text "1", Text "of", Text "N", `NextPage` (button). N is the number of printed pages in the current layout. The preview itself (`PART_PageView`) is an image with no readable text.
- **Export links**: `Hyperlink` elements (all with automation id `TheHyperlink`) named "Rich Text Format (Word, etc.)", "Text Document", "Web Page (HTML)", "XPS Document", and "Microsoft Word" twice (send to a new document / paste into an open one).
- **Print section**: printer combo box, "Properties…", page size combo boxes, `PrintButton`.

## How the range box behaves

- **Setting the text does nothing by itself.** Logos only acts on Enter.
- **Enter rewrites the text at once,** before any layout starts:
  - A range that runs past the last page is shortened: "Pages 1101–1200" becomes "Pages 1101–1108". This is how the script finds the end of the book.
  - A range that starts past the last page is rejected, and the previous valid range is put back. "Pages 1201–1300" turns back into "Pages 1101–1108".
  - Logos never complains about either case; only the text changes.
- **Enter then starts a layout, but not always.** If Enter comes right after the text was set, it sometimes only rewrites the text and lays nothing out; a second Enter is needed. A pause of about 3 seconds between setting the text and pressing Enter makes a single Enter work in most cases. The script waits 3 seconds, checks whether a layout started, and presses Enter again if not.
- **During a layout** the page counter and the NextPage button disappear from the UI tree, and Logos's UI thread is often busy: UI Automation calls block until it is free. 100 book pages of BDAG take about 15–20 seconds. A short range (8 pages) finishes in about a second, so the disappearance is easy to miss; then a changed page count is the only sign.
- **After a layout** the range box sometimes shows the preview position ("Page 3", "Page 103") instead of the range. Read the rewritten range right after Enter, not later.
- **The cap**: Logos lays out at most 100 printed pages. Asking for the whole book ("Pages 1–99999") gave "1 of 100" and kept Logos busy for about 5 minutes. Never ask for huge ranges. A count of 100 means the chunk was probably cut off; the script then retries with smaller chunks.
- **Printed pages per book page** depend on the book's page size: 100 BDAG pages gave 92–93 printed pages, 100 Louw-Nida pages about 50.
- **Selecting the range option** makes Logos lay out whatever range the box then holds. In Louw-Nida, Logos replaced the box text with the reading position ("Volume 1, Page i"). In BDAG it once restored an earlier huge range, which made the first chunk look like it had hit the cap. The script therefore waits until the counter has been steady for 3 seconds before every range.

## Books with volumes and front matter

Tested with Louw-Nida, "Greek-English Lexicon of the New Testament based on Semantic Domains" (2 volumes).

- **Range syntax**: "Volume 1, Pages 1–100". A range without a volume ("Pages 101–150", "Page 5") is ignored.
- **Recognising such a book**: after selecting the range option, the box showed "Volume 1, Page i".
- **Rewritten ranges use a short form**: "Volume 1, Pages 801–2000" became "Vol. 1, pp 801–841"; "Volume 2, Pages 301–2000" became "Vol. 2, pp 301–334".
- **Front matter**: roman page numbers work like numbers. "Volume 1, Pages i–c" (c = 100) became "Vol. 1, pp i–xxiii".
- **Ranges that don't exist are ignored here**: the text stays as typed and nothing is laid out. Examples: "Volume 1, Pages 3000–3100", "Volume 2, Pages i–c" (volume 2 has no front matter), "Volume 3, Pages 1–100" (no volume 3). BDAG behaves differently and puts the previous range back. An ignored range can't be told apart from an Enter that came too early, so the script tries twice before concluding the pages don't exist.
- **Sections tree**: the "Use these sections" tree lists Title Page, Preface, Introduction, Bibliography, the 93 domains of volume 1, and one "Volume 2" item.

## Books that skip page numbers

Tested with NIDNTTE, "New International Dictionary of New Testament Theology and Exegesis, Volumes 1–5" (one resource, five volumes).

Logos narrows every range to the pages that exist:
- "Volume 1, Pages 1–100" became "Volume 1, Pages 1–3": page 4 doesn't exist, and the range stops before it.
- "Volume 1, Pages 4–100" became "Volume 1, Pages 5–100": a start on a missing page moves forward.
- "Volume 2, Pages 1–100" became "Volume 2, Page 3" (a single page: the volume's title page), and "Volume 2, Pages 4–100" became "Volume 2, Pages 73–100": the text of volume 2 starts at page 73.
- "Volume 5, Pages 1–100" became "Volume 5, Page 5". "Volume 6, Pages 1–100" and "Volume 1, Pages 801–2000" were ignored.

So a shortened range can mean a gap as well as the end, and a single "Page N" can be a valid answer. Only an ignored range (or one replaced by an earlier range) means there are no pages there, and only if no page exists anywhere in it: a gap longer than the requested range looks the same as the end.

The page markers in the HTML of a multi-volume book read "V 1, p 2".

## Saving

- **Invoking "Web Page (HTML)"** (InvokePattern) opens the Save As dialog. The Invoke call doesn't return until the dialog closes, so the script makes it on a separate thread. winapp's `invoke` reports "does not support any invoke pattern" here, even though the dialog does open.
- **The file name box** (`Edit`, control id 1001) has keyboard focus when the dialog opens. Its default text is the book title plus ".html", for example "A Greek-English Lexicon of the New Testament and Other Early Christian Literature, 3rd ed..html". The script logs this, and `-ExpectTitle` checks against it.
- **Searching the Save As dialog with UI Automation is very slow.** Its tree includes the folder's file list, so a winapp search or set-value there took 60 seconds or never finished. The script finds the focused box with Win32 (GetGUIThreadInfo), sets the full path with WM_SETTEXT, and posts Enter to it.
- **Writing the file** takes about 10 seconds for 100 pages, roughly 3.5–4 MB of HTML. The file is complete when it ends with `</html>`.
- **Output format**: Word-style HTML with inline styles, `ref.ly` links, and a footer naming the book and the export time ("Exported from Logos Bible Study, 7:01 PM October 27, 2026"). Apart from that timestamp, two exports of the same range are identical.

## Keystrokes

- **Enter**: the script posts WM_KEYDOWN/WM_KEYUP to the window after focusing the control with UI Automation. Posted keys go only to that window and don't mix with keys held down on the keyboard.
- **Ctrl+P**: this only works with real keyboard input (SendKeys or SendInput), because WPF checks the actual state of the Ctrl key. Keys held down physically mix with it: once, Caps Lock was held (screen readers use it as a modifier) and Ctrl+P did nothing. If it fails, the user can press Ctrl+P on the book themselves.

## Opening and closing books without the mouse

Tested while the user was away; it all needs Logos in the foreground and real keystrokes, so do it only when nobody is typing.

- **Close the open book**: Ctrl+F4 in the main window. The close buttons on the book tabs don't support UI Automation.
- **Close Print/Export**: WindowPattern.Close() on its window.
- **Open a book**: Alt+C puts the cursor in the Command box (automation id `CommandTextBox`; the toolbar button "Command:" itself doesn't support UI Automation). Set the book's title in the box with ValuePattern; if it already holds that text, clear it first, or no suggestions appear. A list of suggestions follows (ListItems whose names are an internal type name; their text is in Text children, plus an empty Text). Select the one reading "Open <title>" with SelectionItemPattern and press Enter. The first suggestion is often a Factbook entry, so don't just press Enter.
- **Ctrl+P** then opens Print/Export for the new book, which is the only one open.

## Inspecting the UI yourself

Without installing anything, PowerShell can dump the tree of the Print/Export window:

```powershell
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$logos = (Get-Process Logos | Select-Object -First 1).Id
$panel = $AE::RootElement.FindAll('Children', (New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $logos))) |
    Where-Object { $_.Current.Name -eq 'Print/Export' }
$walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
function Show-Tree($Element, $Depth) {
    for ($child = $walker.GetFirstChild($Element); $child; $child = $walker.GetNextSibling($child)) {
        '{0}{1} name="{2}" id="{3}"' -f ('  ' * $Depth), $child.Current.ControlType.ProgrammaticName, $child.Current.Name, $child.Current.AutomationId
        Show-Tree $child ($Depth + 1)
    }
}
Show-Tree $panel 0
```

Microsoft's winapp CLI (`winget install Microsoft.WinAppCLI`) is handier for exploring:

```powershell
winapp ui list-windows -a Logos                 # window handles (HWND) and titles
winapp ui inspect -w <panel-hwnd> -d 12          # full tree of the Print/Export window
winapp ui get-value NavigatorTextBox -w <panel-hwnd>
winapp ui search "Web Page (HTML)" -w <panel-hwnd>
winapp ui screenshot -w <panel-hwnd>             # PNG of the window, to look at yourself
```

Don't run `inspect` or `search` on the Save As dialog; use `winapp ui get-focused -w <dialog-hwnd>` instead.
