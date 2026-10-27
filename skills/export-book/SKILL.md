---
name: export-book
description: Export a whole book (resource) from Logos Bible Software or Verbum to HTML files, by driving Logos's Print/Export panel with Windows UI Automation from a bundled PowerShell script. It works by page range (about 100 pages per file, handling volumes, roman-numbered front matter and skipped page numbers) or by section (for example one file per Bible book of a commentary or apparatus). Use this whenever the user wants to export, save, back up, convert or get the full text of a Logos book, lexicon, commentary, apparatus or other resource as HTML, mentions the Print/Export 100-page limit, or asks to automate Print/Export in Logos, even if they don't say "HTML" or "skill".
---

# Export a Logos book to HTML

Logos's Print/Export panel lays out at most 100 printed pages at a time, so exporting a whole book by hand means repeating the same routine dozens of times: type a page range, press Enter, wait, click "Web Page (HTML)", name the file, save. `scripts/export-book.ps1` does that routine reliably. Your part is the judgement around it: find out how the book is built, run the export in batches, check each batch, steer when something looks off, and keep the user informed.

The user may not be able to see the screen (for example, they use a screen reader). Report every stage in words: what is happening, which pages are done, what failed and why. Don't rely on the user looking at Logos to tell what happened, and give them the log's meaning, not raw log lines.

## The script

```
powershell -NoProfile -File "<skill-dir>/scripts/export-book.ps1" -OutDir "<folder>" [options]
```

It needs only Windows PowerShell 5.1 (built into Windows) or PowerShell 7, and Logos. Nothing has to be installed. Every line it prints is also appended to `export-log.txt` in the output folder.

If PowerShell refuses to run scripts (the default execution policy on Windows client editions), explain this to the user and let them choose: allow local scripts for their account once with `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, or add `-ExecutionPolicy Bypass` to this one command. Don't change the policy on your own.

## Before you start

1. **Which book.** Logos must be running. The best way is to name the book with `-Book`: its title, its abbreviated title ("BDAG", "Louw-Nida", "NIDNTTE") or its resource ID. The script then opens the book itself, through Logos's own `logosres:` link, and opens Print/Export on it, even when other books are open. A Print/Export window left open for another book is closed first. To find the exact title, search the library with `-ListBooks "<part of the title>"`; if several books match, read them to the user and ask which one. Without `-Book`, the script uses the open Print/Export window, or opens one with Ctrl+P on the one visible book.
2. **Output folder.** Ask where the files should go unless the user said so. Page files are named after their range: `i-xxv.html` (front matter), `1-100.html`, …, `1101-1108.html`; books with volumes get a prefix: `v1-i-xxiii.html`, `v1-1-100.html`, `v2-1-100.html`. Section files are named after the section: `08 ΚΑΤΑ ΜΑΘΘΑΙΟΝ.html`. Existing files in the folder count as done and are skipped, so an interrupted export continues where it stopped.
3. **Tell the user** that Logos will be busy (roughly 15–40 seconds per chunk) and shouldn't be used meanwhile; typing in Logos can disturb the automation.

## Explore, then export in batches

### 1. Find out how the book is built

For a book you haven't exported before, look before you export:

- The script reads Logos's library catalog and logs how the book is numbered: by page, by volume and page, or not at all. A book without page numbers (Bibles, some datasets) can only be exported by section; the script says so and stops. `-ListBooks` shows the same for every match.
- `-ListSections` prints the book's top-level sections, numbered. It shows the structure (volumes, front matter, indexes, Bible books) and whether exporting by section makes more sense. If it might, ask the user.
- `-Probe "Pages 1-100"` or `-Probe "Volume 1, Pages 1-100"` types one range into Print/Export and reports how Logos rewrote it and how many printed pages it gave, without saving anything. Logos narrows a range to the pages that exist, so probes reveal volumes, gaps and dense pages. For example, "Volume 2, Pages 4-100" became "Volume 2, Pages 73-100" in a book whose volume 2 text starts at page 73.
- The title often tells you things: "Volumes 1–5" means `-Volumes 5`.

### 2. Export in batches and steer

Run the export in the foreground with `-MaxChunks 10` (about 3–6 minutes per batch). The run ends with a status line:

- `Status: paused` (exit code 2): the batch is done. Tell the user what's done ("volume 2 up to page 503, 14 files so far"), then run the same command again for the next batch.
- `Status: done` (exit code 0): finished, and the final check passed.
- `Status: stopped` or `Status: done, with problems` (exit code 1): read the `Stopped:` and `Problem:` lines.

What the script has established (where a volume's pages end, that a volume has no front matter, that a volume doesn't exist) is kept in `export-progress.txt`, so later batches don't ask Logos again. Skipped page numbers go to `export-gaps.txt`.

Between batches, read the `Saved` lines and look for anything unusual. These are the moments where your judgement helps:

- **Printed pages per chunk jump** (say from 40 to 90): the pages got dense, for example an index. The script shrinks chunks by itself when one hits 100; just mention it to the user.
- **"Could not read the range back"**: the script took the real last page from the page numbers in the saved file. Check the next `Saved` line continues right after it.
- **"Logos ignored … twice"** right after a full chunk in the middle of a volume, followed by the end of that volume: normally the volume ended exactly at a chunk boundary, but a volume that ends suspiciously early deserves a `-Probe` a bit further on.
- **Single-page chunks** (`v2-3-3.html`) are normal for title pages between gaps.
- **When you already know the answer, tell the script**, so it doesn't spend about 30 seconds per check finding out: `-Volumes N` for the number of volumes, `-NoFrontMatter` when the book has no roman-numbered front matter, `-NoLookAhead` when you're sure that an ignored range means the end of a volume (normally the script looks ten chunks further ahead once).

For a book you already know well, a single run without `-MaxChunks` in the background is fine; report when it's done.

### 3. Finish

The final check looks at the whole folder. Within each volume, the front matter and the numbered pages must each follow on without gaps, except for the page numbers Logos reported as absent. Every file must end with `</html>`. When the book shows page numbers in its text, it also checks that each file holds the pages its name promises. Report the result, including any skipped page numbers, so the user knows those are absent from the book and not lost in the export.

The export contains what Logos shows, so display settings carry over. In particular, page numbers appear in the text (as `p 1102` or `V 1, p 802`) only when they are switched on in Logos. If the user wants them in the output, or wants the stronger check, suggest switching them on before exporting.

## Exporting by section

Some books are better split by their own structure, for example the NA28 apparatus per Bible book.

1. `-ListSections` lists the top-level sections, numbered.
2. Tell the user the list in words and ask which sections to export. Always ask, even if the answer seems obvious: the user wants to choose.
3. Export them with `-Sections "8-34"` (numbers and ranges, such as `"1,3,8-10"`), together with `-OutDir`. `-MaxChunks` works here too.

A section that fills Logos's 100 printed pages is split into halves of its subsections, as often as needed, named after the first and last subsection: `10 ΚΑΤΑ ΛΟΥΚΑΝ, Inscriptio to Chapter 11.html`. Splitting goes one level deep; a single subsection that is still too big is reported at the end and needs a page-range export instead. The script leaves the last exported section ticked in Print/Export.

## Options

- `-OutDir <folder>`: where the files go (not needed with `-ListBooks`, `-ListSections` or `-Probe`).
- `-Book "<title, abbreviated title or resource ID>"`: the book to open and export (also works with `-ListSections`, `-Probe` and `-Sections`). It must match exactly one book in the library.
- `-ListBooks "<text>"`: list the library books whose title, abbreviated title or resource ID contains the text, with how each is numbered.
- `-ExpectTitle "<part of the title>"`: stop before saving anything unless the book in Print/Export has this in its title. Logos puts the title in the Save As dialog's suggested file name; the script logs it on the first save. Without this option the script compares with the title of the book it opened or found, so a Print/Export window left open for another book is always caught.
- `-MaxChunks N`: stop after N new files (`Status: paused`).
- `-Volume 2`: only volume 2. `-Volumes 5`: the book has 5 volumes; don't look for more.
- `-Start 501`, `-End 600`: limit the numbered pages of each volume (front matter only when starting at 1).
- `-NoFrontMatter`, `-NoLookAhead`: skip those checks (see above).
- `-Chunk 100`: pages per file. The script shrinks chunks by itself when they reach the printed page limit.
- `-Sections "8-34"`, `-ListSections`, `-Probe "<range>"`: see above.
- `-Overwrite`: export again even if files exist, and ignore `export-progress.txt`.
- `-VerifyOnly`: only run the final check on `-OutDir`; doesn't touch Logos.
- `-ProcessName Verbum`: for Verbum, Logos's Catholic edition.
- `-PrintedPageLimit 20`: pretend Logos's limit is lower, only to test splitting.

## How a chunk is exported

1. Put the range in the range box (automation id `NavigatorTextBox`), e.g. "Volume 1, Pages 101–200", wait 3 seconds, press Enter. Pressing Enter too soon after typing is why Logos sometimes needs Enter twice.
2. Read the range back at once: Logos rewrites it to the pages that exist. A start on a missing page moves forward, and the range stops at the next missing page or the end ("Pages 1101–1200" becomes "Pages 1101–1108"). The script saves what Logos confirmed and continues right after it. When Logos ignores a range (or puts back an earlier one) twice, there are no pages there.
3. Wait for the layout: the preview's "1 of N" counter disappears from the UI tree while Logos works. N is the number of printed pages. At 100 the chunk may have been cut off, so the script retries with fewer pages. A single page that fills the limit by itself is reported as a problem, to be exported by section instead.
4. Click "Web Page (HTML)", put the full path into the Save As dialog, press Enter, and wait until the file is completely written.

## When something goes wrong

The script stops with a `Stopped:` line that says which step failed. Finished files are kept; after fixing the cause, the same command continues.

- **Ctrl+P didn't open Print/Export**: a key held down physically (screen readers use Caps Lock or Insert as a modifier) can combine with the simulated keystroke. Ask the user to press Ctrl+P on the book, then rerun.
- **No book or several books visible**: name the book with `-Book`.
- **Title mismatch**: the Print/Export window belonged to another book. Run again with `-Book`, which closes it and opens the right one.
- **The catalog can't be read** (very old Windows without `winsqlite3.dll`, or a changed Logos data folder): the script says so and carries on without it, finding out the numbering by trying ranges.
- **Claude Code stopped a long background run because memory was low**: that's not the script's fault. Don't restart on your own; tell the user. Batches with `-MaxChunks` in the foreground avoid it.

If controls can't be found, or Logos behaves differently than described, read `references/print-export-ui.md` before changing the script. It describes the Print/Export window, how every control behaves, the quirks found so far, and how to inspect the UI. Don't search inside the Save As dialog with UI Automation: its tree includes the folder's file list and takes minutes to walk.

## Not covered yet

- **Books whose sections are too big even when split one level deep** need a page-range export of that part, which isn't possible for books without page numbers. The script reports such parts at the end.
- **The book files themselves** (`.logos4` in Logos's data folder) are encrypted and tied to the user's license. Don't try to read or decrypt them, even when the user owns the books: that would bypass Logos's copy protection. Print/Export is the supported way to get the text out, and it's what this skill uses.
