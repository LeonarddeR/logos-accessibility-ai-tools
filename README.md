# Logos accessibility AI tools

AI tools for Logos Bible Software on Windows. Some work in Logos is tedious or hard to check without sight. These tools let Claude do that work and report back in words.

The tools come as a Claude Code plugin, `logos-accessibility-skills`.

More skills will follow; each one lives in its own folder under `skills`.

## Skills

### export-book

Exports a whole Logos book to HTML files.

Logos can export what you read through its Print/Export panel, but at most 100 printed pages at a time. Exporting a whole book by hand means repeating the same steps dozens of times: type a page range, press Enter, wait, choose "Web Page (HTML)", name the file and save it.

The skill lets Claude do this for you. A PowerShell script drives the Print/Export panel through Windows UI Automation. It works out how the book is numbered and saves one file per chunk of pages. When it's done, it checks that nothing is missing.

- **By page range**: files like `1-100.html`, `101-200.html`. Books with volumes get a prefix (`v1-1-100.html`), and front matter with roman page numbers comes first (`v1-i-xxiii.html`). Page numbers the book doesn't have are skipped and listed in `export-gaps.txt`.
- **By section**: one file per top-level section of the book, such as one Bible book of a commentary: `08 ΚΑΤΑ ΜΑΘΘΑΙΟΝ.html`. A section that is too large for one export is split by its subsections.

When page numbers are shown in the book's text in Logos, the final check also compares those numbers with each file name. This catches a chunk that was cut short or saved under the wrong range.

Everything Claude does is reported in words, and the script writes a log (`export-log.txt`) to the output folder. You don't need to watch the screen.

Requirements:
- Windows with Windows PowerShell 5.1 (built in) or PowerShell 7.
- Logos Bible Software (or Verbum), running, with the book open.
- Nothing else needs to be installed.

Example requests:
- "Export the book I have open in Logos to HTML in D:\Books\BDAG."
- "Export the NA28 apparatus per Bible book."

## Installing

Add this folder as a plugin marketplace and install the plugin:

```
/plugin marketplace add <path to this folder>
/plugin install logos-accessibility-skills@logos-accessibility-ai-tools
```

To try it for one session without installing, start Claude Code with `claude --plugin-dir <path to this folder>`.

## Notes

The exports are for your own use of books you have licensed in Logos. The skill uses Logos's own export feature and saves the same files you would get by hand.

## License

MIT; see `LICENSE`.
