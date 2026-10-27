# Logos accessibility AI tools

AI tools for Logos Bible Software on Windows. Some work in Logos is tedious or hard to check without sight. These tools let Claude do that work and report back in words.

The tools come as a Claude Code plugin, `logos-accessibility-skills`.

More skills will follow; each one lives in its own folder under `skills`.

## Skills

### export-book

Exports a whole Logos book to HTML files.

Logos can export what you read through its Print/Export panel, but at most 100 printed pages at a time. Exporting a whole book by hand means repeating the same steps dozens of times: type a page range, press Enter, wait, choose "Web Page (HTML)", name the file and save it.

The skill lets Claude do this for you. You name the book, and Claude finds it in your library and opens it in Logos. A PowerShell script then drives the Print/Export panel through Windows UI Automation. It reads in Logos's library catalog how the book is numbered, and saves one file per chunk of pages. When it's done, it checks that nothing is missing.

- **By page range**: files like `1-100.html`, `101-200.html`. Books with volumes get a prefix (`v1-1-100.html`), and front matter with roman page numbers comes first (`v1-i-xxiii.html`). Page numbers the book doesn't have are skipped and listed in `export-gaps.txt`.
- **By section**: one file per top-level section of the book, such as one Bible book of a commentary: `08 ΚΑΤΑ ΜΑΘΘΑΙΟΝ.html`. A section that is too large for one export is split by its subsections.

When page numbers are shown in the book's text in Logos, the final check also compares those numbers with each file name. This catches a chunk that was cut short or saved under the wrong range.

Everything Claude does is reported in words, and the script writes a log (`export-log.txt`) to the output folder. You don't need to watch the screen.

Requirements:
- Windows 10 or 11, with Windows PowerShell 5.1 (built in) or PowerShell 7.
- Logos Bible Software (or Verbum), running.
- Nothing else needs to be installed.

Example requests:
- "Export BDAG to HTML in D:\Books\BDAG."
- "Export the book I have open in Logos."
- "Export the NA28 apparatus per Bible book."

The skill doesn't read Logos's book files directly: they are encrypted and tied to your license. It only uses Logos's own export feature, and reads the unencrypted library catalog to learn how a book is numbered.

## Installing

In Claude Code, add this repository as a plugin marketplace and install the plugin:

```
/plugin marketplace add LeonarddeR/logos-accessibility-ai-tools
/plugin install logos-accessibility-skills@logos-accessibility-ai-tools
```

To get later versions, run `/plugin marketplace update logos-accessibility-ai-tools`.

### From a local copy

If you work on the plugin itself, clone the repository and add your copy instead:

```
/plugin marketplace add <path to your copy>
/plugin install logos-accessibility-skills@logos-accessibility-ai-tools
```

To try a local copy for one session without installing it, start Claude Code with `claude --plugin-dir <path to your copy>`.

## Notes

The exports are for your own use of books you have licensed in Logos. The skill uses Logos's own export feature and saves the same files you would get by hand.

## License

MIT; see `LICENSE`.
