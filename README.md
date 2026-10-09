<h1 align="center">VibeOffice</h1>

<p align="center">
  <b>A word processor, a spreadsheet and a presentation editor that run in your browser.</b><br>
  They open and save real .docx, .xlsx and .pptx files. Nothing to install, no account, and your files never leave your computer.
</p>

<p align="center">
  <a href="https://vibeoffice.work"><b>▶&nbsp;&nbsp;Try it now at vibeoffice.work</b></a>
</p>

<table>
  <tr>
    <td align="center" width="33%">
      <a href="https://vibeoffice.work/quire/?template=flyer"><img src="docs/images/quire.png" alt="Quire, the word processor, editing an event flyer"></a><br>
      <b><a href="https://vibeoffice.work/quire/">Quire</a></b><br>Word processor
    </td>
    <td align="center" width="33%">
      <a href="https://vibeoffice.work/ledger/?template=loan"><img src="docs/images/ledger.png" alt="Ledger, the spreadsheet, with a loan amortization schedule"></a><br>
      <b><a href="https://vibeoffice.work/ledger/">Ledger</a></b><br>Spreadsheet
    </td>
    <td align="center" width="33%">
      <a href="https://vibeoffice.work/lectern/"><img src="docs/images/lectern.png" alt="Lectern, the presentation editor, with a cycle diagram, WordArt and a chart on a slide"></a><br>
      <b><a href="https://vibeoffice.work/lectern/">Lectern</a></b><br>Presentations
    </td>
  </tr>
</table>

## Why VibeOffice

- **Familiar.** The menus, toolbars and task panes of the classic Office 2003 look, so there is nothing new to learn.
- **Real files.** Open what colleagues send you, and save files built to open in Microsoft Office and LibreOffice. Each app is tested against thousands of real documents from public test collections; the results are in its doc.
- **Private.** Everything runs on your machine. There is no server behind it, no upload and no sign-in.
- **Faithful.** When a file uses a feature VibeOffice can't edit yet (VBA macros, for example), it keeps that part and writes it back unchanged when you save. If a save would change something it can't keep, it tells you first.
- **Installable.** In Chrome or Edge, install it from the Start Center. It then opens .docx, .xlsx and .pptx files from your file manager.

## The apps

| App | | Opens and saves |
|---|---|---|
| [Quire](docs/quire.md) | Word processor | .docx, .docm, .dotx, .dotm, .rtf, .htm, .txt, Markdown (.md, or .zip with pictures); PDF |
| [Ledger](docs/ledger.md) | Spreadsheet | .xlsx, .xlsm, .xltx, .xltm, .csv, XML Spreadsheet 2003, .htm; PDF |
| [Lectern](docs/lectern.md) | Presentations | .pptx, .pptm, .ppsx, .ppsm, .potx, .potm; PDF, PNG, .htm |

Start at [vibeoffice.work](https://vibeoffice.work), the Start Center: open a file, pick one from Recent Files, or create a document, spreadsheet or presentation from a blank page or a template.

## Feedback and contributing

Found a file that doesn't open or save the way it should? Please [open an issue](https://github.com/tcchen2026/VibeOffice/issues).

VibeOffice is plain HTML and JavaScript with no build step. Read [AGENTS.md](AGENTS.md) for the design rules and [docs/testing.md](docs/testing.md) for running it locally and testing.

## License

Apache License 2.0, see [LICENSE](LICENSE). The proofing word lists and thesaurus in `public/common/dict/` carry their own licences, listed in `public/common/dict/LICENSES.txt`.

VibeOffice is an independent project. Microsoft, Word, Excel, PowerPoint and Office are trademarks of Microsoft Corporation; VibeOffice is not affiliated with Microsoft.
