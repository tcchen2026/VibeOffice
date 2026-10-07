# VibeOffice

A free office suite that runs in your browser, in the style of Office 2003. Nothing to install, no account, no server: documents are opened, edited and saved on your own computer.

| App | | Opens and saves |
|---|---|---|
| [Quire](docs/quire.md) | Word processor | .docx, .dotx, .rtf, .htm, .txt, Markdown (.md, or .zip with pictures); PDF |
| [Ledger](docs/ledger.md) | Spreadsheet | .xlsx, .xlsm, .xltx, .csv, XML Spreadsheet 2003, .htm; PDF |
| [Lectern](docs/lectern.md) | Presentations | .pptx, .ppsx, .potx; PDF, PNG, .htm |

Each app was tested against thousands of real files from public test corpora; the results are in its doc.

## Running it

Plain HTML and JavaScript, no build step. Serve `public/` with any static web server:

    python3 tools/serve.py

and open http://127.0.0.1:8760/, the Start Center: open a file, pick one from Recent Files, or create a document, spreadsheet or presentation. In Chromium-based browsers it can be installed as an app, which also lets it open .docx, .xlsx and .pptx files from your computer's file manager.

## Contributing

Read [AGENTS.md](AGENTS.md) for the design rules and [docs/testing.md](docs/testing.md) for the tests.

## License

Apache License 2.0, see [LICENSE](LICENSE). The proofing word lists and thesaurus in `public/common/dict/` carry their own licences, listed in `public/common/dict/LICENSES.txt`.

VibeOffice is an independent project. Microsoft, Word, Excel, PowerPoint and Office are trademarks of Microsoft Corporation; VibeOffice is not affiliated with Microsoft.
