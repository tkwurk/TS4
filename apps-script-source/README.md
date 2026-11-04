# Apps Script source (reference only)

Drop the original Google Apps Script project files here (`.gs`, `.js`, `.html`,
`appsscript.json`) when you're ready to migrate them. Nothing in this folder is
bundled by Vite — it exists purely as a reference while porting logic into
`src/features/`.

## How to export your Apps Script project

Using [clasp](https://github.com/google/clasp):

```bash
npm install -g @google/clasp
clasp login
clasp clone <scriptId>   # or: clasp pull inside an existing .clasp.json project
```

Then copy the pulled files into this folder.

Alternatively, in the Apps Script editor: **Project Settings** gives you the
script ID, and you can copy each file's contents manually.

Once files are here, follow the mapping table in
[`../src/features/README.md`](../src/features/README.md) to port each module.
