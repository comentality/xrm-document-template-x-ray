# Document Template X-Ray

An [XrmToolBox](https://www.xrmtoolbox.com/) tool that extracts and displays all Dynamics 365 field references from Word (.docx) document templates.

When you build Word templates for Dynamics 365, it's easy to lose track of which entity fields, relationships, and repeating sections are actually used. Document Template X-Ray reads the underlying XML content controls and presents every field in a clear flat list or tree view — no need to click through the template one control at a time.

## Features

- **Fetch templates from Dynamics 365** — connects via XrmToolBox and lists all Word document templates in your environment
- **Browse local files** — open any `.docx` template from disk
- **Drag & drop** — drop `.docx` files directly onto the tool
- **Flat list view** — shows every field reference with table, column, tag, alias, repeating section, and location (document body / header / footer)
- **Tree view** — groups fields by their entity/relationship path for a structural overview
- **Display name resolution** — resolves logical names to display names using Dataverse metadata (when connected)
- **Repeating section detection** — identifies and highlights repeating sections and their child fields

## How It Works

Word document templates for Dynamics 365 store field bindings as structured document tags (content controls) in the underlying XML. Each content control has a `w:dataBinding` element with an XPath like:

```
/ns0:DocumentTemplate[1]/account[1]/name[1]
```

The plugin opens the `.docx` as a ZIP archive, reads `word/document.xml` and any `header*.xml` / `footer*.xml` parts, then extracts these XPath bindings and converts them into readable field paths like `account/name`.

When connected to Dynamics 365, it also fetches entity metadata to resolve logical names (e.g., `account/name`) to display names (e.g., Account / Account Name).

## License

MIT
