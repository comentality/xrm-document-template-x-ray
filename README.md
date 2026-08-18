# Document Template XRay

[![NuGet](https://img.shields.io/nuget/v/DocumentTemplateXRay)](https://www.nuget.org/packages/DocumentTemplateXRay)
[![XrmToolBox](https://img.shields.io/badge/XrmToolBox-Plugin-blue)](https://www.xrmtoolbox.com/plugins/plugininfo/?id=6284fa8b-ac36-f111-9a90-7ced8d45b89f)

An [XrmToolBox](https://www.xrmtoolbox.com/) tool that extracts and displays all Dynamics 365 field references from Word (.docx) document templates.

When you build Word templates for Dynamics 365, it's easy to lose track of which entity fields, relationships, and repeating sections are actually used. Document Template XRay reads the underlying XML content controls and presents every field in a clear flat list or tree view — no need to click through the template one control at a time.

## The problem it solves

Here is a real template open in Word:

![A Dynamics 365 template open in Word](https://raw.githubusercontent.com/comentality/xrm-document-template-x-ray/main/docs/screenshots/word-01-duplicate-columns.png)

Three of those lines say `description`. Two say `name`. `address1_city` appears four times. Column names repeat across tables — every table in Dataverse has a `description` and a `createdon` — so the moment a template pulls in a related record, the page fills up with controls that look identical and read from completely different places. Nothing you can see says which is which: the difference is a data binding buried in the XML, and Word will not show it to you without clicking each control in turn.

Here is the same file in Document Template XRay:

![The same template in Document Template XRay](https://raw.githubusercontent.com/comentality/xrm-document-template-x-ray/main/docs/screenshots/tool-01-flat-list.png)

The three `description` fields turn out to be `account/description`, `account/account_parent_account/description` and `account/account_primary_contact/description` — the account, its parent account, and its primary contact. The Table and Column columns name each one the way the business does, so `telephone1` on the account is *Main Phone* while `telephone1` on the contact is *Business Phone*.

## Features

- **Fetch templates from Dynamics 365** — connects via XrmToolBox and lists all Word document templates in your environment
- **Browse local files** — open any `.docx` template from disk
- **Drag & drop** — drop `.docx` files directly onto the tool
- **Flat list view** — shows every field reference with table, column, tag, alias, repeating section, and location (document body / header / footer)
- **Tree view** — groups fields by their entity/relationship path for a structural overview
- **Display name resolution** — resolves logical names to display names using Dataverse metadata (when connected)
- **Repeating section detection** — identifies and highlights repeating sections and their child fields

### Tree view

Tree view makes the same point structurally — three `description` leaves hanging off three different parents:

![Tree view](https://raw.githubusercontent.com/comentality/xrm-document-template-x-ray/main/docs/screenshots/tool-01-tree-view.png)

### Repeating sections and document parts

Repeating sections appear in bold with `(section)`, and every field inside one is coloured and tagged with the section it repeats under — so you can tell the `description` that prints once from the one that prints per contact and the one that prints per task:

![Repeating sections](https://raw.githubusercontent.com/comentality/xrm-document-template-x-ray/main/docs/screenshots/tool-02-repeating-sections.png)

The Location column covers the fields that are easiest to miss entirely: the ones in the header and the footer rather than the body.

![Header and footer fields](https://raw.githubusercontent.com/comentality/xrm-document-template-x-ray/main/docs/screenshots/tool-03-header-footer.png)

## How It Works

Word document templates for Dynamics 365 store field bindings as structured document tags (content controls) in the underlying XML. Each content control has a `w:dataBinding` element with an XPath like:

```
/ns0:DocumentTemplate[1]/account[1]/name[1]
```

The plugin opens the `.docx` as a ZIP archive, reads `word/document.xml` and any `header*.xml` / `footer*.xml` parts, then extracts these XPath bindings and converts them into readable field paths like `account/name`.

When connected to Dynamics 365, it also fetches entity metadata to resolve logical names (e.g., `account/name`) to display names (e.g., Account / Account Name).

A template built against an environment it no longer matches still reads: every binding is listed, and only the display names it cannot resolve are left blank.

![A stale template](https://raw.githubusercontent.com/comentality/xrm-document-template-x-ray/main/docs/screenshots/tool-04-unresolvable.png)

## Testing

`fixtures/` holds a Word template per behaviour worth checking — duplicate column names, repeating sections, header and footer fields, stale bindings, and a document with no bindings at all — together with the script that generates them. See [fixtures/README.md](https://github.com/comentality/xrm-document-template-x-ray/blob/main/fixtures/README.md) for what each one proves. Every screenshot above is one of them, loaded against a live environment.

`xtb.ps1` builds the tool and launches a private XrmToolBox containing nothing but it, connected to the current `pac auth` environment, so you can try a change against a real environment without disturbing the XrmToolBox you use for work.

## License

MIT
