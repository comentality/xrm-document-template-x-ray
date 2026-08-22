---
name: xtb-slow-network
description: Document Template XRay's own slow-link suite and the specifics the general procedure needs
---

The general procedure is the user-level `xtb-slow-network` skill. This is what it needs to
know about this repo.

## Run it

From `tests`, in pwsh:

```
.\slow.ps1                          # every scenario
.\slow.ps1 -Scenario cancel         # one, or a comma separated list
.\slow.ps1 -NoBuild                 # reuse the last build
```

Give it a generous timeout — the suite is about a minute of deliberate waiting, and
`slow.ps1` kills the run at five. The exit code is the number of scenarios with findings.
Windows open past the edge of the desktop and take no focus; `DTX_HARNESS_ONSCREEN=1` puts
them in front to watch one play out.

Output lands in `tests\.slow`: a folder per scenario holding a screenshot per gesture, plus
`report.txt` with the same lines the console printed. When a scenario fails, look at the shot
named for the beat before the failing check.

## The parts

| | |
|---|---|
| `tests\DocumentTemplateXRay.SlowHarness\Scenarios.cs` | The nine scenarios. One per failure mode; each says in its summary which one. |
| `tests\DocumentTemplateXRay.SlowHarness\SlowService.cs` | The fake environment. Answers the two things this tool asks, with per-call latency, scripted failures and a call log. |
| `tests\DocumentTemplateXRay.SlowHarness\Probe.cs` | The control from the outside. Every private field name the suite depends on is here, so a rename breaks one file. |
| `tests\DocumentTemplateXRay.SlowHarness\Bench.cs` | Generic: the timeline, the stall watch, the dialog sweeper, the connection setter. Destined for XtbSandbox. |
| `tests\harness\` | `Sample.cs`, `Capture.cs`, `Quiet.cs`. |

## Specifics

- The control is `DocumentTemplateXRay.DocumentTemplateXRayControl`.
- What a fake must answer, and the name each is logged under: `RetrieveMultiple` on
  `documenttemplate` with the `content` column (`templates`), and `RetrieveEntityRequest`
  (`entity`, one per table). The second is the interesting one: resolving one template walks
  its field paths table by table, a heavy round trip each, one after another.
- `SlowService.Latency` and `.Fails` take the call and its per-kind index, so
  `Slow("entity", 1200)` makes every metadata hop take 1.2s and everything else 20ms.
- The templates the fake serves are the real fixtures from `fixtures\`, copied beside the
  harness at build time and handed over as base64 bytes. The tool unzips and scans them with
  `DocxFieldExtractor`, so a scenario asserts on what the tool would really have found.
  `Sample.Metadata()` covers exactly the tables, columns and relationship schema names those
  fixtures name — `04-unresolvable-fields.docx` stays unresolvable here too.
- `Sample.BigTemplate()` builds a template that is big rather than complicated: four fields in
  a document part padded to ~24MB of incompressible text. Big and simple on purpose, so the
  `big-template` scenario measures the reading rather than the drawing. It is cached in
  `%TEMP%\dtx-slow-harness`, so the first run of that scenario is a few seconds longer.
- `Stall` is the measurement `big-template` turns on: a gesture is timed directly (a click
  handler that does not return for a second is a window frozen for a second), and between
  gestures a timer watches for a callback holding the thread. Screenshots are taken inside
  `Stall.Ours`, which does not blame the tool for them. Before the fix, opening the big
  template froze the window for ~400ms; the scenario allows 150.

## Where the answers live

`README.md`, section **Testing**. The control carries the reasoning in the doc comments on
`_fetchGeneration`, `_readGeneration`, `_fetchTrouble`, `_fieldNote`, `_metadata` and `Gate`,
and in `Reselect`, `Read` and `UpdateState`. `MetadataResolver.Resolve` says why it returns
words instead of writing into the fields.

## Do not

- Let the resolver write into `FieldInfo` objects again. It runs on a worker and those fields
  are being drawn on the UI thread at the same time; `Resolve` takes paths and gives back
  words for the UI thread to apply.
- Read a `.docx` on the UI thread. Decoding, writing out, unzipping and scanning all happen in
  `Read`, on a worker — a Word file is as likely to be on a share as on the local disk.
- Throw the metadata cache away with the resolver. `_metadata` outlives it and is handed back
  in, because a RetrieveEntityRequest carries every attribute and every relationship of a
  table and three templates used to pay for the same four tables three times.
- Rebuild the template list without putting the selection back. `Reselect` is what keeps the
  list and the pane describing the same thing.
- Add a test hook to the control. `Probe` reaches private state by reflection on purpose.
