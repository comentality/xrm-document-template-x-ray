# Icon

The shipped icon is `DocumentTemplateXRay/icon.svg`. Everything else — the two
`ExportMetadata` base64 blobs in `DocumentTemplateXRayPlugin.cs` and the nuspec's
`icon.png` — is rendered from it at 32 px and 80 px, and all three have to be
updated together or XrmToolBox shows a mix.

## What it draws

A page of a template, mid-scan.

| Part | Colour | Meaning |
|---|---|---|
| the tile | olive `#798156` down to deep olive `#3E4426` | the document itself |
| lines below the cut | ink at 60% | text the tool has not read yet |
| the cut | lilac `#B3A9FF`, hard edge | the ray's leading edge, travelling **down** |
| the band above it | violet `#7C6DF2`, fading out upward | where the ray has already been |
| lines above the cut | pink `#FF3D8B`, per-word shading | the field references it found |

Two details carry most of the reading and are easy to undo by accident:

- **The band is sharp on the bottom and fades on the top.** A hard edge is where
  something is arriving, a fade is where it has been; reverse them and the icon
  reads as a scan travelling upward.
- **Words are packed shoulder to shoulder and separated by brightness, not by
  gaps.** A round-capped stroke extends half its width past each end, so a gap
  the same size as the stroke closes to nothing and the line fuses into one bar.
  If the words ever merge, widen the shading spread before touching the spacing.

Line ends are deliberately uneven — 46, 48, 38, 49, 45 in a 64-unit box. An
evenly increasing set reads as a bar chart rather than as text.

## The tile, and why this one

Pink on the light tent-olive tile measures **1.24:1** at the top of the gradient,
and the violet band measures **1.05:1** — the brand guide's own warning is that
pink on olive "depends on the accent being a large, solid shape. Do not use it
for thin strokes, small marks, or anything carrying text." This icon is nothing
but text, so the tile runs down into deep olive, where the pink measures 3.06:1
and the ray's leading edge 4.87:1. (The band itself only reaches 2.59:1 there,
which is why the edge and not the band is what has to carry the reading.)

The alternates here were built and looked at in a real XrmToolBox tool list
beside Events2Code before this one was chosen:

| File | Tile | Why it lost |
|---|---|---|
| `alt-o1-light-tile.svg` | Events2Code's olive exactly | the dim words in each line sink into the tile; the band all but disappears |
| `alt-o4-mid-tile.svg` | olive_lo → deep olive | reads well, but sits furthest from any scheme in `brand.py` |


## The tile XrmToolBox draws around it

The plugin also sets the surface — the card behind the icon on the Tools list:

```csharp
[ExportMetadata("BackgroundColor", "#F2F3EC")]
[ExportMetadata("PrimaryFontColor", "#222512")]
[ExportMetadata("SecondaryFontColor", "#5F673E")]
```

That is the brand's `tint` surface (`python brand.py --surface tint` prints these
lines), the same one Events2Code uses, so the two cards read as a set instead of
this one sitting on stock white beside it.

The rule that matters if it is ever changed: **a surface and the icon ground on
it must never share a value.** `tint` is far lighter than this icon's olive-to-deep
tile, so the rounded square keeps its edge. A surface close to the icon's own
green would dissolve that edge and leave the pink text floating in a field.

## Re-rendering it

The glyph now lives in [`comentality-brand`](https://github.com/comentality/comentality-brand)`/brand.py`
as `documenttemplatexray`, so it can be rendered the way Events2Code can:

```powershell
python brand.py --glyph documenttemplatexray                       # every scheme, 80px and 32px
python brand.py --glyph documenttemplatexray --base64 full-olive   # the two ExportMetadata blobs
```

The tile it ships on is the kit's `full-olive` scheme, added with the glyph —
`#798156` down to `#3E4426`, the one scheme whose top stays light enough to read
as paper while its bottom is dark enough to carry text.

`icon.svg` here is still the source of record for what shipped, and the two
renderers do not agree to the byte: the Python render differs from it by a mean
of 4/255 across the tile, all of it anti-aliasing along the word ends. If the
icon is ever regenerated from `brand.py`, re-render the base64 blobs and the
nuspec `icon.png` from the same run so all three stay in step.
