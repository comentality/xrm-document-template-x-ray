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
but text, so the tile runs down into deep olive, where the pink measures 3.97:1
and the band 3.36:1.

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

## Still to do

The glyph is not yet in [`comentality-brand`](https://github.com/comentality/comentality-brand)`/brand.py`,
so it cannot be re-rendered in the other schemes the way Events2Code can. Porting
it means one glyph function: a list of word lengths per line, the shading ladder,
and two rectangles for the band and the cut.
