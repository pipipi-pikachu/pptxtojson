# pptxtojson User Guide

[简体中文](USER_GUIDE_zh.md) | English

For **pptxtojson 3.x**. This guide describes the public API and runtime output of **3.0.0**.

pptxtojson converts `.pptx` files into structured data for extraction, conversion, and further processing. It is a **parser, not a PowerPoint renderer**: it extracts slides, text, geometry, charts, and assets. Your application is responsible for drawing slides, editing rich text, loading fonts, playing media, and typesetting equations.

## Contents

- [Installation and runtime environments](#installation)
- [Common usage](#examples)
- [API and options](#api)
- [Shared data conventions](#conventions)
- [Presentation and slides](#presentation)
- [Fills, borders, and shadows](#appearance)
- [Text and shapes](#text-and-shape)
- [Images](#images)
- [Tables](#tables)
- [Charts](#charts)
- [Video and audio](#media)
- [Equations](#math)
- [Groups and SmartArt](#groups)
- [Integration and security](#integration)
- [Migration from earlier versions](#migration)
- [Frequently asked questions](#faq)

<a id="installation"></a>
## Installation and runtime environments

```bash
npm install pptxtojson
```

### Browsers and frontend build tools

In a build environment with ESM support, such as Vite:

```js
import { parse } from 'pptxtojson'
```

The browser must support common APIs such as `ArrayBuffer` and Promise. The examples below also use `File.arrayBuffer()`, `fetch()`, and Blob URLs; provide compatibility implementations if you target older environments.

Without a build tool, load the UMD bundle:

```html
<script src="https://cdn.jsdelivr.net/npm/pptxtojson@3.0.0/dist/index.umd.js"></script>
<script>
  // The API is the same; access it through the global object.
  // const result = await window.pptxtojson.parse(arrayBuffer)
</script>
```

Pin a version in production to avoid unexpected data-structure changes. Loading the library from a CDN is not a conversion service: PPTX data passed to `parse()` is processed in the current environment. The library does not automatically upload it to a conversion server.

### Node.js

Node.js support remains experimental. You can explicitly select a packaged entry point instead of relying on how a tool interprets `main` and `module`:

```js
// CommonJS: save as .cjs, or use in a CommonJS project.
const { parse } = require('pptxtojson/dist/index.cjs')
```

```js
// ESM: save as .mjs, or use in a type: module project.
import { parse } from 'pptxtojson/dist/index.js'
```

Use a supported modern Node.js LTS release. Media `blob` mode also requires `Blob` and `URL.createObjectURL()`. A Blob URL created in Node.js cannot be used directly on another machine or in a browser. For server-side JSON export, use image `base64` mode and leave video/audio loading disabled.

### Supported files

- Input is the binary content of a `.pptx` file, not a file path, URL string, or previously parsed JSON.
- Legacy binary `.ppt`, PDF, and Keynote files are not directly supported.
- Encrypted, damaged, or incomplete packages may fail. Open protected files in an authorized Office environment, decrypt them, and save ordinary PPTX files first.
- Some advanced Office content can only be partially extracted. Successful parsing does not imply identical appearance.

<a id="examples"></a>
## Common usage

These examples cover parsing and reading data, not application UI or rendering. In the browser examples, `file` is an existing `File` object.

### Browser: parse a local file

```js
import { parse } from 'pptxtojson'

try {
  const result = await parse(await file.arrayBuffer(), {
    imageMode: 'base64',
  })
  console.log(result.slides)
  const json = JSON.stringify(result, null, 2)
  // Save or download json in your application.
}
catch (error) {
  console.error('Parsing failed', error)
}
```

Images default to Data URLs, which can be saved with JSON; audio and video content is not loaded by default. To extract text, read Text/Shape `content`, TableCell `text`, Math `latex`, and slide `note`. The first two contain HTML, not plain text. Traverse `elements` recursively for grouped content.

### Browser: media modes and resource cleanup

```js
const result = await parse(await file.arrayBuffer(), {
  imageMode: 'blob',
  videoMode: 'blob',
  audioMode: 'blob',
})

// Use image.blob, video.blob, audio.blob, or math.picBlob.
// These are URL strings; provide your own players and image components.
```

Release Blob URLs only after every consumer has stopped using the media. This function recursively collects and deduplicates URLs, including those in groups, image fills, and equation fallback images:

```js
function releaseResources(result) {
  const urls = new Set()
  function collect(value) {
    if (!value || typeof value !== 'object') return
    for (const [key, item] of Object.entries(value)) {
      if ((key === 'blob' || key === 'picBlob') &&
        typeof item === 'string' && item.startsWith('blob:')) urls.add(item)
      else collect(item)
    }
  }
  collect(result)
  urls.forEach(url => URL.revokeObjectURL(url))
}
```

- Call `releaseResources(result)` at an appropriate point, such as after removing previews or replacing a document—not immediately after parsing.
- Blob URLs belong to the current runtime. Saving JSON does not save their binary content; store media separately for persistence.
- Media may still fail to play if the browser does not support its encoding. Parsing success does not guarantee playback.
- There is no `dispose()` or cancellation API. If parsing fails partway through, assets may already have been created without being returned. Prefer `none` when media is unnecessary; use an isolated page or a suitable Worker environment if you need strict lifecycle isolation.

### Node.js: read a file and save JSON

ESM example: save as `convert.mjs` and run `node convert.mjs`:

```js
import { readFile, writeFile } from 'node:fs/promises'
import { parse } from 'pptxtojson/dist/index.js'

const buffer = await readFile('input.pptx')
// Slice the Buffer's valid range; its underlying ArrayBuffer may contain other data.
const data = buffer.buffer.slice(
  buffer.byteOffset, buffer.byteOffset + buffer.byteLength,
)
const result = await parse(data, { imageMode: 'base64' })
await writeFile('presentation.json', JSON.stringify(result, null, 2), 'utf8')
```

Handle file-reading, parsing, and writing errors in your application. `base64` already contains the `data:image/...;base64,` prefix; do not add it again. Large images can substantially increase JSON size.

### Remote PPTX files and TypeScript

```ts
import { parse, type Options } from 'pptxtojson'

const response = await fetch('/files/example.pptx')
if (!response.ok) throw new Error(`Download failed: HTTP ${response.status}`)
const options: Options = { imageMode: 'none' }
const result = await parse(await response.arrayBuffer(), options)
// The result type is inferred from parse. Element, Chart, and other types are also exported.
```

Browser downloads are subject to CORS and authentication. These are file-retrieval concerns, not responsibilities of `parse()`.

<a id="api"></a>
## API and options

### parse

```ts
parse(file: ArrayBuffer, options?: Options): Promise<{
  slides: Slide[]
  themeColors: string[]
  usedFonts: string[]
  size: { width: number; height: number }
}>
```

- `file`: the complete binary content of a PPTX file.
- `options`: optional configuration, using the defaults below when omitted.
- Success: resolves to the entire presentation's data.
- Failure: rejects the Promise; handle it with `try/catch`. Stable error codes and partial-slide results are not guaranteed.
- There are currently no per-slide callbacks, progress callbacks, `AbortSignal` support, or pagination parameters.

### Options

| Field | Accepted values | Default | Purpose |
|---|---|---|---|
| `imageMode` | `'base64' \| 'blob' \| 'both' \| 'none'` | `'base64'` | Controls image elements, image fills, and equation fallback images |
| `videoMode` | `'blob' \| 'none'` | `'none'` | Whether supported embedded video assets produce Blob URLs |
| `audioMode` | `'blob' \| 'none'` | `'none'` | Whether supported embedded audio assets produce Blob URLs |
| `singleLineSpacingFactor` | A finite positive number | `1` | Multiplier for percentage-based line spacing; also used in the approximate pt conversion of line-relative paragraph spacing. Fixed point spacing is unaffected |

How `imageMode` affects asset fields:

| Value | `base64` / `picBase64` | `blob` / `picBlob` |
|---|---|---|
| `base64` | Data URL | `''` |
| `blob` | `''` | Blob URL |
| `both` | Data URL | Blob URL |
| `none` | `''` | `''` |

Missing relationships or fallback images can also result in empty strings. If the file explicitly references a missing ZIP asset, enabling content loading may cause parsing to fail rather than consistently returning an empty string. `none` does not mean removing the element: position, styles, and `ref` are still returned where possible.

For example, set `singleLineSpacingFactor` to `1.15` to approximate PowerPoint's font-dependent single-line height. It is not a font-size scale and does not guarantee pixel equivalence across fonts. Invalid, non-finite, or non-positive values use `1`. Do not rely on unlisted option values.

<a id="conventions"></a>
## Shared data conventions

### Units and coordinates

| Data | Unit or meaning |
|---|---|
| Slide and element position/size, border width, shadow offsets/blur, text insets, table row heights and column widths | `pt`; 1 inch = 72pt |
| HTML font size, letter spacing, absolute line height, and paragraph spacing | CSS strings with units such as `pt` |
| `rotate`, gradient `rot` | Degrees, not radians or raw OOXML integers |
| `autoFit.fontScale`, image crop `rect` | Percentages: 80 means 80%, not 0.8 |
| Image-fill `opacity` | Ratio: 1 means fully opaque |
| Chart `x/y/size` | Original data values or indices, not layout lengths |
| `transition.duration`, `autoNextAfter` | Milliseconds |
| `keypoints` | Shape adjustment parameters, not lengths; see shape fields |

The origin is the upper-left corner: `left` increases to the right, and `top` increases downward. Negative coordinates and elements extending beyond the slide are possible; do not automatically clamp them to zero.

On a 96dpi CSS-pixel canvas, `px = pt × 96 / 72`. Alternatively, scale an internal canvas that uses pt throughout. Avoid scaling only bounding boxes while forgetting fonts and borders. Do not apply pt conversion to percentages, angles, or chart values.

### Ordering and identifiers

- `slides` follows the presentation's slide list, not the file-name order within the ZIP.
- `id` is a local identifier from OOXML. It is not guaranteed to be globally unique across slides, masters, groups, or SmartArt. It may be an empty string when unavailable, and some SmartArt children may lack it entirely. Generate your own global IDs for an editor or application.
- `order` is a drawing-order key, not a contiguous array index or persistent identifier. Higher values are drawn later. Do not compare absolute values across documents or parsing calls.
- Top-level `layoutElements` usually have negative `order` values, placing master/layout graphics behind current-slide elements. They are not a complete export of the original master tree.

Combine current-slide and inherited graphics for painting:

```js
function getSlidePaintOrder(slide) {
  return [...slide.layoutElements, ...slide.elements]
    .sort((a, b) => a.order - b.order)
}
```

Sort within each group separately, then apply its transform. Do not flatten every nested element into a single global `order` sequence.

### Defaults, empty values, and colors

- Optional fields can be absent or `undefined`; JSON serialization removes object properties whose value is `undefined`.
- Asset strings are usually `''` when a mode is disabled or a reference is missing. Shape/Text `fill` can be `null`; do not directly access `element.fill.type` without checking.
- `borderWidth: 0` means no visible outline. A `borderColor` value alone does not mean the border should be drawn.
- Colors are commonly `#RRGGBB`, but may use short hexadecimal forms or include transparency. Do not assume a fixed length or letter case.
- Unresolved chart-color entries can be `undefined`, becoming `null` after JSON serialization. Choose a theme or default color separately for each missing entry.
- Type declarations provide development assistance. When importing arbitrary documents, still guard against nulls, missing values, and invalid numbers.

### Common element fields

| Field | Type | Description |
|---|---|---|
| `type` | String union | Element kind; determines the remaining structure |
| `id` | `string` | Local source ID, subject to the limitations above |
| `left`, `top` | `number` | Bounding-box position in pt; group children are relative to their immediate container |
| `width`, `height` | `number` | Bounding-box size in pt; zero sizes are possible |
| `order` | `number` | Drawing-order key within the current container |

`name` is mainly provided by Text/Shape, not every element type. Rotation, flipping, and borders are not universal either. Treat missing rotation as 0 and missing flips as false for consumption, without assuming that every original effect has been represented.

<a id="presentation"></a>
## Presentation and slides

### Top-level result

| Field | Type | Description |
|---|---|---|
| `slides` | `Slide[]` | Slide array; may be empty |
| `size.width`, `size.height` | `number` | Presentation slide size in pt, not browser preview dimensions |
| `themeColors` | `string[]` | Presentation accent1–accent6 colors; there may be fewer than six, or none |
| `usedFonts` | `string[]` | Deduplicated font names from the embedded-font list; not every font in use, and not font binaries |

Example structure:

```js
{
  slides: [
    {
      fill: { type: 'color', value: '#FFFFFF' },
      elements: [],
      layoutElements: [],
      note: '',
      transition: null,
    },
  ],
  size: { width: 720, height: 405 },
  themeColors: ['#4472C4', '#ED7D31'],
  usedFonts: [],
}
```

A font name in HTML does not mean the browser has that font installed. A presentation can use many fonts even when `usedFonts=[]`.

### Slide

| Field | Type | Description |
|---|---|---|
| `fill` | `Fill` | Resolved slide background; see fill structures |
| `elements` | `Element[]` | Current-slide elements |
| `layoutElements` | `Element[]` | Non-placeholder master/layout graphics included according to visibility settings; may be empty |
| `note` | `string` | Speaker-note text, usually `''` when absent; not HTML |
| `transition` | Object, `null`, or absent | Slide transition; usually `null` when not specified |

Do not treat `layoutElements` as another slide or manually copy the same master graphics again, which would draw them twice.

### transition

| Field | Type | Description |
|---|---|---|
| `type` | `string` | OOXML transition name, such as `fade` or `push`; `none` means no specific effect was identified |
| `duration` | `number` | Duration in milliseconds; may be derived from speed or a default when no explicit duration exists |
| `direction` | `string \| null` | Source direction such as `l/r/u/d`, not a normalized player enum |
| `autoNextAfter` | Optional `number` | Automatic slide-advance delay in milliseconds. 0 is valid, independently of click-to-advance behavior |

Only data is returned; the library does not start animations or timers. There is no per-element animation timeline or separate field fully exposing whether click-to-advance is enabled.

<a id="appearance"></a>
## Fills, borders, and shadows

### Fill

`fill.type` distinguishes four structures. Text/Shape with no fill may return `null` directly.

| `type` | `value` | Usage and limitations |
|---|---|---|
| `color` | Color string | Solid background or fill |
| `image` | `{ ref, base64, blob, opacity }` | Image asset plus opacity; asset fields depend on `imageMode` |
| `gradient` | `{ path, rot, colors }` | Gradient data, not a ready-to-use complete CSS string |
| `pattern` | `{ type, foregroundColor, backgroundColor }` | Office preset pattern and foreground/background colors; implement pattern rendering yourself |

Image-fill values:

| Field | Type | Description |
|---|---|---|
| `ref` | `string` | Package asset reference such as `ppt/media/image1.png`, not an HTTP URL |
| `base64` | `string` | Complete image Data URL, or `''` |
| `blob` | `string` | Blob URL, or `''`; not a `Blob` object |
| `opacity` | `number` | Usually defaults to 1; a ratio, not a percentage |

Gradient values:

| Field | Type | Description |
|---|---|---|
| `path` | `'line' \| 'circle' \| 'rect' \| 'shape'` | Linear or path-gradient kind |
| `rot` | `number` | Source gradient angle in degrees. CSS uses a different angle convention; do not treat them as equivalent without conversion |
| `colors` | `{ pos: string; color: string }[]` | Color stops. `pos` is usually `'0%'`, `'100%'`, etc.; an omitted source position may appear as an empty string |

```js
const fill = {
  type: 'gradient',
  value: {
    path: 'line',
    rot: 0,
    colors: [
      { pos: '0%', color: '#FF0000' },
      { pos: '100%', color: '#0000FF' },
    ],
  },
}
```

Theme colors and common theme solid/gradient references are resolved where possible. Complete support for all theme images, patterns, effects, and path gradients is not guaranteed.

### Outlines and line ends

Text, Shape, and Image provide these fields directly:

| Field | Type | Description |
|---|---|---|
| `borderColor` | `string` | Outline color |
| `borderWidth` | `number` | Width in pt; 0 means do not draw |
| `borderType` | `'solid' \| 'dashed' \| 'dotted'` | Broad line-style category |
| `borderStrokeDasharray` | `string` | Finer dash pattern, such as `'5, 5, 1, 5'`; usually `'0'` for solid lines. Not an array |

Table `borders` is `{ top?, right?, bottom?, left? }`. Each edge is a `Border` object with at least `borderColor/borderWidth/borderType`; runtime output may also contain **`strokeDasharray`**. This name differs from the element-level **`borderStrokeDasharray`**.

Shapes may include `headEnd` and `tailEnd`:

| Subfield | Type | Description |
|---|---|---|
| `type` | `none/triangle/stealth/diamond/oval/arrow` | Line-end style |
| `width` | Optional `sm/med/lg` | Relative width category, not pt |
| `length` | Optional `sm/med/lg` | Relative length category, not pt |

Not every shape is a straight line. Determine direction from `shapType`, `path`, flipping, and rotation rather than simply connecting the bounding box's top-left and bottom-right corners.

### Shadow

When present, `shadow` has this structure:

| Field | Type | Description |
|---|---|---|
| `h`, `v` | `number` | Horizontal and vertical offsets in pt; can be negative |
| `blur` | `number`, or an empty string in some default cases | Blur radius in pt; use a fallback such as `Number(shadow.blur) || 0` |
| `color` | `string` | Shadow color, possibly with transparency |

Element shadows and text shadows inside HTML apply at different levels. Avoid applying the same shadow twice.

<a id="text-and-shape"></a>
## Text and shapes

### Shared Text and Shape styles

Both include the common geometry fields and the applicable fields below:

| Field | Type | Description |
|---|---|---|
| `name` | `string` | Source element name, not a unique ID |
| `content` | `string` | HTML rich text; may be empty |
| `rotate` | `number` | Element angle; default to 0 when consuming a missing value |
| `isFlipH`, `isFlipV` | `boolean` | Horizontal and vertical flips |
| `fill` | `Fill \| null` | Element background or shape fill |
| Outline fields | See above | Text can have a background and outline too |
| `shadow` | Optional `Shadow` | Element shadow |
| `vAlign` | `'up' \| 'mid' \| 'down'` | Vertical text alignment: top, middle, bottom |
| `wrap` | `boolean` | Whether automatic wrapping is allowed; explicit `<br>` is separate |
| `textDirection` | Optional `TextDirection` | Internal text direction, usually omitted for horizontal text; independent of element rotation |
| `autoFit` | Optional `AutoFit` | Fitting strategy; see below |
| `textInset` | Optional `{ l, t, r, b }` | Four text insets in pt. Default body properties commonly correspond to 7.2pt left/right and 3.6pt top/bottom |
| `link` | Optional `string` | Element-level HTTP(S) hyperlink, separate from inline HTML links |

`type='text'` identifies text elements; `type='shape'` identifies shapes. **Shapes can contain text too**: do not extract text only from `type='text'` elements.

### content structure and consumption

Typical content includes paragraphs, nested spans, links, soft breaks, and lists:

```html
<p style="white-space: pre-wrap;">
  <span style="font-size: 24pt;font-weight: bold;">
    <span style="font-family: Arial;">ABC</span>
    <span style="font-family: SimSun;">中文</span>
  </span>
</p>
```

- Common styles such as font size/family, color, bold, italic, underline, strikethrough, letter spacing, text shadow, and superscript/subscript use HTML/CSS, not separate `fontSize/bold` JSON fields.
- Paragraph styles can include alignment, line height, paragraph spacing, indentation, and first-line indentation. Line-relative spacing uses approximate font metrics, not guaranteed Office pixel equivalence.
- A run can be split into inner spans for Latin and East Asian text. The language of a digit-only or punctuation-only run may affect font selection. Do not use the first span's styles as the styles for the entire paragraph.
- Preserve inheritance and HTML entities. Do not arbitrarily collapse CSS semicolons with regular expressions: in `&quot;;`, the first semicolon belongs to the entity and the second to CSS.
- `white-space: pre-wrap`, `pre`, `tab-size`, and `<br>` affect repeated spaces, tabs, and line breaks. Do not replace them all with spaces or `&nbsp;`.
- Lists may use nested `ul/ol/li`, `start`, `li[value]`, and `data-pptx-level`. The last attribute retains the source level; actual HTML nesting can also reflect explicit indentation overrides.
- Bullets depending on symbol fonts such as Wingdings/Symbol, or characters that cannot be represented reliably, fall back to a default dot. Some special numbering formats are only approximated; complete Office list layout is not promised.
- Span counts, style ordering, and HTML whitespace are not stable contracts. Base application logic on semantics and the DOM, not exact string matching.

Font names are not font files. Your application must load system fonts or appropriately licensed Web Fonts. `usedFonts` does not automatically resolve missing glyphs, fallback fonts, or character-width differences.

### Direction and automatic fitting

`TextDirection` can be:

| Value | Meaning |
|---|---|
| `horz` or absent | Horizontal text |
| `vert`, `vert270` | Two rotated vertical text directions |
| `eaVert` | East Asian vertical text |
| `wordArtVert`, `wordArtVertRtl` | WordArt-related vertical arrangements |
| `mongolianVert` | Mongolian vertical text |

Only Text provides `isVertical`, retained for compatibility with `eaVert`. `false` does not rule out directions such as `vert/vert270`. Apply the direction to the text container rather than simply adding it to the whole shape's `rotate`.

`autoFit`:

| Structure | Description |
|---|---|
| `{ type: 'shape' }` | The text container can adapt to its content |
| `{ type: 'text', fontScale?: number }` | The container size is fixed, and text uses the given scale |
| Absent | No fitting strategy was returned; do not automatically assume geometry must change |

**`fontScale=80` means 80% of the original font size**. Divide by 100 when applying it; treat an omitted scale as 100%. Font sizes in `content` do not already include this scale. Do not apply it twice.

```js
function getFontScale(element) {
  return element.autoFit?.type === 'text'
    ? (element.autoFit.fontScale ?? 100) / 100
    : 1
}
```

### Shape-specific fields

| Field | Type | Description |
|---|---|---|
| `shapType` | `string` | Office preset name, such as `rect/ellipse/triangle/line`, or `custom` |
| `path` | Optional `string` | SVG path `d`, possibly with multiple subpaths; an empty path can be `''` |
| `pathViewBox` | Optional `{ x, y, width, height }` | Path coordinate system, not slide position. Do not reinterpret path values as EMU |
| `strokeOnly` | Optional `boolean` | Outline only; do not automatically add a fill |
| `headEnd`, `tailEnd` | Optional `LineEnd` | End styles; see outline fields |
| `keypoints` | Optional `Record<string, number>` | Adjustment-name/value map. Current values use the source `val / 50000`; they are not universally in the 0–1 range and must not all be clamped |

Prefer `pathViewBox` when mapping a path to the element's width and height. Do not infer bounds from maximum path coordinates and stretch it again. Fall back safely when the path or size is unusable. Preset and custom geometry support is not a complete Office geometry evaluator.

<a id="images"></a>
## Images

`type='image'`, with common geometry fields and:

| Field | Type | Description |
|---|---|---|
| `ref` | `string` | Source asset reference, usually an internal package path; may retain XML entities and should not be a business identifier |
| `base64` | `string` | Data URL, or `''` |
| `blob` | `string` | Blob URL, or `''` |
| `rotate` | `number` | Element rotation |
| `isFlipH`, `isFlipV` | `boolean` | Horizontal and vertical flips |
| Outline fields | See [outlines](#appearance) | Image border |
| `geom` | `string` | Crop shape, potentially a preset name or `custom:` identifier; not a complete clipping algorithm |
| `rect` | Optional `{ l?, t?, r?, b? }` | Amount cropped from the left, top, right, and bottom, in percent. For example, 10 means cropping 10% from that edge; not two corner coordinates |
| `filters` | Optional object | Image-effect parameters below |
| `link` | Optional `string` | Element-level HTTP(S) hyperlink |

`rect` values are not guaranteed to stay within 0–100. Negative crops can express expansion or other framing in the source. Original image bytes are not recropped, and filters are not baked into them; apply these values in your renderer.

`filters`:

| Field | Type | Current value semantics |
|---|---|---|
| `sharpen` | `number` | Sharpening strength: source value divided by 100000 |
| `soften` | May be present at runtime as `number` | Softening strength: absolute value of a non-positive source sharpen/soften value divided by 100000; currently missing from the declaration |
| `colorTemperature` | `number` | Original color-temperature integer, interpreted as K; not normalized |
| `saturation` | `number` | Saturation ratio; 1 corresponds to 100% |
| `brightness` | `number` | Signed source brightness offset divided by 100000; 0 is neutral, not a CSS brightness multiplier |
| `contrast` | `number` | Signed source contrast offset divided by 100000; 0 is neutral, not a CSS contrast multiplier |

The library does not apply or clamp these effects or convert them into CSS filters. Standalone Image elements do not have a general `opacity` field; image-fill opacity is at `fill.value.opacity`.

<a id="tables"></a>
## Tables

### Table

`type='table'`, with common geometry fields and:

| Field | Type | Description |
|---|---|---|
| `data` | `TableCell[][]` | Cells in source row/column order; merge-continuation cells may remain |
| `rowHeights` | `number[]` | Row heights in pt, not ratios |
| `colWidths` | `number[]` | Column widths in pt, not ratios; divide by the total if you need ratios |
| `borders` | Four-sided `Border` object | Table-level borders, not a replacement for cell borders |

### TableCell

| Field | Type | Description |
|---|---|---|
| `text` | `string` | Cell HTML rich text, not plain text |
| `rowSpan`, `colSpan` | Optional `number` | Main-cell span counts; default to 1 |
| `hMerge`, `vMerge` | Optional `number` | Horizontal/vertical merge-continuation flags, usually 1; not main-cell spans |
| `fillColor` | Optional `string` | Cell background; explicitly unfilled cells omit this field in the final output |
| `fontColor` | Optional `string` | Supplemental text color from cell/table styles |
| `fontBold` | Optional `boolean` | Supplemental bold setting from table styles, which may not be repeated in every HTML paragraph |
| `vAlign` | `'up' \| 'mid' \| 'down'` | Cell vertical alignment |
| `borders` | Four-sided `Border` object | Cell borders; can be `{}`, with individual edges missing |

Handle both rich text and table-style fields. Do not discard other text fragments by reading only `querySelector('span')`. Avoid replacing an explicitly unfilled background with your own default fill.

### Reading merged cells

Skip continuation cells and default main-cell spans to 1:

```js
const cells = table.data.map(row => row
  .filter(cell => !cell.hMerge && !cell.vMerge)
  .map(cell => ({
    html: cell.text,
    rowSpan: cell.rowSpan || 1,
    colSpan: cell.colSpan || 1,
  })))
```

Do not also advance through the source array by `colSpan`; that would skip cells twice. If rendering HTML, sanitize `html` and let your table component handle spans and styles. Files without complete built-in style definitions can retain text and explicit properties, but theme styling may be limited.

<a id="charts"></a>
## Charts

### Common fields and types

`type='chart'`, with common geometry fields and:

| Field | Type | Description |
|---|---|---|
| `chartType` | `ChartType` | Determines the `data` / `groups` structure |
| `colors` | `string[]` | Chart colors. Supply fallbacks for missing entries; JSON may contain nulls |
| `data` | Depends on chart type | See common charts, scatter charts, and bubble charts below. Combination charts do not have this top-level field |
| `groups` | `ChartGroup[]` | Only for combination charts |

Recognized type names:

| Type name | Meaning |
|---|---|
| `barChart`, `bar3DChart` | Column/bar charts; see `barDir` for direction |
| `lineChart`, `line3DChart` | Line charts |
| `areaChart`, `area3DChart` | Area charts |
| `pieChart`, `pie3DChart` | Pie charts |
| `doughnutChart` | Doughnut charts |
| `scatterChart` | Scatter charts |
| `bubbleChart` | Bubble charts |
| `radarChart` | Radar charts |
| `surfaceChart`, `surface3DChart` | Surface charts |
| `stockChart` | Stock charts |
| `comboChart` | Multiple supported chart groups in one chart |

Recognizing a type and extracting series does not provide every rendering option. For example, the common model does not fully expose 3D views, legend layout, detailed axis definitions, or all data labels.

### Common charts: ChartItem[]

For charts other than scatter, bubble, and combination charts, `data` is:

```js
[
  {
    key: 'Sales',
    values: [
      { x: '0', y: 10 },
      { x: '1', y: 20 },
    ],
    xlabels: { '0': 'January', '1': 'February' },
  },
]
```

| Field | Type | Description |
|---|---|---|
| `key` | `string` | Series name. Runtime output may fall back to a numeric index when absent; use `String(key)` for display |
| `values` | `{ x: string; y: number }[]` | Points; may be empty |
| `values[].x` | `string` | Category/point-cache index, not label text or a numeric scatter coordinate |
| `values[].y` | `number` | Numeric value |
| `xlabels` | `Record<string, string>` | Index-to-category-label map; may be empty |

Match labels by index instead of blindly pairing `Object.values(xlabels)` and `values` by array position:

```js
const points = item.values.map(point => ({
  label: item.xlabels[point.x] ?? point.x,
  value: point.y,
}))
```

The parser reads chart XML reference caches or literal data. It **does not execute formulas or recalculate embedded Excel workbooks**. If the workbook changed without refreshing the chart cache, the output may still contain stale values.

### Scatter charts: ScatterSeries[]

```js
[
  { key: 'Series A', values: [{ x: 1, y: 11 }, { x: 2, y: 22 }] },
  { key: 'Series B', values: [{ x: 100, y: 44 }, { x: 200, y: 55 }] },
]
```

- `key` is the series name; `values[].x/y` are both numbers.
- Each series retains its own X values. Do not reuse the first series' X values for others.
- Points are paired by their source indices and returned in index order. Missing pairs or non-finite values are not replaced with zero. Public point objects do not include an additional source `idx`.
- Points are not guaranteed to be sorted by X. Do not arbitrarily reorder them to draw lines.

For a charting library that accepts `[x,y]` arrays:

```js
const series = chart.data.map(item => ({
  name: item.key,
  points: item.values.map(point => [point.x, point.y]),
}))
```

### Bubble charts: BubbleSeries[]

```js
[
  {
    key: 'Sample A',
    values: [
      { x: 0.7, y: 2.7, size: 10 },
      { x: 1.8, y: 3.2, size: 4 },
    ],
  },
]
```

`size` is the original size value, not a pixel radius, diameter, or normalized area. Apply your renderer's scaling rules together with `sizeRepresents`. Zero and negative values may be retained. X/Y/size are matched by the same source point index.

Optional fields:

| Field | Type | Description |
|---|---|---|
| `bubbleScale` | `number` | Bubble-size scale percentage; 100 means 100% |
| `showNegBubbles` | `boolean` | Source setting for displaying negative bubbles |
| `sizeRepresents` | `'area' \| 'w'` | Whether size represents area or width |

These are returned only when the source provides corresponding information. Distinguish an omitted field from explicit false or 0.

### Combination charts: groups

```js
{
  type: 'chart',
  chartType: 'comboChart',
  left: 72, top: 72, width: 400, height: 240,
  order: 10, id: '6',
  colors: ['#4472C4', '#ED7D31'],
  groups: [
    {
      chartType: 'barChart', colors: ['#4472C4'],
      data: [{ key: 'Sales', values: [{ x: '0', y: 10 }], xlabels: { '0': 'January' } }],
      barDir: 'col', grouping: 'clustered',
      axisIds: ['100', '200'], seriesIndices: [0],
    },
    {
      chartType: 'lineChart', colors: ['#ED7D31'],
      data: [{ key: 'Share', values: [{ x: '0', y: 0.5 }], xlabels: { '0': 'January' } }],
      axisIds: ['100', '300'], seriesIndices: [1],
    },
  ],
}
```

- This remains one chart element, with position and size at the top level. Groups are not separate slide elements.
- There is **no top-level `data`**. Each group has its own `chartType/data/colors` and applicable options.
- `axisIds: string[]`: axes referenced by the group; may be empty. These are not complete axis configurations.
- `seriesIndices: number[]`: source series indices, not point indices within each group.
- Groups retain source order. Top-level colors aggregate the groups' colors.
- Do not merge distinct axis IDs into one axis. However, the available fields alone cannot guarantee exact secondary-axis scales and positioning.

Read each group's own data structure:

```js
const groups = chart.chartType === 'comboChart' ? chart.groups : [chart]
for (const group of groups) {
  console.log(group.chartType, group.data)
}
```

Do not concatenate different groups' `data` as though all point structures were identical. Read each group according to its type, as described above.

### Other chart options

| Field | Commonly applies to | Description |
|---|---|---|
| `barDir` | Column/bar charts | `col` means vertical columns; `bar` means horizontal bars |
| `grouping` | Bar, line, area, etc. | Values such as `clustered/standard/stacked/percentStacked`. Returned values are not already stacked or converted to percentages |
| `marker` | Common line charts | Basic marker hint, not complete per-point marker styling |
| `holeSize` | Doughnut charts | Percentage as a string, such as `'50'`; not `'50%'` or px |
| `style` | Scatter, radar, etc. | Source style name, such as scatter `lineMarker` or radar `filled`; not CSS |

Chart titles, units, date axes, 3D settings, financial/surface-chart options, and internal formatting may be incomplete. Decide which fields your application can map and which need fallback behavior. Do not treat every type as a bar chart or every `data` field as a two-dimensional array.

<a id="media"></a>
## Video and audio

`type='video'` or `type='audio'`, with common geometry fields and:

| Field | Type | Description |
|---|---|---|
| `ref` | `string` | Internal media path or external video reference |
| `blob` | `string` | Blob URL, or `''`; controlled by the corresponding media mode |
| `rotate` | May be present at runtime as `number` | Angle; default to 0 when missing. Currently absent from the declaration |

- Video loading mainly covers mp4/webm/ogg; audio covers mp3/wav/ogg. Other formats may return only `ref` and an empty `blob`.
- External videos may have only `ref`. `videoMode='blob'` does not mean the external URL is downloaded. It may be a web page rather than a playable `<video src>` URL; do not automatically fetch or embed it without checking.
- There is no general base64 output for audio/video, and no `Blob` instances are returned.
- External references can retain entity/escaping forms. External audio has no dedicated handling equivalent to video; direct playback is not promised.
- Some audio/video elements require preview-image relationships to be recognized. Elements without relationships may be skipped, and missing asset content can cause loading to fail.
- Players, codec compatibility, and autoplay permission are browser/application concerns. Complete Office playback controls, trimming, and timelines are not provided.

<a id="math"></a>
## Equations

`type='math'`, with common geometry fields and:

| Field | Type | Description |
|---|---|---|
| `latex` | `string` | Conversion of common OMML structures; not an array and not automatically wrapped in `$...$` delimiters |
| `picRef` | `string` | Equation fallback-image reference; `''` when absent |
| `picBase64` | `string` | Fallback image Data URL, controlled by `imageMode` |
| `picBlob` | `string` | Fallback image Blob URL, controlled by `imageMode` |
| `text` | Optional or empty `string` | Mixed-content HTML that has not been incorporated into LaTeX |

Common conversions include fractions, scripts, radicals, matrices, delimiters, some integrals/sums, functions, limits, and accents. Validate complex OMML and command compatibility with your chosen renderer on real presentations.

Multiple expressions in one text box:

- A single expression remains a string, without an extra array wrapper.
- Inline expressions may be joined with `\quad`.
- Multiple lines may use `gathered`, or a single-column `array` for explicit left/right alignment.
- Text mixed with multiple expressions may be included in `\text{...}`. Text already incorporated into LaTeX is not also returned through `text`.
- The image and element bounds remain a single whole, not split by equation count.

```js
const expressions = [
  'x_{i}^{2}',
  '\\begin{gathered}x=1 \\\\ y=2\\end{gathered}',
]
```

For appearance-sensitive use, consider displaying the original fallback image first, but not every equation has one. If using KaTeX/MathJax, configure it yourself, render safely, and verify command support. A nonempty LaTeX string is not a guarantee of exact reconstruction.

<a id="groups"></a>
## Groups and SmartArt

### Group

| Field | Type | Description |
|---|---|---|
| `type` | `'group'` | Multiple grouped elements |
| Common geometry fields | See [shared data conventions](#conventions) | Parent-group position and size in pt |
| `rotate` | `number` | Parent-group rotation |
| `isFlipH`, `isFlipV` | `boolean` | Parent-group flips |
| `elements` | Element array | May include nested groups; runtime data is broader than the current `BaseElement[]` declaration |

Coordinate handling:

- Child `left/top` values are relative to the immediate parent group, not absolute slide coordinates.
- Source group-coordinate scaling and offsets are already folded into child `left/top/width/height` where possible. **Do not rescale again from the original chExt or child bounds.**
- Parent rotation and flips are not uniformly baked into each child. Keeping a group container and applying transforms around its center is usually preferable to flattening.
- Only without rotation/flipping can flattening simply add parent and child `left/top`; repeat recursively for nested groups.
- Consistent matrix conversion of all properties is not guaranteed for complex nonuniform scaling, deep groups, or tables/SmartArt inside groups. Table row/column sizes, path viewBoxes, and outlines cannot be assumed fully scaled just because a bounding box was scaled.

### Diagram / SmartArt

| Field | Type | Description |
|---|---|---|
| `type` | `'diagram'` | SmartArt data |
| Common geometry fields | See [shared data conventions](#conventions) | Overall position and size |
| `elements` | `(Shape \| Text)[]` | Internal elements extracted where a readable cached drawing is available |
| `textList` | `string[]` | Text extracted from the diagram data model |

`elements` may be empty while `textList` is not; both can also be empty. This is not a full SmartArt layout engine. It mainly reads direct shapes from cached drawings and does not guarantee recursive reconstruction of every group, image, and connector. Child coordinates come from the saved drawing, without the outer position added or a uniform fit to the outer dimensions. Handle positioning and any coordinate-system adaptation yourself; do not assume Group's folded-in scaling rules apply. `textList` can duplicate `elements[].content`; choose a primary text source to avoid double counting.

<a id="integration"></a>
## Integration and security

### Integrating a rich-text editor

- Read rich text through the DOM or the editor's parser. Do not try to "repair" entities or nested spans with string replacements.
- Check support for font names/sizes, bold, colors, links, nested lists, numbering starts, and soft breaks. Define fallbacks for unsupported styles.
- Editors may normalize HTML, whitespace, and default paragraph alignment. Do not require byte-for-byte identical serialization.
- Font names containing spaces or Chinese characters may be quoted and HTML-escaped. This is valid syntax, not corrupted text.
- Copying only `content` while ignoring `autoFit`, insets, vertical alignment, and text direction can still cause oversized text, overflow, or layout differences.

### HTML, links, and user files

PPTX files may come from untrusted sources. **The output is not security-audited HTML and is not a substitute for sanitization.**

- Use `textContent` when displaying JSON or plain text. Before rendering rich text with `innerHTML` or an editor, sanitize it with a maintained HTML sanitizer and suitable CSS/link allowlists.
- Restrict link protocols and navigation behavior, and use an appropriate `rel` when opening new windows. Do not automatically visit every `link`.
- Asset `ref` is not a trusted filesystem path. Do not use it to access server files directly.
- Enforce application-level limits for input size, expanded ZIP size, parsing time, and memory use. Do not process arbitrary uploads without limits in a high-concurrency service.
- Parsing may occupy the main thread. Consider a Worker for large files, but first verify bundling, Blob URLs, and asset transfer for your target workflow.

### Persisting JSON

- `base64` Data URLs can be saved with JSON but increase its size.
- `blob` URLs belong to the current runtime and are not persistent storage.
- `ref` points into the source package. Saving JSON alone and discarding the PPTX/assets does not let the library automatically recover media later.
- For long-term image/audio/video storage, extract and store assets yourself, then map JSON references to application-managed URLs.

### Capabilities that are not fully guaranteed

Do not treat this parser as a lossless Office replacement. Assess complete animation timelines, complex 3D/effects, all preset geometry formulas, every chart axis/annotation option, special fonts and complex-script shaping, picture bullets, full SmartArt layout, embedded objects, and external resources separately. Determine actual support from the target document's output and rendering results.

<a id="migration"></a>
## Migration from earlier versions

### Upgrading from 2.2.0 or earlier to 3.x

Check the following:

- **Scatter/bubble `data` is no longer a two-dimensional array.** Use series of point objects; bubble points contain original `size` values, and each series retains independent X values.
- **Combination charts use `chartType='comboChart'` and `groups`.** Do not read top-level `data` or assume that the last group represents the whole chart.
- **Common-chart `ChartItem[]` is unchanged.** Do not migrate every chart into scatter-point data.
- **Text HTML can be more complex.** Allow nested font spans, real lists, and soft breaks. Do not depend on previous span counts or read only the first span.
- **Use `textDirection` for internal text direction.** `isVertical` retains only limited compatibility behavior.
- **Use `order` to position inherited graphics correctly.** It need not start at 0, and local IDs are not global IDs.

If older JSON already discarded independent X values or bubble sizes, it cannot be converted losslessly using that JSON alone. Parse the original PPTX again.

### Older versions

- Version 0.x used px for lengths; current lengths use pt.
- Migrate Image `src`, old image-fill `picBase64`, and similar fields from 1.x or earlier to the current `ref/base64/blob` structure.
- Equation assets use `picRef/picBase64/picBlob`; audio/video use `ref/blob`. Do not continue relying on old `src` fields.

<a id="faq"></a>
## Frequently asked questions

### Why do fonts differ even though parsing succeeded?

Output records font selection; it does not install fonts. Missing fonts, different versions, browser fallback, discarded editor marks, or incorrect entity replacement can alter glyphs and widths. Inspect the original `content`, then the imported DOM and computed styles, to distinguish parsing, conversion, and system-font issues.

### Why is there still so much data with imageMode set to none?

It disables image-content loading only, not text, position, style, or reference extraction. It does not disable chart or SmartArt parsing either.

### Why does chart data differ from Excel?

The parser uses chart caches/literal data; it does not run Excel calculations. Refresh charts in an authorized Office environment, save, and parse again.

### Why are element IDs duplicated and order values not 0, 1, 2?

They are local source identifiers and sorting keys, not database primary keys and array indices. Generate application IDs and sort within each container.

### Can I pass the JSON directly to PPTist, ECharts, or any editor?

Usually you need an adapter. Tools differ in coordinate units, chart data, rich-text schemas, shape paths, group transforms, and asset lifecycles. In particular, do not compress independent-X scatter series into shared-X data and still claim lossless fidelity.

### Can this library directly output a PNG, PDF, or complete HTML for each slide?

There is no general slide-rendering/export API. You can render the parsed output and then export it yourself. For high-fidelity full-slide conversion, choose a suitable Office-rendering solution rather than equating content parsing with pixel reconstruction.

---

API type reference: [TypeScript definitions](https://github.com/pipipi-pikachu/pptxtojson/blob/master/types/index.d.ts). The npm package provides type hints through `dist/index.d.ts`.
