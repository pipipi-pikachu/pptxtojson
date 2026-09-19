# 🎨 pptxtojson

[![npm-version](https://img.shields.io/npm/v/pptxtojson)](https://www.npmjs.com/package/pptxtojson)
[![npm download](https://img.shields.io/npm/dm/pptxtojson)](https://www.npmjs.com/package/pptxtojson)
[![GitHub issues](https://img.shields.io/github/issues-closed/pipipi-pikachu/pptxtojson)](https://github.com/pipipi-pikachu/pptxtojson/issues)
[![license](https://img.shields.io/github/license/pipipi-pikachu/pptxtojson)](https://www.github.com/pipipi-pikachu/pptxtojson/blob/master/LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/pipipi-pikachu/pptxtojson)](https://www.github.com/pipipi-pikachu/pptxtojson/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/pipipi-pikachu/pptxtojson)](https://www.github.com/pipipi-pikachu/pptxtojson/network/members)

[简体中文](README_zh.md) | English

**pptxtojson is a browser-first PPTX parser that converts `.pptx` files into clean, readable, structured JSON.**

Live demo: https://pipipi-pikachu.github.io/pptxtojson/

> China mirrors, regularly synced: [Gitee](https://gitee.com/pptist/pptxtojson), [GitCode](https://gitcode.com/pipipi-pikachu/pptxtojson)

# ✨ Core Capabilities

- **Browser-side parsing**: Read `.pptx` files directly in the browser, making it suitable for handling local user files without uploading them to a server for conversion.
- **Readable JSON output**: Organize results by slides, elements, assets, themes, and related structures instead of exposing raw Office XML as low-level objects.
- **Built for secondary processing**: Compared with simply converting PPTX to HTML, pptxtojson focuses more on data readability and programmability, making it easier to plug into downstream workflows.

# 🚀 Use Cases

- **Web editor import**: Convert PPTX slides and elements into an editable data model, such as [PPTist](https://github.com/pipipi-pikachu/PPTist).
- **Content extraction**: Extract text, speaker notes, media assets, and other information for search, archiving, review, or data analysis.
- **AI document understanding**: Turn presentation content into structured input for summarization, question answering, knowledge-base ingestion, and similar workflows.
- **Custom rendering**: Build your own preview, thumbnails, editing canvas, or conversion pipeline from the JSON result.

# 🔨 Installation
```
npm install pptxtojson
```

# 💿 Usage
```javascript
parse(file, options = {})
```

### Browser Example
```html
<input type="file" accept="application/vnd.openxmlformats-officedocument.presentationml.presentation"/>
```

```javascript
import { parse } from 'pptxtojson'

document.querySelector('input').addEventListener('change', evt => {
	const file = evt.target.files[0]
	
	const reader = new FileReader()
	reader.onload = async e => {
		const json = await parse(e.target.result, {
			imageMode: 'base64',
			videoMode: 'none',
			audioMode: 'none',
		})
		console.log(json)
	}
	reader.readAsArrayBuffer(file)
})
```

### Node.js Example (experimental, v1.5.0+)
```javascript
const pptxtojson = require('pptxtojson/dist/index.cjs')
const fs = require('fs')

async function func() {
  const buffer = fs.readFileSync('test.pptx')

  const json = await pptxtojson.parse(buffer.buffer, {
    imageMode: 'base64',
    videoMode: 'none',
    audioMode: 'none',
  })
  console.log(json)
}

func()
```

### Output Example
```javascript
{
	"slides": [
		{
			"fill": {
				"type": "color",
				"value": "#FF0000"
			},
			"elements": [
				{
					"left":	0,
					"top": 0,
					"width": 72,
					"height":	72,
					"borderColor": "#1F4E79",
					"borderWidth": 1,
					"borderType": "solid",
					"borderStrokeDasharray": 0,
					"fill": {
						"type": "color",
						"value": "#FF0000"
					},
					"content": "<p style=\"text-align: center;\"><span style=\"font-size: 18pt;font-family: Calibri;\">TEST</span></p>",
					"isFlipV": false,
					"isFlipH": false,
					"rotate": 0,
					"vAlign": "mid",
					"name": "Rectangle 1",
					"type": "shape",
					"shapType": "rect"
				},
				// more...
			],
			"layoutElements": [
				// more...
			],
			"note": "Speaker notes..."
		},
		// more...
	],
	"themeColors": ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'],
	"size": {
		"width": 960,
		"height": 540
	}
}
```

# 🎲 Options
> `options` is optional. Default values are used when it is not provided.

- `imageMode`: Controls how image assets are parsed. Available values: `base64`, `blob`, `both`, `none`. Default: `base64`.
	- `base64` means only `base64` is parsed.
	- `blob` means only `blob` is parsed.
	- `both` means both `base64` and `blob` are parsed.
	- `none` means image content is not parsed.

- `videoMode`: Controls how video assets are parsed. Available values: `blob`, `none`. Default: `none`.
  - `blob` means video `blob` is parsed.
  - `none` means video content is not parsed.

- `audioMode`: Controls how audio assets are parsed. Available values: `blob`, `none`. Default: `none`.
  - `blob` means audio `blob` is parsed.
  - `none` means audio content is not parsed.

- `singleLineSpacingFactor`: Multiplier applied to percentage-based line spacing (`spcPct`). Default: `1`. Set it to a value such as `1.2` to better approximate PowerPoint's font-dependent single-line spacing in HTML rendering. Absolute point spacing (`spcPts`) is unaffected.

# 🎯 Notes

The current parsing result can achieve roughly 80%+ overall fidelity in layout and styling compared with the source file. For PPTX files manually created and edited from scratch by ordinary users, common page structures and basic styles can even reach 95%+ fidelity.

However, if the file comes from a complex online template, or if it was created by a highly skilled PowerPoint user with many "advanced techniques" such as complex masters, deeply nested groups, special shape effects, complex gradients, non-standard shapes, or complex SmartArt, the parsing difficulty increases significantly and the fidelity will decrease accordingly. These files are better treated as complex samples for separate evaluation.

### Length Units
All numeric length values in the output JSON use `pt` (point) as the unit.

### Legacy Version Notes
- In version 0.x, all output length values used px (pixels).
- In version 1.x and earlier:
	- Image elements used the `src` field to return base64 data.
	- Image fills only returned `picBase64`.
	- Video elements might return `blob` or `src`.
	- Audio elements only returned `blob`.
	- Formula images only returned `picBase64`.

# 📕 Parsed Properties

- Slide theme colors `themeColors`

- Embedded font list `usedFonts`

- Slide size `size`
	- Width `width`
	- Height `height`

- Slides `slides`

	- Speaker notes `note`

	- Slide background fill (color, image, gradient, pattern) `fill`
		- Solid color fill `type='color'`
		- Image fill `type='image'`
		- Gradient fill `type='gradient'`
		- Pattern fill `type='pattern'`

	- Slide transition `transition`
		- Type `type`
		- Duration `duration`
		- Direction `direction`

	- Slide elements `elements` / master layout elements `layoutElements`
		- OOXML ID `id`
		- Text
			- Type `type='text'`
			- Horizontal coordinate `left`
			- Vertical coordinate `top`
			- Width `width`
			- Height `height`
			- Border color `borderColor`
			- Border width `borderWidth`
			- Border type (solid, dotted, dashed) `borderType`
			- Non-solid border style `borderStrokeDasharray`
			- Shadow `shadow`
			- Fill (color, image, gradient, pattern) `fill`
			- Text content (HTML rich text) `content`:
				- Inline styles/structure: font family, font size, color, gradient, underline, strikethrough, italic, bold, character spacing, shadow, superscript/subscript, hyperlink
				- Block-level styles/structure: horizontal alignment, line spacing, paragraph spacing, indentation, first-line indentation, bullets, numbered lists
			- Vertical flip `isFlipV`
			- Horizontal flip `isFlipH`
			- Rotation angle `rotate`
			- Vertical alignment `vAlign`
			- Whether text wraps automatically `wrap`
			- Whether it is vertical text `isVertical`
			- Element name `name`
			- Auto fit `autoFit`
				- Type `type`
					- `shape`: the text box height automatically adjusts according to the text content
					- `text`: the text box size is fixed, and the font size is automatically scaled to fit the text box (note: when `autoFit` does not exist, the text box size is also fixed, but the font size is not scaled)
				- Font scale ratio (only for `type='text'`, default is 1) `fontScale`
			- Text inset on four sides `textInset`
			- Hyperlink `link`

		- Image
			- Type `type='image'`
			- Horizontal coordinate `left`
			- Vertical coordinate `top`
			- Width `width`
			- Height `height`
			- Border color `borderColor`
			- Border width `borderWidth`
			- Border type (solid, dotted, dashed) `borderType`
			- Non-solid border style `borderStrokeDasharray`
			- Crop shape `geom`
			- Crop rectangle `rect`
			- Asset reference path `ref`
			- Image base64 `base64`
			- Image blob `blob`
			- Rotation angle `rotate`
			- Filters `filters`
			- Hyperlink `link`

		- Shape
			- Type `type='shape'`
			- Horizontal coordinate `left`
			- Vertical coordinate `top`
			- Width `width`
			- Height `height`
			- Border color `borderColor`
			- Border width `borderWidth`
			- Border type (solid, dotted, dashed) `borderType`
			- Non-solid border style `borderStrokeDasharray`
			- Shadow `shadow`
			- Fill (color, image, gradient, pattern) `fill`
			- Stroke only (no fill) `strokeOnly`
			- Text content (HTML rich text, same as text elements) `content`
			- Vertical flip `isFlipV`
			- Horizontal flip `isFlipH`
			- Rotation angle `rotate`
			- Shape type `shapType`
			- Vertical alignment `vAlign`
			- Whether text wraps automatically `wrap`
			- Shape path `path`
			- Shape path viewBox `pathViewBox`
			- Line head endpoint `headEnd`
			- Line tail endpoint `tailEnd`
			- Shape adjustment parameters `keypoints`
			- Element name `name`
			- Auto fit `autoFit`
			- Text inset on four sides `textInset`
			- Hyperlink `link`

		- Table
			- Type `type='table'`
			- Horizontal coordinate `left`
			- Vertical coordinate `top`
			- Width `width`
			- Height `height`
			- Borders on four sides `borders`
			- Cell styles and data `data`
			- Row heights `rowHeights`
			- Column widths `colWidths`

		- Chart
			- Type `type='chart'`
			- Horizontal coordinate `left`
			- Vertical coordinate `top`
			- Width `width`
			- Height `height`
			- Chart data `data`
			- Chart theme colors `colors`
			- Chart type `chartType`
			- Bar chart direction `barDir`
			- Whether markers are enabled `marker`
			- Doughnut chart hole size `holeSize`
			- Grouping mode `grouping`
			- Chart style `style`

		- Video
			- Type `type='video'`
			- Horizontal coordinate `left`
			- Vertical coordinate `top`
			- Width `width`
			- Height `height`
			- Asset reference path `ref`
			- Video blob `blob`

		- Audio
			- Type `type='audio'`
			- Horizontal coordinate `left`
			- Vertical coordinate `top`
			- Width `width`
			- Height `height`
			- Asset reference path `ref`
			- Audio blob `blob`

		- Formula
			- Type `type='math'`
			- Horizontal coordinate `left`
			- Vertical coordinate `top`
			- Width `width`
			- Height `height`
			- Formula image reference path `picRef`
			- Formula image base64 `picBase64`
			- Formula image blob `picBlob`
			- LaTeX expression (only common structures are supported) `latex`
			- Text (exists when text and formulas are mixed) `text`

		- SmartArt
			- Type `type='diagram'`
			- Horizontal coordinate `left`
			- Vertical coordinate `top`
			- Width `width`
			- Height `height`
			- Child elements `elements`
			- Text list (text content list in SmartArt) `textList`

		- Group
			- Type `type='group'`
			- Horizontal coordinate `left`
			- Vertical coordinate `top`
			- Width `width`
			- Height `height`
			- Child elements `elements`

### See more detailed types here 👇
[https://github.com/pipipi-pikachu/pptxtojson/blob/master/dist/index.d.ts](https://github.com/pipipi-pikachu/pptxtojson/blob/master/dist/index.d.ts)

# 🙏 Acknowledgements
pptxtojson refers heavily to the implementations of [PPTX2HTML](https://github.com/g21589/PPTX2HTML) and [PPTXjs](https://github.com/meshesha/PPTXjs).

Unlike those projects, pptxtojson does not aim to convert PPT files into HTML pages. Instead, it outputs clean, readable JSON that is easier to process further, and it includes many optimizations and additions to improve the completeness and accuracy of extracted information.

Issues and PRs with more PPTX samples, parsing scenarios, and improvement suggestions are welcome.

# 📄 License
MIT License | Copyright © 2020-PRESENT [pipipi-pikachu](https://github.com/pipipi-pikachu)
