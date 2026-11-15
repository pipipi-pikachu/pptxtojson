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
```html
<input type="file" accept=".pptx">
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

For browser parsing, resource cleanup, Node.js, and TypeScript examples, see the [user guide](USER_GUIDE.md#examples).

# 🎯 Notes

The current parsing result can achieve roughly 85%+ overall fidelity in layout and styling compared with the source file. For PPTX files manually created and edited from scratch by ordinary users, common page structures and basic styles can even reach 95%+ fidelity.

However, if the file comes from a complex online template, or if it was created by a highly skilled PowerPoint user with many "advanced techniques" such as complex masters, deeply nested groups, special shape effects, complex gradients, non-standard shapes, or complex SmartArt, the parsing difficulty increases significantly and the fidelity will decrease accordingly. These files are better treated as complex samples for separate evaluation.

# 📕 User Documentation

See the [user guide](USER_GUIDE.md) for the full API, output fields, units, data examples, rendering considerations, and migration notes.

- [API and options](USER_GUIDE.md#api)
- [Ordering, coordinates, and data conventions](USER_GUIDE.md#conventions)
- [Presentation and element fields](USER_GUIDE.md#presentation)
- [Integration and security](USER_GUIDE.md#integration)
- [Migration](USER_GUIDE.md#migration)

TypeScript definitions: [types/index.d.ts](https://github.com/pipipi-pikachu/pptxtojson/blob/master/types/index.d.ts).

# 🙏 Acknowledgements
pptxtojson refers heavily to the implementations of [PPTX2HTML](https://github.com/g21589/PPTX2HTML) and [PPTXjs](https://github.com/meshesha/PPTXjs).

Unlike those projects, pptxtojson does not aim to convert PPT files into HTML pages. Instead, it outputs clean, readable JSON that is easier to process further, and it includes many optimizations and additions to improve the completeness and accuracy of extracted information.

Issues and PRs with more PPTX samples, parsing scenarios, and improvement suggestions are welcome.

# 📄 License
MIT License | Copyright © 2020-PRESENT [pipipi-pikachu](https://github.com/pipipi-pikachu)
