# 🎨 pptxtojson

[![npm-version](https://img.shields.io/npm/v/pptxtojson)](https://www.npmjs.com/package/pptxtojson)
[![npm download](https://img.shields.io/npm/dm/pptxtojson)](https://www.npmjs.com/package/pptxtojson)
[![GitHub issues](https://img.shields.io/github/issues-closed/pipipi-pikachu/pptxtojson)](https://github.com/pipipi-pikachu/pptxtojson/issues)
[![license](https://img.shields.io/github/license/pipipi-pikachu/pptxtojson)](https://www.github.com/pipipi-pikachu/pptxtojson/blob/master/LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/pipipi-pikachu/pptxtojson)](https://www.github.com/pipipi-pikachu/pptxtojson/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/pipipi-pikachu/pptxtojson)](https://www.github.com/pipipi-pikachu/pptxtojson/network/members)

简体中文 | [English](README.md)

**pptxtojson 是一个浏览器优先的 PPTX 解析库，可以将 `.pptx` 文件转换为干净、可读、结构化的 JSON。**

在线体验：https://pipipi-pikachu.github.io/pptxtojson/

> 国内镜像（定期同步）：[Gitee](https://gitee.com/pptist/pptxtojson)、[GitCode](https://gitcode.com/pipipi-pikachu/pptxtojson)

# ✨ 核心能力

- **浏览器端解析**：直接读取 `.pptx` 文件，适合在前端处理用户本地文件，无需上传服务端转换。
- **可读 JSON 输出**：按页面、元素、资源、主题等维度组织结果，而不是把 Office XML 原样翻译成低层结构。
- **面向二次处理**：相比单纯转成 HTML 页面，更关注数据可读性和可编程性，方便接入后续业务流程。

# 🚀 适用场景

- **Web 编辑器导入**：将 PPTX 页面和元素转换为可编辑的数据模型（如 [PPTist](https://github.com/pipipi-pikachu/PPTist)）。
- **内容提取**：抽取文本、备注、媒体资源，用于搜索、归档、审核或数据分析。
- **AI 文档理解**：把幻灯片内容整理为结构化输入，用于总结、问答、知识库入库等流程。
- **自定义渲染**：基于 JSON 结果生成自己的预览、缩略图、编辑画布或转换流程。

# 🔨安装
```
npm install pptxtojson
```

# 💿用法
```javascript
parse(file, options = {})
```

### 浏览器示例
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

### Node.js 示例（实验性，1.5.0以上版本）
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

### 输出示例
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
					"name": "矩形 1",
					"type": "shape",
					"shapType": "rect"
				},
				// more...
			],
			"layoutElements": [
				// more...
			],
			"note": "演讲者备注内容..."
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

# 🎲 Options 配置说明
> options 为可选参数，不传时使用默认配置。

- `imageMode`：控制图片资源的解析方式，可选值为 `base64`、`blob`、`both`、`none`，默认值为 `base64`。
	- `base64` 表示仅解析 `base64`。
	- `blob` 表示仅解析 `blob`。
	- `both` 表示同时解析 `base64` 和 `blob`。
	- `none` 表示不解析图片内容。

- `videoMode`：控制视频资源的解析方式，可选值为 `blob`、`none`，默认值为 `none`。
  - `blob` 表示解析视频 `blob`。
  - `none` 表示不解析视频内容。

- `audioMode`：控制音频资源的解析方式，可选值为 `blob`、`none`，默认值为 `none`。
  - `blob` 表示解析音频 `blob`。
  - `none` 表示不解析音频内容。

- `singleLineSpacingFactor`：对百分比行距（`spcPct`）应用的倍率，默认值为 `1`。HTML 渲染时可设为 `1.2` 等值，以近似 PowerPoint 与字体相关的单倍行距。绝对磅值行距（`spcPts`）不受影响。

# 🎯 注意事项

目前解析结果与源文件在排版和样式上的综合还原度约为 80%+。对于普通用户手动从头创建和编辑的 PPTX，常见页面结构与基础样式甚至可以达到 95%+ 的还原度。

但是如果文件来自网上下载的复杂模板，或由 PPT 水平较高的用户大量使用“高级技巧”制作（如复杂母版、深层组合、特殊图形效果、复杂渐变、非标准形状、复杂 SmartArt 等），解析难度会明显上升，还原度也会相应降低。此类文件更适合作为复杂样本单独评估。

### 长度值单位
输出的 JSON 中，所有数值长度值单位都为 `pt`（point）。

### 旧版本说明
- 在0.x版本中，所有输出的长度值单位都是px（像素）
- 在1.x及以下版本：
	- 图片元素使用 `src` 字段返回 base64 数据；
	- 图片填充仅返回 `picBase64`；
	- 视频元素可能返回 `blob` 或 `src`；
	- 音频元素仅返回 `blob`；
	- 公式图片仅返回 `picBase64`；

# 📕 解析属性

- 幻灯片主题色 `themeColors`

- 内嵌字体清单 `usedFonts`

- 幻灯片尺寸 `size`
	- 宽度 `width`
	- 高度 `height`

- 幻灯片页面 `slides`

	- 页面备注 `note`

	- 页面背景填充（颜色、图片、渐变、图案） `fill`
		- 纯色填充 `type='color'`
		- 图片填充 `type='image'`
		- 渐变填充 `type='gradient'`
		- 图案填充 `type='pattern'`

	- 页面切换动画 `transition`
		- 类型 `type`
		- 持续时间 `duration`
		- 方向 `direction`

	- 页面内元素 `elements` / 母版元素 `layoutElements`
		- OOXML ID `id`
		- 文字
			- 类型 `type='text'`
			- 水平坐标 `left`
			- 垂直坐标 `top`
			- 宽度 `width`
			- 高度 `height`
			- 边框颜色 `borderColor`
			- 边框宽度 `borderWidth`
			- 边框类型（实线、点线、虚线） `borderType`
			- 非实线边框样式 `borderStrokeDasharray`
			- 阴影 `shadow`
			- 填充（颜色、图片、渐变、图案） `fill`
			- 内容文字（HTML富文本） `content`：
				- 行内样式/结构：字体、字号、颜色、渐变、下划线、删除线、斜体、加粗、字间距、阴影、角标、超链接
				- 块级样式/结构：水平对齐、行距、段间距、缩进、首行缩进、项目符号、编号列表
			- 垂直翻转 `isFlipV`
			- 水平翻转 `isFlipH`
			- 旋转角度 `rotate`
			- 垂直对齐方向 `vAlign`
			- 是否自动换行 `wrap`
			- 是否为竖向文本 `isVertical`
			- 元素名 `name`
			- 自动调整大小 `autoFit`
				- 类型 `type`
					- `shape`：文本框高度会根据文本内容自动调整
					- `text`：文本框大小固定，字号会自动缩放以适应文本框（注：autoFit不存在时，也会固定文本框大小，但字号不会缩放）
				- 字体缩放比例（type='text'专有，默认为1） `fontScale`
			- 文本内边距（4边） `textInset`
			- 超链接 `link`

		- 图片
			- 类型 `type='image'`
			- 水平坐标 `left`
			- 垂直坐标 `top`
			- 宽度 `width`
			- 高度 `height`
			- 边框颜色 `borderColor`
			- 边框宽度 `borderWidth`
			- 边框类型（实线、点线、虚线） `borderType`
			- 非实线边框样式 `borderStrokeDasharray`
			- 裁剪形状 `geom`
			- 裁剪范围 `rect`
			- 资源引用路径 `ref`
			- 图片base64 `base64`
			- 图片blob `blob`
			- 旋转角度 `rotate`
			- 滤镜 `filters`
			- 超链接 `link`

		- 形状
			- 类型 `type='shape'`
			- 水平坐标 `left`
			- 垂直坐标 `top`
			- 宽度 `width`
			- 高度 `height`
			- 边框颜色 `borderColor`
			- 边框宽度 `borderWidth`
			- 边框类型（实线、点线、虚线） `borderType`
			- 非实线边框样式 `borderStrokeDasharray`
			- 阴影 `shadow`
			- 填充（颜色、图片、渐变、图案） `fill`
			- 仅描边（无填充） `strokeOnly`
			- 内容文字（HTML富文本，与文字元素一致） `content`
			- 垂直翻转 `isFlipV`
			- 水平翻转 `isFlipH`
			- 旋转角度 `rotate`
			- 形状类型 `shapType`
			- 垂直对齐方向 `vAlign`
			- 是否自动换行 `wrap`
			- 形状路径 `path`
			- 形状路径 viewBox `pathViewBox`
			- 线条起点端点样式 `headEnd`
			- 线条终点端点样式 `tailEnd`
			- 形状调整参数 `keypoints`
			- 元素名 `name`
			- 自动调整大小 `autoFit`
			- 文本内边距（4边） `textInset`
			- 超链接 `link`

		- 表格
			- 类型 `type='table'`
			- 水平坐标 `left`
			- 垂直坐标 `top`
			- 宽度 `width`
			- 高度 `height`
			- 边框（4边） `borders`
			- 单元格样式与数据 `data`
			- 行高 `rowHeights`
			- 列宽 `colWidths`

		- 图表
			- 类型 `type='chart'`
			- 水平坐标 `left`
			- 垂直坐标 `top`
			- 宽度 `width`
			- 高度 `height`
			- 图表数据 `data`
			- 图表主题色 `colors`
			- 图表类型 `chartType`
			- 柱状图方向 `barDir`
			- 是否带数据标记 `marker`
			- 环形图尺寸 `holeSize`
			- 分组模式 `grouping`
			- 图表样式 `style`

		- 视频
			- 类型 `type='video'`
			- 水平坐标 `left`
			- 垂直坐标 `top`
			- 宽度 `width`
			- 高度 `height`
			- 资源引用路径 `ref`
			- 视频blob `blob`

		- 音频
			- 类型 `type='audio'`
			- 水平坐标 `left`
			- 垂直坐标 `top`
			- 宽度 `width`
			- 高度 `height`
			- 资源引用路径 `ref`
			- 音频blob `blob`

		- 公式
			- 类型 `type='math'`
			- 水平坐标 `left`
			- 垂直坐标 `top`
			- 宽度 `width`
			- 高度 `height`
			- 公式图片引用路径 `picRef`
			- 公式图片base64 `picBase64`
			- 公式图片blob `picBlob`
			- LaTeX表达式（仅支持常见结构） `latex`
			- 文本（文本和公式混排时存在） `text`

		- Smart图
			- 类型 `type='diagram'`
			- 水平坐标 `left`
			- 垂直坐标 `top`
			- 宽度 `width`
			- 高度 `height`
			- 子元素集合 `elements`
			- 文本列表（Smart图中的文字内容清单） `textList`

		- 多元素组合
			- 类型 `type='group'`
			- 水平坐标 `left`
			- 垂直坐标 `top`
			- 宽度 `width`
			- 高度 `height`
			- 子元素集合 `elements`

### 更详细类型请参考 👇
[https://github.com/pipipi-pikachu/pptxtojson/blob/master/dist/index.d.ts](https://github.com/pipipi-pikachu/pptxtojson/blob/master/dist/index.d.ts)

# 🙏 感谢
本仓库大量参考了 [PPTX2HTML](https://github.com/g21589/PPTX2HTML) 和 [PPTXjs](https://github.com/meshesha/PPTXjs) 的实现。

与它们不同的是：pptxtojson 的目标不是将 PPT 文件转换为 HTML 页面，而是输出干净、易读、便于二次处理的 JSON 数据，并在原有基础上进行了大量优化补充，提升了提取信息的完整度和准确度。

欢迎通过 Issue / PR 反馈更多 PPTX 样本、解析场景和改进建议。

# 📄 开源协议
MIT License | Copyright © 2020-PRESENT [pipipi-pikachu](https://github.com/pipipi-pikachu)
