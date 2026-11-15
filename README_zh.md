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

浏览器解析、资源释放、Node.js 及 TypeScript 用法，见 [使用手册](USER_GUIDE_zh.md#examples)。

# 🎯 注意事项

目前解析结果与源文件在排版和样式上的综合还原度约为 85%+。对于普通用户手动从头创建和编辑的 PPTX，常见页面结构与基础样式甚至可以达到 95%+ 的还原度。

但是如果文件来自网上下载的复杂模板，或由 PPT 水平较高的用户大量使用“高级技巧”制作（如复杂母版、深层组合、特殊图形效果、复杂渐变、非标准形状、复杂 SmartArt 等），解析难度会明显上升，还原度也会相应降低。此类文件更适合作为复杂样本单独评估。

# 📕 使用文档

完整接口、输出字段、单位、数据示例、渲染注意事项和版本迁移，请阅读 [使用手册](USER_GUIDE_zh.md)。

- [API 与配置](USER_GUIDE_zh.md#api)
- [页面、顺序与坐标约定](USER_GUIDE_zh.md#conventions)
- [各类元素与字段](USER_GUIDE_zh.md#presentation)
- [接入与安全注意事项](USER_GUIDE_zh.md#integration)
- [旧版本迁移](USER_GUIDE_zh.md#migration)

TypeScript 类型参考：[types/index.d.ts](https://github.com/pipipi-pikachu/pptxtojson/blob/master/types/index.d.ts)。

# 🙏 感谢
本仓库大量参考了 [PPTX2HTML](https://github.com/g21589/PPTX2HTML) 和 [PPTXjs](https://github.com/meshesha/PPTXjs) 的实现。

与它们不同的是：pptxtojson 的目标不是将 PPT 文件转换为 HTML 页面，而是输出干净、易读、便于二次处理的 JSON 数据，并在原有基础上进行了大量优化补充，提升了提取信息的完整度和准确度。

欢迎通过 Issue / PR 反馈更多 PPTX 样本、解析场景和改进建议。

# 📄 开源协议
MIT License | Copyright © 2020-PRESENT [pipipi-pikachu](https://github.com/pipipi-pikachu)
