# pptxtojson 使用手册

简体中文 | [English](USER_GUIDE.md)

适用于 **pptxtojson 3.x**，本文按 **3.0.0** 的公开接口和实际输出说明。

pptxtojson 将 `.pptx` 转成便于读取、转换和二次处理的结构化数据。它是**解析器，不是 PowerPoint 渲染器**：负责提取页面、文字、几何、图表和资源；页面绘制、富文本编辑、字体加载、媒体播放、公式排版等由使用方完成。

## 目录

- [安装与运行环境](#installation)
- [常用用法](#examples)
- [API 与配置](#api)
- [公共数据约定](#conventions)
- [顶层结果与页面](#presentation)
- [填充、边框与阴影](#appearance)
- [文字与形状](#text-and-shape)
- [图片](#images)
- [表格](#tables)
- [图表](#charts)
- [视频与音频](#media)
- [公式](#math)
- [组合与 SmartArt](#groups)
- [接入与安全注意事项](#integration)
- [从旧版本迁移](#migration)
- [常见问题](#faq)

<a id="installation"></a>
## 安装与运行环境

```bash
npm install pptxtojson
```

### 浏览器与前端构建工具

在 Vite 等支持 ESM 的构建环境中：

```js
import { parse } from 'pptxtojson'
```

浏览器需要支持 `ArrayBuffer`、Promise 等常见 API。下文使用 `File.arrayBuffer()`、`fetch()` 和 Blob URL；较旧环境需自行提供相应兼容实现。

不使用构建工具时，可以引入 UMD 文件：

```html
<script src="https://cdn.jsdelivr.net/npm/pptxtojson@3.0.0/dist/index.umd.js"></script>
<script>
  // 接口与 ESM 一致，只是从全局对象读取。
  // const result = await window.pptxtojson.parse(arrayBuffer)
</script>
```

生产环境建议固定版本，避免升级时数据结构变化。CDN 加载不等于文件转换：传给 `parse()` 的 PPTX 仍在当前运行环境内解析，库不会主动将其上传到转换服务。

### Node.js

Node.js 使用仍属于实验性支持。可显式选择包中的入口，避免依赖不同工具对 `main` / `module` 的解释：

```js
// CommonJS：保存为 .cjs，或在 CommonJS 项目中使用
const { parse } = require('pptxtojson/dist/index.cjs')
```

```js
// ESM：保存为 .mjs，或在 type: module 的项目中使用
import { parse } from 'pptxtojson/dist/index.js'
```

建议使用受支持的现代 Node.js LTS。若启用媒体 `blob` 模式，运行环境还需要 `Blob` 与 `URL.createObjectURL()`；Node.js 生成的 Blob URL 不能交给另一台机器或浏览器直接使用。服务端导出 JSON 通常使用图片 `base64`，并关闭音视频读取。

### 文件范围

- 输入是 `.pptx` 的二进制内容，不是文件路径、URL 字符串或已经解析的 JSON。
- 不直接支持旧二进制 `.ppt`、PDF、Keynote 文件。
- 加密、损坏或不完整的包可能解析失败；受保护文件应先在有权限的 Office 环境中解密并保存为普通 PPTX。
- 部分高级 Office 内容只能提取一部分数据；解析成功不代表视觉完全一致。

<a id="examples"></a>
## 常用用法

以下示例只演示解析与数据读取，不包含页面 UI 或渲染器。浏览器示例中的 `file` 指已取得的 `File` 对象。

### 浏览器：解析本地文件

```js
import { parse } from 'pptxtojson'

try {
  const result = await parse(await file.arrayBuffer(), {
    imageMode: 'base64',
  })
  console.log(result.slides)
  const json = JSON.stringify(result, null, 2)
  // 将 json 交给业务侧保存或下载。
}
catch (error) {
  console.error('解析失败', error)
}
```

图片默认返回可随 JSON 保存的 Data URL；音视频默认不读取内容。提取文字时，读取 Text/Shape 的 `content`、TableCell 的 `text`、Math 的 `latex` 及页面 `note`；前两者是 HTML，不是纯文本。组合内容需递归遍历 `elements`。

### 浏览器：媒体模式与资源释放

```js
const result = await parse(await file.arrayBuffer(), {
  imageMode: 'blob',
  videoMode: 'blob',
  audioMode: 'blob',
})

// 使用 image.blob、video.blob、audio.blob 或 math.picBlob。
// 它们是 URL 字符串；播放器、图片组件由业务侧提供。
```

所有使用者都不再引用媒体后，再释放 Blob URL。以下函数递归收集并去重，覆盖组合、图片填充及公式后备图：

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

- 在移除预览、替换文稿等适当时机调用 `releaseResources(result)`，不要刚解析完就释放。
- Blob URL 只属于当前运行环境，保存 JSON 不会保存其二进制内容；持久化时应自行存储媒体。
- 浏览器不支持的媒体编码仍可能播放失败，解析成功不等于可播放。
- 当前没有 `dispose()` 或取消 API。若解析中途失败，可能已有资源被创建却未返回结果；不需要媒体时优先使用 `none`，需严格隔离生命周期时使用独立页面或合适的 Worker 环境。

### Node.js：读取文件并保存 JSON

ESM 示例，保存为 `convert.mjs` 后运行 `node convert.mjs`：

```js
import { readFile, writeFile } from 'node:fs/promises'
import { parse } from 'pptxtojson/dist/index.js'

const buffer = await readFile('input.pptx')
// 只截取 Buffer 的有效范围，避免带入底层 ArrayBuffer 中的其他数据。
const data = buffer.buffer.slice(
  buffer.byteOffset, buffer.byteOffset + buffer.byteLength,
)
const result = await parse(data, { imageMode: 'base64' })
await writeFile('presentation.json', JSON.stringify(result, null, 2), 'utf8')
```

实际应用中请捕获文件读取、解析与写入错误。`base64` 已包含 `data:image/...;base64,` 前缀，不要再加一次；大图片会显著增加 JSON 大小。

### 远程 PPTX 与 TypeScript

```ts
import { parse, type Options } from 'pptxtojson'

const response = await fetch('/files/example.pptx')
if (!response.ok) throw new Error(`下载失败：HTTP ${response.status}`)
const options: Options = { imageMode: 'none' }
const result = await parse(await response.arrayBuffer(), options)
// result 的类型由 parse 自动推导，也可直接导入 Element、Chart 等类型。
```

浏览器下载远程 PPTX 受 CORS 和鉴权限制；这些是获取文件阶段的问题，不是 `parse()` 负责的内容。

<a id="api"></a>
## API 与配置

### parse

```ts
parse(file: ArrayBuffer, options?: Options): Promise<{
  slides: Slide[]
  themeColors: string[]
  usedFonts: string[]
  size: { width: number; height: number }
}>
```

- `file`：完整 PPTX 二进制内容。
- `options`：可选配置，不传时使用下表默认值。
- 成功：返回整个演示文稿的数据。
- 失败：Promise reject，通过 `try/catch` 处理。没有承诺稳定的错误码，也不保证失败时返回部分页面。
- 当前没有逐页回调、进度回调、`AbortSignal` 或分页解析参数。

### Options

| 字段 | 可选值 | 默认值 | 作用 |
|---|---|---|---|
| `imageMode` | `'base64' \| 'blob' \| 'both' \| 'none'` | `'base64'` | 图片元素、图片填充及公式后备图像的资源读取方式 |
| `videoMode` | `'blob' \| 'none'` | `'none'` | 支持的内嵌视频资源是否生成 Blob URL |
| `audioMode` | `'blob' \| 'none'` | `'none'` | 支持的内嵌音频资源是否生成 Blob URL |
| `singleLineSpacingFactor` | 有限正数 | `1` | 百分比行距倍率，也参与按行段前/段后间距的 pt 近似换算；固定磅值不受影响 |

`imageMode` 对资源字段的影响：

| 值 | `base64` / `picBase64` | `blob` / `picBlob` |
|---|---|---|
| `base64` | Data URL | `''` |
| `blob` | `''` | Blob URL |
| `both` | Data URL | Blob URL |
| `none` | `''` | `''` |

关系信息或后备图缺省时，相关字符串也可能为空。若文件明确引用了不存在的 ZIP 资源，启用内容读取仍可能使解析失败，而不是统一退回空字符串。`none` 不代表删除元素：仍会尽可能返回位置、样式和 `ref`。

`singleLineSpacingFactor` 例如设为 `1.15`，用于近似不同字体在 PowerPoint 中的单倍行高。它不是字号缩放，也不保证跨字体像素一致。无效、非有限或非正值按 `1` 处理；不要依赖未列出的其他配置值。

<a id="conventions"></a>
## 公共数据约定

### 单位与坐标

| 数据 | 单位/语义 |
|---|---|
| 页面、元素位置尺寸，边框宽度，阴影偏移/模糊，文本内边距，表格行高列宽 | `pt`；1 英寸 = 72pt |
| HTML 文字字号、字间距、绝对行距和段间距 | 带 `pt` 等单位的 CSS 字符串 |
| `rotate`、渐变 `rot` | 角度（度），不是弧度或 OOXML 原始整数 |
| `autoFit.fontScale`、图片裁剪 `rect` | 百分数；80 表示 80%，不是 0.8 |
| 图片填充 `opacity` | 比例；1 表示完全不透明 |
| 图表 `x/y/size` | 原始数据值或索引，不是页面长度 |
| `transition.duration`、`autoNextAfter` | 毫秒 |
| `keypoints` | 形状调整参数，非长度；详见形状字段 |

左上角为原点，`left` 向右、`top` 向下。负坐标和超出页面边界的元素是可能的，不要无条件裁成零。

在 96dpi CSS 像素画布中，`px = pt × 96 / 72`。也可在一个使用 pt 的内部画布上整体缩放，避免只缩放包围盒却忘记字号、边框等。不要把百分比、角度或图表数据当作 pt 换算。

### 顺序与标识

- `slides` 按演示文稿页面清单的顺序输出，不按 ZIP 中的文件名排序。
- `id` 是来源 OOXML 的局部标识，不保证跨页、母版、组合或 SmartArt 全局唯一；取不到时可能是空字符串，部分 SmartArt 子元素也可能没有该值。业务编辑器应生成自己的全局 ID。
- `order` 是绘制顺序键，不是连续数组下标或持久化 ID。数值较大的后绘制；不要比较不同文稿或不同解析调用的绝对值。
- 顶层 `layoutElements` 通常使用负 `order`，表示母版/布局图形在当前页图形之下；它不是原始母版树的完整导出。

合并当前页与继承背景的绘制队列：

```js
function getSlidePaintOrder(slide) {
  return [...slide.layoutElements, ...slide.elements]
    .sort((a, b) => a.order - b.order)
}
```

组合内部应分别排序，再应用该组合的变换；不要把所有嵌套元素按一个全局 `order` 混排。

### 缺省值、空值与颜色

- 可选字段可能省略，也可能是 `undefined`；JSON 序列化会移除对象中的 `undefined`。
- 未启用资源模式或缺少资源引用时，字符串通常为 `''`；无形状填充时 `fill` 可能是 `null`。不要直接写 `element.fill.type`。
- `borderWidth: 0` 表示无可见描边；存在 `borderColor` 并不代表边框可见。
- 颜色通常是 `#RRGGBB`，也可能有短十六进制、透明度等形式；不要假定固定长度或大小写。
- 图表颜色数组中，未解析出的项可能是 `undefined`，序列化后成为 `null`。渲染时为每个缺失项单独选择主题色或默认色。
- 类型声明用于开发提示；导入任意来源文稿时，仍应对空值、缺失值和无效数值做防御处理。

### 元素公共字段

| 字段 | 类型 | 说明 |
|---|---|---|
| `type` | 字符串联合类型 | 元素种类，决定其他字段结构 |
| `id` | `string` | 来源局部 ID；限制见上文 |
| `left`、`top` | `number` | 包围盒左上位置，pt；组合子元素相对其直接容器 |
| `width`、`height` | `number` | 包围盒尺寸，pt；允许零尺寸 |
| `order` | `number` | 当前容器内的绘制顺序键 |

`name` 主要由 Text/Shape 提供，不是所有元素都有。`rotate`、翻转和边框也不是所有类型共有；缺失旋转可按 0 处理，缺失翻转可按 false 处理，不能据此推断原文件所有效果都已被表达。

<a id="presentation"></a>
## 顶层结果与页面

### 顶层结果

| 字段 | 类型 | 说明 |
|---|---|---|
| `slides` | `Slide[]` | 页面数组，可为空 |
| `size.width`、`size.height` | `number` | 演示文稿页面尺寸，pt，不是浏览器预览尺寸 |
| `themeColors` | `string[]` | 演示文稿主题的 accent1～accent6 颜色；可能不足六项或为空 |
| `usedFonts` | `string[]` | 文稿内嵌字体清单中的去重字体名称；不是全部实际使用字体，也不包含字体二进制 |

结构示意：

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

字体名称出现在 HTML 中，不表示当前浏览器安装了该字体；即使 `usedFonts=[]`，文稿仍可能使用很多字体。

### Slide

| 字段 | 类型 | 说明 |
|---|---|---|
| `fill` | `Fill` | 页面最终背景，见填充结构 |
| `elements` | `Element[]` | 当前幻灯片元素 |
| `layoutElements` | `Element[]` | 按显示设置解析的母版、布局非占位背景元素；可能为空 |
| `note` | `string` | 演讲者备注文字；无备注通常为 `''`，不是 HTML |
| `transition` | 对象、`null` 或省略 | 页面切换信息；无切换时通常为 `null` |

不要把 `layoutElements` 再当成独立一页，也不要再从母版手工复制同样背景，否则会重复绘制。

### transition

| 字段 | 类型 | 说明 |
|---|---|---|
| `type` | `string` | OOXML 切换类型名，如 `fade`、`push`；`none` 表示未识别到具体效果 |
| `duration` | `number` | 时长，毫秒；未提供明确时长时可能由速度或默认值换算 |
| `direction` | `string \| null` | 来源方向值，如 `l/r/u/d`；不是已经标准化的播放器枚举 |
| `autoNextAfter` | `number`，可选 | 自动换页等待时间，毫秒；0 是有效值，与是否允许点击换页独立 |

只返回数据，不会启动动画或定时器。没有逐元素动画时间轴，也没有完整暴露“单击换页是否启用”的独立字段。

<a id="appearance"></a>
## 填充、边框与阴影

### Fill

`fill` 用 `type` 区分以下四种结构。Text/Shape 没有填充时可能直接返回 `null`。

| `type` | `value` | 用法与限制 |
|---|---|---|
| `color` | 颜色字符串 | 纯色背景或填充 |
| `image` | `{ ref, base64, blob, opacity }` | 图片资源加透明度；资源字段受 `imageMode` 控制 |
| `gradient` | `{ path, rot, colors }` | 渐变，不是直接可用的完整 CSS 字符串 |
| `pattern` | `{ type, foregroundColor, backgroundColor }` | Office 预设图案标识及前景/背景色，需自行实现图案渲染 |

图片填充：

| 字段 | 类型 | 说明 |
|---|---|---|
| `ref` | `string` | 包内资源引用，如 `ppt/media/image1.png`；不是 HTTP URL |
| `base64` | `string` | 完整图片 Data URL 或 `''` |
| `blob` | `string` | Blob URL 或 `''`，不是 `Blob` 对象 |
| `opacity` | `number` | 默认通常为 1，取值是比例，不是百分数 |

渐变：

| 字段 | 类型 | 说明 |
|---|---|---|
| `path` | `'line' \| 'circle' \| 'rect' \| 'shape'` | 线性或路径渐变类型 |
| `rot` | `number` | 来源渐变角度，度；CSS 渐变角度约定不同，不能不经转换直接等同 |
| `colors` | `{ pos: string; color: string }[]` | 色标数组；`pos` 通常如 `'0%'`、`'100%'`，来源缺省时可能是空字符串 |

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

主题颜色和常见主题纯色/渐变引用已尽量解析成具体值；不保证所有主题图片、图案、效果及路径渐变都能完整还原。

### 描边与线端点

Text、Shape、Image 直接提供：

| 字段 | 类型 | 说明 |
|---|---|---|
| `borderColor` | `string` | 描边颜色 |
| `borderWidth` | `number` | pt；0 表示不绘制 |
| `borderType` | `'solid' \| 'dashed' \| 'dotted'` | 概括性的线型 |
| `borderStrokeDasharray` | `string` | 更细的虚线序列，如 `'5, 5, 1, 5'`，实线通常为 `'0'`；不是数组 |

表格的 `borders` 是 `{ top?, right?, bottom?, left? }`，每条边是 `Border` 对象，至少包含 `borderColor/borderWidth/borderType`；实际还可能有 **`strokeDasharray`**。注意它与元素字段 **`borderStrokeDasharray`** 的名字不同。

形状可提供 `headEnd`、`tailEnd`：

| 子字段 | 类型 | 说明 |
|---|---|---|
| `type` | `none/triangle/stealth/diamond/oval/arrow` | 线端样式 |
| `width` | `sm/med/lg`，可选 | 相对宽度等级，不是 pt |
| `length` | `sm/med/lg`，可选 | 相对长度等级，不是 pt |

不是所有 shape 都是直线。应结合 `shapType`、`path` 和翻转/旋转决定方向，不能只连接包围盒的左上、右下点。

### Shadow

`shadow` 存在时结构为：

| 字段 | 类型 | 说明 |
|---|---|---|
| `h`、`v` | `number` | 水平、垂直偏移，pt，可能为负 |
| `blur` | `number`，部分缺省情况为空字符串 | 模糊半径，pt；消费时可用 `Number(shadow.blur) || 0` 兜底 |
| `color` | `string` | 阴影颜色，可能包含透明度 |

元素阴影和 HTML 内文字阴影是不同层面的效果，不要把同一个阴影重复应用。

<a id="text-and-shape"></a>
## 文字与形状

### Text 与 Shape 的共用样式

二者均包含公共位置字段，及下表中适用字段：

| 字段 | 类型 | 说明 |
|---|---|---|
| `name` | `string` | 来源元素名称，不是唯一 ID |
| `content` | `string` | HTML 富文本，可能为空 |
| `rotate` | `number` | 元素角度；消费时缺失按 0 处理 |
| `isFlipH`、`isFlipV` | `boolean` | 水平、垂直翻转 |
| `fill` | `Fill \| null` | 元素背景或形状填充 |
| 描边字段 | 见上节 | Text 也可能有背景和边框 |
| `shadow` | `Shadow`，可选 | 元素阴影 |
| `vAlign` | `'up' \| 'mid' \| 'down'` | 文字垂直对齐：上、中、下 |
| `wrap` | `boolean` | 是否允许自动换行；显式 `<br>` 与自动换行是两回事 |
| `textDirection` | `TextDirection`，可选 | 内部文字方向；水平通常省略，与元素旋转独立 |
| `autoFit` | `AutoFit`，可选 | 自动适应策略，见下文 |
| `textInset` | `{ l, t, r, b }`，可选 | 四边文字内边距，pt；缺省体属性常对应左右 7.2pt、上下 3.6pt |
| `link` | `string`，可选 | 元素级 HTTP(S) 超链接；与 HTML 内文字超链接不同 |

`type='text'` 表示文字元素；`type='shape'` 表示形状。**形状也可以包含文字**，提取文字时不能只扫描 `type='text'`。

### content 的结构与使用

常见结构是段落、嵌套 span、链接、软换行和列表：

```html
<p style="white-space: pre-wrap;">
  <span style="font-size: 24pt;font-weight: bold;">
    <span style="font-family: Arial;">ABC</span>
    <span style="font-family: SimSun;">中文</span>
  </span>
</p>
```

- 字号、字体、颜色、加粗、斜体、下划线、删除线、字间距、文字阴影及上下标等常见样式通过 HTML/CSS 表达，不是独立的 `fontSize/bold` JSON 字段。
- 段落可能包含对齐、行高、段间距、缩进及首行缩进等样式。按行间距是字体度量近似，不保证与 Office 像素一致。
- 一个 run 可按 Latin/东亚脚本拆成多个内部 span；纯数字/标点 run 的语言也可能影响选用字体。不要只读取第一个 span 的样式作为整段样式。
- 保留继承关系和 HTML 实体。不要用正则随意合并 CSS 分号；例如 `&quot;;` 的第一个分号属于实体，第二个才属于 CSS。
- `white-space: pre-wrap`、`pre`、`tab-size`、`<br>` 关系到连续空格、制表符和换行；不要统一替换为空格或 `&nbsp;`。
- 列表可能使用嵌套 `ul/ol/li`、`start`、`li[value]` 及 `data-pptx-level`。最后一个是原始级别信息，实际 HTML 嵌套还可能反映显式缩进覆盖。
- 依赖 Wingdings/Symbol 等特殊字体或不可可靠表示的符号会回退为默认圆点；部分特殊编号格式也只能近似，不承诺完整 Office 列表排版。
- 不承诺 span 数量、样式排列顺序、HTML 空白写法稳定。业务逻辑应依赖语义和 DOM，而不是精确字符串匹配。

字体名称不等于字体文件。加载系统字体或有授权的 Web Font 是使用方的责任，`usedFonts` 不会自动解决缺字、回退或字宽变化。

### 方向与自动适应

`TextDirection` 可为：

| 值 | 语义 |
|---|---|
| `horz` 或省略 | 水平文字 |
| `vert`、`vert270` | 两种竖向旋转文字方向 |
| `eaVert` | 东亚竖排 |
| `wordArtVert`、`wordArtVertRtl` | 艺术字相关竖向排列 |
| `mongolianVert` | 蒙古文竖排 |

`isVertical` 仅 Text 提供，保留对 `eaVert` 的兼容判断；`false` 不代表没有 `vert/vert270` 等方向。方向应应用到文字容器，不能简单加到整个形状的 `rotate` 上。

`autoFit`：

| 结构 | 说明 |
|---|---|
| `{ type: 'shape' }` | 文本容器可随文字内容适应尺寸 |
| `{ type: 'text', fontScale?: number }` | 容器尺寸固定，文字使用给定缩放比例 |
| 省略 | 没有返回自动适应策略；不要自行认定必定要改变几何尺寸 |

**`fontScale=80` 表示原字号的 80%**。使用时除以 100；缺省可按 100% 处理。`content` 中字号不会替你应用这个缩放，不应重复乘两次。

```js
function getFontScale(element) {
  return element.autoFit?.type === 'text'
    ? (element.autoFit.fontScale ?? 100) / 100
    : 1
}
```

### Shape 专有字段

| 字段 | 类型 | 说明 |
|---|---|---|
| `shapType` | `string` | Office 预设形状名，如 `rect/ellipse/triangle/line`，或 `custom` |
| `path` | `string`，可选 | SVG path 的 `d`，可能由多个子路径组成；空路径可为 `''` |
| `pathViewBox` | `{ x, y, width, height }`，可选 | 路径坐标系；不是页面位置，不应再次将路径数字当作 EMU |
| `strokeOnly` | `boolean`，可选 | 仅描边，不应自动补充填充色 |
| `headEnd`、`tailEnd` | `LineEnd`，可选 | 线端点样式，见描边节 |
| `keypoints` | `Record<string, number>`，可选 | 调整参数名称到数值的映射；当前按原始 `val / 50000` 输出，不是通用的 0～1 比例，不可一律 clamp |

绘制路径时，优先使用 `pathViewBox` 将路径映射到元素宽高；不应再按路径最大坐标重新推断并拉伸一次。缺少可用路径或尺寸时，应有安全降级策略。预设形状与复杂自定义几何的解析覆盖不等同于完整 Office 几何求值器。

<a id="images"></a>
## 图片

`type='image'`，包含公共位置字段和：

| 字段 | 类型 | 说明 |
|---|---|---|
| `ref` | `string` | 来源资源引用，通常为包内路径；可能仍含 XML 实体，不宜作业务唯一键 |
| `base64` | `string` | Data URL 或 `''` |
| `blob` | `string` | Blob URL 或 `''` |
| `rotate` | `number` | 元素旋转角度 |
| `isFlipH`、`isFlipV` | `boolean` | 水平、垂直翻转 |
| 描边字段 | 见[描边说明](#appearance) | 图片边框 |
| `geom` | `string` | 裁剪形状，可能是预设名或 `custom:` 标识；不是完整裁剪算法 |
| `rect` | `{ l?, t?, r?, b? }`，可选 | 左、上、右、下裁剪量，百分数，例如 10 表示该边裁去 10%；不是两个角点坐标 |
| `filters` | 对象，可选 | 下表中的图片效果参数 |
| `link` | `string`，可选 | 元素级 HTTP(S) 链接 |

`rect` 值不保证被限制在 0～100；来源文稿可能通过负裁剪表达扩展或不同构图。图片原始字节不会被重新裁剪或烘焙滤镜，渲染器应自行应用这些信息。

`filters`：

| 字段 | 类型 | 当前数值语义 |
|---|---|---|
| `sharpen` | `number` | 锐化强度，来源值除以 100000 |
| `soften` | `number`，运行时可能提供 | 柔化强度，来源非正锐化/柔化值取绝对值后除以 100000；当前声明中未列出 |
| `colorTemperature` | `number` | 色温原始整数，按 K 理解，不做比例换算 |
| `saturation` | `number` | 饱和度比例，1 对应 100% |
| `brightness` | `number` | 来源亮度有符号偏移除以 100000，0 为中性，不是 CSS brightness 的倍率 |
| `contrast` | `number` | 来源对比度有符号偏移除以 100000，0 为中性，不是 CSS contrast 的倍率 |

库不负责实际应用、钳制或转换为 CSS 滤镜。独立图片元素没有统一的 `opacity` 输出字段；图片填充透明度位于 `fill.value.opacity`。

<a id="tables"></a>
## 表格

### Table

`type='table'`，包含公共位置字段和：

| 字段 | 类型 | 说明 |
|---|---|---|
| `data` | `TableCell[][]` | 按源行列顺序保留的单元格数组；合并续接格仍可能存在 |
| `rowHeights` | `number[]` | 行高，pt，不是比例 |
| `colWidths` | `number[]` | 列宽，pt，不是比例；需要比例时除以总和 |
| `borders` | 四边 `Border` 对象 | 整表边框信息，不替代单元格各边框 |

### TableCell

| 字段 | 类型 | 说明 |
|---|---|---|
| `text` | `string` | 单元格 HTML 富文本，不是纯文本 |
| `rowSpan`、`colSpan` | `number`，可选 | 主格跨行/列数量，缺省按 1 |
| `hMerge`、`vMerge` | `number`，可选 | 横向/纵向合并续接标记，通常为 1；不是主格跨度 |
| `fillColor` | `string`，可选 | 单元格底色；显式无填充最终省略该字段 |
| `fontColor` | `string`，可选 | 来自单元格/表格样式的文字颜色补充信息 |
| `fontBold` | `boolean`，可选 | 来自表格样式的加粗补充信息；可能未写进每段 HTML |
| `vAlign` | `'up' \| 'mid' \| 'down'` | 单元格垂直对齐 |
| `borders` | 四边 `Border` 对象 | 单元格边框；可为 `{}`，各边可缺省 |

消费时需要同时处理富文本及表格样式字段，不能只用 `querySelector('span')` 丢弃其他文字片段。显式无填充不应被应用自定的默认底色覆盖。

### 读取合并格

遍历时跳过续接格，主格的跨度缺省按 1 处理：

```js
const cells = table.data.map(row => row
  .filter(cell => !cell.hMerge && !cell.vMerge)
  .map(cell => ({
    html: cell.text,
    rowSpan: cell.rowSpan || 1,
    colSpan: cell.colSpan || 1,
  })))
```

不要根据 `colSpan` 再次跳过数组中的列，否则会与续接标记重复跳格。若渲染为 HTML，应先净化 `html`，再由表格组件处理跨度和样式。对内置样式定义不完整的文件，文字和显式属性可保留，但主题样式还原可能有限。

<a id="charts"></a>
## 图表

### 公共字段与类型

`type='chart'`，包含公共位置字段和：

| 字段 | 类型 | 说明 |
|---|---|---|
| `chartType` | `ChartType` | 决定 `data` / `groups` 结构 |
| `colors` | `string[]` | 图表颜色；缺失项应自行回退，JSON 中也可能出现 null |
| `data` | 依图表类型变化 | 普通图、散点图、气泡图结构见下文；组合图没有该顶层字段 |
| `groups` | `ChartGroup[]` | 仅组合图提供 |

支持识别的类型名：

| 类型名 | 含义 |
|---|---|
| `barChart`、`bar3DChart` | 柱/条形图；方向看 `barDir` |
| `lineChart`、`line3DChart` | 折线图 |
| `areaChart`、`area3DChart` | 面积图 |
| `pieChart`、`pie3DChart` | 饼图 |
| `doughnutChart` | 环形图 |
| `scatterChart` | 散点图 |
| `bubbleChart` | 气泡图 |
| `radarChart` | 雷达图 |
| `surfaceChart`、`surface3DChart` | 曲面图 |
| `stockChart` | 股价图 |
| `comboChart` | 多个受支持图表组构成的组合图 |

识别类型和提取系列不等于包含该图的完整渲染参数；例如 3D 视角、图例布局、坐标轴详细定义及所有数据标签不在统一输出模型内。

### 普通图表：ChartItem[]

除散点、气泡和组合图外，`data` 为：

```js
[
  {
    key: '销量',
    values: [
      { x: '0', y: 10 },
      { x: '1', y: 20 },
    ],
    xlabels: { '0': '一月', '1': '二月' },
  },
]
```

| 字段 | 类型 | 说明 |
|---|---|---|
| `key` | `string` | 系列名称；来源缺省时实际可能回退为数字索引，展示前可 `String(key)` |
| `values` | `{ x: string; y: number }[]` | 数据点数组，可能为空 |
| `values[].x` | `string` | 类别/点缓存索引，不是类别文字，更不是散点数值坐标 |
| `values[].y` | `number` | 数值 |
| `xlabels` | `Record<string, string>` | 索引到类别标签的映射，可为空 |

正确按索引关联标签，而不是盲目将 `Object.values(xlabels)` 与 `values` 按数组位置拼接：

```js
const points = item.values.map(point => ({
  label: item.xlabels[point.x] ?? point.x,
  value: point.y,
}))
```

当前读取图表 XML 中的引用缓存或内联数据，**不执行公式、不重新计算内嵌 Excel 工作簿**。工作簿已改而缓存未刷新时，输出仍可能是旧缓存值。

### 散点图：ScatterSeries[]

```js
[
  { key: '系列 A', values: [{ x: 1, y: 11 }, { x: 2, y: 22 }] },
  { key: '系列 B', values: [{ x: 100, y: 44 }, { x: 200, y: 55 }] },
]
```

- `key` 是系列名；`values[].x/y` 均为数值。
- 各系列独立保存 X，不能复用第一个系列的 X。
- 内部按来源点索引配对并升序输出；缺失配对或非有限值不补零。公开点结构不再额外包含来源 `idx`。
- 不保证按 X 值排序，不能为了画线擅自重排。

给接收 `[x,y]` 点数组的绘图库转换：

```js
const series = chart.data.map(item => ({
  name: item.key,
  points: item.values.map(point => [point.x, point.y]),
}))
```

### 气泡图：BubbleSeries[]

```js
[
  {
    key: '样本 A',
    values: [
      { x: 0.7, y: 2.7, size: 10 },
      { x: 1.8, y: 3.2, size: 4 },
    ],
  },
]
```

`size` 是原始大小数据，不是 px 半径、直径或已归一化的面积。渲染器应结合自己的缩放规则和 `sizeRepresents` 使用；零值、负值可能被原样保留。X/Y/size 按同一来源点索引配对。

可选字段：

| 字段 | 类型 | 说明 |
|---|---|---|
| `bubbleScale` | `number` | 气泡大小缩放百分数，例如 100 表示 100% |
| `showNegBubbles` | `boolean` | 来源是否显示负值气泡的设置 |
| `sizeRepresents` | `'area' \| 'w'` | 大小按面积或宽度解释 |

这些字段只在来源有相应信息时输出，不应将“未输出”与“显式 false/0”混为一谈。

### 组合图：groups

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
      data: [{ key: '销量', values: [{ x: '0', y: 10 }], xlabels: { '0': '一月' } }],
      barDir: 'col', grouping: 'clustered',
      axisIds: ['100', '200'], seriesIndices: [0],
    },
    {
      chartType: 'lineChart', colors: ['#ED7D31'],
      data: [{ key: '占比', values: [{ x: '0', y: 0.5 }], xlabels: { '0': '一月' } }],
      axisIds: ['100', '300'], seriesIndices: [1],
    },
  ],
}
```

- 整体仍是一个图表元素，位置尺寸在顶层；子组不是独立页面元素。
- **没有顶层 `data`**。每组有自己的 `chartType/data/colors` 和适用选项。
- `axisIds: string[]`：该组引用的轴标识，可为空；不能当作现成的完整坐标轴配置。
- `seriesIndices: number[]`：来源系列索引，不是各组内数据点索引。
- 分组按来源顺序保留，顶层颜色为各组颜色的汇总。
- 不同轴 ID 不应合并为同一轴；但仅凭现有字段无法保证精确恢复全部次轴刻度和位置。

读取组合图时，逐组处理各自的数据结构：

```js
const groups = chart.chartType === 'comboChart' ? chart.groups : [chart]
for (const group of groups) {
  console.log(group.chartType, group.data)
}
```

不要将不同子图的 `data` 直接拼接成同一种点结构；按上文对应类型读取即可。

### 其他图表选项

| 字段 | 常见适用类型 | 说明 |
|---|---|---|
| `barDir` | 柱/条形 | `col` 为竖柱、`bar` 为横条 |
| `grouping` | 柱、线、面积等 | 如 `clustered/standard/stacked/percentStacked`；不是已经堆叠或百分比换算后的数值 |
| `marker` | 常见折线图 | 基础标记提示，不包含逐点完整标记设置 |
| `holeSize` | 环形图 | 百分数的字符串，例如 `'50'`；不是 `'50%'` 或 px |
| `style` | 散点、雷达等 | 来源样式名，如散点的 `lineMarker`、雷达的 `filled`；不是 CSS |

图表标题、单位、日期轴、3D、金融/曲面特殊配置以及图表内部格式可能未完整暴露。接入时应明确哪些字段能映射，哪些需要降级，不能把所有类型都当柱状图或把所有 `data` 都当二维数组。

<a id="media"></a>
## 视频与音频

`type='video'` 或 `type='audio'`，包含公共位置字段和：

| 字段 | 类型 | 说明 |
|---|---|---|
| `ref` | `string` | 包内媒体路径，或视频外部链接引用 |
| `blob` | `string` | Blob URL 或 `''`；由对应媒体模式控制 |
| `rotate` | `number`，运行时可能提供 | 旋转角度；缺省按 0，当前类型声明中未列出 |

- 视频内容加载主要覆盖 mp4/webm/ogg，音频覆盖 mp3/wav/ogg；其他格式可能只返回 `ref` 和空 `blob`。
- 外链视频可能只有 `ref`，即使 `videoMode='blob'` 也不代表会下载该外链。它可能是网页地址，不是 `<video src>` 可播放的直链；不要未经检查自动访问或嵌入。
- 不提供统一 base64 音视频输出，也不返回 `Blob` 实例。
- 外链引用可能保留实体/转义形式；外链音频不具备与视频相同的专门处理路径，不承诺可直接播放。
- 某些音视频依赖文件中的预览图片关系才能被识别；缺少关系的元素可能被跳过，缺失的资源内容也可能使读取失败。
- 播放器、编码兼容、自动播放权限由浏览器/业务侧处理；没有完整的 Office 播放控制、剪辑或时间轴输出。

<a id="math"></a>
## 公式

`type='math'`，包含公共位置字段和：

| 字段 | 类型 | 说明 |
|---|---|---|
| `latex` | `string` | 常见 OMML 结构转换结果；不是数组，不包含自动添加的 `$...$` 分隔符 |
| `picRef` | `string` | 文件中后备公式图片的引用，没有时为 `''` |
| `picBase64` | `string` | 后备图像 Data URL，受 `imageMode` 控制 |
| `picBlob` | `string` | 后备图像 Blob URL，受 `imageMode` 控制 |
| `text` | `string`，可选或为空 | 尚未合入 LaTeX 的混排文字 HTML |

支持转换的常见结构包括分数、上下标、根式、矩阵、括号、部分积分/求和、函数、极限及重音等；复杂 OMML 与具体渲染器的命令兼容仍需按实际文稿验证。

同一文本框中的多个公式：

- 单个公式保持一个字符串，不为凑数组额外包装。
- 同行多个表达式可用 `\quad` 连接。
- 多行可使用 `gathered`，明确左/右对齐时可使用单列 `array`。
- 多公式的混排文字可合入 `\text{...}`；已合入的文字不会再通过 `text` 重复提供。
- 图片与元素边界仍作为一个整体返回，不按公式数量拆元素。

```js
const expressions = [
  'x_{i}^{2}',
  '\\begin{gathered}x=1 \\\\ y=2\\end{gathered}',
]
```

若对外观要求高，可优先显示原文件后备图像；但并非每个公式都有后备图。如果使用 KaTeX/MathJax，应自行配置、安全渲染，并检查命令支持情况。不要认为只要拿到非空 LaTeX 就必定能完全还原。

<a id="groups"></a>
## 组合与 SmartArt

### Group

| 字段 | 类型 | 说明 |
|---|---|---|
| `type` | `'group'` | 多元素组合 |
| 公共位置字段 | 见[公共数据约定](#conventions) | 父组合位置尺寸，pt |
| `rotate` | `number` | 父组合旋转 |
| `isFlipH`、`isFlipV` | `boolean` | 父组合翻转 |
| `elements` | 元素数组 | 可包含嵌套组合；实际运行时比当前 `BaseElement[]` 声明更宽 |

坐标处理规则：

- 子元素 `left/top` 相对其直接父组合，不是幻灯片绝对位置。
- 来源组合坐标系的缩放和偏移已尽量折算进子元素的 `left/top/width/height`。**不要再按原始 chExt 或子元素包围盒重复缩放一次。**
- 父组的旋转和翻转没有统一烘焙到每个子元素中。保留组容器并围绕组中心应用变换通常比直接扁平化更合适。
- 只在没有旋转/翻转时，扁平化位置才能简单等于“父 left/top + 子 left/top”；嵌套时递归累加。
- 复杂非均匀缩放、深层组合和组内表格/SmartArt 的内部尺寸，不保证全部属性都完成一致的矩阵转换。表格行列尺寸、路径 viewBox、描边等不可仅凭包围盒推断已完整缩放。

### Diagram / SmartArt

| 字段 | 类型 | 说明 |
|---|---|---|
| `type` | `'diagram'` | SmartArt 数据 |
| 公共位置字段 | 见[公共数据约定](#conventions) | 图形整体位置尺寸 |
| `elements` | `(Shape \| Text)[]` | 文件包含可读的缓存 drawing 时，尝试返回内部元素 |
| `textList` | `string[]` | 从图表数据模型中提取的文本清单 |

`elements` 可能为空而 `textList` 非空；两者都可能为空。库不是完整 SmartArt 自动布局引擎：当前主要读取缓存 drawing 中的直接形状，不保证递归还原其中所有组、图片和连接线。子坐标来自保存的 drawing，没有加外框位置，也没有统一拟合到外框宽高；下游应负责定位及必要的坐标系适配，不能直接套用 Group 已折算缩放的规则。`textList` 可能与 `elements[].content` 重复，提取文字时选择一种主来源，避免重复计数。

<a id="integration"></a>
## 接入与安全注意事项

### 接入富文本编辑器

- 优先通过 DOM / 编辑器 parser 读取富文本，不用字符串替换“修复”HTML 实体或嵌套 span。
- 检查编辑器是否支持字体名、字号、加粗、颜色、链接、列表嵌套、编号起始值和软换行；不支持的样式要有明确降级。
- 编辑器可能规范化 HTML、空白或默认段落对齐；不应要求序列化后字符串逐字相同。
- 字体名含空格/中文时可能带引号及 HTML 转义，这是合法表达，不是乱码。
- 使用者若仅复制 `content` 而忽略 `autoFit`、内边距、垂直对齐、文字方向，仍可能出现字体过大、溢出或排版差异。

### HTML、链接与用户文件

PPTX 可能来自不可信来源。**输出不是经过安全审计的 HTML，也不是净化器的替代品。**

- 展示 JSON/文字时使用 `textContent`；需要渲染富文本时，先用维护良好的 HTML 净化库与适当的 CSS/链接白名单清理，再交给 `innerHTML` 或编辑器。
- 限制链接协议和跳转方式，开启新窗口时使用恰当的 `rel`；不要仅因数据里有 `link` 就自动访问。
- 资源 `ref` 不是可靠的本地文件路径，不要直接拿它访问服务器文件系统。
- 文件大小、ZIP 解压后体积、解析时长和内存开销应有业务侧限制；不要在高并发服务中无约束处理任意上传文件。
- 解析可能占用主线程；较大文件可考虑 Worker，但应先确认打包环境、Blob URL 与资源传递方式支持目标流程。

### JSON 持久化

- `base64` Data URL 可以随 JSON 保存，但体积较大。
- `blob` URL 只属于当前运行环境，不能作为持久化存储方案。
- `ref` 只是来源包中的引用；仅保存 JSON 而丢掉 PPTX/资源，不能再由它自动恢复媒体。
- 若要长期保存图片/音视频，应自行提取并存储资源，之后将 JSON 中引用映射到自己管理的地址。

### 不承诺的完整能力

不应将解析器视作 Office 的无损替代。常见需要另行评估的内容包括：完整动画时间轴、复杂 3D/特效、所有预设几何公式、全部图表轴与标注、特殊字体和复杂脚本塑形、图片项目符号、完整 SmartArt 布局、嵌入对象和外部资源。实际支持应以目标文件输出及渲染验证为准。

<a id="migration"></a>
## 从旧版本迁移

### 2.2.0 及以下升级到 3.x

重点检查：

- **散点、气泡 `data` 不再是二维数组。** 使用系列点结构；气泡点包含原始 `size`，每系列独立 X。
- **组合图使用 `chartType='comboChart'` 和 `groups`。** 不读取顶层 `data`，不假定最后一个图表组就是全部内容。
- **普通图表的 `ChartItem[]` 结构保持不变。** 不需要将所有图表统一迁移成散点数据。
- **文字 HTML 可能更复杂。** 允许嵌套字体 span、真实列表与软换行；不要依赖旧 span 数量或只读取首 span。
- **使用 `textDirection` 判断内部文字方向。** 旧 `isVertical` 只保留部分兼容语义。
- **按 `order` 正确处理继承图形。** 不要求其从 0 开始，也不把局部 ID 当全局 ID。

旧二维数组若原本丢失了各系列独立 X 或气泡大小，不能仅靠旧 JSON 无损推回新结构；需要重新解析原始 PPTX。

### 更早版本

- 0.x 长度曾使用 px；当前长度单位为 pt。
- 1.x 及以下图片元素的 `src`、旧图片填充 `picBase64` 等资源字段，应迁移到当前 `ref/base64/blob` 结构。
- 公式资源使用 `picRef/picBase64/picBlob`；音视频使用 `ref/blob`，不要继续依赖旧 `src`。

<a id="faq"></a>
## 常见问题

### 为什么解析成功但字体仍不同？

输出保存的是字体选择，不负责安装字体。缺字体、字体版本不同、浏览器回退、编辑器丢弃 marks、错误的实体替换，都可能改变字形和字宽。先检查原始 `content`，再检查导入后的 DOM 和计算样式，区分解析、转换和系统字体问题。

### 为什么 imageMode 为 none 仍有很多数据？

它只关闭图片资源内容读取，不关闭文字、位置、样式和引用提取；也不会关闭图表或 SmartArt 数据解析。

### 为什么图表数据与 Excel 不同？

当前使用图表缓存/内联数据，不启动 Excel 计算。请在有权限的 Office 环境中刷新图表并保存，然后重新解析。

### 为什么元素 ID 重复，order 又不是 0、1、2？

它们是来源局部标识和排序键，不是业务数据库主键和数组下标。请自行生成业务 ID，并在当前容器内排序。

### 能否直接把 JSON 交给 PPTist、ECharts 或任意编辑器？

通常需要适配。各工具的坐标单位、图表数据、富文本 schema、形状 path、分组变换和资源生命周期并不相同。尤其不要把独立 X 的散点系列压成共享 X 后仍宣称数据完全保真。

### 是否可以只靠此库得到每页 PNG、PDF 或完整 HTML？

没有统一的页面渲染/导出 API。可以基于输出自己实现渲染，再导出；若目标是高保真整页转换，应选择适合的 Office 渲染方案，而不是把内容解析等同于像素还原。

---

API 类型参考：[TypeScript 类型定义](https://github.com/pipipi-pikachu/pptxtojson/blob/master/types/index.d.ts)。npm 安装包通过 `dist/index.d.ts` 提供类型提示。
