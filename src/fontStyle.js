import { getTextByPathList } from './utils'
import { getShadow } from './shadow'
import { getFillType, getGradientFill, getSolidFill } from './fill'

function pushStyleNode(styleNodes, styleNode) {
  if (styleNode) styleNodes.push(styleNode)
}

function getLevelPath(lvl) {
  return `a:lvl${lvl}pPr`
}

function appendTextBodyStyleNodes(styleNodes, textBodyNode, lvl) {
  if (!textBodyNode) return

  const lvlPath = getLevelPath(lvl)
  pushStyleNode(styleNodes, getTextByPathList(textBodyNode, ['a:lstStyle', lvlPath, 'a:defRPr']))
}

function appendShapeStyleNodes(styleNodes, shapeNode, lvl) {
  if (!shapeNode) return

  const lvlPath = getLevelPath(lvl)
  pushStyleNode(styleNodes, getTextByPathList(shapeNode, ['p:txBody', 'a:lstStyle', lvlPath, 'a:defRPr']))
  pushStyleNode(styleNodes, getTextByPathList(shapeNode, ['p:txBody', 'a:p', 'a:pPr', 'a:defRPr']))
}

function appendMasterTextStyleNodes(styleNodes, type, lvl, slideMasterTextStyles) {
  if (!slideMasterTextStyles) return

  const lvlPath = getLevelPath(lvl)

  if (type === 'title' || type === 'ctrTitle' || type === 'subTitle') {
    pushStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:titleStyle', lvlPath, 'a:defRPr']))
    if (type === 'subTitle') {
      pushStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', lvlPath, 'a:defRPr']))
    }
  }
  else if (type === 'body') {
    pushStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', lvlPath, 'a:defRPr']))
  }
  else {
    pushStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:otherStyle', lvlPath, 'a:defRPr']))
  }
}

function appendDefaultTextStyleNodes(styleNodes, lvl, defaultTextStyle) {
  if (!defaultTextStyle) return

  const lvlPath = getLevelPath(lvl)
  pushStyleNode(styleNodes, getTextByPathList(defaultTextStyle, [lvlPath, 'a:defRPr']))
  pushStyleNode(styleNodes, getTextByPathList(defaultTextStyle, ['a:defPPr', 'a:defRPr']))
}

function getBaseFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, lvl) {
  const styleNodes = []
  const runStyleNode = getTextByPathList(node, ['a:rPr'])

  pushStyleNode(styleNodes, runStyleNode)
  if (!runStyleNode) {
    pushStyleNode(styleNodes, getTextByPathList(pNode, ['a:endParaRPr']))
  }
  pushStyleNode(styleNodes, getTextByPathList(pNode, ['a:pPr', 'a:defRPr']))

  appendTextBodyStyleNodes(styleNodes, textBodyNode, lvl)
  appendShapeStyleNodes(styleNodes, slideLayoutSpNode, lvl)
  appendShapeStyleNodes(styleNodes, slideMasterSpNode, lvl)

  return styleNodes
}

function getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getBaseFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, lvl)
  appendMasterTextStyleNodes(styleNodes, type, lvl, slideMasterTextStyles)

  return styleNodes
}

function getFontAttr(styleNodes, attrName) {
  for (const styleNode of styleNodes) {
    const attrValue = getTextByPathList(styleNode, ['attrs', attrName])
    if (attrValue !== undefined && attrValue !== '') return attrValue
  }

  return ''
}

function getFontTypeface(styleNodes) {
  for (const styleNode of styleNodes) {
    const eaTypeface = getTextByPathList(styleNode, ['a:ea', 'attrs', 'typeface'])
    const latinTypeface = getTextByPathList(styleNode, ['a:latin', 'attrs', 'typeface'])
    // 优先使用 a:ea（东亚字体），对中文内容更准确
    const typeface = eaTypeface || latinTypeface
    if (typeface) return buildFontFamily(typeface)
  }

  return ''
}

// ---- 字体分类 ----
const FontCategory = {
  SANS_SERIF: 'sans-serif',
  SERIF: 'serif',
  MONOSPACE: 'monospace',
  CURSIVE: 'cursive',
  FANTASY: 'fantasy',
}

// 已知字体分类映射
const FONT_CATEGORY_MAP = {
  // 无衬线
  '微软雅黑': FontCategory.SANS_SERIF, 'Microsoft YaHei': FontCategory.SANS_SERIF,
  'MicrosoftYaHei': FontCategory.SANS_SERIF, 'Microsoft YaHei Light': FontCategory.SANS_SERIF,
  '黑体': FontCategory.SANS_SERIF, 'SimHei': FontCategory.SANS_SERIF,
  '华文细黑': FontCategory.SANS_SERIF, 'STXiHei': FontCategory.SANS_SERIF,
  '等线': FontCategory.SANS_SERIF, 'DengXian': FontCategory.SANS_SERIF,
  '思源黑体': FontCategory.SANS_SERIF, 'Source Han Sans': FontCategory.SANS_SERIF,
  '阿里巴巴普惠体': FontCategory.SANS_SERIF, 'MiSans': FontCategory.SANS_SERIF,
  '得意黑': FontCategory.SANS_SERIF, '优设标题黑': FontCategory.SANS_SERIF,
  '峰广明锐体': FontCategory.SANS_SERIF, '摄图摩登小方体': FontCategory.SANS_SERIF,
  '素材集市酷方体': FontCategory.SANS_SERIF, '锐字真言体': FontCategory.SANS_SERIF,
  'Arial': FontCategory.SANS_SERIF, 'Helvetica': FontCategory.SANS_SERIF,
  'Segoe UI': FontCategory.SANS_SERIF, 'Calibri': FontCategory.SANS_SERIF,
  'Inter': FontCategory.SANS_SERIF, 'Roboto': FontCategory.SANS_SERIF,
  'Open Sans': FontCategory.SANS_SERIF, 'Montserrat': FontCategory.SANS_SERIF,
  // 衬线
  '宋体': FontCategory.SERIF, 'SimSun': FontCategory.SERIF,
  'NSimSun': FontCategory.SERIF, '新宋体': FontCategory.SERIF,
  '楷体': FontCategory.SERIF, 'KaiTi': FontCategory.SERIF,
  'KaiTi_GB2312': FontCategory.SERIF,
  '仿宋': FontCategory.SERIF, 'FangSong': FontCategory.SERIF,
  'FangSong_GB2312': FontCategory.SERIF,
  '华文楷体': FontCategory.SERIF, 'STKaiti': FontCategory.SERIF,
  '华文宋体': FontCategory.SERIF, 'STSong': FontCategory.SERIF,
  '华文仿宋': FontCategory.SERIF, 'STFangSong': FontCategory.SERIF,
  '思源宋体': FontCategory.SERIF, 'Source Han Serif': FontCategory.SERIF,
  '文鼎PL宋体': FontCategory.SERIF, '文鼎PL楷体': FontCategory.SERIF,
  '朱雀仿宋': FontCategory.SERIF, '霞鹜文楷': FontCategory.SERIF,
  'Times New Roman': FontCategory.SERIF, 'Georgia': FontCategory.SERIF,
  'Merriweather': FontCategory.SERIF, 'Literata': FontCategory.SERIF,
  // 等宽
  'Courier New': FontCategory.MONOSPACE, 'Consolas': FontCategory.MONOSPACE,
  'JetBrains Mono': FontCategory.MONOSPACE,
  // 手写/装饰
  '仓耳小丸子': FontCategory.CURSIVE, '喵喵奶糖': FontCategory.CURSIVE,
  '糯米奶团体': FontCategory.CURSIVE, '站酷快乐体': FontCategory.CURSIVE,
  '字制区喜脉体': FontCategory.CURSIVE, '素材集市康康体': FontCategory.CURSIVE,
  '途牛类圆体': FontCategory.FANTASY,
}

// ---- 字体别名映射 ----
// key: PPT中的原始字体名
// value: 风格最接近的可用 web font 名（即 PPTist 中 @font-face 注册的字体名）
// 用途: 当某个字体没有本地安装时，优先用这个 web font 替代，而不是只按分类回退
const FONT_ALIAS_MAP = {
  // WPS 装饰字体 → 风格接近的开源字体
  '喵喵奶糖': 'SucaiJishiKangkang',
  '糯米奶团体': 'SucaiJishiKangkang',
  '素材集市康康体': 'SucaiJishiKangkang',
  '苍耳粗体': 'SucaiJishiKangkang',       // 和喵喵奶糖风格接近
  '素材集市酷方体': 'SucaiJishiCoolSquare',
  '仓耳小丸子': 'CangerXiaowanzi',
  '途牛类圆体': 'TuniuRounded',
  '站酷快乐体': 'ZcoolHappy',
  '字制区喜脉体': 'ZizhiQuXiMai',
  '锐字真言体': 'RuiziZhenyan',
  // 英文字体替代
  'Times New Roman': 'Merriweather',
  'Arial': 'Inter',
  'Calibri': 'OpenSans',
  'Courier New': 'JetBrainsMono',
}

// 每个分类的回退字体链
const CATEGORY_FALLBACK = {
  [FontCategory.SANS_SERIF]: '"Microsoft YaHei", "SourceHanSans", sans-serif',
  [FontCategory.SERIF]: '"SimSun", "SourceHanSerif", serif',
  [FontCategory.MONOSPACE]: '"JetBrains Mono", "Consolas", monospace',
  [FontCategory.CURSIVE]: '"LXGWWenKai", "KaiTi", cursive',
  [FontCategory.FANTASY]: '"CangerXiaowanzi", "Microsoft YaHei", fantasy',
}

// 推断字体分类
function inferFontCategory(fontName) {
  if (FONT_CATEGORY_MAP[fontName]) return FONT_CATEGORY_MAP[fontName]
  if (/黑|hei|sans|gothic|grotesk/i.test(fontName)) return FontCategory.SANS_SERIF
  if (/宋|song|serif|roman|明朝|mincho/i.test(fontName)) return FontCategory.SERIF
  if (/楷|kai|仿|fang/i.test(fontName)) return FontCategory.SERIF
  if (/mono|consol|courier|code/i.test(fontName.toLowerCase())) return FontCategory.MONOSPACE
  if (/手写|行|草|cursive|hand|script|calligrap/i.test(fontName)) return FontCategory.CURSIVE
  if (/圆|艺术|装饰|fantasy|fun|happy|可爱|胖|童|cool/i.test(fontName)) return FontCategory.FANTASY
  return FontCategory.SANS_SERIF
}

// 生成完整的 CSS font-family 回退链
// 流程: 原始字体名 → 别名web font → 分类回退链
// 例: "苍耳粗体" → "苍耳粗体", "SucaiJishiKangkang", "LXGWWenKai", "KaiTi", cursive
function buildFontFamily(fontName) {
  if (!fontName) return ''
  // 跳过 theme 占位符
  if (fontName.startsWith('+')) return ''
  const category = inferFontCategory(fontName)
  const fallback = CATEGORY_FALLBACK[category]
  const alias = FONT_ALIAS_MAP[fontName]
  if (alias) {
    return `"${fontName}", "${alias}", ${fallback}`
  }
  return `"${fontName}", ${fallback}`
}

function getColorFromNode(node, warpObj) {
  if (!node) return ''

  const fillType = getFillType(node)
  if (fillType === 'SOLID_FILL') {
    return getSolidFill(node['a:solidFill'], undefined, undefined, warpObj)
  }
  if (fillType === 'GRADIENT_FILL') {
    return getGradientFill(node['a:gradFill'], warpObj)
  }

  return ''
}

function getFontColorFromStyleNodes(styleNodes, warpObj) {
  for (const styleNode of styleNodes) {
    const color = getColorFromNode(styleNode, warpObj)
    if (color) return color
  }

  return ''
}

function getTextShadowFromStyleNodes(styleNodes, warpObj) {
  for (const styleNode of styleNodes) {
    const txtShadow = getTextByPathList(styleNode, ['a:effectLst', 'a:outerShdw'])
    if (!txtShadow) continue

    const shadow = getShadow(txtShadow, warpObj)
    if (shadow) return shadow
  }

  return null
}

export function getFontType(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, warpObj) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const typeface = getFontTypeface(styleNodes)

  if (!typeface) {
    const fontSchemeNode = getTextByPathList(warpObj['themeContent'], ['a:theme', 'a:themeElements', 'a:fontScheme'])

    if (fontSchemeNode) {
      let schemeTypeface = ''
      // 优先取东亚字体，对中文内容更准确
      const eaMajor = getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:ea', 'attrs', 'typeface'])
      const eaMinor = getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:ea', 'attrs', 'typeface'])
      const ltMajor = getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:latin', 'attrs', 'typeface'])
      const ltMinor = getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface'])

      if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
        schemeTypeface = eaMajor || ltMajor || ''
      }
      else {
        schemeTypeface = eaMinor || ltMinor || ''
      }

      if (schemeTypeface) return buildFontFamily(schemeTypeface)
    }

    return ''
  }

  return typeface
}

export function getFontColor(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, pFontStyle, warpObj) {
  const styleNodes = getBaseFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, lvl)
  let color = getFontColorFromStyleNodes(styleNodes, warpObj)

  if (!color) {
    if (pFontStyle) color = getSolidFill(pFontStyle, undefined, undefined, warpObj)
    if (!color) {
      const layoutFontStyle = getTextByPathList(slideLayoutSpNode, ['p:style', 'a:fontRef'])
      if (layoutFontStyle) color = getSolidFill(layoutFontStyle, undefined, undefined, warpObj)
    }
    if (!color) {
      const masterFontStyle = getTextByPathList(slideMasterSpNode, ['p:style', 'a:fontRef'])
      if (masterFontStyle) color = getSolidFill(masterFontStyle, undefined, undefined, warpObj)
    }
  }

  if (!color) {
    appendMasterTextStyleNodes(styleNodes, type, lvl, slideMasterTextStyles)
    color = getFontColorFromStyleNodes(styleNodes, warpObj)
  }

  return color || ''
}

export function getFontSize(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, defaultTextStyle) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  appendDefaultTextStyleNodes(styleNodes, lvl, defaultTextStyle)
  const sz = getFontAttr(styleNodes, 'sz')
  let fontSize = sz ? parseInt(sz) / 100 : undefined

  if ((isNaN(fontSize) || !fontSize) && (type === 'dt' || type === 'sldNum')) fontSize = 12

  fontSize = (isNaN(fontSize) || !fontSize) ? 18 : fontSize

  return fontSize + 'pt'
}

export function getFontBold(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  return getFontAttr(styleNodes, 'b') === '1' ? 'bold' : ''
}

export function getFontItalic(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  return getFontAttr(styleNodes, 'i') === '1' ? 'italic' : ''
}

export function getFontDecoration(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  return getFontAttr(styleNodes, 'u') === 'sng' ? 'underline' : ''
}

export function getFontDecorationLine(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  return getFontAttr(styleNodes, 'strike') === 'sngStrike' ? 'line-through' : ''
}

export function getFontSpace(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const spc = getFontAttr(styleNodes, 'spc')
  return (spc && parseInt(spc) !== 0) ? (parseInt(spc) / 100 + 'pt') : ''
}

export function getFontSubscript(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const baseline = getFontAttr(styleNodes, 'baseline')
  if (!baseline || parseInt(baseline) === 0) return ''
  return parseInt(baseline) > 0 ? 'super' : 'sub'
}

export function getFontShadow(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, warpObj) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const shadow = getTextShadowFromStyleNodes(styleNodes, warpObj)
  if (shadow) {
    const { h, v, blur, color } = shadow
    if (!isNaN(v) && !isNaN(h)) {
      return h + 'pt ' + v + 'pt ' + (blur ? blur + 'pt' : '') + ' ' + color
    }
  }
  return ''
}
