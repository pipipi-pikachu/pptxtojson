import { getHorizontalAlign, getParagraphSpacing, getParagraphIndent, getParagraphListInfo } from './paragraph'
import { decodeXmlEntities, escapeHtml, getResourcesForSource, getTextByPathList } from './utils'

import {
  getFontFamilies,
  getFontColor,
  getFontSize,
  getFontBold,
  getFontItalic,
  getFontDecoration,
  getFontDecorationLine,
  getFontSpace,
  getFontSubscript,
  getFontShadow,
} from './fontStyle'

export function getTextNodeValue(node) {
  if (typeof node === 'string') return node
  if (node && typeof node.value === 'string') return node.value
  return undefined
}

function characterScript(token) {
  const character = decodeXmlEntities(token)
  if (character.startsWith('&') && character !== '&') return ''
  const code = character.codePointAt(0)
  if ((code >= 0x2e80 && code <= 0x9fff) || (code >= 0xac00 && code <= 0xd7af) ||
    (code >= 0xf900 && code <= 0xfaff) || (code >= 0xfe10 && code <= 0xfe6f) ||
    (code >= 0xff00 && code <= 0xffef) || (code >= 0x20000 && code <= 0x323af) ||
    (code >= 0x1100 && code <= 0x11ff)) return 'eastAsian'
  if (/[A-Za-zÀ-˿Ͱ-ԯ]/.test(character)) return 'latin'
  return ''
}

function scriptFontContent(styleInfo, value) {
  const families = styleInfo.fontFamilies
  if (!families.eastAsian || families.eastAsian === families.latin) return value
  const tokens = value.match(/&(?:#[xX][0-9a-fA-F]+|#\d+|[a-zA-Z][a-zA-Z0-9]*);|[\s\S]/gu) || []
  const scripts = tokens.map(characterScript)
  let currentScript = scripts.find(Boolean) || 'latin'
  const pieces = []
  for (let i = 0; i < tokens.length; i++) {
    currentScript = scripts[i] || currentScript
    const family = currentScript === 'eastAsian' ? families.eastAsian : families.latin
    const previous = pieces[pieces.length - 1]
    if (previous && previous.family === family) previous.value += tokens[i]
    else pieces.push({ family, value: tokens[i] })
  }
  if (pieces.every(piece => piece.family === families.latin)) return value
  const children = pieces.map(piece => {
    const cssFamily = cssFontFamily(piece.family)
    return `<span style="${escapeHtml('font-family: ' + cssFamily + ';')}">${piece.value}</span>`
  }).join('')
  return children
}

function cssFontFamily(value) {
  return /^[a-zA-Z][a-zA-Z0-9_-]*$/.test(value) ? value : JSON.stringify(value)
}

function renderStyledSpan(styleInfo, value) {
  return `<span style="${styleInfo.styleText}">${scriptFontContent(styleInfo, value)}</span>`
}

function renderLinkedSpan(styleInfo) {
  return `<span style="${styleInfo.styleText}"><a href="${escapeHtml(decodeXmlEntities(styleInfo.linkURL))}" target="_blank">${scriptFontContent(styleInfo, styleInfo.text)}</a></span>`
}

function listOpening(info) {
  const start = info.type === 'ol' ? ` start="${info.start}"` : ''
  const format = info.format || ''
  let markerStyle = ''
  if (info.type === 'ol') {
    const style = format.startsWith('alphaLc') ? 'lower-alpha' : format.startsWith('alphaUc') ? 'upper-alpha' : format.startsWith('romanLc') ? 'lower-roman' : format.startsWith('romanUc') ? 'upper-roman' : 'decimal'
    markerStyle = `list-style-type: ${style};`
  }
  else {
    const marker = info.marker === '•' || info.marker === '·' ? 'disc' : JSON.stringify(info.marker + ' ')
    markerStyle = `list-style-type: ${marker};`
  }
  return `<${info.type}${start} style="${escapeHtml('margin: 0;padding-left: 1.5em;' + markerStyle)}">`
}

export function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj, source = 'slide') {
  if (!textBodyNode) return ''

  let text = ''

  const pFontStyle = getTextByPathList(spNode, ['p:style', 'a:fontRef'])
  const slideMasterTextStyles = spNode && spNode['a:tcPr'] ? undefined : warpObj['slideMasterTextStyles']
  const defaultTextStyle = spNode && spNode['a:tcPr'] ? warpObj['defaultTextStyle'] : undefined

  const pNode = textBodyNode['a:p']
  const pNodes = pNode.constructor === Array ? pNode : [pNode]

  const listStyleType = type === 'obj' && getTextByPathList(spNode, ['p:nvSpPr', 'p:nvPr', 'p:ph']) ? 'body' : type
  const listStack = []
  const closeList = () => {
    const closed = listStack.pop()
    text += `</li></${closed.type}>`
  }

  for (const pNode of pNodes) {
    const inlineNodes = []
    for (const key of ['a:r', 'a:fld', 'a:br']) {
      const nodes = pNode[key]
      if (!nodes) continue
      const list = Array.isArray(nodes) ? nodes : [nodes]
      for (const node of list) inlineNodes.push({ node, isBreak: key === 'a:br' })
    }
    inlineNodes.sort((a, b) => {
      const aOrder = getTextByPathList(a.node, ['attrs', 'order'])
      const bOrder = getTextByPathList(b.node, ['attrs', 'order'])
      return (aOrder === undefined ? 0 : aOrder) - (bOrder === undefined ? 0 : bOrder)
    })

    const align = getHorizontalAlign(pNode, spNode, type, slideLayoutSpNode, slideMasterSpNode, warpObj)
    const spacing = getParagraphSpacing(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj)
    const indent = getParagraphIndent(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj)
    const listInfo = getParagraphListInfo(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, listStyleType, slideMasterTextStyles, warpObj)
    const listType = listInfo.type
    const listLevel = Math.max(0, getListLevel(pNode))

    const noWrap = getTextByPathList(textBodyNode, ['a:bodyPr', 'attrs', 'wrap']) === 'none'
    let styleText = `text-align: ${align};white-space: ${noWrap ? 'pre' : 'pre-wrap'};tab-size: 4;`
    if (spacing) {
      if (spacing.lineSpacing) styleText += `line-height: ${spacing.lineSpacing};`
      if (spacing.spaceBefore) styleText += `margin-top: ${spacing.spaceBefore};`
      if (spacing.spaceAfter) styleText += `margin-bottom: ${spacing.spaceAfter};`
    }
    if (indent) {
      if (!listType && indent.marginLeft) styleText += `margin-left: ${indent.marginLeft};`
      if (!listType && indent.textIndent) styleText += `text-indent: ${indent.textIndent};`
    }

    if (listType) {
      while (listStack.length && listStack[listStack.length - 1].level > listLevel) closeList()
      let previous = listStack[listStack.length - 1]
      if (previous && previous.level === listLevel && previous.signature !== listInfo.signature) {
        closeList()
        previous = listStack[listStack.length - 1]
      }
      const sameList = previous && previous.level === listLevel
      if (sameList) text += '</li>'
      else {
        text += listOpening(listInfo)
        listStack.push({ type: listType, level: listLevel, signature: listInfo.signature })
      }
      const itemValue = sameList && listType === 'ol' && listInfo.explicitStart ? ` value="${listInfo.start}"` : ''
      text += `<li${itemValue} data-pptx-level="${listLevel}"><p style="${styleText}">`
    }
    else {
      while (listStack.length) closeList()
      text += `<p style="${styleText}">`
    }
    
    if (inlineNodes.length === 0) {
      text += genSpanElement(pNode, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, defaultTextStyle, warpObj, source)
    }
    else {
      let prevStyleInfo = null
      let accumulatedText = ''
      const flushText = () => {
        if (accumulatedText && prevStyleInfo) text += renderStyledSpan(prevStyleInfo, accumulatedText)
        accumulatedText = ''
        prevStyleInfo = null
      }

      for (const item of inlineNodes) {
        if (item.isBreak) {
          flushText()
          text += '<br>'
          continue
        }
        const styleInfo = getSpanStyleInfo(item.node, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, defaultTextStyle, warpObj, source)
        if (styleInfo.hasLink) {
          flushText()
          text += renderLinkedSpan(styleInfo)
        }
        else {
          if (prevStyleInfo && (prevStyleInfo.styleText !== styleInfo.styleText || prevStyleInfo.fontFamilies.eastAsian !== styleInfo.fontFamilies.eastAsian)) flushText()
          prevStyleInfo = styleInfo
          accumulatedText += styleInfo.text
        }
      }
      flushText()
      if (inlineNodes[inlineNodes.length - 1].isBreak) text += '&nbsp;'
    }

    text += '</p>'
  }
  while (listStack.length) closeList()
  return text
}

export function getListType(node) {
  const pPrNode = node['a:pPr']
  if (!pPrNode) return ''

  if (pPrNode['a:buChar']) return 'ul'
  if (pPrNode['a:buAutoNum']) return 'ol'
  
  return ''
}
export function getListLevel(node) {
  const pPrNode = node['a:pPr']
  if (!pPrNode) return -1

  const lvlNode = getTextByPathList(pPrNode, ['attrs', 'lvl'])
  if (lvlNode !== undefined) return parseInt(lvlNode)

  return 0
}

export function genSpanElement(node, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, defaultTextStyle, warpObj, source = 'slide') {
  const styleInfo = getSpanStyleInfo(node, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, defaultTextStyle, warpObj, source)
  const content = renderStyledSpan(styleInfo, styleInfo.text)
  if (styleInfo.hasLink) return renderLinkedSpan(styleInfo)
  return content
}

export function getSpanStyleInfo(node, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, defaultTextStyle, warpObj, source = 'slide') {
  let lvl = 1
  const pPrNode = pNode['a:pPr']
  const lvlNode = getTextByPathList(pPrNode, ['attrs', 'lvl'])
  if (lvlNode !== undefined) lvl = parseInt(lvlNode) + 1

  let text = getTextNodeValue(node['a:t'])
  if (typeof text !== 'string') text = getTextNodeValue(getTextByPathList(node, ['a:fld', 'a:t']))
  if (typeof text !== 'string') text = '&nbsp;'

  let styleText = ''
  const fontColor = getFontColor(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, pFontStyle, warpObj)
  const fontSize = getFontSize(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, defaultTextStyle)
  const fontFamilies = getFontFamilies(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, warpObj)
  const fontType = fontFamilies.latin
  const fontBold = getFontBold(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const fontItalic = getFontItalic(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const fontDecoration = getFontDecoration(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const fontDecorationLine = getFontDecorationLine(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const fontSpace = getFontSpace(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const shadow = getFontShadow(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, warpObj)
  const subscript = getFontSubscript(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)

  if (fontColor) {
    if (typeof fontColor === 'string') styleText += `color: ${fontColor};`
    else if (fontColor.colors) {
      const { colors, rot } = fontColor
      const stops = colors.map(item => `${item.color} ${item.pos}`).join(', ')
      const gradientStyle = `linear-gradient(${rot + 90}deg, ${stops})`
      styleText += `background: ${gradientStyle}; background-clip: text; color: transparent;`
    }
  }
  if (fontSize) styleText += `font-size: ${fontSize};`
  if (fontType) styleText += escapeHtml(`font-family: ${cssFontFamily(fontType)};`)
  if (fontBold) styleText += `font-weight: ${fontBold};`
  if (fontItalic) styleText += `font-style: ${fontItalic};`
  if (fontDecoration) styleText += `text-decoration: ${fontDecoration};`
  if (fontDecorationLine) styleText += `text-decoration-line: ${fontDecorationLine};`
  if (fontSpace) styleText += `letter-spacing: ${fontSpace};`
  if (subscript) styleText += `vertical-align: ${subscript};`
  if (shadow) styleText += `text-shadow: ${shadow};`

  const linkID = getTextByPathList(node, ['a:rPr', 'a:hlinkClick', 'attrs', 'r:id'])
  const resources = getResourcesForSource(warpObj, source)
  const resource = linkID ? resources[linkID] : null
  const hasLink = !!resource && !!resource.target

  return {
    styleText,
    text,
    fontFamilies,
    hasLink,
    linkURL: hasLink ? resource.target : null
  }
}
