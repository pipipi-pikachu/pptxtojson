import tinycolor from 'tinycolor2'
import { getSolidFill } from './fill'
import { getSchemeColorFromTheme } from './schemeColor'
import { getTextByPathList } from './utils'

function getLineEnd(node) {
  const attrs = getTextByPathList(node, ['attrs'])
  if (!attrs) return undefined

  const lineEnd = { type: attrs.type || 'none' }
  if (attrs.w) lineEnd.width = attrs.w
  if (attrs.len) lineEnd.length = attrs.len
  return lineEnd
}

export function getShapeBorder(node, elType, warpObj, { slideLayoutSpNode, slideMasterSpNode } = {}) {
  // Inherit only a missing whole outline. A local line (including noFill), or
  // a local theme-line reference, must not be overwritten by an ancestor.
  for (const candidate of [node, slideLayoutSpNode, slideMasterSpNode]) {
    if (getTextByPathList(candidate, ['p:spPr', 'a:ln']) || getTextByPathList(candidate, ['p:style', 'a:lnRef'])) {
      return getBorder(candidate, elType, warpObj)
    }
  }
  return getBorder(node, elType, warpObj)
}

export function getBorder(node, elType, warpObj) {
  let lineNode = getTextByPathList(node, ['p:spPr', 'a:ln'])
  const lnRefNode = getTextByPathList(node, ['p:style', 'a:lnRef'])
  const phClr = getSolidFill(lnRefNode, undefined, undefined, warpObj)
  if (!lineNode && lnRefNode) {
    const lnIdx = Number(getTextByPathList(lnRefNode, ['attrs', 'idx']))
    const lines = getTextByPathList(warpObj['themeContent'], ['a:theme', 'a:themeElements', 'a:fmtScheme', 'a:lnStyleLst', 'a:ln'])
    if (lnIdx === 0) lineNode = { 'a:noFill': {} }
    else if (lines) lineNode = Array.isArray(lines) ? lines[lnIdx - 1] : lnIdx === 1 ? lines : undefined
  }
  if (!lineNode) lineNode = node

  const isNoFill = getTextByPathList(lineNode, ['a:noFill'])

  let borderWidth = isNoFill ? 0 : (parseInt(getTextByPathList(lineNode, ['attrs', 'w'])) / 12700)
  if (isNaN(borderWidth)) {
    if (lineNode) borderWidth = 0
    else if (elType !== 'obj') borderWidth = 0
    else borderWidth = 1
  }

  const solidFill = getTextByPathList(lineNode, ['a:solidFill'])
  let borderColor = getSolidFill(solidFill, undefined, phClr, warpObj)

  if (!borderColor) {
    const schemeClrNode = getTextByPathList(node, ['p:style', 'a:lnRef', 'a:schemeClr'])
    const schemeClr = 'a:' + getTextByPathList(schemeClrNode, ['attrs', 'val'])
    borderColor = getSchemeColorFromTheme(schemeClr, warpObj)

    if (borderColor) {
      let shade = getTextByPathList(schemeClrNode, ['a:shade', 'attrs', 'val'])

      if (shade) {
        shade = parseInt(shade) / 100000
        
        const color = tinycolor('#' + borderColor).toHsl()
        borderColor = tinycolor({ h: color.h, s: color.s, l: color.l * shade, a: color.a }).toHex()
      }
    }
  }

  if (!borderColor) borderColor = '#000000'
  else if (!borderColor.startsWith('#')) borderColor = `#${borderColor}`

  const type = getTextByPathList(lineNode, ['a:prstDash', 'attrs', 'val'])
  let borderType = 'solid'
  let strokeDasharray = '0'
  switch (type) {
    case 'solid':
      borderType = 'solid'
      strokeDasharray = '0'
      break
    case 'dash':
      borderType = 'dashed'
      strokeDasharray = '5'
      break
    case 'dashDot':
      borderType = 'dashed'
      strokeDasharray = '5, 5, 1, 5'
      break
    case 'dot':
      borderType = 'dotted'
      strokeDasharray = '1, 5'
      break
    case 'lgDash':
      borderType = 'dashed'
      strokeDasharray = '10, 5'
      break
    case 'lgDashDotDot':
      borderType = 'dotted'
      strokeDasharray = '10, 5, 1, 5, 1, 5'
      break
    case 'sysDash':
      borderType = 'dashed'
      strokeDasharray = '5, 2'
      break
    case 'sysDashDot':
      borderType = 'dotted'
      strokeDasharray = '5, 2, 1, 5'
      break
    case 'sysDashDotDot':
      borderType = 'dotted'
      strokeDasharray = '5, 2, 1, 5, 1, 5'
      break
    case 'sysDot':
      borderType = 'dotted'
      strokeDasharray = '2, 5'
      break
    default:
  }

  const headEnd = getLineEnd(getTextByPathList(lineNode, ['a:headEnd']))
  const tailEnd = getLineEnd(getTextByPathList(lineNode, ['a:tailEnd']))

  return {
    borderColor,
    borderWidth,
    borderType,
    strokeDasharray,
    headEnd,
    tailEnd,
  }
}
