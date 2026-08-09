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

export function getBorder(node, elType, warpObj) {
  let defaultLineNode = null
  const lnRefNode = getTextByPathList(node, ['p:style', 'a:lnRef'])
  if (lnRefNode) {
    const lnIdx = getTextByPathList(lnRefNode, ['attrs', 'idx'])
    defaultLineNode = warpObj['themeContent']['a:theme']['a:themeElements']['a:fmtScheme']['a:lnStyleLst']['a:ln'][Number(lnIdx) - 1]
  }
  const defaultBorderWidth = defaultLineNode && !getTextByPathList(defaultLineNode, ['a:noFill']) ? (parseInt(getTextByPathList(defaultLineNode, ['attrs', 'w'])) / 12700) : 0
  let lineNode = getTextByPathList(node, ['p:spPr', 'a:ln'])
  // if (!lineNode) {
  //   const lnRefNode = getTextByPathList(node, ['p:style', 'a:lnRef'])
  //   if (lnRefNode) {
  //     const lnIdx = getTextByPathList(lnRefNode, ['attrs', 'idx'])
  //     lineNode = warpObj['themeContent']['a:theme']['a:themeElements']['a:fmtScheme']['a:lnStyleLst']['a:ln'][Number(lnIdx) - 1]
  //   }
  // }
  if (!lineNode) lineNode = node

  const isNoFill = getTextByPathList(lineNode, ['a:noFill'])

  const lineWidth = getTextByPathList(lineNode, ['attrs', 'w'])
  let borderWidth = isNoFill ? 0 : lineWidth ? (parseInt(lineWidth) / 12700) : defaultBorderWidth
  if (isNaN(borderWidth)) {
    if (lineNode) borderWidth = 0
    else if (elType !== 'obj') borderWidth = 0
    else borderWidth = 1
  }

  const solidFill = getTextByPathList(lineNode, ['a:solidFill'])
  let borderColor = getSolidFill(solidFill, undefined, undefined, warpObj)

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
