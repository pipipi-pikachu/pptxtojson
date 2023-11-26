import { getTextByPathList } from './utils'

export function getSchemeColorFromTheme(schemeClr, warpObj) {
  switch (schemeClr) {
    case 'tx1':
      schemeClr = 'a:dk1'
      break
    case 'tx2':
      schemeClr = 'a:dk2'
      break
    case 'bg1':
      schemeClr = 'a:lt1'
      break
    case 'bg2':
      schemeClr = 'a:lt2'
      break
    default:
      break
  }
  const refNode = getTextByPathList(warpObj['themeContent'], ['a:theme', 'a:themeElements', 'a:clrScheme', schemeClr])
  let color = getTextByPathList(refNode, ['a:srgbClr', 'attrs', 'val'])
  if (!color && refNode) color = getTextByPathList(refNode, ['a:sysClr', 'attrs', 'lastClr'])
  return color
}