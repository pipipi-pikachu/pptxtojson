import { getTextByPathList } from './utils'

export function getHorizontalAlign(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) {
  let algn = getTextByPathList(node, ['a:pPr', 'attrs', 'algn'])

  if (!algn) algn = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:p', 'a:pPr', 'attrs', 'algn'])
  if (!algn) algn = getTextByPathList(slideMasterSpNode, ['p:txBody', 'a:p', 'a:pPr', 'attrs', 'algn'])
  if (!algn) {
    switch (type) {
      case 'title':
      case 'subTitle':
      case 'ctrTitle':
        algn = getTextByPathList(slideMasterTextStyles, ['p:titleStyle', 'a:lvl1pPr', 'attrs', 'alng'])
        break
      default:
        algn = getTextByPathList(slideMasterTextStyles, ['p:otherStyle', 'a:lvl1pPr', 'attrs', 'alng'])
    }
  }
  if (!algn) {
    if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') return 'center'
    else if (type === 'sldNum') return 'right'
  }
  return algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : 'left'
}

export function getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode) {
  let anchor = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'anchor'])
  if (!anchor) {
    anchor = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:bodyPr', 'attrs', 'anchor'])
    if (!anchor) {
      anchor = getTextByPathList(slideMasterSpNode, ['p:txBody', 'a:bodyPr', 'attrs', 'anchor'])
      if (!anchor) anchor = 't'
    }
  }
  return (anchor === 'ctr') ? 'mid' : ((anchor === 'b') ? 'down' : 'up')
}