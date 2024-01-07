import { getSolidFill } from './fill'

export function getShadow(node, warpObj) {
  const slideFactor = warpObj.options.slideFactor
  const chdwClrNode = getSolidFill(node, undefined, undefined, warpObj)
  const outerShdwAttrs = node['attrs']
  const dir = (outerShdwAttrs['dir']) ? (parseInt(outerShdwAttrs['dir']) / 60000) : 0
  const dist = parseInt(outerShdwAttrs['dist']) * slideFactor
  const blurRad = outerShdwAttrs['blurRad'] ? parseFloat((parseInt(outerShdwAttrs['blurRad']) * slideFactor).toFixed(2)) : ''
  const vx = dist * Math.sin(dir * Math.PI / 180)
  const hx = dist * Math.cos(dir * Math.PI / 180)

  return {
    h: parseFloat(hx.toFixed(2)),
    v: parseFloat(vx.toFixed(2)),
    blur: blurRad,
    color: '#' + chdwClrNode,
  }
}