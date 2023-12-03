import * as txml from 'txml/dist/txml.mjs'

let cust_attr_order = 0

export function simplifyLostLess(children, parentAttributes = {}) {
  const out = {}
  if (!children.length) return out

  if (children.length === 1 && typeof children[0] === 'string') {
    return Object.keys(parentAttributes).length ? {
      attrs: { order: cust_attr_order++, ...parentAttributes },
      value: children[0],
    } : children[0]
  }
  for (const child of children) {
    if (typeof child !== 'object') return
    if (child.tagName === '?xml') continue

    if (!out[child.tagName]) out[child.tagName] = []

    const kids = simplifyLostLess(child.children || [], child.attributes)
    out[child.tagName].push(kids)

    if (Object.keys(child.attributes).length) {
      kids.attrs = { order: cust_attr_order++, ...child.attributes }
    }
  }
  for (const child in out) {
    if (out[child].length === 1) out[child] = out[child][0]
  }

  return out
}

export async function readXmlFile(zip, filename) {
  try {
    const data = await zip.file(filename).async('string')
    return simplifyLostLess(txml.parse(data))
  }
  catch {
    return null
  }
}