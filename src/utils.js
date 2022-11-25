export function base64ArrayBuffer(arrayBuffer) {
  const encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
  const bytes = new Uint8Array(arrayBuffer)
  const byteLength = bytes.byteLength
  const byteRemainder = byteLength % 3
  const mainLength = byteLength - byteRemainder
  
  let base64 = ''
  let a, b, c, d
  let chunk

  for (let i = 0; i < mainLength; i = i + 3) {
    chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2]
    a = (chunk & 16515072) >> 18
    b = (chunk & 258048) >> 12
    c = (chunk & 4032) >> 6
    d = chunk & 63
    base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d]
  }

  if (byteRemainder === 1) {
    chunk = bytes[mainLength]
    a = (chunk & 252) >> 2
    b = (chunk & 3) << 4
    base64 += encodings[a] + encodings[b] + '=='
  } 
  else if (byteRemainder === 2) {
    chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1]
    a = (chunk & 64512) >> 10
    b = (chunk & 1008) >> 4
    c = (chunk & 15) << 2
    base64 += encodings[a] + encodings[b] + encodings[c] + '='
  }

  return base64
}

export function extractFileExtension(filename) {
  return filename.substr((~-filename.lastIndexOf('.') >>> 0) + 2)
}

export function eachElement(node, doFunction) {
  if (!node) return node

  let result = ''
  if (node.constructor === Array) {
    for (let i = 0; i < node.length; i++) {
      result += doFunction(node[i], i)
    }
  } 
  else result += doFunction(node, 0)

  return result
}

export function getTextByPathList(node, path) {
  if (path.constructor !== Array) throw Error('Error of path type! path is not array.')

  if (!node) return node

  for (const key of path) {
    node = node[key]
    if (!node) return node
  }

  return node
}

export function angleToDegrees(angle) {
  if (!angle) return 0
  return Math.round(angle / 60000)
}