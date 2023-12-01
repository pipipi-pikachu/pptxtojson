import tinycolor from 'tinycolor2'

export function hueToRgb(t1, t2, hue) {
  if (hue < 0) hue += 6
  if (hue >= 6) hue -= 6
  if (hue < 1) return (t2 - t1) * hue + t1
  else if (hue < 3) return t2
  else if (hue < 4) return (t2 - t1) * (4 - hue) + t1
  return t1
}

export function hslToRgb(hue, sat, light) {
  let t2
  hue = hue / 60
  if (light <= 0.5) {
    t2 = light * (sat + 1)
  } 
  else {
    t2 = light + sat - (light * sat)
  }
  const t1 = light * 2 - t2
  const r = hueToRgb(t1, t2, hue + 2) * 255
  const g = hueToRgb(t1, t2, hue) * 255
  const b = hueToRgb(t1, t2, hue - 2) * 255
  return { r, g, b }
}

export function applyShade(rgbStr, shadeValue, isAlpha) {
  const color = tinycolor(rgbStr).toHsl()
  if (shadeValue >= 1) shadeValue = 1
  const cacl_l = Math.min(color.l * shadeValue, 1)
  if (isAlpha) {
    return tinycolor({
      h: color.h,
      s: color.s,
      l: cacl_l,
      a: color.a
    }).toHex8()
  }

  return tinycolor({
    h: color.h,
    s: color.s,
    l: cacl_l,
    a: color.a,
  }).toHex()
}

export function applyTint(rgbStr, tintValue, isAlpha) {
  const color = tinycolor(rgbStr).toHsl()
  if (tintValue >= 1) tintValue = 1
  const cacl_l = color.l * tintValue + (1 - tintValue)
  if (isAlpha) {
    return tinycolor({
      h: color.h,
      s: color.s,
      l: cacl_l,
      a: color.a
    }).toHex8()
  }

  return tinycolor({
    h: color.h,
    s: color.s,
    l: cacl_l,
    a: color.a
  }).toHex()
}

export function applyLumOff(rgbStr, offset, isAlpha) {
  const color = tinycolor(rgbStr).toHsl()
  const lum = offset + color.l
  if (lum >= 1) {
    if (isAlpha) {
      return tinycolor({
        h: color.h,
        s: color.s,
        l: 1,
        a: color.a
      }).toHex8()
    }
      
    return tinycolor({
      h: color.h,
      s: color.s,
      l: 1,
      a: color.a
    }).toHex()
  }
  if (isAlpha) {
    return tinycolor({
      h: color.h,
      s: color.s,
      l: lum,
      a: color.a
    }).toHex8()
  }

  return tinycolor({
    h: color.h,
    s: color.s,
    l: lum,
    a: color.a
  }).toHex()
}

export function applyLumMod(rgbStr, multiplier, isAlpha) {
  const color = tinycolor(rgbStr).toHsl()
  let cacl_l = color.l * multiplier
  if (cacl_l >= 1) cacl_l = 1
  if (isAlpha) {
    return tinycolor({
      h: color.h,
      s: color.s,
      l: cacl_l,
      a: color.a
    }).toHex8()
  }

  return tinycolor({
    h: color.h,
    s: color.s,
    l: cacl_l,
    a: color.a
  }).toHex()
}

export function applyHueMod(rgbStr, multiplier, isAlpha) {
  const color = tinycolor(rgbStr).toHsl()
  let cacl_h = color.h * multiplier
  if (cacl_h >= 360) cacl_h = cacl_h - 360
  if (isAlpha) {
    return tinycolor({
      h: cacl_h,
      s: color.s,
      l: color.l,
      a: color.a
    }).toHex8()
  }

  return tinycolor({
    h: cacl_h,
    s: color.s,
    l: color.l,
    a: color.a
  }).toHex()
}

export function applySatMod(rgbStr, multiplier, isAlpha) {
  const color = tinycolor(rgbStr).toHsl()
  let cacl_s = color.s * multiplier
  if (cacl_s >= 1) cacl_s = 1
  if (isAlpha) {
    return tinycolor({
      h: color.h,
      s: cacl_s,
      l: color.l,
      a: color.a
    }).toHex8()
  }

  return tinycolor({
    h: color.h,
    s: cacl_s,
    l: color.l,
    a: color.a
  }).toHex()
}