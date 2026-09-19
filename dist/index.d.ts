export interface Shadow {
  h: number
  v: number
  blur: number
  color: string
}

export interface ColorFill {
  type: 'color'
  value: string
}

export interface ImageFill {
  type: 'image'
  value: {
    ref: string
    base64: string
    blob: string
    opacity: number
  }
}

export interface GradientFill {
  type: 'gradient'
  value: {
    path: 'line' | 'circle' | 'rect' | 'shape'
    rot: number
    colors: {
      pos: string
      color: string
    }[]
  }
}

export interface PatternFill {
  type: 'pattern'
  value: {
    type: string
    foregroundColor: string
    backgroundColor: string
  }
}

export type Fill = ColorFill | ImageFill | GradientFill | PatternFill

export interface Border {
  borderColor: string
  borderWidth: number
  borderType:'solid' | 'dashed' | 'dotted'
}

export interface AutoFit {
  type: 'shape' | 'text'
  fontScale?: number
}

export interface TextInset {
  l: number
  t: number
  r: number
  b: number
}

export interface PathViewBox {
  x: number
  y: number
  width: number
  height: number
}

export interface LineEnd {
  type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow'
  width?: 'sm' | 'med' | 'lg'
  length?: 'sm' | 'med' | 'lg'
}

export interface Shape {
  type: 'shape'
  id: string
  left: number
  top: number
  width: number
  height: number
  borderColor: string
  borderWidth: number
  borderType: 'solid' | 'dashed' | 'dotted'
  borderStrokeDasharray: string
  shadow?: Shadow
  fill: Fill
  content: string
  isFlipV: boolean
  isFlipH: boolean
  rotate: number
  shapType: string
  vAlign: string
  wrap: boolean
  path?: string
  pathViewBox?: PathViewBox
  headEnd?: LineEnd
  tailEnd?: LineEnd
  strokeOnly?: boolean
  keypoints?: Record<string, number>
  name: string
  order: number
  autoFit?: AutoFit
  textInset?: TextInset
  link?: string
}

export interface Text {
  type: 'text'
  id: string
  left: number
  top: number
  width: number
  height: number
  borderColor: string
  borderWidth: number
  borderType: 'solid' | 'dashed' | 'dotted'
  borderStrokeDasharray: string
  shadow?: Shadow
  fill: Fill
  isFlipV: boolean
  isFlipH: boolean
  isVertical: boolean
  rotate: number
  content: string
  vAlign: string
  wrap: boolean
  name: string
  order: number
  autoFit?: AutoFit
  textInset?: TextInset
  link?: string
}

export interface Image {
  type: 'image'
  id: string
  left: number
  top: number
  width: number
  height: number
  ref: string
  base64: string
  blob: string
  rotate: number
  isFlipH: boolean
  isFlipV: boolean
  order: number
  rect?: {
    t?: number
    b?: number
    l?: number
    r?: number
  }
  geom: string
  borderColor: string
  borderWidth: number
  borderType: 'solid' | 'dashed' | 'dotted'
  borderStrokeDasharray: string
  filters?: {
    sharpen?: number
    colorTemperature?: number
    saturation?: number
    brightness?: number
    contrast?: number
  }
  link?: string
}

export interface TableCell {
  text: string
  rowSpan?: number
  colSpan?: number
  vMerge?: number
  hMerge?: number
  fillColor?: string
  fontColor?: string
  fontBold?: boolean
  vAlign: string
  borders: {
    top?: Border
    bottom?: Border
    left?: Border
    right?: Border
  }
}
export interface Table {
  type: 'table'
  id: string
  left: number
  top: number
  width: number
  height: number
  data: TableCell[][]
  borders: {
    top?: Border
    bottom?: Border
    left?: Border
    right?: Border
  }
  order: number
  rowHeights: number[]
  colWidths: number[]
}

export type ChartType = 'lineChart' |
  'line3DChart' |
  'barChart' |
  'bar3DChart' |
  'pieChart' |
  'pie3DChart' |
  'doughnutChart' |
  'areaChart' |
  'area3DChart' |
  'scatterChart' |
  'bubbleChart' |
  'radarChart' |
  'surfaceChart' |
  'surface3DChart' |
  'stockChart'

export interface ChartValue {
  x: string
  y: number
}
export interface ChartXLabel {
  [key: string]: string
}
export interface ChartItem {
  key: string
  values: ChartValue[]
  xlabels: ChartXLabel
}
export type ScatterChartData = number[][]
export interface CommonChart {
  type: 'chart'
  id: string
  left: number
  top: number
  width: number
  height: number
  data: ChartItem[]
  colors: string[]
  chartType: Exclude<ChartType, 'scatterChart' | 'bubbleChart'>
  barDir?: 'bar' | 'col'
  marker?: boolean
  holeSize?: string
  grouping?: string
  style?: string
  order: number
}
export interface ScatterChart {
  type: 'chart'
  id: string
  left: number
  top: number
  width: number
  height: number
  data: ScatterChartData
  colors: string[]
  chartType: 'scatterChart' | 'bubbleChart'
  order: number
}
export type Chart = CommonChart | ScatterChart

export interface Video {
  type: 'video'
  id: string
  left: number
  top: number
  width: number
  height: number
  ref: string
  blob: string
  order: number
}

export interface Audio {
  type: 'audio'
  id: string
  left: number
  top: number
  width: number
  height: number
  ref: string
  blob: string
  order: number
}

export interface Diagram {
  type: 'diagram'
  id: string
  left: number
  top: number
  width: number
  height: number
  elements: (Shape | Text)[]
  textList: string[]
  order: number
}

export interface Math {
  type: 'math'
  id: string
  left: number
  top: number
  width: number
  height: number
  latex: string
  picRef: string
  picBase64: string
  picBlob: string
  order: number
  text?: string
}

export type BaseElement = Shape | Text | Image | Table | Chart | Video | Audio | Diagram | Math

export interface Group {
  type: 'group'
  id: string
  left: number
  top: number
  width: number
  height: number
  rotate: number
  elements: BaseElement[]
  order: number
  isFlipH: boolean
  isFlipV: boolean
}
export type Element = BaseElement | Group

export interface SlideTransition {
  type: string
  duration: number
  direction: string | null
}

export interface Slide {
  fill: Fill
  elements: Element[]
  layoutElements: Element[]
  note: string
  transition?: SlideTransition | null
}

export interface Options {
  imageMode?: 'base64' | 'blob' | 'both' | 'none'
  videoMode?: 'blob' | 'none'
  audioMode?: 'blob' | 'none'
  singleLineSpacingFactor?: number
}

export const parse: (file: ArrayBuffer, options?: Options) => Promise<{
  slides: Slide[]
  themeColors: string[]
  usedFonts: string[]
  size: {
    width: number
    height: number
  }
}>
