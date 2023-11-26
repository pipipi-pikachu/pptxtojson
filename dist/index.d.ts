export interface Shape {
  type: 'shape'
  left: number
  top: number
  width: number
  height: number
  cx: number
  cy: number
  borderColor: string
  borderWidth: number
  borderType: 'solid' | 'dashed' | 'dotted'
  borderStrokeDasharray: string
  fillColor: string
  content: string
  isFlipV: boolean
  isFlipH: boolean
  rotate: number
  shapType: string
  vAlign: string
  path?: string
  id: string
  name: string
}

export interface Text {
  type: 'text'
  left: number
  top: number
  width: number
  height: number
  borderColor: string
  borderWidth: number
  borderType: 'solid' | 'dashed' | 'dotted'
  borderStrokeDasharray: string
  fillColor: string
  isFlipV: boolean
  isFlipH: boolean
  isVertical: boolean
  rotate: number
  content: string
  vAlign: string
  id: string
  name: string
}

export interface Image {
  type: 'image'
  left: number
  top: number
  width: number
  height: number
  src: string
  rotate: number
}

export interface TableCell {
  text: string
  rowSpan?: number
  colSpan?: number
  vMerge?: boolean
  hMerge?: boolean
}
export interface Table {
  type: 'table'
  left: number
  top: number
  width: number
  height: number
  data: TableCell[][]
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
export type ScatterChartData = [number[], number[]]
export interface CommonChart {
  type: 'chart'
  left: number
  top: number
  width: number
  height: number
  data: ChartItem[]
  chartType: ChartType
  barDir?: 'bar' | 'col'
  marker?: boolean
  holeSize?: string
  grouping?: string
  style?: string
}
export interface ScatterChart {
  type: 'chart'
  left: number
  top: number
  width: number
  height: number
  data: ScatterChartData,
  chartType: 'scatterChart'
}
export type Chart = CommonChart | ScatterChart

export interface Video {
  type: 'video'
  left: number
  top: number
  width: number
  height: number
  blob?: string
  src?: string
}

export interface Audio {
  type: 'audio'
  left: number
  top: number
  width: number
  height: number
  blob: string
}

export interface Diagram {
  type: 'diagram'
  left: number
  top: number
  width: number
  height: number
}

export type BaseElement = Shape | Text | Image | Table | Chart | Diagram | Video | Audio

export interface Group {
  type: 'group'
  left: number
  top: number
  width: number
  height: number
  elements: BaseElement[]
}
export type Element = BaseElement | Group

export interface SlideColorFill {
  type: 'color'
  value: string
}

export interface SlideImageFill {
  type: 'image'
  value: {
    picBase64: string
    opacity: number
  }
}

export interface SlideGradientFill {
  type: 'gradient'
  value: {
    rot: number
    colors: {
      pos: string
      color: string
    }[]
  }
}

export type SlideFill = SlideColorFill | SlideImageFill | SlideGradientFill

export interface Slide {
  fill: SlideFill
  elements: Element[]
}

export const parse: (file: ArrayBuffer) => Promise<{
  slides: Slide[]
  size: {
    width: number
    height: number
  }
}>