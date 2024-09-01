# 🎨 pptxtojson
一个运行在浏览器中，可以将 .pptx 文件转为可读的 JSON 数据的 JavaScript 库。

> 与其他的pptx文件解析工具的最大区别在于：
> 1. 直接运行在浏览器端；
> 2. 解析结果是**可读**的 JSON 数据，而不仅仅是把 XML 文件内容原样翻译成难以理解的 JSON。

在线DEMO：https://pipipi-pikachu.github.io/pptxtojson/

# 🪧 注意事项
### ⚒️ 使用场景
本仓库诞生于项目 [PPTist](https://github.com/pipipi-pikachu/PPTist) ，希望为其“导入 .pptx 文件功能”提供一个参考示例。不过就目前来说，解析出来的PPT信息与源文件在样式上还是存在不少差距，还不足以直接运用到生产环境中。

但如果你只是需要提取PPT文件的文本内容和媒体资源信息，对排版精准度/样式信息没有特别高的要求，那么 pptxtojson 可能会对你有一些帮助。

### 📏 长度值单位
输出的JSON中，所有数值长度值单位都为`pt`（point）
> 注意：在0.x版本中，所有输出的长度值单位都是px（像素）

# 🔨安装
```
npm install pptxtojson
```

# 💿用法
```html
<input type="file" accept="application/vnd.openxmlformats-officedocument.presentationml.presentation"/>
```

```js
import { parse } from 'pptxtojson'

document.querySelector('input').addEventListener('change', evt => {
	const file = evt.target.files[0]
	
	const reader = new FileReader()
	reader.onload = async e => {
		const json = await parse(e.target.result)
		console.log(json)
	}
	reader.readAsArrayBuffer(file)
})
```

```js
// 输出示例
{
	"slides": {
		"fill": {
			"type": "color",
			"value": "#FF0000"
		},
		"elements": [
			{
				"left":	0,
				"top": 0,
				"width": 72,
				"height":	72,
				"borderColor": "#1f4e79",
				"borderWidth": 1,
				"borderType": "solid",
				"borderStrokeDasharray": 0,
				"fillColor": "#5b9bd5",
				"content": "<p style=\"text-align: center;\"><span style=\"font-size: 18pt;font-family: Calibri;\">TEST</span></p>",
				"isFlipV": false,
				"isFlipH": false,
				"rotate": 0,
				"vAlign": "mid",
				"name": "矩形 1",
				"type": "shape",
				"shapType": "rect"
			},
			// more...
		],
	},
	"size": {
		"width": 960,
		"height": 540
	}
}
```

# 📕 功能支持

### 幻灯片尺寸
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| width                  | number                         | 宽度            
| height                 | number                         | 高度  

### 页面背景
| prop                   | type                            | 描述            
|------------------------|---------------------------------|---------------
| type                   | 'color' 丨 'image' 丨 'gradient' | 背景类型            
| value                  | SlideColorFill 丨 SlideImageFill 丨 SlideGradientFill| 背景值  

### 页内元素
#### 文字
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'text'                         | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| borderColor            | string                         | 边框颜色          
| borderWidth            | number                         | 边框宽度          
| borderType             | 'solid' 丨 'dashed' 丨 'dotted' | 边框类型          
| borderStrokeDasharray  | string                         | 非实线边框样式       
| shadow                 | Shadow                         | 阴影            
| fillColor              | string                         | 填充色           
| content                | string                         | 内容文字（HTML富文本） 
| isFlipV                | boolean                        | 垂直翻转          
| isFlipH                | boolean                        | 水平翻转          
| rotate                 | number                         | 旋转角度          
| vAlign                 | string                         | 垂直对齐方向        
| isVertical             | boolean                        | 是否为竖向文本        
| name                   | string                         | 元素名  

#### 图片
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'image'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| src                    | string                         | 图片地址（base64）    
| rotate                 | number                         | 旋转角度  

#### 形状
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'shape'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| borderColor            | string                         | 边框颜色          
| borderWidth            | number                         | 边框宽度          
| borderType             | 'solid' 丨 'dashed' 丨 'dotted' | 边框类型          
| borderStrokeDasharray  | string                         | 非实线边框样式       
| shadow                 | Shadow                         | 阴影            
| fillColor              | string                         | 填充色           
| content                | string                         | 内容文字（HTML富文本） 
| isFlipV                | boolean                        | 垂直翻转          
| isFlipH                | boolean                        | 水平翻转          
| rotate                 | number                         | 旋转角度          
| shapType               | string                         | 形状类型          
| vAlign                 | string                         | 垂直对齐方向        
| path                   | string                         | 路径（仅自定义形状存在）         
| name                   | string                         | 元素名   

#### 表格
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'table'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度               
| borderColor            | string                         | 边框颜色          
| borderWidth            | number                         | 边框宽度          
| borderType             | 'solid' 丨 'dashed' 丨 'dotted' | 边框类型           
| data                   | TableCell[][]                  | 表格数据

#### 图表
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'chart'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| data                   | ChartItem[] 丨 ScatterChartData | 图表数据    
| chartType              | ChartType                      | 图表类型    
| barDir                 | 'bar' 丨 'col'                  | 柱状图方向    
| marker                 | boolean                        | 是否带数据标记    
| holeSize               | string                         | 环形图尺寸    
| grouping               | string                         | 分组模式    
| style                  | string                         | 图表样式 

#### 视频
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'video'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| blob                   | string                         | 视频blob    
| src                    | string                         | 视频src 

#### 音频
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'audio'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| blob                   | string                         | 音频blob   

#### Smart图
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'diagram'                      | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| elements               | (Shape 丨 Text)[]               | 子元素集合  

#### 多元素组合
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'group'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| elements               | Element[]                      | 子元素集合  

### 更多类型请参考 👇
[https://github.com/pipipi-pikachu/pptxtojson/blob/master/dist/index.d.ts](https://github.com/pipipi-pikachu/pptxtojson/blob/master/dist/index.d.ts)

# 🙏 感谢
本仓库大量参考了 [PPTX2HTML](https://github.com/g21589/PPTX2HTML) 和 [PPTXjs](https://github.com/meshesha/PPTXjs) 的实现。
> 与它们不同的是，PPTX2HTML 和 PPTXjs 是将PPT文件转换为能够运行的 HTML 页面，而 pptxtojson 做的是将PPT文件转换为干净的 JSON 数据

# 📄 开源协议
MIT License | Copyright © 2020-PRESENT [pipipi-pikachu](https://github.com/pipipi-pikachu)