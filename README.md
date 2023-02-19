# 🎨 PPTX2JSON
这是一个派生于 [PPTX2HTML](https://github.com/g21589/PPTX2HTML) 的工具。可以将 .pptx 文件解析为 JSON 数据。目前还不足以用于生产环境。

在线DEMO：https://pipipi-pikachu.github.io/pptx2json/

相较于原版：
- 使用更现代的语法和依赖重写（原项目年代较久远），方便阅读和理解；
- 删除了所有非核心代码，仅关注 XML 的解析过程；
- 输出 JSON 格式的解析结果；

# 🔨安装
> npm install pptxtojson

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


# 📄 开源协议
GPL-3.0 LICENSE © [pipipi-pikachu](https://github.com/pipipi-pikachu)

仅供学习，禁止商用