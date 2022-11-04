这是一个派生于 [PPTX2HTML](https://github.com/g21589/PPTX2HTML) 的案例。主要用于实验和学习 PPTX 文件到 JSON 数据的转换过程，目前还远不足以用于生产环境。

在线 demo 演示（上传 PPTX 文件后打开控制台可以看到 JSON 数据）：https://pipipi-pikachu.github.io/pptx2json/

相较于原版：
- 使用更现代的语法重写（原项目年代较久远），方便阅读和理解；
- 删除了所有非核心代码，仅关注 XML 的解析过程；
- 输出 JSON 格式的解析结果；

注：该项目支持的形状非常少，有需要的可以参考 [PPTXjs](https://github.com/meshesha/PPTXjs) 这个项目进行扩充。 PPTXjs 本身也是参考了 PPTX2HTML 的一个项目，但是更强大更完善，特别是形状部分。