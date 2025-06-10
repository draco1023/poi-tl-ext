# poi-tl-ext
![GitHub](https://img.shields.io/github/license/draco1023/poi-tl-ext) ![JDK](https://img.shields.io/badge/jdk-1.8-blue)

# Maven

poi 4.x poi-tl 1.11 以前的版本

```xml
<dependency>
    <groupId>io.github.draco1023</groupId>
    <artifactId>poi-tl-ext</artifactId>
    <version>0.4.24</version>
</dependency>
```

poi 5.x poi-tl 1.11.0+

```xml
<dependency>
    <groupId>io.github.draco1023</groupId>
    <artifactId>poi-tl-ext</artifactId>
    <version>0.4.24-poi5</version>
</dependency>
```

# 扩展功能

在 [poi-tl](https://github.com/Sayi/poi-tl) 的基础上扩展了如下功能：

- 支持渲染`HTML`字符串，插件`HtmlRenderPolicy`的使用方法如下（也可参考[文档](http://deepoove.com/poi-tl/#_%E4%BD%BF%E7%94%A8%E6%8F%92%E4%BB%B6)）

  ```java
  HtmlRenderPolicy htmlRenderPolicy = new HtmlRenderPolicy();
  Configure configure = Configure.builder()
          .bind("key", htmlRenderPolicy)
          .build();
  Map<String, Object> data = new HashMap<>();
  data.put("key", "<p>Hello <b>world</b>!</p>");
  XWPFTemplate.compile("input.docx", configure).render(data).writeToFile("output.docx");
  ```
  
  `HtmlRenderPolicy`可以通过`HtmlRenderConfig`进行如下设置：
  - `globalFont` 全局默认字体（用于归一化处理，而不是用于样式兜底）
  - `globalFontSize` 全局默认字号（用于归一化处理，而不是用于样式兜底）
  - `showDefaultTableBorderInTableCell` 是否显示嵌套表格的边框（`poi`生成嵌套表格时默认不显示边框，见[#12](https://github.com/draco1023/poi-tl-ext/issues/12)）
  - `numberingIndent` 多级列表项缩进长度，默认值360
  - `numberingHanging` 列表项悬挂长度，默认值360，CSS样式`list-style-position`为`inside`时该参数无效
  - `numberingSpacing` 列表编号与内容之间的间隔类型，`STLevelSuffix.NOTHING`/`STLevelSuffix.SPACE`/`STLevelSuffix.TAB`
  
  自定义`<latex>`标签，允许渲染嵌入在`HTML`中的`LaTeX`，字符串格式可参考[文档](https://www2.ph.ed.ac.uk/snuggletex/documentation/supported-latex.html)。
  
  _目前实现了富文本编辑器可实现的大部分效果，后续继续改进..._

- 支持渲染`MathML`字符串，插件类为`MathMLRenderPolicy`
- 支持渲染`LaTeX`字符串，插件类为`LaTeXRenderPolicy`

## 支持我
如果您觉得这个插件节省了您的时间和精力，或者解决了您的难题，可以考虑支持一下我的工作，感谢！ ⚡⚡⚡

![wechat_sp](https://pub-1085551b511e4719a277177fe3c8f95b.r2.dev/wechat_sp.jpg)