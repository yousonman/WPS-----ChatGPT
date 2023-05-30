# ColorFormat 对象

## 简介

代表单色对象的颜色、带有渐变或图案填充的对象的前景色或背景色，或者指针的颜色。可以将颜色设为明确的红-绿-蓝值（使用 RGB 属性）或设为配色方案中的一种颜色（使用 SchemeColor 属性）。

## 说明

使用下表中列出的属性之一返回 ColorFormat 对象。

| 使用此属性     | 对此对象          | 返回一个代表以下颜色的 ColorFormat 对象      |
|----------------|-------------------|----------------------------------------------|
| DimColor       | AnimationSettings | 变暗对象使用的颜色                           |
| BackColor      | FillFormat        | 背景填充颜色（在底纹或带图案的填充中使用）   |
| ForeColor      | FillFormat        | 前景填充颜色（对于纯色填充，则代表填充颜色） |
| Color          | Font              | 项目符号或字符颜色                           |
| BackColor      | LineFormat        | 背景线条颜色（在带图案线条中使用）           |
| ForeColor      | LineFormat        | 前景线条颜色（或只是实线的线条颜色）         |
| ForeColor      | ShadowFormat      | 阴影颜色                                     |
| PointerColor   | SlideShowSettings | 演示文稿的默认指针颜色                       |
| PointerColor   | SlideShowView     | 幻灯片放映视图中的临时指针颜色               |
| ExtrusionColor | ThreeDFormat      | 延伸对象的侧面颜色                           |

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>Application</span> 对象，该对象表示指定对象的创建者。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.ObjectThemeColor">ObjectThemeColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定 ColorFormat 对象的主题颜色。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定对象的父对象。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.RGB">RGB</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定颜色的 红-绿-蓝(RGB)值（RGB 值：RGB 函数的返回值；以红、绿、蓝的数值组合来指定颜色，其数值范围是由 0（零）到 255 的整数。） Long 类型，可读/写 。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.SchemeColor">SchemeColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置已经应用的配色方案中的颜色，该配色方案与指定对象相关联。可读/写 PpColorSchemeIndex 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.TintAndShade">TintAndShade</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Single 类型的值，该值代表使指定形状的颜色是变亮，还是变暗。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.Type">Type</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 MsoColorType 常量，该常量代表颜色的类型。只读。</p></td>
</tr>
</tbody>
</table>

## 成员属性

### ColorFormat.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 ColorFormat 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### ColorFormat.Creator

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

``` JavaScript
/*以下示例显示关于 myObject 创建者的消息。*/
function test(){
　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
　　　　if(myObject.Creator == parseInt(50575054,16)) {
　　　　    alert("This is a WPP object")
　　　　}
　　　　else {
　　　　    alert("This is not a WPP object")
　　　　}
}
```

### ColorFormat.ObjectThemeColor

返回或设置指定 ColorFormat 对象的主题颜色。可读/写。

**语法**

**express.ObjectThemeColor**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

**ObjectThemeColor** 属性的值可以是下列 常量之一。

**示例**

下例显示了如何使用 **ObjectThemeColor** 属性获取活动演示文稿中第一张幻灯片上第一个形状的前景填充的主题颜色。

``` JavaScript
function ObjectThemeColor_Example() {
    Debug.Print(ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.ForeColor.ObjectThemeColor)  
}
```

### ColorFormat.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 ColorFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test() {
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

### ColorFormat.RGB

返回或设置指定颜色的 红-绿-蓝(RGB)值（RGB 值：RGB 函数的返回值；以红、绿、蓝的数值组合来指定颜色，其数值范围是由 0（零）到 255 的整数。） **Long** 类型，可读/写 。

**语法**

**express.RGB**

\*express是一个代表 ColorFormat 对象的变量。

**示例**

``` JavaScript
/*本示例为活动演示文稿的第三个配色方案设定背景色，并将配色方案应用于基于幻灯片母版的演示文稿中所有的幻灯片。*/
function test(){
　　　　let a = Application.ActivePresentation
　　　　let cs1 = a.ColorSchemes.Item(3)
　　　　cs1.Colors(ppBackground).RGB = 128, 128, 0
　　　　a.SlideMaster.ColorScheme = cs1
}
```

### ColorFormat.SchemeColor

返回或设置已经应用的配色方案中的颜色，该配色方案与指定对象相关联。可读/写 **PpColorSchemeIndex** 类型。

**语法**

**express.SchemeColor**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

| PpColorSchemeIndex 可以是下列 PpColorSchemeIndex 常量之一。 |
|-------------------------------------------------------------|
| ppAccent1                                                   |
| ppAccent2                                                   |
| ppAccent3                                                   |
| ppBackground                                                |
| ppFill                                                      |
| ppForeground                                                |
| ppNotSchemeColor                                            |
| ppSchemeColorMixed                                          |
| ppShadow                                                    |
| ppTitle                                                     |

**示例**

``` JavaScript
/*本示例切换活动演示文稿中第一张幻灯片的两种背景色，一种是明确的红-绿-蓝值所定义的颜色，另一种是配色方案的背景色。*/
function test(){
　　　　let slides = Application.ActivePresentation.Slides.Item(1)
　　　　slides.FollowMasterBackground = false
　　　　let color = slides.Background.Fill.ForeColor
    if(color.Type == msoColorTypeScheme) {
　　　　    color.RGB = 0, 128, 128
　　　　}
　　　　else {
　　　　    color.SchemeColor = ppBackground
　　　　}
}
```

### ColorFormat.TintAndShade

返回一个 **Single** 类型的值，该值代表使指定形状的颜色是变亮，还是变暗。可读/写。

**语法**

**express.TintAndShade**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

**TintAndShade** 属性值的输入范围介于 -1（最暗）到 1（最亮）之间，0（零）为适中。

**示例**

本示例在活动文档中新建一个形状，并设置填充颜色，然后使彩色底纹变亮。

``` JavaScript
function PrinterPlate() {
    let shpHeart = Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShapeHeart, 150, 150, 250, 250)
    let color = shpHeart.Fill.ForeColor
    color.CMYK = 16111872
    color.TintAndShade = 0.3
    color.OverPrint = msoTrue
}
```

### ColorFormat.Type

table { word-break:break-all; }

返回一个 **MsoColorType** 常量，该常量代表颜色的类型。只读。

**语法**

**express.Type**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

| MsoColorType 可以是下列 MsoColorType 常量之一。 |
|-------------------------------------------------|
| msoColorTypeCMS                                 |
| msoColorTypeCMYK                                |
| msoColorTypeInk                                 |
| msoColorTypeMixed                               |
| msoColorTypeRGB                                 |
| msoColorTypeScheme                              |

> WPS 开发文档： [WPS 基础接口/演示 API 参考/ColorFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
