# FillFormat 对象

## 简介

代表形状的填充格式。

## 说明

形状可以有纯色、渐变、纹理、图案、图片或半透明填充。

FillFormat 对象的很多属性是只读的。要设置这些属性中每一个，必须使用相应的方法。

使用 Fill 属性可返回 FillFormat 对象。

``` JavaScript
/*下例向 myDocument 中添加矩形并且设置矩形填充的渐变和颜色。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let shapes2 = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill
shapes2.ForeColor.RGB = (0, 128, 128)
shapes2.OneColorGradient(msoGradientHorizontal, 1, 1)
}
```

## 方法

| 序号 | 名称                                             | 说明                                                                                  |
|:----:|:-------------------------------------------------|:--------------------------------------------------------------------------------------|
|  1   | [OneColorGradient](#FillFormat.OneColorGradient) | 将指定填充设为单色渐变。                                                              |
|  2   | [Patterned](#FillFormat.Patterned)               | 将指定填充设为一种图案。                                                              |
|  3   | [PresetGradient](#FillFormat.PresetGradient)     | 将指定填充设为一个预设的渐变。                                                        |
|  4   | [PresetTextured](#FillFormat.PresetTextured)     | 将指定填充格式设为预设纹理。                                                          |
|  5   | [Solid](#FillFormat.Solid)                       | 将指定填充设为统一的颜色。使用此方法可将渐变、纹理、图案或背景填充转换为纯色填充。    |
|  6   | [TwoColorGradient](#FillFormat.TwoColorGradient) | 将指定填充设为双色渐变。                                                              |
|  7   | [UserPicture](#FillFormat.UserPicture)           | 用图像填充指定的形状。                                                                |
|  8   | [UserTextured](#FillFormat.UserTextured)         | 用平铺的小图像填充指定形状。如果要用一个大图像来填充该形状，请使用 UserPicture 方法。 |

## 属性

| 序号 | 名称                                                         | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FillFormat.Application)                       | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [BackColor](#FillFormat.BackColor)                           | 返回或设置一个 ColorFormat 对象，它代表指定的填充背景色。                                                                                                                                                                       |
|  3   | [Creator](#FillFormat.Creator)                               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [ForeColor](#FillFormat.ForeColor)                           | 返回或设置一个 ColorFormat 对象，它代表指定的前景填充色或纯色。                                                                                                                                                                 |
|  5   | [GradientAngle](#FillFormat.GradientAngle)                   | 返回或设置指定填充格式的渐变填充角度。可读/写。                                                                                                                                                                                 |
|  6   | [GradientColorType](#FillFormat.GradientColorType)           | 返回指定填充的渐变颜色类型。 MsoGradientColorType ，只读。                                                                                                                                                                      |
|  7   | [GradientDegree](#FillFormat.GradientDegree)                 | 以浮点数值的方式返回单色阴影填充的渐变程度，数值大小介于 0.0（暗）到 1.0（亮）之间。 Single 型，只读。                                                                                                                          |
|  8   | [GradientStops](#FillFormat.GradientStops)                   | 返回渐变填充的终点。只读。                                                                                                                                                                                                      |
|  9   | [GradientStyle](#FillFormat.GradientStyle)                   | 返回指定填充的渐变样式。 MsoGradientStyle ，只读。                                                                                                                                                                              |
|  10  | [GradientVariant](#FillFormat.GradientVariant)               | 以整数值（1 至 4）形式返回指定填充的底纹变量。此属性的值对应于 “填充效果” 对话框中 “渐变” 选项卡的渐变变量（从左到右、从上到下编号）。 Long 型，只读。                                                                          |
|  11  | [Parent](#FillFormat.Parent)                                 | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  12  | [Pattern](#FillFormat.Pattern)                               | 返回或设置一个代表填充图案的 MsoPatternType 值。                                                                                                                                                                                |
|  13  | [PictureEffects](#FillFormat.PictureEffects)                 | 返回一个对象，该对象表示指定的填充格式的图片或纹理填充。只读                                                                                                                                                                    |
|  14  | [PresetGradientType](#FillFormat.PresetGradientType)         | 返回指定填充的预设渐变类型。 MsoPresetGradientType ，只读。                                                                                                                                                                     |
|  15  | [PresetTexture](#FillFormat.PresetTexture)                   | 返回指定填充的预设纹理。 MsoPresetTexture ，只读。                                                                                                                                                                              |
|  16  | [RotateWithObject](#FillFormat.RotateWithObject)             | 返回或设置是否随对象旋转。可读/写 MsoTriState 类型。                                                                                                                                                                            |
|  17  | [TextureAlignment](#FillFormat.TextureAlignment)             | 返回或设置指定 FillFormat 对象的文本对齐方式。可读/写。                                                                                                                                                                         |
|  18  | [TextureHorizontalScale](#FillFormat.TextureHorizontalScale) | 返回或设置 FillFormat 对象的横向文本缩放值。可读/写 Single 类型。                                                                                                                                                               |
|  19  | [TextureName](#FillFormat.TextureName)                       | 返回指定填充的自定义纹理文件名称。 String 型，只读。                                                                                                                                                                            |
|  20  | [TextureOffsetX](#FillFormat.TextureOffsetX)                 | 返回指定填充的偏移 X 值。可读/写 Single 类型。                                                                                                                                                                                  |
|  21  | [TextureOffsetY](#FillFormat.TextureOffsetY)                 | 返回指定填充的偏移 Y 值。可读/写 Single 类型。                                                                                                                                                                                  |
|  22  | [TextureTile](#FillFormat.TextureTile)                       | 返回指定填充的纹理平铺样式。可读/写 MsoTriState 类型。                                                                                                                                                                          |
|  23  | [TextureType](#FillFormat.TextureType)                       | 返回指定填充的纹理类型。 MsoTextureType ，只读。                                                                                                                                                                                |
|  24  | [TextureVerticalScale](#FillFormat.TextureVerticalScale)     | 返回指定填充的纹理垂直缩放。可读/写 Single 类型。                                                                                                                                                                               |
|  25  | [Transparency](#FillFormat.Transparency)                     | 返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。 Double 型，可读写。                                                                                                                                    |
|  26  | [Type](#FillFormat.Type)                                     | 返回代表填充类型的 MsoFillType 值。                                                                                                                                                                                             |
|  27  | [Visible](#FillFormat.Visible)                               | 返回或设置一个 MsoTriState 值，它确定对象是否可见。可读写。                                                                                                                                                                     |

## 成员方法

### FillFormat.OneColorGradient

将指定填充设为单色渐变。

**语法**

**express.OneColorGradient ( *Style* , *Variant* , *Degree* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型         | 说明                                                                                                                                                                          |
|-----------|-----------|------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Style*   | 必选      | MsoGradientStyle | 渐变样式。                                                                                                                                                                    |
| *Variant* | 必选      | Integer          | 渐变变量。取值范围为 1 到 4 之间，分别与“填充效果”对话框中“渐变”选项卡上的四个渐变变量相对应。如果 GradientStyle 设为 msoGradientFromCenter，则 Variant 参数只能设为 1 或 2。 |
| *Degree*  | 必选      | Single           | 渐变程度。可以为 0.0（暗）到 1.0（亮）之间的值。                                                                                                                              |

### FillFormat.Patterned

将指定填充设为一种图案。

**语法**

**express.Patterned ( *Pattern* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型       | 说明         |
|-----------|-----------|----------------|--------------|
| *Pattern* | 必选      | MsoPatternType | 图案的类型。 |

### FillFormat.PresetGradient

将指定填充设为一个预设的渐变。

**语法**

**express.PresetGradient ( *Style* , *Variant* , *PresetGradientType* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型              | 说明                                                                                                                                                                          |
|----------------------|-----------|-----------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Style*              | 必选      | MsoGradientStyle      | 渐变样式。                                                                                                                                                                    |
| *Variant*            | 必选      | Integer               | 渐变变量。取值范围为 1 到 4 之间，分别与“填充效果”对话框中“渐变”选项卡上的四个渐变变量相对应。如果 GradientStyle 设为 msoGradientFromCenter，则 Variant 参数只能设为 1 或 2。 |
| *PresetGradientType* | 必选      | MsoPresetGradientType | 预设渐变类型。                                                                                                                                                                |

### FillFormat.PresetTextured

将指定填充格式设为预设纹理。

**语法**

**express.PresetTextured ( *PresetTexture* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型         | 说明               |
|-----------------|-----------|------------------|--------------------|
| *PresetTexture* | 必选      | MsoPresetTexture | 要应用的纹理类型。 |

**示例**

``` JavaScript
/*此示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillTextured) {
    let c2f = Charts.Item(2).ChartArea.Fill
    c2f.Visible = true
    if(c1f.TextureType == msoTexturePreset) {
        c2f.PresetTextured(c1f.PresetTexture)
    }
    else {
         c2f.UserTextured(c1f.TextureName)
    }
}
}
```

### FillFormat.Solid

将指定填充设为统一的颜色。使用此方法可将渐变、纹理、图案或背景填充转换为纯色填充。

**语法**

**express.Solid ()**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*此示例将 myDocument 中的所有填充转换为统一的红色填充。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
for(let s = 1; s <= myDocument.Shapes.Count; s++) {
    let s2 = myDocument.Shapes.Item(s).Fill
        s2.Solid()
        s2.ForeColor.RGB = (255, 0, 0)
}
}
```

### FillFormat.TwoColorGradient

将指定填充设为双色渐变。

**语法**

**express.TwoColorGradient ( *Style* , *Variant* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型         | 说明                                                                                                                                                                  |
|-----------|-----------|------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Style*   | 必选      | MsoGradientStyle | 渐变样式。                                                                                                                                                            |
| *Variant* | 必选      | Integer          | 渐变变量。取值范围为 1 到 4 之间，分别与“填充效果”对话框中“渐变”选项卡上的四个渐变变量相对应。如果 Style 设为 msoGradientFromCenter，则 Variant 参数只能设为 1 或 2。 |

### FillFormat.UserPicture

用图像填充指定的形状。

**语法**

**express.UserPicture ( *PictureFile* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明         |
|---------------|-----------|----------|--------------|
| *PictureFile* | 必选      | String   | 图片文件名。 |

### FillFormat.UserTextured

用平铺的小图像填充指定形状。如果要用一个大图像来填充该形状，请使用 **UserPicture** 方法。

**语法**

**express.UserTextured ( *TextureFile* )**

\*express是一个代表 FillFormat 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明         |
|---------------|-----------|----------|--------------|
| *TextureFile* | 必选      | String   | 图片文件名。 |

**示例**

``` JavaScript
/*本示例为第二个图表设置填充格式。*/
Application.Charts.Item(2).ChartArea.Fill.UserTextured("brick.gif")
```

## 成员属性

### FillFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 FillFormat 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### FillFormat.BackColor

返回或设置一个 **ColorFormat** 对象，它代表指定的填充背景色。

**语法**

**express.BackColor**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FillFormat 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### FillFormat.ForeColor

返回或设置一个 **ColorFormat** 对象，它代表指定的前景填充色或纯色。

**语法**

**express.ForeColor**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.GradientAngle

返回或设置指定填充格式的渐变填充角度。可读/写。

**语法**

**express.GradientAngle**

\*express是一个代表 FillFormat 对象的变量。

**说明**

可以在图表中各种元素（形状）的格式设置中指定渐变填充。例如，可使用 **“设置数据系列格式”** 对话框将 **“列”** 图表中各列的格式设置为渐变填充。在此情况下， **GradientAngle** 属性对应于 **“设置数据系列格式”** 对话框内 **“填充”** 类别中的 **“角度”** 框的设置。 **GradientAngle** 属性值的有效范围是 0 到 359.9。

**示例**

``` JavaScript
/*下面的代码示例将 Chart 1 中第 1 个系列的渐变填充角度设置为 45 度。*/
function test(){
Application.ActiveSheet.ChartObjects("Chart 1").Activate()
Application.ActiveChart.SeriesCollection(1).Format.Fill.GradientAngle = 45
}
```

### FillFormat.GradientColorType

返回指定填充的渐变颜色类型。 **MsoGradientColorType** ，只读。

**语法**

**express.GradientColorType**

\*express是一个代表 FillFormat 对象的变量。

**说明**

**MsoGradientColorMixed** 是一个返回值，它仅指明指定区域中其他状态的组合。使用 **OneColorGradient** 、 **PresetGradient** 或 **TwoColorGradient** 方法可设置填充的渐变类型。

**示例**

``` JavaScript
/*此示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillGradient && c1f.GradientColorType == msoGradientOneColor) {
    let c2f = Charts.Item(2).ChartArea.Fill
    c2f.Visible = true
    c2f.OneColorGradient(c1f.GradientStyle, c1f.GradientVariant, c1f.GradientDegree)
}
}
```

### FillFormat.GradientDegree

以浮点数值的方式返回单色阴影填充的渐变程度，数值大小介于 0.0（暗）到 1.0（亮）之间。 **Single** 型，只读。

**语法**

**express.GradientDegree**

\*express是一个代表 FillFormat 对象的变量。

**说明**

此属性为只读。使用 **OneColorGradient** 方法可设置填充的渐变程度。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillGradient && c1f.GradientColorType == msoGradientOneColor) {
    let c2f = Charts.Item(2).ChartArea.Fill
    c2f.Visible = true
    c2f.OneColorGradient(c1f.GradientStyle, c1f.GradientVariant, c1f.GradientDegree)
}
}
```

### FillFormat.GradientStops

返回渐变填充的终点。只读。

**语法**

**express.GradientStops**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.GradientStyle

返回指定填充的渐变样式。 **MsoGradientStyle** ，只读。

**语法**

**express.GradientStyle**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.GradientVariant

以整数值（1 至 4）形式返回指定填充的底纹变量。此属性的值对应于 **“填充效果”** 对话框中 **“渐变”** 选项卡的渐变变量（从左到右、从上到下编号）。 **Long** 型，只读。

**语法**

**express.GradientVariant**

\*express是一个代表 FillFormat 对象的变量。

**说明**

此属性为只读。使用 **OneColorGradient** 或 **TwoColorGradient** 方法可设置填充的渐变变量。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillGradient && c1f.GradientColorType == msoGradientOneColor) {
    let c2f = Charts.Item(2).ChartArea.Fill
    c2f.Visible = true
    c2f.OneColorGradient(c1f.GradientStyle, c1f.GradientVariant, c1f.GradientDegree)
}
}
```

### FillFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.Pattern

返回或设置一个代表填充图案的 **MsoPatternType** 值。

**语法**

**express.Pattern**

\*express是一个代表 FillFormat 对象的变量。

**说明**

使用 **Patterned** 方法可设置填充的图案类型。

### FillFormat.PictureEffects

返回一个对象，该对象表示指定的填充格式的图片或纹理填充。只读

**语法**

**express.PictureEffects**

\*express是一个代表 FillFormat 对象的变量。

**说明**

可以在图表中各种元素（形状）的格式设置中指定图片或纹理填充。例如，可使用 **“设置数据系列格式”** 对话框将 **“列”** 图表中各列的格式设置为图片或纹理填充。在此情况下， **PictureEffects** 属性返回的 **PictureEffects** 集合对应于 **“设置数据系列格式”** 对话框的 **“填充”** 类别中 **“图片或纹理填充”** 选项的关联设置。

### FillFormat.PresetGradientType

返回指定填充的预设渐变类型。 **MsoPresetGradientType** ，只读。

**语法**

**express.PresetGradientType**

\*express是一个代表 FillFormat 对象的变量。

**说明**

使用 **PresetGradient** 方法可设置填充的预设渐变类型。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillGradient) {
    let c2f = Charts.Item(2).ChartArea.Fill
    c2f.Visible = true
    c2f.PresetGradient(c1f.GradientStyle, c1f.GradientVariant, c1f.PresetGradientType)
}
}
```

### FillFormat.PresetTexture

返回指定填充的预设纹理。 **MsoPresetTexture** ，只读。

**语法**

**express.PresetTexture**

\*express是一个代表 FillFormat 对象的变量。

**说明**

使用 **PresetTextured** 方法可设置填充的预设纹理。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillTextured) {
    let c2f = Charts.Item(2).ChartArea.Fill
        c2f.Visible = true
        if(c1f.TextureType == msoTexturePreset) {
            c2f.PresetTextured(c1f.PresetTexture)
        }
        else {
            c2f.UserTextured(c1f.TextureName)
        }
}
}
```

### FillFormat.RotateWithObject

返回或设置是否随对象旋转。可读/写 **MsoTriState** 类型。

**语法**

**express.RotateWithObject**

\*express是一个代表 FillFormat 对象的变量。

**说明**

|                                                   |
|---------------------------------------------------|
| MsoTriState 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue**                                      |
| **msoFalse**                                      |
| **msoTriStateMixed**                              |
| **msoTriStateToggle**                             |
| **msoTrue**                                       |

### FillFormat.TextureAlignment

返回或设置指定 **FillFormat** 对象的文本对齐方式。可读/写。

**语法**

**express.TextureAlignment**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.TextureHorizontalScale

返回或设置 **FillFormat** 对象的横向文本缩放值。可读/写 **Single** 类型。

**语法**

**express.TextureHorizontalScale**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.TextureName

返回指定填充的自定义纹理文件名称。 **String** 型，只读。

**语法**

**express.TextureName**

\*express是一个代表 FillFormat 对象的变量。

**说明**

使用 **UserPicture** 或 **UserTextured** 方法可设置填充的纹理文件。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillTextured) {
    let c2f = Charts.Item(2).ChartArea.Fill
        c2f.Visible = true
        if(c1f.TextureType == msoTexturePreset) {
            c2f.PresetTextured(c1f.PresetTexture)
        }
        else {
            c2f.UserTextured(c1f.TextureName)
        }
}
}
```

### FillFormat.TextureOffsetX

返回指定填充的偏移 X 值。可读/写 **Single** 类型。

**语法**

**express.TextureOffsetX**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.TextureOffsetY

返回指定填充的偏移 Y 值。可读/写 **Single** 类型。

**语法**

**express.TextureOffsetY**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.TextureTile

返回指定填充的纹理平铺样式。可读/写 **MsoTriState** 类型。

**语法**

**express.TextureTile**

\*express是一个代表 FillFormat 对象的变量。

**说明**

|                                                   |
|---------------------------------------------------|
| MsoTriState 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue**                                      |
| **msoFalse**                                      |
| **msoTriStateMixed**                              |
| **msoTriStateToggle**                             |
| **msoTrue**                                       |

### FillFormat.TextureType

返回指定填充的纹理类型。 **MsoTextureType** ，只读。

**语法**

**express.TextureType**

\*express是一个代表 FillFormat 对象的变量。

**说明**

使用 **UserTextured** 方法可设置填充的纹理类型。

**示例**

``` JavaScript
/*本示例以第一个图表中使用的填充样式对第二个图表的填充格式进行设置。*/
function test(){
let c1f = Application.Charts.Item(1).ChartArea.Fill
if(c1f.Type == msoFillTextured) {
    let c2f = Charts.Item(2).ChartArea.Fill
        c2f.Visible = true
        if(c1f.TextureType == msoTexturePreset) {
            c2f.PresetTextured(c1f.PresetTexture)
        }
        else {
            c2f.UserTextured(c1f.TextureName)
        }
}
}
```

### FillFormat.TextureVerticalScale

返回指定填充的纹理垂直缩放。可读/写 **Single** 类型。

**语法**

**express.TextureVerticalScale**

\*express是一个代表 FillFormat 对象的变量。

### FillFormat.Transparency

返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。 **Double** 型，可读写。

**语法**

**express.Transparency**

\*express是一个代表 FillFormat 对象的变量。

**说明**

此属性的值仅影响均一颜色填充和线条的外观。对图案线条及渐变、图案、图片和纹理填充的外观无效。

### FillFormat.Type

返回代表填充类型的 **MsoFillType** 值。

**语法**

**express.Type**

\*express是一个代表 FillFormat 对象的变量。

**说明**

此属性也可返回 **xlAutomatic** 或 **xlNone** 。

在 ET 中不使用 **msoFillBackground** 常量。

### FillFormat.Visible

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 FillFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/FillFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
