# CommandEffect 对象

## 简介

代表动画动作的命令效果。可使用此对象向嵌入对象发送事件、调用函数以及发送 OLE 谓词。

## 说明

使用 AnimationBehavior 对象的 CommandEffect 属性返回 CommandEffect 对象。使用 CommandEffect 对象的 Command 和 Type 属性可以更改命令效果。

## 属性

| 序号 | 名称                                      | 说明                                                                        |
|:----:|:------------------------------------------|:----------------------------------------------------------------------------|
|  1   | [Application](#CommandEffect.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。                     |
|  2   | [Command](#CommandEffect.Command)         | 设置或返回一个 String 类型值，该值代表为获得命令效果而执行的命令。可读/写。 |
|  3   | [Parent](#CommandEffect.Parent)           | 返回指定对象的父对象。                                                      |
|  4   | [Type](#CommandEffect.Type)               | 返回或设置一个 MsoAnimType 常量，该常量代表动画的类型。可读/写。            |

## 成员属性

### CommandEffect.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 CommandEffect 对象的变量。

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

### CommandEffect.Command

设置或返回一个 **String** 类型值，该值代表为获得命令效果而执行的命令。可读/写。

**语法**

**express.Command**

\*express是一个代表 CommandEffect 对象的变量。

**说明**

可以使用该属性向嵌入对象发送 OLE 动作。

如果形状是 OLE 对象，则 OLE 对象将在识别该动作的情况下执行相应的命令。

如果形状是媒体对象（声音/视频），WPP就会识别下面的谓词：play、stop、pause、togglepause、resume 和 playfrom。发送给形状的其他命令全都会被忽略。

**示例**

### CommandEffect.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 CommandEffect 对象的变量。

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

### CommandEffect.Type

返回或设置一个 **MsoAnimType** 常量，该常量代表动画的类型。可读/写。

**语法**

**express.Type**

\*express是一个代表 CommandEffect 对象的变量。

**说明**

| MsoAnimType 可以是下列 MsoAnimType 常量之一。 |
|-----------------------------------------------|
| MsoAnimTypeColor                              |
| MsoAnimTypeMixed                              |
| MsoAnimTypeMotion                             |
| MsoAnimTypeNone                               |
| MsoAnimTypeProperty                           |
| MsoAnimTypeRoatation                          |
| MsoAnimTypeScale                              |
| MsoAnimTypeTransition                         |

> WPS 开发文档： [WPS 基础接口/演示 API 参考/CommandEffect](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
