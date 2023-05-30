# PlaceholderFormat 对象

## 简介

包含专门应用于占位符的属性，例如占位符类型

## 属性

| 序号 | 名称                                          | 说明                                                            |
|:----:|:----------------------------------------------|:----------------------------------------------------------------|
|  1   | [Application](#PlaceholderFormat.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。         |
|  2   | [Name](#PlaceholderFormat.Name)               | 返回或设置指定对象的名称。可读/写。                             |
|  3   | [Parent](#PlaceholderFormat.Parent)           | 返回指定对象的父对象。                                          |
|  4   | [Type](#PlaceholderFormat.Type)               | 返回一个 PpPlaceholderType 常量，该常量代表占位符的类型。只读。 |

## 成员属性

### PlaceholderFormat.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 PlaceholderFormat 对象的变量。

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

### PlaceholderFormat.Name

返回或设置指定对象的名称。可读/写。

**语法**

**express.Name**

\*express是一个代表 PlaceholderFormat 对象的变量。

### PlaceholderFormat.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 PlaceholderFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
　　　　let myPicture = myShapes.AddShape(msoShapeOval,50,50,300,150).TextFrame
   　　　　 myPicture.TextRange.Text = "Test text"
   　　　　 myPicture.Parent.Rotation = 45
}
```

### PlaceholderFormat.Type

返回一个 **PpPlaceholderType** 常量，该常量代表占位符的类型。只读。

**语法**

**express.Type**

\*express是一个代表 PlaceholderFormat 对象的变量。

**说明**

| PpPlaceholderType 可以是下列 PpPlaceholderType 常量之一。 |
|-----------------------------------------------------------|
| ppPlaceholderBitmap                                       |
| ppPlaceholderBody                                         |
| ppPlaceholderCenterTitle                                  |
| ppPlaceholderChart                                        |
| ppPlaceholderDate                                         |
| ppPlaceholderFooter                                       |
| ppPlaceholderHeader                                       |
| ppPlaceholderMediaClip                                    |
| ppPlaceholderMixed                                        |
| ppPlaceholderObject                                       |
| ppPlaceholderOrgChart                                     |
| ppPlaceholderSlideNumber                                  |
| ppPlaceholderSubtitle                                     |
| ppPlaceholderTable                                        |
| ppPlaceholderTitle                                        |
| ppPlaceholderVerticalBody                                 |
| ppPlaceholderVerticalTitle                                |

> WPS 开发文档： [WPS 基础接口/演示 API 参考/PlaceholderFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
