# Hyperlinks 对象

## 简介

幻灯片或母版上所有 Hyperlink 对象的集合。

## 方法

| 序号 | 名称                     | 说明                       |
|:----:|:-------------------------|:---------------------------|
|  1   | [Item](#Hyperlinks.Item) | 从指定集合中返回单个对象。 |

## 属性

| 序号 | 名称                                   | 说明                                                    |
|:----:|:---------------------------------------|:--------------------------------------------------------|
|  1   | [Application](#Hyperlinks.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#Hyperlinks.Count)             | 返回指定集合中的对象数目                                |
|  3   | [Parent](#Hyperlinks.Parent)           | 返回指定对象的父对象。                                  |

## 成员方法

### Hyperlinks.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Hyperlinks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明 |
|---------|-----------|----------|------|
| *Index* | 必选      | Long     | Long |

**返回值**

Hyperlink

## 成员属性

### Hyperlinks.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Hyperlinks 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPre) {
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

### Hyperlinks.Count

返回指定集合中的对象数目

**语法**

**express.Count**

\*express是一个代表 Hyperlinks 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除窗口1外的所有窗口。*/
function test(){
　　　　for(let i=2; i<=Application.Windows.Count; i++) {
　　　　    Application.Windows.Item(i).Close()
　　　　}
}
```

### Hyperlinks.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Hyperlinks 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
　　　　let tf = myShapes.AddShape(msoShapeOval,50,50,300,150).TextFrame
　　　　tf.TextRange.Text = "Test text"
　　　　tf.Parent.Rotation = 45
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Hyperlinks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
