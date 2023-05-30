# Designs 对象

## 简介

代表幻灯片设计模板的集合。

## 说明

使用 Presentation 对象的 Designs 属性引用一个设计模板。

若要添加或克隆单个的设计模板，请分别使用 Designs 集合的 Add 或 Clone 方法。若要引用单个的设计模板，请使用 Item 方法。

若要加载一个设计模板，请使用 Load 方法。

## 方法

| 序号 | 名称                    | 说明                                                                         |
|:----:|:------------------------|:-----------------------------------------------------------------------------|
|  1   | [Add](#Designs.Add)     | 返回一个代表新幻灯片设计的 Design 对象。                                     |
|  2   | [Clone](#Designs.Clone) | 创建 Design 对象的一个副本。                                                 |
|  3   | [Item](#Designs.Item)   | 从指定集合中返回单个对象。                                                   |
|  4   | [Load](#Designs.Load)   | 返回一个 Design 对象，该对象代表一个已加载到指定演示文稿的母版列表中的设计。 |

## 属性

| 序号 | 名称                                | 说明                                                    |
|:----:|:------------------------------------|:--------------------------------------------------------|
|  1   | [Application](#Designs.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#Designs.Count)             | 返回指定集合中的对象数目。                              |
|  3   | [Parent](#Designs.Parent)           | 返回指定对象的父对象。                                  |

## 成员方法

### Designs.Add

返回一个代表新幻灯片设计的 **Design** 对象。

**语法**

**express.Add ( *designName* , *Index* )**

\*express是一个代表 Designs 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                 |
|--------------|-----------|----------|------------------------------------------------------------------------------------------------------|
| *designName* | 必选      | String   | 设计的名称。                                                                                         |
| *Index*      | 必选      | Integer  | 设计的索引号。其默认值为 -1，表示如果省略 Index 参数，则新的幻灯片设计将添加到现有幻灯片设计的末尾。 |

**返回值**

Design

### Designs.Clone

创建 **Design** 对象的一个副本。

**语法**

**express.Clone ( *pOriginal* , *Index* )**

\*express是一个代表 Designs 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                               |
|-------------|-----------|----------|----------------------------------------------------------------------------------------------------|
| *pOriginal* | 必选      | Design   | Design 对象。原始设计。                                                                            |
| *Index*     | 必选      | Long     | 设计将要复制到的 Designs 集合中的索引位置。如果忽略 Index，复制的设计则添加到 Designs 集合的最后。 |

**返回值**

Design

**示例**

### Designs.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Designs 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

**返回值**

Design

### Designs.Load

返回一个 **Design** 对象，该对象代表一个已加载到指定演示文稿的母版列表中的设计。

**语法**

**express.Load ( *TemplateName* , *Index* )**

\*express是一个代表 Designs 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                                              |
|----------------|-----------|----------|---------------------------------------------------------------------------------------------------|
| *TemplateName* | 必选      | String   | 设计模板的路径。                                                                                  |
| *Index*        | 必选      | Long     | 设计模板在设计模板集合中的索引号。默认值为 -1，表示将设计模板添加到演示文稿中设计模板列表的末尾。 |

## 成员属性

### Designs.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Designs 对象的变量。

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

### Designs.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 Designs 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭第二个窗口。*/
function test(){
　　　　let dCount = Application.Windows
    for (let i = 2; i <= dCount.Count; i++) {
　　　　    dCount.Item(2).Close()
　　　　}
}
```

### Designs.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Designs 对象的变量。

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

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Designs](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
