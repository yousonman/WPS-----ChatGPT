# SmartArtNodes 对象

## 简介

代表 Smart Art 图表内部的节点的集合。

## 说明

这些节点直接对应于包含在图形的数据模型内部的语义元素。

``` JavaScript
/*下面的代码返回 Smart Art 图表中的节点数目。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.AllNodes.Count
```

## 方法

| 序号 | 名称                        | 说明                                                         |
|:----:|:----------------------------|:-------------------------------------------------------------|
|  1   | [Add](#SmartArtNodes.Add)   | 将新的 SmartArtNode 对象添加到具有指定文本的图表中。         |
|  2   | [Item](#SmartArtNodes.Item) | 在指定的索引处或通过指定的唯一 ID 来检索 SmartArtNode 对象。 |

## 属性

| 序号 | 名称                                      | 说明                                                                          |
|:----:|:------------------------------------------|:------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtNodes.Application) | 获取一个 Application 对象，该对象代表 SmartArtNodes 对象的容器应用程序。只读  |
|  2   | [Count](#SmartArtNodes.Count)             | 检索包含在 SmartArtNodes 集合中的 SmartArtNode 对象数。只读                   |
|  3   | [Creator](#SmartArtNodes.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtNodes 对象的应用程序。只读 |
|  4   | [Parent](#SmartArtNodes.Parent)           | 返回调用对象。只读                                                            |

## 成员方法

### SmartArtNodes.Add

将新的 **SmartArtNode** 对象添加到具有指定文本的图表中。

**语法**

**express.Add ()**

\*express是一个代表 SmartArtNodes 对象的变量。

**返回值**

SmartArtNode

**说明**

这个新节点将添加到位于此节点集合最顶层的数据模型的底部。如果最高级别是 2，则这个新节点的级别也将是 2。

**示例**

``` JavaScript
/*下面的代码将一个 SmartArtNode 添加到 SmartArtNodes 集合中。*/
function test(){
    var saNodes = Application.ActiveWorkbook.ActiveChart.Shapes.Item(1).SmartArt.AllNodes.Item(1);
    saNodes.AddNode();
}
```

### SmartArtNodes.Item

在指定的索引处或通过指定的唯一 ID 来检索 **SmartArtNode** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SmartArtNodes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                             |
|---------|-----------|----------|------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 指定一个代表索引的整数或一个代表 SmartArtNode 对象位置的字符串。 |

**返回值**

SmartArtNode

## 成员属性

### SmartArtNodes.Application

获取一个 **Application** 对象，该对象代表 **SmartArtNodes** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 SmartArtNodes 对象的变量。

### SmartArtNodes.Count

检索包含在 **SmartArtNodes** 集合中的 **SmartArtNode** 对象数。只读

**语法**

**express.Count**

\*express是一个代表 SmartArtNodes 对象的变量。

### SmartArtNodes.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtNodes** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 SmartArtNodes 对象的变量。

### SmartArtNodes.Parent

返回调用对象。只读

**语法**

**express.Parent**

\*express是一个代表 SmartArtNodes 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtNodes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
