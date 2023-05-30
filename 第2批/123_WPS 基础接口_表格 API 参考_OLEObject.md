# OLEObject 对象

## 简介

代表工作表上的一个 ActiveX 控件或链接或嵌入的 OLE 对象。

## 说明

OLEObject 对象是 OLEObjects 集合的成员。 OLEObjects 集合在一张工作表上包含所有的 OLE 对象。

使用 OLEObjects ( *index* )（其中 *index* 是对象名称或编号）可返回一个 OLEObject 对象。下例删除 Sheet1 上的 OLE 对象一。

``` JavaScript
Application.Worksheets.Item("sheet1").OLEObjects(1).Delete()
```

下例删除名为“ListBox1”的 OLE 对象。

``` JavaScript
Application.Worksheets.Item("sheet1").OLEObjects("ListBox1").Delete()
```

工作表上的 ActiveX 控件的 OLEObject 对象的属性和方法是相同的。这样，通过使用控件名称，Visual Basic 代码即可访问这些属性。下例选中复选框控件“MyCheckBox”，将其设为与活动单元格对齐，然后激活此控件。

``` JavaScript
function test(){
    let mcbx = MyCheckBox
    mcbx.Value = true
    mcbx.Top = ActiveCell.Top
    mcbx.Activate()
}
```

## 方法

| 序号 | 名称                                    | 说明                                                           |
|:----:|:----------------------------------------|:---------------------------------------------------------------|
|  1   | [Activate](#OLEObject.Activate)         | 激活对象。返回 Variant值                                       |
|  2   | [BringToFront](#OLEObject.BringToFront) | 将对象放到 z-次序前面。返回 Variant值                          |
|  3   | [Copy](#OLEObject.Copy)                 | 将对象复制到剪贴板。返回 Variant值                             |
|  4   | [CopyPicture](#OLEObject.CopyPicture)   | 将所选对象作为图片复制到剪贴板。返回 Variant值                 |
|  5   | [Cut](#OLEObject.Cut)                   | 将对象剪切到剪贴板，或者将其粘贴到指定的目的地。返回 Variant值 |
|  6   | [Delete](#OLEObject.Delete)             | 删除对象。返回 Variant值                                       |
|  7   | [Duplicate](#OLEObject.Duplicate)       | 复制对象，并返回对新复制对象的引用。                           |
|  8   | [Select](#OLEObject.Select)             | 选择对象。返回 Variant值                                       |
|  9   | [SendToBack](#OLEObject.SendToBack)     | 将对象放到 z-次序的后面。返回 Variant值                        |
|  10  | [Update](#OLEObject.Update)             | 更新链接。返回Variant值                                        |
|  11  | [Verb](#OLEObject.Verb)                 | 向指定的 OLE 对象服务器发送动词。返回 Variant值                |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#OLEObject.Application)         | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [AutoLoad](#OLEObject.AutoLoad)               | 如果打开包含指定 OLE 对象的工作簿时自动载入该 OLE 对象，则为 True 。 Boolean 类型，可读写。                                                                                                                                     |
|  3   | [AutoUpdate](#OLEObject.AutoUpdate)           | 如果数据源改变时 OLE 对象将自动更新，则为 True 。仅当对象是链接方式时有效（该对象的 OLEType 属性必须设为 xlOLELink ）。 Boolean 类型，只读。                                                                                    |
|  4   | [Border](#OLEObject.Border)                   | 返回一个 Border 对象，它代表对象的边框。                                                                                                                                                                                        |
|  5   | [BottomRightCell](#OLEObject.BottomRightCell) | 返回一个 Range 对象，它代表位于该对象右下角下方的单元格。只读。                                                                                                                                                                 |
|  6   | [Creator](#OLEObject.Creator)                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  7   | [Enabled](#OLEObject.Enabled)                 | 如果启用对象，则为 True 。 Boolean 类型，可读写。                                                                                                                                                                               |
|  8   | [Height](#OLEObject.Height)                   | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |
|  9   | [Index](#OLEObject.Index)                     | 返回 Long 值，它代表对象在其同类对象所组成的集合内的索引号。                                                                                                                                                                    |
|  10  | [Interior](#OLEObject.Interior)               | 返回一个 Interior 对象，它代表指定对象的内部。                                                                                                                                                                                  |
|  11  | [Left](#OLEObject.Left)                       | 返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                       |
|  12  | [LinkedCell](#OLEObject.LinkedCell)           | 返回或设置指向控制值的工作表区域。如果为这些单元格赋值，则指定控制也会取得相应的值。与此类似，如果更改控制的值，则单元格的值也作相应变动。 String 型，可读写。                                                                  |
|  13  | [ListFillRange](#OLEObject.ListFillRange)     | 返回或设置用于填充指定列表框的工作表区域。对该属性进行设置将破坏列表框中的所有列表项。 String 型，可读写。                                                                                                                      |
|  14  | [Locked](#OLEObject.Locked)                   | 返回或设置一个 Boolean 值，它指明对象是否已被锁定。                                                                                                                                                                             |
|  15  | [Name](#OLEObject.Name)                       | 返回或设置一个 String 值，它代表对象的名称。                                                                                                                                                                                    |
|  16  | [OLEType](#OLEObject.OLEType)                 | 返回 OLE 对象类型。可为以下 XlOLEType 常量之一： xlOLELink 或 xlOLEEmbed 。如果对象是链接的（对象存储于文件之外），则本属性返回 xlOLELink ，如果对象是内嵌的（对象完全包含于文件之内），则返回 xlOLEEmbed 。 Long 类型，只读。  |
|  17  | [Object](#OLEObject.Object)                   | 返回与此 OLE 对象相联系的 OLE 自动化对象。 Object 型，只读。                                                                                                                                                                    |
|  18  | [Parent](#OLEObject.Parent)                   | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  19  | [Placement](#OLEObject.Placement)             | 返回或设置一个包含 XlPlacement 常量的 Variant 值，它代表对象附加到它所在的单元格的方式。                                                                                                                                        |
|  20  | [PrintObject](#OLEObject.PrintObject)         | 如果打印文档时也打印指定对象，则为 True 。 Boolean 类型，可读写。                                                                                                                                                               |
|  21  | [Shadow](#OLEObject.Shadow)                   | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                               |
|  22  | [ShapeRange](#OLEObject.ShapeRange)           | 返回一个 ShapeRange 对象，它代表指定的一个或多个对象。只读。                                                                                                                                                                    |
|  23  | [SourceName](#OLEObject.SourceName)           | 返回或设置一个 String 值，它代表指定对象的链接源名称。                                                                                                                                                                          |
|  24  | [Top](#OLEObject.Top)                         | 返回或设置一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                      |
|  25  | [TopLeftCell](#OLEObject.TopLeftCell)         | 返回一个 Range 对象，它代表位于指定对象左上角下方的单元格。只读。                                                                                                                                                               |
|  26  | [Visible](#OLEObject.Visible)                 | 返回或设置一个 Boolean 值，它确定对象是否可见。可读写。                                                                                                                                                                         |
|  27  | [Width](#OLEObject.Width)                     | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |
|  28  | [ZOrder](#OLEObject.ZOrder)                   | 返回指定对象的 z-次序位置。 Long 型，只读。                                                                                                                                                                                     |
|  29  | [progId](#OLEObject.progId)                   | 返回对象的程序标识符。 String 型，只读。                                                                                                                                                                                        |

## 成员方法

### OLEObject.Activate

激活对象。返回 Variant值

**语法**

**express.Activate ()**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.BringToFront

将对象放到 z-次序前面。返回 Variant值

**语法**

**express.BringToFront ()**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Copy

将对象复制到剪贴板。返回 Variant值

**语法**

**express.Copy ()**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.CopyPicture

将所选对象作为图片复制到剪贴板。返回 **Variant值**

**语法**

**express.CopyPicture ( *Appearance* , *Format* )**

\*express是一个代表 OLEObject 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型            | 说明                                    |
|--------------|-----------|---------------------|-----------------------------------------|
| *Appearance* | 可选      | XlPictureAppearance | 指定图片的复制方式。默认值为 xlScreen。 |
| *Format*     | 可选      | XlCopyPictureFormat | 图片的格式。默认值为 xlPicture。        |

### OLEObject.Cut

将对象剪切到剪贴板，或者将其粘贴到指定的目的地。返回 Variant值

**语法**

**express.Cut ()**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Delete

删除对象。返回 Variant值

**语法**

**express.Delete ()**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Duplicate

复制对象，并返回对新复制对象的引用。

**语法**

**express.Duplicate ()**

\*express是一个代表 OLEObject 对象的变量。

**说明**

返回 Object值

### OLEObject.Select

选择对象。返回 Variant值

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 OLEObject 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                            |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | 如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

### OLEObject.SendToBack

将对象放到 z-次序的后面。返回 Variant值

**语法**

**express.SendToBack ()**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Update

更新链接。返回Variant值

**语法**

**express.Update ()**

\*express是一个代表 OLEObject 对象的变量。

**示例**

``` JavaScript
/*此示例更新工作表 Sheet1 中第一个 OLE 对象的链接。*/
Application.Worksheets.Item("Sheet1").OLEObjects(1).Update()
```

### OLEObject.Verb

向指定的 OLE 对象服务器发送动词。返回 Variant值

**语法**

**express.Verb ( *Verb* )**

\*express是一个代表 OLEObject 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                                                                                                                                                                     |
|--------|-----------|-----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Verb* | 可选      | XlOLEVerb | OLE 对象服务器将执行其操作的动词。如果省略此参数，则发送默认动词。对象的源应用程序决定哪些动词可用。OLE 对象的典型动词为 Open 和 Primary（由 XlOLEVerb 常量 xlOpen 和 xlPrimary 表示）。 |

**示例**

``` JavaScript
/*本示例向工作表 Sheet1 的第一个 OLE 对象的服务器发送默认动词。*/
Application.Worksheets.Item("Sheet1").OLEObjects(1).Verb()
```

## 成员属性

### OLEObject.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 OLEObject 对象的变量。

**示例**

``` JavaScript
function test(){
  /*本示例显示一条有关创建 myObject 的应用程序的消息。*/
  let myObject = Application.ActiveWorkbook
  if (myObject.Application.Value == "ET"){
      alert("This is an ET Application object.")
  }
  else{
      alert("This is not an ET Application object.")
  }
}
```

### OLEObject.AutoLoad

如果打开包含指定 OLE 对象的工作簿时自动载入该 OLE 对象，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoLoad**

\*express是一个代表 OLEObject 对象的变量。

**说明**

ActiveX 忽略此属性。打开一个工作簿时总会载入 ActiveX 控件。

对于大多数 OLE 对象类型，此属性不能设为 **True** 。对于新 OLE 对象，默认情况下其 **AutoLoad** 属性设为 **False** ；当 ET 载入工作簿时，设为“False”可节省时间和内存。自动载入 OLE 对象的好处在于，对于代表易变动的数据的对象，可立即重建到数据源的链接，而且，如果需要，可对这些对象进行重新映射。

**示例**

``` JavaScript
/*此示例对活动工作表中第一个 OLE 对象的 AutoLoad 属性进行设置。*/
Application.ActiveSheet.OLEObjects(1).AutoLoad = false
```

### OLEObject.AutoUpdate

如果数据源改变时 OLE 对象将自动更新，则为 **True** 。仅当对象是链接方式时有效（该对象的 **OLEType** 属性必须设为 **xlOLELink** ）。 **Boolean** 类型，只读。

**语法**

**express.AutoUpdate**

\*express是一个代表 OLEObject 对象的变量。

**示例**

``` JavaScript
function test(){
  /*如果数据源改变时 OLE 对象将自动更新，则为 True。仅当对象是链接方式时有效（该对象的 OLEType 属性必须设为 xlOLELink）。Boolean 类型，只读。*/
  Application.Worksheets.Item("Sheet1").Activate()
  Range("A1").Value2 = "Name"
  Range("B1").Value2 = "Link Status"
  Range("C1").Value2 = "AutoUpdate Status"
  let i = 2
  let obj = Application.ActiveSheet.OLEObjects()
  for (let x =1; x <= obj.Count; x++){
      Cells.Item(i, 1).Value2 = obj.Item(x).Name
      if(obj.Item(x).OLEType == xlOLELink){
          Cells.Item(i, 2).Value2 = "Linked"
          Cells.Item(i, 3).Value2 = obj.Item(x).AutoUpdate
      }
      else{
          Cells.Item(i, 2).Value2 = "Embedded"
      }
      i = i + 1
  }
  
}
```

### OLEObject.Border

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.BottomRightCell

返回一个 **Range** 对象，它代表位于该对象右下角下方的单元格。只读。

**语法**

**express.BottomRightCell**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 OLEObject 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### OLEObject.Enabled

如果启用对象，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Enabled**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Height

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Index

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Interior

返回一个 **Interior** 对象，它代表指定对象的内部。

**语法**

**express.Interior**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Left

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.LinkedCell

返回或设置指向控制值的工作表区域。如果为这些单元格赋值，则指定控制也会取得相应的值。与此类似，如果更改控制的值，则单元格的值也作相应变动。 **String** 型，可读写。

**语法**

**express.LinkedCell**

\*express是一个代表 OLEObject 对象的变量。

**说明**

不能将该属性应用于多选列表框。

### OLEObject.ListFillRange

返回或设置用于填充指定列表框的工作表区域。对该属性进行设置将破坏列表框中的所有列表项。 **String** 型，可读写。

**语法**

**express.ListFillRange**

\*express是一个代表 OLEObject 对象的变量。

**说明**

ET 阅读区域中的每一单元格的内容，并将单元格值插入到列表框中。列表对该区域中单元格的修订进行追踪。

如果列表框中的列表是使用 **AddItem** 方法创建的，则此属性返回一个空字符串 ("")。

### OLEObject.Locked

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

\*express是一个代表 OLEObject 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True** ；如果在工作表处于受保护状态时仍能修改对象，则返回 **False** 。

### OLEObject.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.OLEType

返回 OLE 对象类型。可为以下 **XlOLEType** 常量之一： **xlOLELink** 或 **xlOLEEmbed** 。如果对象是链接的（对象存储于文件之外），则本属性返回 **xlOLELink** ，如果对象是内嵌的（对象完全包含于文件之内），则返回 **xlOLEEmbed** 。 **Long** 类型，只读。

**语法**

**express.OLEType**

\*express是一个代表 OLEObject 对象的变量。

**示例**

``` JavaScript
function test(){
  /*本示例创建工作表 Sheet1 上 OLE 对象的链接类型列表。该列表将出现在本示例新建的工作表中。*/
  let newSheet = Application.Worksheets.Add()
  let i = 2
  let obj = Application.Worksheets.Item("Sheet1").OLEObjects()
  newSheet.Range("A1").Value2 = "Name"
  newSheet.Range("B1").Value2 = "Link Type"
  for (let x = 1; x <= obj.Count; x++){
      newSheet.Cells.Item(i, 1).Value2 = obj.Item(x).Name
      if (obj.Item(x).OLEType == xlOLELink){
          newSheet.Cells.Item(i, 2).Value2 = "Linked"
      }
      else{
          newSheet.Cells.Item(i, 2).Value2 = "Embedded"
      }
      i = i + 1
  }
}
```

### OLEObject.Object

返回与此 OLE 对象相联系的 OLE 自动化对象。 **Object** 型，只读。

**语法**

**express.Object**

\*express是一个代表 OLEObject 对象的变量。

**示例**

``` JavaScript
function test(){  
  /*此示例在 sheet1 的内嵌 WPS 文档对象的开始处插入文字。注意，在 With 控制结构内的三条语句为 WordBasic 语句。*/
  let wordObj = Application.Worksheets.Item("Sheet1").OLEObjects(1)
  wordObj.Activate()
  let wbc = wordObj.Object.Application.WordBasic
  wbc.StartOfDocument
  wbc.Insert ("This is the beginning")
  wbc.InsertPara
}
```

### OLEObject.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Placement

返回或设置一个包含 **XlPlacement** 常量的 **Variant** 值，它代表对象附加到它所在的单元格的方式。

**语法**

**express.Placement**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.PrintObject

如果打印文档时也打印指定对象，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintObject**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.ShapeRange

返回一个 **ShapeRange** 对象，它代表指定的一个或多个对象。只读。

**语法**

**express.ShapeRange**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.SourceName

返回或设置一个 **String** 值，它代表指定对象的链接源名称。

**语法**

**express.SourceName**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Top

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.TopLeftCell

返回一个 **Range** 对象，它代表位于指定对象左上角下方的单元格。只读。

**语法**

**express.TopLeftCell**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Visible

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.Width

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 OLEObject 对象的变量。

### OLEObject.ZOrder

返回指定对象的 z-次序位置。 **Long** 型，只读。

**语法**

**express.ZOrder**

\*express是一个代表 OLEObject 对象的变量。

**说明**

在任何对象集合中，z-次序尾端的对象为 *collection* (1)，z-次序前端的对象为 *collection* ( *collection* . **Count** )。例如，如果活动工作表中有嵌入图表，z-次序尾端的图表为 ` ActiveSheet.ChartObjects(1) ` ，z-次序前端的图表为 ` ActiveSheet.ChartObjects(ActiveSheet.ChartObjects.Count) ` 。

**示例**

``` JavaScript
/*此示例显示 Sheet1 上嵌入的第一张图表的 z-次序位置。*/
alert("The chart's z-order position is " + Application.Worksheets.Item("Sheet1").ChartObjects(1).ZOrder)
```

### OLEObject.progId

返回对象的程序标识符。 **String** 型，只读。

**语法**

**express.progId**

\*express是一个代表 OLEObject 对象的变量。

**示例**

``` JavaScript
function test(){
  /*此示例为第一张工作表中所有 OLE 对象创建程序标识符列表。*/
  let rw = 0
  let o = Application.Worksheets.Item(1).OLEObjects
  for (let i = 1; i <= o.Count; i++){
      let wss = Worksheets.Item(2)
      rw = rw + 1
      wss.cells.Item(rw, 1).Value2 = o.Item(i).ProgId
  }
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/OLEObject](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
