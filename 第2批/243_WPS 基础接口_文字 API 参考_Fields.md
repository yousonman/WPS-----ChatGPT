# Fields 对象

## 简介

Field 对象的集合，这些对象代表所选内容、范围或文档中的所有域。

## 方法

| 序号 | 名称                                       | 说明                                                                        |
|:----:|:-------------------------------------------|:----------------------------------------------------------------------------|
|  1   | [Add](#Fields.Add)                         | 将 Field 对象添加到 Fields 集合。返回指定区域内的 Field 对象。              |
|  2   | [Item](#Fields.Item)                       | 返回集合中的单个 Field 对象。                                               |
|  3   | [ToggleShowCodes](#Fields.ToggleShowCodes) | 在域代码与域结果之间切换域的显示。可以使用 ShowCodes 属性控制单个域的显示。 |
|  4   | [Unlink](#Fields.Unlink)                   | 将 Fields 集合中的所有域替换为它们的最新结果。                              |
|  5   | [Update](#Fields.Update)                   | 更新域对象的结果。返回 Number 类型。                                        |
|  6   | [UpdateSource](#Fields.UpdateSource)       | 将对 INCLUDETEXT 域所做的更改保存回源文档。                                 |

## 属性

| 序号 | 名称                               | 说明                                                                         |
|:----:|:-----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Fields.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Fields.Count)             | 返回一个 Long 类型的值，该值代表集合中域的数量。只读。                       |
|  3   | [Creator](#Fields.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Locked](#Fields.Locked)           | 如果 Fields 集合中的所有域都已被锁定，则该属性值为 True 。可读写 Long 类型。 |
|  5   | [Parent](#Fields.Parent)           | 返回一个 Object 类型值，该值代表指定 Fields 对象的父对象。                   |

## 成员方法

### Fields.Add

将 **Field** 对象添加到 **Fields** 集合。返回指定区域内的 **Field** 对象。

**语法**

**express.Add ( *Range* , *Type* , *Text* , *PreserveFormatting* )**

\*express是一个代表 Fields 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                                                                         |
|----------------------|-----------|----------|----------------------------------------------------------------------------------------------|
| *Range*              | 必选      | Range    | 需要添加域的区域。如果该区域未折叠，那么域将替换该区域。                                     |
| *Type*               | 可选      | Variant  | 可以是任意 WdFieldType 常量。有关有效的常量列表，请参阅“对象浏览器”。默认值为 wdFieldEmpty。 |
| *Text*               | 可选      | Variant  | 域所需的其他文本。例如，如果要给域指定一个开关，可将其添加到此处。                           |
| *PreserveFormatting* | 可选      | Variant  | 如果该属性值为 True，则更新时保留域所应用的格式。                                            |

**说明**

不能用 **Fields** 集合的 **Add** 方法来插入诸如 **wdFieldOCX** 和 **wdFieldFormCheckBox** 之类的域，而应使用诸如 **FormFields** 集合的 **AddOLEControl** 方法和 **Add** 方法。

**示例**

``` JavaScript
/* 选区开始插入一个日期域 */
function test() {
    Application.Selection.Collapse(wdCollapseStart)
    let myField = Application.ActiveDocument.Fields.Add(Application.Selection.Range,wdFieldDate)
}
```

### Fields.Item

返回集合中的单个 **Field** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Fields 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                           |
|---------|-----------|----------|----------------------------------------------------------------|
| *Index* | 可选      | Number   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

### Fields.ToggleShowCodes

在域代码与域结果之间切换域的显示。可以使用 **ShowCodes** 属性控制单个域的显示。

**语法**

**express.ToggleShowCodes ()**

\*express是一个代表 Fields 对象的变量。

**示例**

``` JavaScript
/*本示例打开或关闭所选内容中域的显示（相当于按 Shift+F9）。*/
Selection.Fields.ToggleShowCodes()
```

``` JavaScript
/*本示例打开或关闭活动文档中域的显示（相当于按 Alt+F9）。*/
ActiveDocument.Fields.ToggleShowCodes()
```

### Fields.Unlink

将 **Fields** 集合中的所有域替换为它们的最新结果。

**语法**

**express.Unlink ()**

\*express是一个代表 Fields 对象的变量。

**说明**

如果断开一个域的链接，则当前结果会转换为文字或图形，并且无法再自动更新。注意，某些域（例如，XE（索引项）域和 SEQ（序列）域）不能断开链接。

**示例**

``` JavaScript
/*以下示例更新并断开活动文档第一节中的所有域的链接。*/
function test()
{
ActiveDocument.Sections.Item(1).Range.Fields.Update()
ActiveDocument.Sections.Item(1).Range.Fields.Unlink()
}
```

### Fields.Update

更新域对象的结果。返回 Number 类型。

**语法**

**express.Update ()**

\*express是一个代表 Fields 对象的变量。

**说明**

如果更新域时没有出错，则返回 0（零）；或者返回一个 **Long** 类型的值，该值代表第一个出错域的索引。

**示例**

``` JavaScript
/* 更新文档第一个域，并且弹框提示是否成功 */
function test() {
    if( Application.ActiveDocument.Fields.Update() == 0) {
        alert("Update Successful")
    }
    else {
        alert( "Field " +  Application.ActiveDocument.Fields.Update() + " has an error")
    }
}
```

### Fields.UpdateSource

将对 INCLUDETEXT 域所做的更改保存回源文档。

**语法**

**express.UpdateSource ()**

\*express是一个代表 Fields 对象的变量。

**说明**

源文档必须是 WPS 文档格式。

## 成员属性

### Fields.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Fields 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Fields.Count

返回一个 **Long** 类型的值，该值代表集合中域的数量。只读。

**语法**

**express.Count**

\*express是一个代表 Fields 对象的变量。

### Fields.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Fields 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Fields.Locked

如果 **Fields** 集合中的所有域都已被锁定，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.Locked**

\*express是一个代表 Fields 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** （集合中的部分域已被锁定，而其他域未锁定）。

**示例**

``` JavaScript
/*以下示例锁定所选内容中的所有域。*/
Selection.Fields.Locked = true
```

``` JavaScript
/*如果活动文档中的部分域已被锁定，则以下示例将显示一条消息。*/
function test()
{
let theFields = ActiveDocument.Fields
if( theFields.Locked == wdUndefined) {
    MsgBox( "Some fields are locked")
}
else if ( theFields.Locked == false) { 
    MsgBox ("No fields are locked")
}
else if( theFields.Locked == true) {
    MsgBox ("All fields are locked")
}
}
```

### Fields.Parent

返回一个 **Object** 类型值，该值代表指定 **Fields** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Fields 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Fields](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
