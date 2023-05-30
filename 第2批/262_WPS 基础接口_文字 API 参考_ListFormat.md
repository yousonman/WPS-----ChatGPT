# ListFormat 对象

## 简介

代表可应用于范围中各段落的列表格式属性。

## 说明

使用 ListFormat 属性可为范围返回 ListFormat 对象。以下示例对所选内容应用默认的项目符号列表格式。

``` JavaScript
Application.Selection.Range.ListFormat.ApplyBulletDefault()
```

应用列表格式的方便途径是使用 ApplyBulletDefault 、 ApplyNumberDefault 和 ApplyOutlineNumberDefault 方法，这些方法分别对应于 “项目符号和编号” 对话框中各选项卡上的第一种列表格式（不包括 “无” ）。

要应用一种非默认格式，请使用 ApplyListTemplate 方法，该方法允许您指定要应用的列表格式（列表模板）。

使用 List 或 ListTemplate 属性从指定范围中的第一段返回列表或列表模板。

可以使用 Range 属性和 ListFormat 对象访问可用于指定区域的列表格式属性和方法。以下示例对活动文档中的第二段应用默认项目符号列表格式。

``` JavaScript
Application.ActiveDocument.Paragraphs.Item(2).Range.ListFormat.ApplyBulletDefault()
```

但是，如果文档中已定义了列表，则可以使用 Lists 属性访问 List 对象。以下示例将在上例中创建的列表的格式更改为 “项目符号和编号” 对话框中 “编号” 选项卡上的第一种编号格式。

``` JavaScript
Application.ActiveDocument.Lists.Item(1).ApplyListTemplate(Application.ListGalleries.Item(2).ListTemplates.Item(1))
```

## 方法

| 序号 | 名称                                                 | 说明                                                                 |
|:----:|:-----------------------------------------------------|:---------------------------------------------------------------------|
|  1   | [ApplyBulletDefault](#ListFormat.ApplyBulletDefault) | 向指定的 ListFormat 对象的区域中的段落添加项目符号和格式。           |
|  2   | [ApplyListTemplate](#ListFormat.ApplyListTemplate)   | 为指定的 ListFormat 对象设置列表格式。                               |
|  3   | [ApplyNumberDefault](#ListFormat.ApplyNumberDefault) | 向指定的 ListFormat 对象的区域中的段落添加默认的编号格式。           |
|  4   | [ListIndent](#ListFormat.ListIndent)                 | 增加指定 ListFormat 对象的区域中的段落的列表级别，增加量为一个级别。 |
|  5   | [ListOutdent](#ListFormat.ListOutdent)               | 降低指定的 ListFormat 对象的区域中段落的列表级别，减量为一个级别。   |
|  6   | [RemoveNumbers](#ListFormat.RemoveNumbers)           | 删除指定列表中的编号或项目符号。                                     |

## 属性

| 序号 | 名称                                               | 说明                                                                                     |
|:----:|:---------------------------------------------------|:-----------------------------------------------------------------------------------------|
|  1   | [List](#ListFormat.List)                           | 返回一个 [List]() 对象，该对象代表在指定 ListFormat 对象中包含的第一个已设置格式的列表。 |
|  2   | [ListLevelNumber](#ListFormat.ListLevelNumber)     | 返回或设置指定的 ListFormat 对象中第一段的列表级别。可读写 Number 类型。                 |
|  3   | [ListPictureBullet](#ListFormat.ListPictureBullet) | 返回一个 InlineShape 对象，该对象代表图片项目符号列表中作为项目符号的图片。              |
|  4   | [ListType](#ListFormat.ListType)                   | 返回指定的 ListFormat 对象区域中所包含的列表类型。 WdListType 类型，只读。               |

## 成员方法

### ListFormat.ApplyBulletDefault

向指定的 **ListFormat** 对象的区域中的段落添加项目符号和格式。

**语法**

**express.ApplyBulletDefault ( *DefaultListBehavior* )**

\*express是一个代表 ListFormat 对象的变量。

**参数**

| 名称                  | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                                                              |
|-----------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *DefaultListBehavior* | 可选      | Variant  | 设置一个值来指定 WPS 是否使用新的面向网站的格式，以使列表具有更好地显示效果。可以是下列常量之一：wdWord8ListBehavior（使用与 WPS 97 兼容的格式）或 wdWord9ListBehavior（使用面向网站的格式）。出于兼容原因，默认常量为 wdWord8ListBehavior。但在新建过程中，要建立缩进式多级项目列表，应当使用 wdWord9ListBehavior 来利用改进过的面向网站的格式。 |

**示例**

``` JavaScript
/* 本示例实现的功能是：为选定内容中的段落添加项目符号并设置格式。如果在选定内容中已经设置了项目符号，本示例将清除其原有的项目符号和格式 */
Application.Selection.Range.ListFormat.ApplyBulletDefault()
```

### ListFormat.ApplyListTemplate

为指定的 **ListFormat** 对象设置列表格式。

**语法**

**express.ApplyListTemplate ( *ListTemplate* , *ContinuePreviousList* , *ApplyTo* , *DefaultListBehavior* )**

\*express是一个代表 ListFormat 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型     | 说明                                                                                                                                                                                                                                                                                                                                                                  |
|------------------------|-----------|--------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ListTemplate*         | 必选      | ListTemplate | 要应用的列表模板。                                                                                                                                                                                                                                                                                                                                                    |
| *ContinuePreviousList* | 可选      | Variant      | 如果为 True，则接续前一列表的编号；如果为 False，则新建一个列表。                                                                                                                                                                                                                                                                                                     |
| *ApplyTo*              | 可选      | Variant      | 要应用列表模板的列表部分。可以是下列 WdListApplyTo 常量之一：wdListSelection、wdListWholeList 或 wdListThisPointForward。                                                                                                                                                                                                                                             |
| *DefaultListBehavior*  | 可选      | Variant      | 设置可指定 WPS 是否使用新的网页格式，以便列表具有更好的显示效果的值。其值可取下列 WdDefaultListBehavior 常量之一：wdWord8ListBehavior（使用与 WPS 97 兼容的格式）或 wdWord9ListBehavior（使用适于网页的格式）。考虑到兼容问题，默认常量为 wdWord8ListBehavior。但是在新建过程中若要建立缩进式多级列表，应当使用 wdWord9ListBehavior，以利用改进过的适用于网页的格式。 |

**示例**

``` JavaScript
/* 将第一段和第二段添加上编号 */
function test() {
    let actDoc = Application.ActiveDocument;
    let rng = actDoc.Range(actDoc.Paragraphs.Item(1).Range.Start, actDoc.Paragraphs.Item(2).Range.End);
    if (rng.ListFormat.ListType === wdListNoNumbering) {
        rng.ListFormat.ApplyListTemplate(Application.ListGalleries.Item(2).ListTemplates.Item(4))
    }
}
```

### ListFormat.ApplyNumberDefault

向指定的 **ListFormat** 对象的区域中的段落添加默认的编号格式。

**语法**

**express.ApplyNumberDefault ( *DefaultListBehavior* )**

\*express是一个代表 ListFormat 对象的变量。

**参数**

| 名称                  | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                                  |
|-----------------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *DefaultListBehavior* | 可选      | Variant  | 设置一个值来指定 WPS 是否使用新的面向网站的格式，以使列表具有更好地显示效果。可以是下列常量之一：wdWord8ListBehavior或 wdWord9ListBehavior（使用面向网站的格式）。出于兼容原因，默认常量为 wdWord8ListBehavior。但在新建过程中，要建立缩进式多级项目列表，应当使用 wdWord9ListBehavior 来利用改进过的面向网站的格式。 |

**说明**

如果段落已设置了编号列表格式，该方法将清除其中已有的编号和格式。

### ListFormat.ListIndent

增加指定 **ListFormat** 对象的区域中的段落的列表级别，增加量为一个级别。

**语法**

**express.ListIndent ()**

\*express是一个代表 ListFormat 对象的变量。

**示例**

``` JavaScript
/* 本示例设置文档 1 的第一个列表中的每一段依次缩进一个级别 */
Application.Documents.Item(1).Lists.Item(1).Range.ListFormat.ListIndent()
```

### ListFormat.ListOutdent

降低指定的 **ListFormat** 对象的区域中段落的列表级别，减量为一个级别。

**语法**

**express.ListOutdent ()**

\*express是一个代表 ListFormat 对象的变量。

**示例**

``` JavaScript
/* 本示例依次降低活动文档中每一段落的缩进量，减量为一个级别 */
Application.Documents.Item(1).Lists.Item(1).Range.ListFormat.ListOutdent()
```

### ListFormat.RemoveNumbers

删除指定列表中的编号或项目符号。

**语法**

**express.RemoveNumbers ( *NumberType* )**

\*express是一个代表 ListFormat 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型     | 说明               |
|--------------|-----------|--------------|--------------------|
| *NumberType* | 可选      | WdNumberType | 要删除的编号类型。 |

## 成员属性

### ListFormat.List

返回一个 **[List]()** 对象，该对象代表在指定 **ListFormat** 对象中包含的第一个已设置格式的列表。

**语法**

**express.List**

\*express是一个代表 ListFormat 对象的变量。

**示例**

``` JavaScript
/* 本示例返回选定内容中的第一个列表的段落数 */
Application.Selection.Range.ListFormat.List.ListParagraphs.Count
```

### ListFormat.ListLevelNumber

返回或设置指定的 **ListFormat** 对象中第一段的列表级别。可读写 **Number** 类型。

**语法**

**express.ListLevelNumber**

\*express是一个代表 ListFormat 对象的变量。

### ListFormat.ListPictureBullet

返回一个 InlineShape 对象，该对象代表图片项目符号列表中作为项目符号的图片。

**语法**

**express.ListPictureBullet**

\*express是一个代表 ListFormat 对象的变量。

### ListFormat.ListType

返回指定的 **ListFormat** 对象区域中所包含的列表类型。 **WdListType** 类型，只读。

**语法**

**express.ListType**

\*express是一个代表 ListFormat 对象的变量。

**说明**

**wdListListNumOnly** 常量引用 LISTNUM 域，可将该域添至段落的文本中。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/ListFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
