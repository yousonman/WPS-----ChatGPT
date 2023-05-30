# CoAuthors 对象

## 简介

文档中所有 CoAuthor 对象的集合。

## 说明

CoAuthors 集合包含文档中的所有共同作者（当前正在编辑文档的作者）。

``` JavaScript
/*下面的代码示例获取活动文档中的共同作者的数量。*/
function test()
{
let i
i = ActiveDocument.CoAuthoring.Authors.Count
MsgBox("The number of co-authors is " + i)
}
```

## 属性

| 序号 | 名称                                  | 说明                                                             |
|:----:|:--------------------------------------|:-----------------------------------------------------------------|
|  1   | [Application](#CoAuthors.Application) | 返回一个代表 WPS 应用程序的 Application 对象。只读。             |
|  2   | [Count](#CoAuthors.Count)             | 返回 CoAuthors 集合中的项目数。只读。                            |
|  3   | [Creator](#CoAuthors.Creator)         | 返回表示用于创建指定对象的应用程序的 32 位整数。只读 Long 类型。 |
|  4   | [Parent](#CoAuthors.Parent)           | 返回一个 Object 类型值，该值代表指定 CoAuthors 对象的父对象。    |

## 成员属性

### CoAuthors.Application

返回一个代表 WPS 应用程序的 Application 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CoAuthors 对象的变量。

### CoAuthors.Count

返回 CoAuthors 集合中的项目数。只读。

**语法**

**express.Count**

\*express是一个代表 CoAuthors 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例显示活动文档中的共同作者的数量。*/
MsgBox("The active document contains " + ActiveDocument.CoAuthoring.Authors.Count + " authors.")
```

### CoAuthors.Creator

返回表示用于创建指定对象的应用程序的 32 位整数。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CoAuthors 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表 **string** “WPS”。该属性主要设计用于 Apple Macintosh 平台，在该平台上，每个应用程序都有一个由四个字符组成的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的详细信息，请参考 WPS OfficeMacintosh Edition 附带的语言参考帮助。

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td>该值也可用常量 <strong>wdCreatorCode</strong> 表示。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

### CoAuthors.Parent

返回一个 **Object** 类型值，该值代表指定 **CoAuthors** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 CoAuthors 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/CoAuthors](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
