# Bookmark 对象

## 简介

代表文档、选定内容或区域中的单个书签。 Bookmark 对象是 Bookmarks 集合的成员。 Bookmarks 集合包括 “书签” 对话框中（ “插入” 菜单）列出的所有书签。

## 说明

使用 Bookmarks ( *索引* ) （其中 *index* 是书签名称或索引号）可返回一个 Bookmark 对象。 您必须完全匹配拼写 （但不必大小写匹配） 的书签名称。 下面的示例选择名为活动文档中的"temp"的书签。

``` JavaScript
Application.ActiveDocument.Bookmarks.Item("temp").Select()
```

索引编号代表书签在选区Selection或者区域Range 对象中的位置。对 Document 对象， 索引编号代表书签在 “书签对话框”中按名称字母排序的位置。 下面的示例显示 书签 集合中的第二个书签的名称。

``` JavaScript
Application.ActiveDocument.Bookmarks.Item(1).Name
```

使用 Add 方法将书签添加到文档范围。 以下示例通过添加名为“Abc”的书签来标记选定内容。

``` JavaScript
Application.ActiveDocument.Bookmarks.Add("Abc", Application.Selection.Range)
```

可将 BookmarkID 属性与区域或选定内容对象配合使用，以返回 Bookmark 对象在 Bookmarks 集合中的索引编号。以下示例显示活动文档中名为“ pe ak ”的书签的索引编号。

``` JavaScript
Application.ActiveDocument.Bookmarks.Item("peak").Range.BookmarkID
```

若要确定所选内容、 范围或文档中是否已存在一个书签的方法。 下面的示例可确保在选定该书签之前名为"temp"的书签存在于活动文档。

``` JavaScript
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    if (bks.Exists("temp1") === true) {
        bks.Item("temp").Select();
    }
}
```

## 方法

| 序号 | 名称                       | 说明                                                               |
|:----:|:---------------------------|:-------------------------------------------------------------------|
|  1   | [Copy](#Bookmark.Copy)     | 将书签复制到 Name 参数中指定的新书签处，并返回一个 Bookmark 对象。 |
|  2   | [Delete](#Bookmark.Delete) | 删除指定的书签。                                                   |
|  3   | [Select](#Bookmark.Select) | 选择指定书签。                                                     |

## 属性

| 序号 | 名称                                 | 说明                                                                         |
|:----:|:-------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Bookmark.Application) | 返回一个 Application 对象，该对象代表 WPS 应用程序。                         |
|  2   | [Column](#Bookmark.Column)           | 如果指定的书签是表格的一列，则该属性值为 True 。 Boolean 类型，只读。        |
|  3   | [Creator](#Bookmark.Creator)         | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。 |
|  4   | [Empty](#Bookmark.Empty)             | 如果指定书签为空，则该属性值为 True 。只读 Boolean 类型。                    |
|  5   | [End](#Bookmark.End)                 | 返回或设置选定内容、区域或书签中结束字符的位置。 Number 类型，可读写。       |
|  6   | [Name](#Bookmark.Name)               | 返回指定对象的名称。 String 类型，只读。                                     |
|  7   | [Parent](#Bookmark.Parent)           | 返回一个 Object 类型的值，该值代表指定 Bookmark 对象的父对象。               |
|  8   | [Range](#Bookmark.Range)             | 返回一个 Range 对象，该对象代表指定对象所含的部分文档。                      |
|  9   | [Start](#Bookmark.Start)             | 返回或设置书签起始字符的位置。 Number 类型，可读写。                         |
|  10  | [StoryType](#Bookmark.StoryType)     | 返回指定范围、选定内容或书签的 文章 类型。 WdStoryType 类型，只读。          |

## 成员方法

### Bookmark.Copy

将书签复制到 **Name** 参数中指定的新书签处，并返回一个 **Bookmark** 对象。

**语法**

**express.Copy ( *Name* )**

\*express是一个代表 Bookmark 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明           |
|--------|-----------|----------|----------------|
| *Name* | 必选      | String   | 新书签的名称。 |

**示例**

``` JavaScript
/* 在书签temp处复制生成一个新的书签 */
Application.ActiveDocument.Bookmarks.Item("temp").Copy("test")
```

### Bookmark.Delete

删除指定的书签。

**语法**

**express.Delete ()**

\*express是一个代表 Bookmark 对象的变量。

**示例**

``` JavaScript
/* 如果书签"temp"存在，则删除掉 */
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    if (bks.Exists("temp") === true) {
        bks.Item("temp").Delete();
    }
}
```

### Bookmark.Select

选择指定书签。

**语法**

**express.Select ()**

\*express是一个代表 Bookmark 对象的变量。

**说明**

使用该方法后，可用 **Selection** 对象来处理选定项目。有关详细信息，请参阅 **Selection** 对象。

## 成员属性

### Bookmark.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Bookmark 对象的变量。

### Bookmark.Column

如果指定的书签是表格的一列，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.Column**

\*express是一个代表 Bookmark 对象的变量。

**示例**

``` JavaScript
/* 获取"peak"书签是否为表格的一列 */
Application.ActiveDocument.Bookmarks.Item("peak").Column
```

### Bookmark.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Bookmark 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

### Bookmark.Empty

如果指定书签为空，则该属性值为 **True** 。只读 **Boolean** 类型。

**语法**

**express.Empty**

\*express是一个代表 Bookmark 对象的变量。

**说明**

空书签可标记位置（折叠选定内容），而不标记任何文本。如果指定书签不存在，则会出错。

**示例**

``` JavaScript
/* 本示例查看名为“temp”的书签是否存在和是否为空。 */
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    if(bks.Exists("temp") === true) {
        if(bks.Item("temp").Empty === true) {
            alert("The Temp bookmark is empty");
        }
    }
}
```

### Bookmark.End

返回或设置选定内容、区域或书签中结束字符的位置。 **Number** 类型，可读写。

**语法**

**express.End**

\*express是一个代表 Bookmark 对象的变量。

**说明**

如果本属性设置的值小于 [**Start**](#Bookmark.Start) 属性的值，则 **Start** 属性将被设成同一值（即 **Start** 与 **End** 属性值相等）。

该属性返回结束字符相对于文档开头部分的位置。文档主要文字部分 (wdMainTextStory) 的起始字符位置为 0。通过设置该属性可以更改书签的大小。

**示例**

``` JavaScript
/* 返回书签temp的长度 */
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    let start = bks.Item("temp").Start;
    let end = bks.Item("temp").End;
    alert(end - start);
}
```

### Bookmark.Name

返回指定对象的名称。 **String** 类型，只读。

**语法**

**express.Name**

\*express是一个代表 Bookmark 对象的变量。

**示例**

``` JavaScript
/* 返回索引为1的书签名字 */
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    let bk = bks.Item(1);
    alert(bk.Name);
}
```

### Bookmark.Parent

返回一个 **Object** 类型的值，该值代表指定 **Bookmark** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Bookmark 对象的变量。

### Bookmark.Range

返回一个 **Range** 对象，该对象代表指定对象所含的部分文档。

**语法**

**express.Range**

\*express是一个代表 Bookmark 对象的变量。

**说明**

有关从文档返回一个区域或从图形集合返回图形区域的信息，请参阅 **Range** 方法。

### Bookmark.Start

返回或设置书签起始字符的位置。 **Number** 类型，可读写。

**语法**

**express.Start**

\*express是一个代表 Bookmark 对象的变量。

**说明**

如果将该属性的值设置为大于 **[End](#Bookmark.End)** 属性值，则 **End** 属性值会设置为与 **Start** 属性值相同。

书签对象包括起始字符和结束字符位置。起始字符位置距文档开头部分最近。

该属性返回起始字符相对于文档开头部分的位置。文字部分 ( **wdMainTextStory** ) 的起始字符位置为 0。通过设置本属性可以更改书签的大小。

**示例**

``` JavaScript
/* 返回书签temp的长度 */
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    let start = bks.Item("temp").Start;
    let end = bks.Item("temp").End;
    alert(end - start);
}
```

### Bookmark.StoryType

返回指定范围、选定内容或书签的 文章 类型。 **WdStoryType** 类型，只读。

**语法**

**express.StoryType**

\*express是一个代表 Bookmark 对象的变量。

**示例**

``` JavaScript
/*如果活动文档的主体部分包括名为“temp”的书签，本示例将选中该书签。*/
function test()
{
if(ActiveDocument.Bookmarks.Exists("temp") == true) {
    let myBookmark = ActiveDocument.Bookmarks.Item("temp")
    if(myBookmark.StoryType == wdMainTextStory) {
        myBookmark.Select()
    }
}
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Bookmark](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
