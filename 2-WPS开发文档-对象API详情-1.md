# VSpace 对象

## 简介

填充垂直方向上无用的空隙。

## 说明

填充垂直方向上无用的空隙。

## 属性

| 序号 | 名称                                 | 说明                                                 |
|:----:|:-------------------------------------|:-----------------------------------------------------|
|  1   | [Height](#VSpace.Height)             | 获取或设置指定对象的高度                             |
|  2   | [HeightPolicy](#VSpace.HeightPolicy) | 获取或设置指定对像的高度策略                         |
|  3   | [Left](#VSpace.Left)                 | 您可以通过该属性指定控件在对象或报表中的位置         |
|  4   | [Name](#VSpace.Name)                 | 您可以通过该属性指定或确定表示对象名称的字符串，只读 |
|  5   | [Top](#VSpace.Top)                   | 您可以通过该属性指定控件在对象或报表中的位置         |
|  6   | [Width](#VSpace.Width)               | 获取或设置指定对象的宽度                             |
|  7   | [WidthPolicy](#VSpace.WidthPolicy)   | 获取或设置指定对象的宽度策略                         |

## 成员属性

### VSpace.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 VSpace 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将VSpacer1控件的高度设置为100*/
function func1()
{
    UserForm1.VSpacer1.Height = 100
}
```

### VSpace.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 VSpace 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 可变宽度 |

**示例**

``` JavaScript
/*将VSpacer1的高度设置为固定*/
function func1()
{
    UserForm1.VSpacer1.HeightPolicy = 2
}
```

### VSpace.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 VSpace 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将VSpacer1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.VSpacer1.Left = 100
}
```

### VSpace.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 VSpace 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### VSpace.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 VSpace 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将VSpacer1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.VSpacer1.Top = 100
}
```

### VSpace.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 VSpace 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将VSpacer1的宽度设置为100*/
function func1()
{
    UserForm1.VSpacer1.Width = 100
}
```

### VSpace.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 VSpace 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将VSpacer1的宽度设置为固定*/
function func1()
{
    UserForm1.VSpacer1.WidthPolicy = 2
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/VSpace](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# AllowEditRange 对象

## 简介

代表受保护的工作表上可进行编辑的单元格。

## 说明

使用 Add 方法或 AllowEditRanges 集合的 Item 属性可返回 AllowEditRange 对象。 返回 AllowEditRange 对象后，可使用 ChangePa ssword 方法更改访问受保护工作表上可编辑区域的密码。

此示例中，ET 允许用户在活动工作表上编辑单元格区域 A1:A4，并通知用户，然后更改此指定区域的密码并将所做的更改通知用户。

``` JavaScript
function test() {
    let wksOne
    let wksPassword
    wksOne = Application.ActiveSheet
　　wksPassword = prompt("Enter password for the worksheet")
    // Establish a range that can allow edits on the protected worksheet.
    wksOne.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), "123")

    alert("Cells A1 to A4 can be edited on the protected worksheet.")

    // Change the password.　　wksPassword = prompt("Enter the new password for the worksheet")
    wksOne.Protection.AllowEditRanges.Item(1).ChangePassword("456")

    alert("The password for these cells has been changed.")
} 
```

## 方法

| 序号 | 名称                                             | 说明                                                                         |
|:----:|:-------------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [ChangePassword](#AllowEditRange.ChangePassword) | 更改受保护的工作表中可以进行编辑的区域的密码。                               |
|  2   | [Delete](#AllowEditRange.Delete)                 | 删除对象。                                                                   |
|  3   | [Unprotect](#AllowEditRange.Unprotect)           | 取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。 |

## 属性

| 序号 | 名称                           | 说明                                                                   |
|:----:|:-------------------------------|:-----------------------------------------------------------------------|
|  1   | [Range](#AllowEditRange.Range) | 返回一个 Range 代表，它代表在受保护工作表上可编辑的区域的子集。        |
|  2   | [Title](#AllowEditRange.Title) | 返回或设置受保护的工作上可编辑的单元格区域的标题。 String 型，可读写。 |
|  3   | [Users](#AllowEditRange.Users) | 返回工作表上受保护区域的一个 UserAccessList 对象。                     |

## 成员方法

### AllowEditRange.ChangePassword

更改受保护的工作表中可以进行编辑的区域的密码。

**语法**

**express.ChangePassword ( *Password* )**

\*express是一个代表 AllowEditRange 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明     |
|------------|-----------|----------|----------|
| *Password* | 必选      | String   | 新密码。 |

**示例**

``` JavaScript
function test() {
    let wksOne
    let wksPassword
    wksOne = Application.ActiveSheet

    // Establish a range that can allow edits on the protected worksheet.
    wksOne.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), "123")

    alert("Cells A1 to A4 can be edited on the protected worksheet.")

    // Change the password.
    wksOne.Protection.AllowEditRanges.Item(1).ChangePassword("456")

    alert("The password for these cells has been changed.")
}
```

### AllowEditRange.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 AllowEditRange 对象的变量。

### AllowEditRange.Unprotect

取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。

**语法**

**express.Unprotect ( *Password* )**

\*express是一个代表 AllowEditRange 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                             |
|------------|-----------|----------|------------------------------------------------------------------------------------------------------------------|
| *Password* | 可选      | Variant  | 一个字符串，它指定用于解除单元格区域保护的密码，此密码是区分大小写的。如果单元格区域不设密码保护，则忽略此参数。 |

**说明**

***如果您忘记了密码，将不能取消工作表或工作簿的保护。建议将密码和对应文档名妥善保存。***

## 成员属性

### AllowEditRange.Range

返回一个 Range 代表，它代表在受保护工作表上可编辑的区域的子集。

**语法**

**express.Range**

\*express是一个代表 AllowEditRange 对象的变量。

### AllowEditRange.Title

返回或设置受保护的工作上可编辑的单元格区域的标题。 **String** 型，可读写。

**语法**

**express.Title**

\*express是一个代表 AllowEditRange 对象的变量。

### AllowEditRange.Users

返回工作表上受保护区域的一个 **UserAccessList** 对象。

**语法**

**express.Users**

\*express是一个代表 AllowEditRange 对象的变量。

**示例**

``` JavaScript
function test() {
    let wksSheet = Application.ActiveSheet

    // Display name of user with access to protected range.
    alert(wksSheet.Protection.AllowEditRanges.Item(1).Users.Item(1).Name)
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/AllowEditRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# AutoFilter 对象

## 简介

代表对指定工作表的自动筛选。

## 说明

使用 AutoFilter 属性可返回 AutoFilter 对象。使用 Filters 属性可返回由各个列筛选组成的集合。使用 Range 属性可返回代表整个筛选区域的 Range 对象。

``` JavaScript
Dim w As Worksheet
Dim filterArray()
Dim currentFiltRange As String

Sub ChangeFilters()

Set w = Worksheets("Crew")
With w.AutoFilter
    currentFiltRange = .Range.Address
    With .Filters
        ReDim filterArray(1 To .Count, 1 To 3)
        For f = 1 To .Count
            With .Item(f)
                If .On Then
                    filterArray(f, 1) = .Criteria1
                    If .Operator Then
                        filterArray(f, 2) = .Operator
                        filterArray(f, 3) = .Criteria2
                    End If
                End If
            End With
        Next
    End With
End With

w.AutoFilterMode = False
w.Range("A1").AutoFilter field:=1, Criteria1:="S"

End Sub
```

要为工作表创建 AutoFilter 对象，必须手动打开或使用 Range 对象的 AutoFilter 方法打开工作表上某个区域上的自动筛选功能。

``` JavaScript
Sub RestoreFilters()
Set w = Worksheets("Crew")
w.AutoFilterMode = False
For col = 1 To UBound(filterArray(), 1)
    If Not IsEmpty(filterArray(col, 1)) Then
        If filterArray(col, 2) Then
            w.Range(currentFiltRange).AutoFilter field:=col, _
                Criteria1:=filterArray(col, 1), _
                    Operator:=filterArray(col, 2), _
                Criteria2:=filterArray(col, 3)
        Else
            w.Range(currentFiltRange).AutoFilter field:=col, _
                Criteria1:=filterArray(col, 1)
        End If
    End If
Next
End 
```

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
<td>When using AutoFilter with dates, the format should be consistent with English date separators ("/") instead of local settings ("."). A valid date would be "2/2/2007", whereas "2.2.2007" is invalid.</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

## 方法

| 序号 | 名称                                   | 说明                                 |
|:----:|:---------------------------------------|:-------------------------------------|
|  1   | [ApplyFilter](#AutoFilter.ApplyFilter) | 应用指定的 Autofilter 对象。         |
|  2   | [ShowAllData](#AutoFilter.ShowAllData) | 显示 AutoFilter 对象返回的所有数据。 |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#AutoFilter.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#AutoFilter.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [FilterMode](#AutoFilter.FilterMode)   | 如果工作表的筛选模式为自动筛选，则返回 True 。只读 Boolean 类型。                                                                                                                                                               |
|  4   | [Filters](#AutoFilter.Filters)         | 返回一个 Filters 集合，该集合表示自动筛选区域中的所有筛选器。只读。                                                                                                                                                             |
|  5   | [Parent](#AutoFilter.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  6   | [Range](#AutoFilter.Range)             | 返回一个 Range 对象，它代表应用指定的“自动筛选”的区域。                                                                                                                                                                         |
|  7   | [Sort](#AutoFilter.Sort)               | 获取 AutoFilter 集合的排序列和排序次序。                                                                                                                                                                                        |

## 成员方法

### AutoFilter.ApplyFilter

应用指定的 **Autofilter** 对象。

**语法**

**express.ApplyFilter ()**

\*express是一个代表 AutoFilter 对象的变量。

### AutoFilter.ShowAllData

显示 **AutoFilter** 对象返回的所有数据。

**语法**

**express.ShowAllData ()**

\*express是一个代表 AutoFilter 对象的变量。

## 成员属性

### AutoFilter.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 AutoFilter 对象的变量。

### AutoFilter.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 AutoFilter 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### AutoFilter.FilterMode

如果工作表的筛选模式为自动筛选，则返回 **True** 。只读 **Boolean** 类型。

**语法**

**express.FilterMode**

\*express是一个代表 AutoFilter 对象的变量。

### AutoFilter.Filters

返回一个 **Filters** 集合，该集合表示自动筛选区域中的所有筛选器。只读。

**语法**

**express.Filters**

\*express是一个代表 AutoFilter 对象的变量。

**示例**

``` JavaScript
/*下例将变量设置为工作表“Crew”上的筛选区域中第一列的筛选的 Criteria1 属性值。*/
function test()
{
  let mySheet = Application.Worksheets.Item("Crew")
  if(mySheet.AutoFilterMode){
      if(mySheet.AutoFilter.Filters.Item(1).On){
          let c1 = mySheet.AutoFilter.Filters.Item(1).Criteria1
      }
  }
}
```

### AutoFilter.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 AutoFilter 对象的变量。

### AutoFilter.Range

返回一个 **Range** 对象，它代表应用指定的“自动筛选”的区域。

**语法**

**express.Range**

\*express是一个代表 AutoFilter 对象的变量。

**示例**

``` JavaScript
/*下例将应用于“Crew”工作表的“自动筛选”的地址保存在变量中。*/
let rAddress = Application.Worksheets.Item("Crew").AutoFilter.Range.Address
/*此示例滚动工作簿窗口，直至超链接区域位于活动窗口的左上角。*/
function test()
{
  Application.Workbooks.Item(1).Activate()
  
  let myRange = Application.ActiveSheet.Hyperlinks.Item(1).Range
  Application.ActiveWindow.ScrollRow = myRange.Row
  Application.ActiveWindow.ScrollColumn = myRange.Column
}
```

### AutoFilter.Sort

获取 **AutoFilter** 集合的排序列和排序次序。

**语法**

**express.Sort**

\*express是一个代表 AutoFilter 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/AutoFilter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ChartArea 对象

## 简介

代表图表的图表区。

## 说明

图表区包含绘图区在内的一切内容。但是，绘图区具有它自己的填充方式，因此填充绘图区并不会填充图表区。

有关设置绘图区格式的信息，请参阅 PlotArea 对象 。

使用 ChartArea 属性可返回 ChartArea 对象。

示例

以下示例关闭名为“Sheet1”的工作表上嵌入图表 1 的图表区边框。

``` JavaScript
Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format.Line.Visible = false
```

## 方法

| 序号 | 名称                                      | 说明                             |
|:----:|:------------------------------------------|:---------------------------------|
|  1   | [Clear](#ChartArea.Clear)                 | 清除整个对象。                   |
|  2   | [ClearContents](#ChartArea.ClearContents) | 清除图表中的数据但保留格式设置。 |
|  3   | [ClearFormats](#ChartArea.ClearFormats)   | 清除对象的格式设置。             |
|  4   | [Copy](#ChartArea.Copy)                   | 将对象复制到剪贴板。             |
|  5   | [Select](#ChartArea.Select)               | 选择对象。                       |

## 属性

| 序号 | 名称                                        | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ChartArea.Application)       | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#ChartArea.Creator)               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Format](#ChartArea.Format)                 | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  4   | [Height](#ChartArea.Height)                 | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |
|  5   | [Left](#ChartArea.Left)                     | 返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                       |
|  6   | [Name](#ChartArea.Name)                     | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  7   | [Parent](#ChartArea.Parent)                 | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  8   | [RoundedCorners](#ChartArea.RoundedCorners) | 如果图表的图表区域使用圆角，则为 True 。可读写 Boolean 类型。                                                                                                                                                                   |
|  9   | [Shadow](#ChartArea.Shadow)                 | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                               |
|  10  | [Top](#ChartArea.Top)                       | 返回一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                            |
|  11  | [Width](#ChartArea.Width)                   | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |

## 成员方法

### ChartArea.Clear

清除整个对象。

**语法**

**express.Clear ()**

\*express是一个代表 ChartArea 对象的变量。

**示例**

``` JavaScript
/*此示例清除 Chart1 中的图表区域（图表数据和格式设置）。*/
Application.Charts.Item("Chart1").ChartArea.Clear()
```

### ChartArea.ClearContents

清除图表中的数据但保留格式设置。

**语法**

**express.ClearContents ()**

\*express是一个代表 ChartArea 对象的变量。

**示例**

``` JavaScript
/*此示例清除 Chart1 的图表数据，但保留其格式设置。*/
Application.Charts.Item("Chart1").ChartArea.ClearContents()
```

### ChartArea.ClearFormats

清除对象的格式设置。

**语法**

**express.ClearFormats ()**

\*express是一个代表 ChartArea 对象的变量。

**示例**

``` JavaScript
/*此示例清除工作表 Sheet1 上第一个嵌入式图表的格式设置。*/
Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.ClearFormats()
```

### ChartArea.Copy

将对象复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 ChartArea 对象的变量。

## 成员属性

### ChartArea.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ChartArea 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    if(Application.ActiveWorkbook.Application.Value == "ET"){
        alert("This is an ET Application object.")
    }
    else{
        alert("This is not an ET Application object.")
    }
}
```

### ChartArea.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ChartArea 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ChartArea.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 ChartArea 对象的变量。

**说明**

**ChartFormat** 对象包含图表区的线条、填充和效果格式。

### ChartArea.Height

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Left

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.RoundedCorners

如果图表的图表区域使用圆角，则为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.RoundedCorners**

\*express是一个代表 ChartArea 对象的变量。

**示例**

``` JavaScript
/*此示例为 Sheet1 上的第一个图表添加圆角。*/
Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.RoundedCorners = true
```

### ChartArea.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Top

返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 ChartArea 对象的变量。

### ChartArea.Width

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 ChartArea 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ChartArea](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ChartObjects 对象

## 简介

由指定的图表工作表、对话框工作表或工作表上的所有 ChartObject 对象组成的集合。

## 说明

每个 ChartObject 对象都代表一个嵌入图表。 ChartObject 对象充当 Chart 对象的容器。 ChartObject 对象的属性和方法控制工作表上嵌入图表的外观和大小。 ChartObjects 集合

示例

使用 ChartObjects 方法返回 ChartObjects 集合。以下示例删除名为“Sheet1”的工作表上的所有嵌入图表。

``` JavaScript
Application.Worksheets.Item("Sheet1").ChartObjects().Delete()
```

不能使用 ChartObjects 集合来调用以下属性和方法：

- Locked 属性
- Placement 属性
- PrintObject 属性

与早期版本不同， ChartObjects 集合现在可以读取表示高度、宽度、左对齐和顶对齐的属性。

使用 Add 方法可创建一个新的空嵌入图表并将它添加到集合中。使用 ChartWizard 方法可添加数据并设置新图表的格式。以下示例创建一个新嵌入图表，然后以折线图形式添加单元格 A1:A20 中的数据。

``` JavaScript
let ch = Application.Worksheets.Item("Sheet1").ChartObjects().Add(100, 30, 400, 250)
ch.Chart.ChartWizard(Worksheets.Item("sheet1").Range("a1:a20"), xlLine
    , null, null, null, null, null, "New Chart", null, null, null)
```

使用 ChartObjects ( *index* )（其中 *index* 是嵌入图表的索引号或名称）可以返回单个对象。以下示例设置名为“Sheet1”的工作表上嵌入图表 Chart 1 中的图表区图案。

``` JavaScript
Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart. 
    ChartArea.Format.Fill.Pattern = msoPatternLightDownwardDiagonal
```

## 方法

| 序号 | 名称                                     | 说明                                 |
|:----:|:-----------------------------------------|:-------------------------------------|
|  1   | [Add](#ChartObjects.Add)                 | 创建新的嵌入式图表。                 |
|  2   | [Copy](#ChartObjects.Copy)               | 将对象复制到剪贴板。                 |
|  3   | [CopyPicture](#ChartObjects.CopyPicture) | 将所选对象作为图片复制到剪贴板。     |
|  4   | [Cut](#ChartObjects.Cut)                 | 将对象剪切到剪贴板。                 |
|  5   | [Delete](#ChartObjects.Delete)           | 删除对象。                           |
|  6   | [Duplicate](#ChartObjects.Duplicate)     | 复制对象，并返回对新复制对象的引用。 |
|  7   | [Item](#ChartObjects.Item)               | 从集合中返回一个对象。               |
|  8   | [Select](#ChartObjects.Select)           | 选择对象。                           |

## 属性

| 序号 | 名称                                                   | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ChartObjects.Application)               | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#ChartObjects.Count)                           | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#ChartObjects.Creator)                       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Height](#ChartObjects.Height)                         | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |
|  5   | [Left](#ChartObjects.Left)                             | 返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                       |
|  6   | [Parent](#ChartObjects.Parent)                         | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  7   | [ProtectChartObject](#ChartObjects.ProtectChartObject) | 如果嵌入式图表框不能通过用户界面移动、调整大小或删除，则为 True 。可读/写 Boolean 类型。                                                                                                                                        |
|  8   | [ShapeRange](#ChartObjects.ShapeRange)                 | 返回一个 ShapeRange 对象，它代表指定的一个或多个对象。只读。                                                                                                                                                                    |
|  9   | [Top](#ChartObjects.Top)                               | 返回或设置一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                      |
|  10  | [Visible](#ChartObjects.Visible)                       | 返回或设置一个 Boolean 值，它确定对象是否可见。可读写。                                                                                                                                                                         |
|  11  | [Width](#ChartObjects.Width)                           | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |

## 成员方法

### ChartObjects.Add

创建新的嵌入式图表。

**语法**

**express.Add ( *Left* , *Width* )**

\*express是一个代表 ChartObjects 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                           |
|---------|-----------|----------|------------------------------------------------------------------------------------------------|
| *Left*  | 必选      | Double   | 以磅为单位给出新对象的初始坐标，该坐标是相对于工作表上单元格 A1 的左上角或图表的左上角的坐标。 |
| *Width* | 必选      | Double   | 以磅为单位给出新对象的初始大小。                                                               |

**示例**

``` JavaScript
/*此示例创建新的嵌入式图表。*/
function test()
{
    let co = Application.Sheets.Item("Sheet1").ChartObjects().Add(50, 40, 200, 100)
    co.Chart.ChartWizard(Worksheets.Item("Sheet1").Range("A1:B2"), 
        xlColumn, 6, xlColumns, 1, 0, 1)
}
```

### ChartObjects.Copy

将对象复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 ChartObjects 对象的变量。

### ChartObjects.CopyPicture

将所选对象作为图片复制到剪贴板。

**语法**

**express.CopyPicture ( *Appearance* , *Format* )**

\*express是一个代表 ChartObjects 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型            | 说明                                    |
|--------------|-----------|---------------------|-----------------------------------------|
| *Appearance* | 可选      | XlPictureAppearance | 指定图片的复制方式。默认值为 xlScreen。 |
| *Format*     | 可选      | XlCopyPictureFormat | 图片的格式。默认值为 xlPicture。        |

### ChartObjects.Cut

将对象剪切到剪贴板。

**语法**

**express.Cut ()**

\*express是一个代表 ChartObjects 对象的变量。

**说明**

只可剪切嵌入式图表。

### ChartObjects.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 ChartObjects 对象的变量。

### ChartObjects.Duplicate

复制对象，并返回对新复制对象的引用。

**语法**

**express.Duplicate ()**

\*express是一个代表 ChartObjects 对象的变量。

**示例**

``` JavaScript
/*此示例复制工作表 Sheet1 上嵌入的第一个图表，然后选择新复制的图表。*/
function test()
{  
    let dChart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Duplicate()
    dChart.Select()
}  
```

### ChartObjects.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 ChartObjects 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**说明**

对象的文本名称就是 **Name** 和 **Value** 属性的值。

**示例**

``` JavaScript
/*此示例激活第一张嵌入式图表。*/
Application.Worksheets.Item("Sheet1").ChartObjects.Item(1).Activate()
```

### ChartObjects.Select

选择对象。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 ChartObjects 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                            |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | 如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

## 成员属性

### ChartObjects.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ChartObjects 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### ChartObjects.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 ChartObjects 对象的变量。

### ChartObjects.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ChartObjects 对象的变量。

### ChartObjects.Height

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 ChartObjects 对象的变量。

### ChartObjects.Left

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 ChartObjects 对象的变量。

### ChartObjects.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ChartObjects 对象的变量。

### ChartObjects.ProtectChartObject

如果嵌入式图表框不能通过用户界面移动、调整大小或删除，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.ProtectChartObject**

\*express是一个代表 ChartObjects 对象的变量。

**说明**

将此属性设置为 **True** ，嵌入式图表框还是可以通过对象模型进行修改。

**示例**

``` JavaScript
Application.Worksheets.Item(1).ChartObjects(1).ProtectChartObject = true
```

### ChartObjects.ShapeRange

返回一个 **ShapeRange** 对象，它代表指定的一个或多个对象。只读。

**语法**

**express.ShapeRange**

\*express是一个代表 ChartObjects 对象的变量。

**示例**

``` JavaScript
/*此示例创建一个形状区域，该区域代表第一张工作表上的嵌入图表。*/
let sr = Application.Worksheets.Item(1).ChartObjects.ShapeRange
```

### ChartObjects.Top

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 ChartObjects 对象的变量。

### ChartObjects.Visible

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 ChartObjects 对象的变量。

### ChartObjects.Width

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 ChartObjects 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ChartObjects](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ConditionValue 对象

## 简介

代表数据条条件格式规则计算最短数据条和最长数据条的方法。

## 说明

ConditionValue 对象是使用 Databar 对象的 MaxPoint 或 MinPoint 属性返回的。

通过使用 Modify 方法，您可以从默认设置（最低值表示最短的数据条，最高值表示最长的数据条）中更改计算类型。

以下示例将创建一个数据范围，然后对该范围应用数据条。您将会发现，由于范围中存在非常低和非常高的值，因此中间值具有长度相似的数据条。为了明确显示中间值，示例代码使用 ConditionValue 对象将阈值的计算方式更改为百分点。

``` JavaScript
function test(){
 //Create a range of data with a couple of extreme values
    Application.ActiveSheet.Range("D1").Value2 = 1
    Application.ActiveSheet.Range("D2").Value2 = 45
    Application.ActiveSheet.Range("D3").Value2 = 50
    Application.ActiveSheet.Range("D2:D3").AutoFill(Range("D2:D8"))
    Application.ActiveSheet.Range("D9").Value2 = 500
    
    Range("D1:D9").Select()
        
    //Create a data bar with default behavior
    let cfDataBar = Application.Selection.FormatConditions.AddDatabar()
    alert("Because of the extreme values, middle data bars are very similar")
    
    //The MinPoint and MaxPoint properties return a ConditionValue object
    //which you can use to change threshold parameters
    cfDataBar.MinPoint.Modify(xlConditionValuePercentile, 5)
    cfDataBar.MaxPoint.Modify(xlConditionValuePercentile, 75)

}
```

## 方法

| 序号 | 名称                             | 说明                                                     |
|:----:|:---------------------------------|:---------------------------------------------------------|
|  1   | [Modify](#ConditionValue.Modify) | 修改数据条条件格式规则计算最长数据条或最短数据条的方法。 |

## 属性

| 序号 | 名称                                       | 说明                                                                                                                                                               |
|:----:|:-------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ConditionValue.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#ConditionValue.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Parent](#ConditionValue.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  4   | [Type](#ConditionValue.Type)               | 返回 XlConditionValueTypes 枚举的常量之一，该常量指定如何确定数据栏、色标或图标集条件格式的阈值。只读。                                                            |
|  5   | [Value](#ConditionValue.Value)             | 针对数据条条件格式返回或设置最短或最长的数据条的阈值。可读/写 Variant 类型。                                                                                       |

## 成员方法

### ConditionValue.Modify

修改数据条条件格式规则计算最长数据条或最短数据条的方法。

**语法**

**express.Modify ( *newtype* , *newvalue* )**

\*express是一个代表 ConditionValue 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型              | 说明                                                                                                                              |
|------------|-----------|-----------------------|-----------------------------------------------------------------------------------------------------------------------------------|
| *newtype*  | 必选      | XlConditionValueTypes | 指定最短数据条或最长数据条的计算方法。最短数据条的默认值是 xlConditionLowestValue；最长数据条的默认值是 xlConditionHighestValue。 |
| *newvalue* | 可选      | Variant               | 分配给最短或最长数据条的值。取决于 newtype 参数，此值可以是数字，也可以是计算结果为数字的公式。                                   |

**说明**

table { word-break:break-all; }

下表描述了每种计算类型的可接受阈值。

| newtype 参数               | newvalue 参数           |
|----------------------------|-------------------------|
| xlConditionLowestValue     | 忽略参数                |
| xlConditionHighestValue    | 忽略参数                |
| xlConditionValueNumber     | 任何数字                |
| xlConditionValuePercent    | 0 到 100 之间的任何数字 |
| xlConditionValuePercentile | 0 到 100 之间的任何数字 |
| xlConditionValueFormula    | 一个返回单一数字的公式  |

## 成员属性

### ConditionValue.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ConditionValue 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### ConditionValue.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ConditionValue 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ConditionValue.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ConditionValue 对象的变量。

### ConditionValue.Type

返回 **XlConditionValueTypes** 枚举的常量之一，该常量指定如何确定数据栏、色标或图标集条件格式的阈值。只读。

**语法**

**express.Type**

\*express是一个代表 ConditionValue 对象的变量。

### ConditionValue.Value

针对数据条条件格式返回或设置最短或最长的数据条的阈值。可读/写 **Variant** 类型。

**语法**

**express.Value**

\*express是一个代表 ConditionValue 对象的变量。

**说明**

只有将条件格式的 **ConditionValue.Type** 属性设置为以下常量之一，您才可以设置值： **xlConditionValueNumber** 、 **xlConditionValuePercent** 、 **xlConditionValuePercentile** 或 **xlConditionValueFormula** 。

如果阈值类型是公式，您可以将公式设置为 **String** 。此公式必须返回单一数字。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ConditionValue](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DefaultWebOptions 对象

## 简介

包含应用程序级的全局属性，当将文档另存为网页或打开网页时，ET 将使用这些属性。您可以在应用程序（全局）级或在工作簿级返回或设置属性。

## 说明

工作簿级的属性设置会重写应用程序级的属性设置。工作簿级的属性设置包含在 WebOptions 对象中。

| 注释                                                                 |
|----------------------------------------------------------------------|
| 由于保存工作簿时的属性值可能不同，所以不同工作簿的属性值也可能不同。 |

使用 DefaultWebOptions 属性可返回 DefaultWebOptions 对象。

``` JavaScript
/*下例将查看是否允许将 PNG（可移植网络图形）作为图像格式使用，同时设置相应的 strImageFileType 变量。*/
function test(){
let objAppWebOptions = Application.DefaultWebOptions
    if(objAppWebOptions.AllowPNG == true) {
        strImageFileType = "PNG"
    }
    else {
        strImageFileType = "JPG"
    }
}
```

## 属性

| 序号 | 名称                                                                            | 说明                                                                                                                                                                                                                                                                                                |
|:----:|:--------------------------------------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AllowPNG](#DefaultWebOptions.AllowPNG)                                         | 如果以网页保存文档时，允许将 PNG（可移植网络图形）作为图像格式使用，则为 True 。如果不允许将 PNG 作为输出格式使用，则为 False 。默认值是 False 。 Boolean 类型，可读写。                                                                                                                            |
|  2   | [AlwaysSaveInDefaultEncoding](#DefaultWebOptions.AlwaysSaveInDefaultEncoding)   | 如果在保存网页或纯文本文档时将使用默认编码，而该默认编码与文件打开时的初始编码是互相独立的，则该属性值为 True 。如果使用文件的初始编码，则该属性值为 False 。默认值为 False 。 Boolean 类型，可读写。                                                                                               |
|  3   | [Application](#DefaultWebOptions.Application)                                   | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。                                                                     |
|  4   | [CheckIfOfficeIsHTMLEditor](#DefaultWebOptions.CheckIfOfficeIsHTMLEditor)       | 如果在启动 ET 时，ET 检查某个 WPS 应用程序是否为默认 HTML 编辑器，则该值为 True 。如果 ET 并不进行此检查，则该值为 False 。默认值为 True 。 Boolean 类型，可读写。                                                                                                                                  |
|  5   | [Creator](#DefaultWebOptions.Creator)                                           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                                                                                          |
|  6   | [DownloadComponents](#DefaultWebOptions.DownloadComponents)                     | 如果必要的“WPS Office Web 组件”在您用 Web 浏览器查看已保存的文档时已经下载，但只有在尚未安装这些组件时，该值为 True 。如果没有下载这些组件，则为 False 。默认值是 False 。 Boolean 类型，可读写。                                                                                                   |
|  7   | [Encoding](#DefaultWebOptions.Encoding)                                         | 当查看已保存的文档时，返回或设置 Web 浏览器使用的文档编码（代码页或字符集）。默认值为系统代码页。 MsoEncoding 类型，可读写。                                                                                                                                                                        |
|  8   | [FolderSuffix](#DefaultWebOptions.FolderSuffix)                                 | 返回文件夹后缀，当您将文档保存为网页、使用长文件名，以及选择将支持文件单独保存在某个文件夹中（即，如果 UseLongFileNames 和 OrganizeInFolder 属性设置为 True ）时，ET 使用该后缀。 String 型，只读。                                                                                                 |
|  9   | [Fonts](#DefaultWebOptions.Fonts)                                               | 返回 WebPageFonts 集合，该集合代表用户在 ET 中打开网页，而该网页中没有指定任何字体信息，或者当前默认字体无法显示网页中的字符集时，ET 将使用的字体集。只读。                                                                                                                                         |
|  10  | [LoadPictures](#DefaultWebOptions.LoadPictures)                                 | 在 ET 中打开某个文档时，如果加载图像，则该值为 True ，通常这些图像和文档并不是在 ET 中创建的。如果未加载图像，则该值为 False 。默认值为 True 。 Boolean 类型，可读写。                                                                                                                              |
|  11  | [LocationOfComponents](#DefaultWebOptions.LocationOfComponents)                 | 返回或设置中央 URL（对于 Intranet 或 Web）或路径（对于本地或网络而言），授权用户可以在查看保存的文档时，从这些位置下载“WPS Office Web 组件”。默认值是 WPS Office 的本地或网络安装路径。 String 型，可读写。                                                                                         |
|  12  | [OrganizeInFolder](#DefaultWebOptions.OrganizeInFolder)                         | 如果为 True ，则在将指定的文档保存为网页时，所有的支持文件（如背景纹理和图形）将组织在一个单独的文件夹中；如果为 False ，则支持文件和网页将保存在同一文件夹中。默认值是 True 。 Boolean 类型，可读写。                                                                                              |
|  13  | [Parent](#DefaultWebOptions.Parent)                                             | 返回指定对象的父对象。只读。                                                                                                                                                                                                                                                                        |
|  14  | [PixelsPerInch](#DefaultWebOptions.PixelsPerInch)                               | 返回或设置网页中图形图像的密度（每英寸的像素数）及表格单元格的密度。密度设置范围通常为 19 到 480 之间，对于普通的屏幕大小，通常的设置一般为 72、96 和 120。默认设置是 96。 Long 型，可读写。                                                                                                        |
|  15  | [RelyOnCSS](#DefaultWebOptions.RelyOnCSS)                                       | 如果在 Web 浏览器中查看保存的文档时，字体格式使用的是级联样式表 (CSS)，则为 True 。ET 会创建一个级联样式表文件，然后依据 OrganizeInFolder 属性的值，将该文件保存到指定文件夹或与网页相同的文件夹中。如果使用了 HTML \<FONT\> 标记和级联样式表，则为 False 。默认值是 True 。 Boolean 类型，可读写。 |
|  16  | [RelyOnVML](#DefaultWebOptions.RelyOnVML)                                       | 如果为 True ，将文档保存为网页时，图形对象不能生成图像文件。如果为 False ，则可生成图像文件。默认值是 False 。 Boolean 类型，可读写。                                                                                                                                                               |
|  17  | [SaveHiddenData](#DefaultWebOptions.SaveHiddenData)                             | 当以网页保存文档时，如果也保存指定区域之外的数据，则该值为 True 。此数据对于维护公式是很有必要的。如果指定区域之外的数据并不与网页一起保存，则该值为 False 。默认值为 True 。 Boolean 类型，可读写。                                                                                                |
|  18  | [SaveNewWebPagesAsWebArchives](#DefaultWebOptions.SaveNewWebPagesAsWebArchives) | 如果新的网页能够保存为 Web 档案，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                                        |
|  19  | [ScreenSize](#DefaultWebOptions.ScreenSize)                                     | 返回或设置在 Web 浏览器中查看已保存文档时应使用的理想的最小屏幕大小（以像素为单位，宽度乘以高度）。可以是 MsoScreenSize 常量之一。默认常量是 msoScreenSize800x600 。 MsoScreenSize 类型，可读写。                                                                                                   |
|  20  | [TargetBrowser](#DefaultWebOptions.TargetBrowser)                               | 返回或设置一个 MsoTargetBrowser 常量，它指明浏览器的版本。可读写。                                                                                                                                                                                                                                  |
|  21  | [UpdateLinksOnSave](#DefaultWebOptions.UpdateLinksOnSave)                       | 如果该属性值为 True ，则将文档保存为网页之前，自动更新超链接和所有指向支持文件的路径，以确保文档保存时包含最新的链接。如果该属性值为 False ，则不更新链接。默认值为 True 。 Boolean 类型，可读写。                                                                                                  |
|  22  | [UseLongFileNames](#DefaultWebOptions.UseLongFileNames)                         | 如果该属性值为 True ，则将文档保存为网页时使用长文件名。如果该属性值为 False ，则不使用长文件名，而使用 DOS 文件名格式（8.3）。默认值是 True 。 Boolean 类型，可读写。                                                                                                                              |

## 成员属性

### DefaultWebOptions.AllowPNG

如果以网页保存文档时，允许将 PNG（可移植网络图形）作为图像格式使用，则为 **True** 。如果不允许将 PNG 作为输出格式使用，则为 **False** 。默认值是 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.AllowPNG**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

如果将图像保存为 PNG 格式而非其他文件格式，则可以提高图像的质量或减小图像文件的大小，因而也就可以减少下载所需的时间，前提是所使用的 Web 浏览器支持 PNG 格式。

**示例**

``` JavaScript
/*或者，新建文档的应用程序可以将 PNG 作为全局默认设置。*/
Application.DefaultWebOptions.AllowPNG = true
```

### DefaultWebOptions.AlwaysSaveInDefaultEncoding

如果在保存网页或纯文本文档时将使用默认编码，而该默认编码与文件打开时的初始编码是互相独立的，则该属性值为 **True** 。如果使用文件的初始编码，则该属性值为 **False** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.AlwaysSaveInDefaultEncoding**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

可以使用 **Encoding** 属性设置默认编码。

**示例**

``` JavaScript
/*本示例将编码设置为默认编码。当用户将文档保存为网页时使用该编码。*/
Application.DefaultWebOptions.AlwaysSaveInDefaultEncoding = true
```

### DefaultWebOptions.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 DefaultWebOptions 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### DefaultWebOptions.CheckIfOfficeIsHTMLEditor

如果在启动 ET 时，ET 检查某个 WPS 应用程序是否为默认 HTML 编辑器，则该值为 **True** 。如果 ET 并不进行此检查，则该值为 **False** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CheckIfOfficeIsHTMLEditor**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

仅当所使用的 Web 浏览器支持 HTML 编辑和 HTML 编辑器时，才使用该属性。

要使用其他 HTML 编辑器，必须将该属性设置为 **False** ，然后将该编辑器注册为默认的系统 HTML 编辑器。

**示例**

``` JavaScript
/*本示例使得 ET 不检查它是否为默认 HTML 编辑器。*/
Application.DefaultWebOptions.CheckIfOfficeIsHTMLEditor = false
```

### DefaultWebOptions.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### DefaultWebOptions.DownloadComponents

如果必要的“WPS Office Web 组件”在您用 Web 浏览器查看已保存的文档时已经下载，但只有在尚未安装这些组件时，该值为 **True** 。如果没有下载这些组件，则为 **False** 。默认值是 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.DownloadComponents**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

可将 **LocationOfComponents** 属性设置为主 URL（在 Intranet 或网站上）或路径（本地或网络），授权用户在查看已保存的文档时，可从该路径所指向的位置下载组件。该路径必须是有效路径，且指向的位置必须包含必要的组件，同时用户还必须拥有有效的 WPS Office 许可。

Office Web 组件可增强保存为网页的文档的交互性。如果在没有安装这些组件的计算机上使用 Web 浏览器查看网页，则该页上的交互部分将表现为静态的（不能交互）。

**示例**

``` JavaScript
/*如果还没有安装“Office Web 组件”，此示例允许“Office Web 组件”与指定的网页一起下载。*/
function test(){
Application.DefaultWebOptions.DownloadComponents = true
Application.DefaultWebOptions.LocationOfComponents = 
    Application.Path + Application.PathSeparator + "foo"
}
```

### DefaultWebOptions.Encoding

当查看已保存的文档时，返回或设置 Web 浏览器使用的文档编码（代码页或字符集）。默认值为系统代码页。 **MsoEncoding** 类型，可读写。

**语法**

**express.Encoding**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

不能使用任何带有 **AutoDetect** 后缀的常量。这些常量由 **ReloadAs** 方法使用。

**示例**

``` JavaScript
/*本示例检查默认编码方式是否为 Western，然后设置相应的 strDocEncoding 字符串。*/
function test(){
if(Application.DefaultWebOptions.Encoding == msoEncodingWestern) {
    strDocEncoding = "Western"
}
else {
    strDocEncoding = "Other"
}
}
```

### DefaultWebOptions.FolderSuffix

返回文件夹后缀，当您将文档保存为网页、使用长文件名，以及选择将支持文件单独保存在某个文件夹中（即，如果 **UseLongFileNames** 和 **OrganizeInFolder** 属性设置为 **True** ）时，ET 使用该后缀。 **String** 型，只读。

**语法**

**express.FolderSuffix**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

新建的文档将使用 **DefaultWebOptions** 对象的 **FolderSuffix** 属性所返回的后缀。如果文档曾在其他语言版本的 ET 中进行过编辑，则 **WebOptions** 对象的 **FolderSuffix** 属性的值可能会与 **DefaultWebOptions** 对象中该属性的值不同。可以使用 **UseDefaultFolderSuffix** 方法将后缀更改为当前在 WPS Office 中所使用的语言。

默认情况下，支持文件夹名由网页的名称加上一个下划线 (\_)、一个句点 (.) 或连字号 (-)，再加上单词“files”（将文件保存为网页的 ET 版本的语言形式）组成。例如，如果使用荷兰语版本的 ET 将名为“Page1”的文件保存为网页，则支持文件夹的默认名称为“Page1_bestanden”。

下表列出 WPS 的各个语言版本及相应的 **LanguageID** 属性值和文件夹后缀。对于表中没有列出的语言，其后缀为“.files”。

| 语言                 | LanguageID | 文件夹后缀    |
|----------------------|------------|---------------|
| 阿拉伯语             | 1025       | .files        |
| 巴斯克语             | 1069       | \_fitxategiak |
| 巴西葡萄牙语         | 1046       | \_arquivos    |
| 保加利亚语           | 1026       | .files        |
| 加泰罗尼亚语         | 1027       | \_fitxers     |
| 中文 - 简体          | 2052       | .files        |
| 中文 - 繁体          | 1028       | .files        |
| 克罗地亚语           | 1050       | \_datoteke    |
| 捷克语               | 1029       | \_soubory     |
| 丹麦语               | 1030       | -filer        |
| 荷兰语               | 1043       | \_bestanden   |
| 英语                 | 1033       | \_files       |
| 爱沙尼亚语           | 1061       | \_failid      |
| 芬兰语               | 1035       | \_tiedostot   |
| 法语                 | 1036       | \_fichiers    |
| 德语                 | 1031       | -Dateien      |
| 希腊语               | 1032       | .files        |
| 希伯来语             | 1037       | .files        |
| 匈牙利语             | 1038       | \_elemei      |
| 意大利语             | 1040       | -file         |
| 日语                 | 1041       | .files        |
| 朝鲜语               | 1042       | .files        |
| 拉脱维亚语           | 1062       | \_fails       |
| 立陶宛语             | 1063       | \_bylos       |
| 挪威语               | 1044       | -filer        |
| 波兰语               | 1045       | \_pliki       |
| 葡萄牙语             | 2070       | \_ficheiros   |
| 罗马尼亚语           | 1048       | .files        |
| 俄语                 | 1049       | .files        |
| 塞尔维亚语(西里尔文) | 3098       | .files        |
| 塞尔维亚语（拉丁文） | 2074       | \_fajlovi     |
| 斯洛伐克语           | 1051       | .files        |
| 斯洛文尼亚语         | 1060       | \_datoteke    |
| 西班牙语             | 3082       | \_archivos    |
| 瑞典语               | 1053       | -filer        |
| 泰语                 | 1054       | .files        |
| 土耳其语             | 1055       | \_dosyalar    |
| 乌克兰语             | 1058       | .files        |
| 越南语               | 1066       | .files        |

### DefaultWebOptions.Fonts

返回 **WebPageFonts** 集合，该集合代表用户在 ET 中打开网页，而该网页中没有指定任何字体信息，或者当前默认字体无法显示网页中的字符集时，ET 将使用的字体集。只读。

**语法**

**express.Fonts**

\*express是一个代表 DefaultWebOptions 对象的变量。

**示例**

``` JavaScript
/*本示例设置英语/西欧/其他拉丁语系字符集的默认的固定宽度字体为 Courier New，14 磅。*/
function test(){
let myFont = Application.DefaultWebOptions 
    .Fonts.Item(msoCharacterSetEnglishWesternEuropeanOtherLatinScript)
        myFont.FixedWidthFont = "Courier New"
        myFont.FixedWidthFontSize = 14
}
```

### DefaultWebOptions.LoadPictures

在 ET 中打开某个文档时，如果加载图像，则该值为 **True** ，通常这些图像和文档并不是在 ET 中创建的。如果未加载图像，则该值为 **False** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.LoadPictures**

\*express是一个代表 DefaultWebOptions 对象的变量。

**示例**

``` JavaScript
/*本示例使得在 ET 中打开文档时要加载图像。*/
Application.DefaultWebOptions.LoadPictures = true
```

### DefaultWebOptions.LocationOfComponents

返回或设置中央 URL（对于 Intranet 或 Web）或路径（对于本地或网络而言），授权用户可以在查看保存的文档时，从这些位置下载“WPS Office Web 组件”。默认值是 WPS Office 的本地或网络安装路径。 **String** 型，可读写。

**语法**

**express.LocationOfComponents**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

如果 **DownloadComponents** 属性设置为 **True** ，WPS Office Web 组件尚未安装，路径是有效路径且指向包含必要组件的位置，同时用户还具有有效的 WPS Office 许可，则会从指定网页自动下载这些组件。

**示例**

``` JavaScript
/*此示例设置路径，用户可以从该位置下载“WPS Office Web 组件”。*/
function test(){
Application.DefaultWebOptions.DownloadComponents = true
Application.DefaultWebOptions.LocationOfComponents = 
    Application.Path + Application.PathSeparator + "foo"
}
```

### DefaultWebOptions.OrganizeInFolder

如果为 **True** ，则在将指定的文档保存为网页时，所有的支持文件（如背景纹理和图形）将组织在一个单独的文件夹中；如果为 **False** ，则支持文件和网页将保存在同一文件夹中。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.OrganizeInFolder**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

在保存网页的文件夹中创建新文件夹，并依据文档命名。如果使用长文件名，则在文件夹名称后加后缀。 **FolderSuffix** 属性返回已选择或已安装的语言支持的文件夹后缀，或返回默认的文件夹后缀。

如果在上次保存某个文档时， **OrganizeInFolder** 属性的值与现在不同，那么现在保存该文档时，ET 会相应地将支持文件自动移进或移出该文件夹。

如果并未使用长文件名（即，如果 **UseLongFileNames** 属性设置为 **False** ），则 ET 会自动将任何支持文件保存到其他文件夹中。这些文件不能保存到与网页相同的文件夹中。

**示例**

``` JavaScript
/*此示例指定在将文档保存为网页时，所有的支持文件和网页保存在同一文件夹中。*/
Application.DefaultWebOptions.OrganizeInFolder = false
```

### DefaultWebOptions.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 DefaultWebOptions 对象的变量。

### DefaultWebOptions.PixelsPerInch

返回或设置网页中图形图像的密度（每英寸的像素数）及表格单元格的密度。密度设置范围通常为 19 到 480 之间，对于普通的屏幕大小，通常的设置一般为 72、96 和 120。默认设置是 96。 **Long** 型，可读写。

**语法**

**express.PixelsPerInch**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

无论何时在 Web 浏览器中查看所保存文档，此属性都可确定指定的网页中相对于文本尺寸的图像以及单元格的大小。图像或单元格的实际最终尺寸等于原始尺寸（以英寸为单位）乘以每英寸中的像素数。

使用 **ScreenSize** 属性可为目标 Web 浏览器设置最佳屏幕大小。

**示例**

此示例根据浏览器的目标屏幕大小来设置像素密度。对于 800x600 像素的屏幕，其密度为 72 像素/英寸。对于 1024x768 像素的屏幕，其密度为 96 像素/英寸。对于所有其他情况，则使用的密度为 120 像素/英寸。

``` JavaScript
function test(){
let myOption = Application.DefaultWebOptions
    switch(myOption.ScreenSize) {
        case msoScreenSize800x600:
            myOption.PixelsPerInch = 72
            break
        case msoScreenSize1024x768:
            myOption.PixelsPerInch = 96
            break
        default:
            myOption.PixelsPerInch = 120
    }
}
```

### DefaultWebOptions.RelyOnCSS

如果在 Web 浏览器中查看保存的文档时，字体格式使用的是级联样式表 (CSS)，则为 **True** 。ET 会创建一个级联样式表文件，然后依据 **OrganizeInFolder** 属性的值，将该文件保存到指定文件夹或与网页相同的文件夹中。如果使用了 HTML \<FONT\> 标记和级联样式表，则为 **False** 。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.RelyOnCSS**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

如果 Web 浏览器支持级联样式表，则应将此属性设置为 **True** ，这样，就可以为该网页上提供更精确的布局和格式控制，网页也会更接近原来的文档（如同在 ET 中显示一样）。

**示例**

``` JavaScript
/*此示例启用级联样式表作为应用程序的全局默认值。*/
Application.DefaultWebOptions.RelyOnCSS = true
```

### DefaultWebOptions.RelyOnVML

如果为 **True** ，将文档保存为网页时，图形对象不能生成图像文件。如果为 **False** ，则可生成图像文件。默认值是 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.RelyOnVML**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

如果您的 Web 浏览器支持 Vector Markup Language (VML)，也可不从图形对象生成图像，这样可减少文件大小。例如，Microsoft Internet Explorer 5 支持此项功能，因此如果要将其作为目标浏览器，应该将 **RelyOnVML** 属性设为 **True** 。如果浏览器不支持 VML，则在查看用此属性保存的网页时，将不会显示图像。

例如，如果网页中使用了已生成的图像文件，并且文档保存的位置与该网页在 Web 服务器上的最终位置不同，则不应生成图像。

**示例**

``` JavaScript
/*此示例指定在将工作表保存到网页上时，生成图像。*/
Application.Workbooks.Item(1).DefaultWebOptions.RelyOnVML = false
```

### DefaultWebOptions.SaveHiddenData

当以网页保存文档时，如果也保存指定区域之外的数据，则该值为 **True** 。此数据对于维护公式是很有必要的。如果指定区域之外的数据并不与网页一起保存，则该值为 **False** 。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SaveHiddenData**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

如果选择不保存指定区域之外的数据，则公式中对这些数据的引用都将变为静态值。如果数据位于另外的工作表或工作簿中，则公式结果将以静态值保存。

**示例**

``` JavaScript
/*本示例防止将隐藏数据与网页一起保存。*/
Application.DefaultWebOptions.SaveHiddenData = false
 
```

### DefaultWebOptions.SaveNewWebPagesAsWebArchives

如果新的网页能够保存为 Web 档案，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SaveNewWebPagesAsWebArchives**

\*express是一个代表 DefaultWebOptions 对象的变量。

**示例**

``` JavaScript
/*在本示例中，ET 确定将新网页存为 Web 档案的设置，并通知用户*/
function test(){
//Determine settings and notify user.
    if(Application.DefaultWebOptions.SaveNewWebPagesAsWebArchives == true) {
        alert("New Web pages will be saved as Web archives.")
    }
    else {
        alert("New Web pages will not be saved as Web archives.")
    }


}
```

### DefaultWebOptions.ScreenSize

返回或设置在 Web 浏览器中查看已保存文档时应使用的理想的最小屏幕大小（以像素为单位，宽度乘以高度）。可以是 **MsoScreenSize** 常量之一。默认常量是 **msoScreenSize800x600** 。 **MsoScreenSize** 类型，可读写。

**语法**

**express.ScreenSize**

\*express是一个代表 DefaultWebOptions 对象的变量。

### DefaultWebOptions.TargetBrowser

返回或设置一个 **MsoTargetBrowser** 常量，它指明浏览器的版本。可读写。

**语法**

**express.TargetBrowser**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

|                                                                                               |
|-----------------------------------------------------------------------------------------------|
| **MsoTargetBrowser** 可为以下这些 **MsoTargetBrowser** 常量之一。                             |
| **msoTargetBrowserIE4** 。Microsoft Internet Explorer 4.0 或更高版本。                        |
| **msoTargetBrowserIE5** 。Microsoft Internet Explorer 5.0 或更高版本。                        |
| **msoTargetBrowserIE6** 。Microsoft Internet Explorer 6.0 或更高版本。                        |
| **msoTargetBrowserV3** 。Microsoft Internet Explorer 3.0、Netscape Navigator 3.0 或更高版本。 |
| **msoTargetBrowserV4** 。Microsoft Internet Explorer 4.0、Netscape Navigator 4.0 或更高版本。 |

**示例**

``` JavaScript
/*在本示例中，ET 确定 Web 选项的浏览器版本是否为 IE5，并通知用户。*/
function test(){
 let wkbOne = Application.Workbooks.Item(1)

    //Determine if IE5 is the target browser.
    if(wkbOne.WebOptions.TargetBrowser == msoTargetBrowserIE5) {
        alert("The target browser is IE5 or later.")
    }
    else {
        alert("The target browser is not IE5 or later.")
    }
}
```

### DefaultWebOptions.UpdateLinksOnSave

如果该属性值为 **True** ，则将文档保存为网页之前，自动更新超链接和所有指向支持文件的路径，以确保文档保存时包含最新的链接。如果该属性值为 **False** ，则不更新链接。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.UpdateLinksOnSave**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

如果文档保存的位置与其最终在 Web 服务器上的位置不同，且支持文件在第一个位置不可用，则应将该属性设为 **False** 。

**示例**

``` JavaScript
/*本示例指定在保存文档前不更新链接。*/
Application.DefaultWebOptions.UpdateLinksOnSave = false
```

### DefaultWebOptions.UseLongFileNames

如果该属性值为 **True** ，则将文档保存为网页时使用长文件名。如果该属性值为 **False** ，则不使用长文件名，而使用 DOS 文件名格式（8.3）。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.UseLongFileNames**

\*express是一个代表 DefaultWebOptions 对象的变量。

**说明**

如果不使用长文件名且文档还具有支持文件，ET 会自动将那些文件组织到单独的文件夹中。另外，使用 **OrganizeInFolder** 属性可确定支持文件是否组织在单独的文件夹中。

**示例**

/\*此示例禁止将长文件名用作应用程序的全局默认设置。\*/

``` JavaScript
Application.DefaultWebOptions.UseLongFileNames = false
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/DefaultWebOptions](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FileExportConverter 对象

## 简介

代表用于保存文件的文件转换器。

## 说明

无法新建文件转换器，也无法向 FileExportConverters 集合中添加文件转换器。 FileExportConverter 对象是在安装 WPS Office 时或安装补充文件转换器时添加的。

使用 FileExportConverters ( *Index* ) 可返回单个 FileExportConverter 对象，其中 *Index* 为整数。

``` JavaScript
/*以下示例显示与该集合中第二个 ET 工作表转换器关联的扩展名。*/
alert(FileExportConverters.Item(2).Extensions)
```

索引号代表文件转换器在 FileExportConverters 集合中的位置。

``` JavaScript
/*以下示例显示该集合中第一个文件转换器的说明。*/
alert(FileExportConvters.Item(1).Description)
```

## 属性

| 序号 | 名称                                            | 说明                                                                            |
|:----:|:------------------------------------------------|:--------------------------------------------------------------------------------|
|  1   | [Application](#FileExportConverter.Application) | 返回代表 ET 应用程序的 Application 对象。只读。                                 |
|  2   | [Creator](#FileExportConverter.Creator)         | 返回表示在其中创建特定对象的应用程序的 32 位整数。 Long 类型，只读。            |
|  3   | [Description](#FileExportConverter.Description) | 返回文件转换器的说明。 String 类型，只读                                        |
|  4   | [Extensions](#FileExportConverter.Extensions)   | 返回与指定 FileExportConverter 对象关联的文件扩展名。 String 类型，只读。       |
|  5   | [FileFormat](#FileExportConverter.FileFormat)   | 返回一个整数，该整数标识与指定 FileExportConverter 对象关联的文件格式。只读。   |
|  6   | [Parent](#FileExportConverter.Parent)           | 返回一个 Object 类型的值，该值代表指定 FileExportConverter 对象的父对象。只读。 |

## 成员属性

### FileExportConverter.Application

返回代表 ET 应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 FileExportConverter 对象的变量。

### FileExportConverter.Creator

返回表示在其中创建特定对象的应用程序的 32 位整数。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 FileExportConverter 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回十六进制值 5843454C，该值表示字符串“XCEL”。 **Creator** 属性设计为在 ET for the Macintosh 中使用，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### FileExportConverter.Description

返回文件转换器的说明。 **String** 类型，只读

**语法**

**express.Description**

\*express是一个代表 FileExportConverter 对象的变量。

**示例**

``` JavaScript
/*下面的示例显示 FileExportConverters 集合中第一个文件转换器的说明。*/
function test(){
let fcTemp = FileExportConverter.Item(1)
alert(fcTemp.Description)
}
```

### FileExportConverter.Extensions

返回与指定 **FileExportConverter** 对象关联的文件扩展名。 **String** 类型，只读。

**语法**

**express.Extensions**

\*express是一个代表 FileExportConverter 对象的变量。

**示例**

``` JavaScript
/*下面的示例显示 FileExportConverters 集合中第一个文件转换器的文件扩展名。*/
function test(){
let fcTemp = FileExportConverters.Item(1)
alert("The file name extensions for the file converter are: " + fcTemp.Extensions)
}
```

### FileExportConverter.FileFormat

返回一个整数，该整数标识与指定 **FileExportConverter** 对象关联的文件格式。只读。

**语法**

**express.FileFormat**

\*express是一个代表 FileExportConverter 对象的变量。

**示例**

``` JavaScript
/*下面的示例显示 FileExportConverters 集合中第一个文件转换器的文件格式标识符。*/
function test(){
let fcTemp = FileExportConverters.Item(1)
alert("The file format identifier for the file converter is: " + fcTemp.FileFormat)
}
```

``` JavaScript
/*下面的示例说明在使用 FileExportConverters 集合中的第一个文件转换器保存文件时，如何将文件格式标识符用作 Workbook 对象的 SaveAs 方法中的参数。*/
Application.ActiveWorkbook.SaveAs("C:\\temp\\myFile.xyz", Application.FileExportConverters.Item(1).FileFormat, null, null, null, false)
```

### FileExportConverter.Parent

返回一个 **Object** 类型的值，该值代表指定 **FileExportConverter** 对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FileExportConverter 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/FileExportConverter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ListObject 对象

## 简介

代表工作表中的表格。

## 说明

ListObject 对象是 ListObjects 集合的成员。 ListObjects 集合包含工作表上所有的列表对象。

使用 Worksheet 对象的 ListObjects 属性可返回一个 ListObjects 集合。

``` JavaScript
/*下例给活动工作簿的第一张工作表的默认 ListObject 对象添加新的 ListRow 对象。*/
function test(){
let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
Set oListCol = wrksht.ListObjects.Item(1).ListRows.Add()
}
```

## 方法

| 序号 | 名称                                       | 说明                                                                                                                                                                        |
|:----:|:-------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Delete](#ListObject.Delete)               | 删除 ListObject 对象并从工作表中清除单元格数据。                                                                                                                            |
|  2   | [ExportToVisio](#ListObject.ExportToVisio) | 将 ListObject 对象导出到 Visio。                                                                                                                                            |
|  3   | [Publish](#ListObject.Publish)             | 将 ListObject 对象发布到运行 Microsoft SharePoint Foundation 的服务器。                                                                                                     |
|  4   | [Refresh](#ListObject.Refresh)             | 从运行 Microsoft SharePoint Foundation 的服务器上检索列表的当前数据和架构。此方法只能被用于链接到 SharePoint 网站的列表。如果 SharePoint 网站不可用，调用此方法将返回错误。 |
|  5   | [Resize](#ListObject.Resize)               | Resize 方法允许 ListObject 对象在新区域内调整大小。不插入或移动单元格。                                                                                                     |
|  6   | [Unlink](#ListObject.Unlink)               | 从列表中删除与 Microsoft SharePoint Foundation 网站的链接。返回 Nothing 。                                                                                                  |
|  7   | [Unlist](#ListObject.Unlist)               | 从 ListObject 对象删除列表功能。使用此方法后，组成列表的单元格区域将变为常规的数据区域。                                                                                    |

## 属性

| 序号 | 名称                                                                   | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Active](#ListObject.Active)                                           | 返回一个 Boolean 类型的值，该值表示工作表中的 ListObject 对象是否是活动的，即活动的单元格是否在 ListObject 对象的区域之内。只读 Boolean 类型。                                                                                  |
|  2   | [AlternativeText](#ListObject.AlternativeText)                         | 返回或设置指定表的描述性（可选）文本字符串。可读/写                                                                                                                                                                             |
|  3   | [Application](#ListObject.Application)                                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  4   | [AutoFilter](#ListObject.AutoFilter)                                   | 使用“自动筛选”筛选一个列表。只读。                                                                                                                                                                                              |
|  5   | [Comment](#ListObject.Comment)                                         | 返回或设置与列表对象关联的批注。可读/写 String 类型。                                                                                                                                                                           |
|  6   | [Creator](#ListObject.Creator)                                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  7   | [DataBodyRange](#ListObject.DataBodyRange)                             | 返回 Range 对象，它代表表中值的范围，但不包括标题行。只读。                                                                                                                                                                     |
|  8   | [DisplayName](#ListObject.DisplayName)                                 | 返回或设置指定的 ListObject 对象的显示名称。可读/写 String 类型。                                                                                                                                                               |
|  9   | [DisplayRightToLeft](#ListObject.DisplayRightToLeft)                   | 如果指定的 ListObject 从右到左而不是从左到右显示，则此属性为 True 。如果对象从左到右显示，则此属性为 False 。 Boolean 类型，只读。                                                                                              |
|  10  | [HeaderRowRange](#ListObject.HeaderRowRange)                           | 返回一个 Range 对象，该对象表示 列表 ?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。） 标题行的区域。 Range 类型，只读。                                                                |
|  11  | [InsertRowRange](#ListObject.InsertRowRange)                           | 返回一个 Range 对象，该对象代表指定的 ListObject 对象的 “插入”行 （如果有）。 Range 类型，只读。                                                                                                                                |
|  12  | [ListColumns](#ListObject.ListColumns)                                 | 返回一个 ListColumns 集合，该集合代表 ListObject 对象中的所有列。只读。                                                                                                                                                         |
|  13  | [ListRows](#ListObject.ListRows)                                       | 返回一个 ListRows 对象，该对象代表 ListObject 对象中的所有数据行。只读。                                                                                                                                                        |
|  14  | [Name](#ListObject.Name)                                               | 返回或设置一个 String 值，它代表 ListObject 对象的名称。                                                                                                                                                                        |
|  15  | [Parent](#ListObject.Parent)                                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  16  | [QueryTable](#ListObject.QueryTable)                                   | 返回 QueryTable 对象，它为 ListObject 对象提供了一个列表服务器的链接。只读。                                                                                                                                                    |
|  17  | [Range](#ListObject.Range)                                             | 返回一个 Range 对象，它代表在以上列表中的指定列表对象所应用的区域。                                                                                                                                                             |
|  18  | [SharePointURL](#ListObject.SharePointURL)                             | 返回一个 String 类型的数值，该数值表示给定 ListObject 对象的 SharePoint 列表的 URL。 String 类型，只读。                                                                                                                        |
|  19  | [ShowAutoFilter](#ListObject.ShowAutoFilter)                           | 返回一个 Boolean 类型的数值，该数值指示是否显示“自动筛选”。 Boolean 类型，可读写。                                                                                                                                              |
|  20  | [ShowHeaders](#ListObject.ShowHeaders)                                 | 返回或设置是否显示指定的 ListObject 对象的标题信息。可读/写 Boolean 类型。                                                                                                                                                      |
|  21  | [ShowTableStyleColumnStripes](#ListObject.ShowTableStyleColumnStripes) | 返回或设置指定的 ListObject 对象是否使用“列条纹”表样式。可读/写 Boolean 类型。                                                                                                                                                  |
|  22  | [ShowTableStyleFirstColumn](#ListObject.ShowTableStyleFirstColumn)     | 返回或设置是否显示指定的 ListObject 对象的第一列。可读/写 Boolean 类型。                                                                                                                                                        |
|  23  | [ShowTableStyleLastColumn](#ListObject.ShowTableStyleLastColumn)       | 返回或设置是否显示指定的 ListObject 对象的最后一列。可读/写 Boolean 类型。                                                                                                                                                      |
|  24  | [ShowTableStyleRowStripes](#ListObject.ShowTableStyleRowStripes)       | 返回或设置指定的 ListObject 对象是否使用“行条纹”表样式。可读/写 Boolean 类型。                                                                                                                                                  |
|  25  | [ShowTotals](#ListObject.ShowTotals)                                   | 获取或设置一个 Boolean 类型的数值以指示“总计”行是否可见。 Boolean 类型，可读写。                                                                                                                                                |
|  26  | [Sort](#ListObject.Sort)                                               | 获取或设置 ListObject 集合的排序列和排序次序。                                                                                                                                                                                  |
|  27  | [SourceType](#ListObject.SourceType)                                   | 返回一个 XlListObjectSourceType 值，它代表列表的当前源。                                                                                                                                                                        |
|  28  | [Summary](#ListObject.Summary)                                         | 如果指定区域为分级显示的汇总行或汇总列，则该值为 True 。该区域应为一行或一列。 Variant 类型，只读。                                                                                                                             |
|  29  | [TableStyle](#ListObject.TableStyle)                                   | 获取或设置指定的 ListObject 对象的表样式。可读/写 Variant 类型。                                                                                                                                                                |
|  30  | [TotalsRowRange](#ListObject.TotalsRowRange)                           | 返回一个 Range ，代表指定的 ListObject 对象的“汇总”行（如果有）。只读。                                                                                                                                                         |
|  31  | [XmlMap](#ListObject.XmlMap)                                           | 返回一个 XmlMap 对象，该对象表示用于指定表格的架构映射。只读。                                                                                                                                                                  |

## 成员方法

### ListObject.Delete

删除 **ListObject** 对象并从工作表中清除单元格数据。

**语法**

**express.Delete ()**

\*express是一个代表 ListObject 对象的变量。

**说明**

如果列表链接到 SharePoint 网站，删除列表不会影响运行 SharePoint Foundation 的服务器上的数据。对本地列表所做的任何未提交的更改不会发送到 SharePoint 列表（系统不会警告这些未提交的更改将会丢失）。

### ListObject.ExportToVisio

将 **ListObject** 对象导出到 Visio。

**语法**

**express.ExportToVisio ()**

\*express是一个代表 ListObject 对象的变量。

### ListObject.Publish

将 **ListObject** 对象发布到运行 Microsoft SharePoint Foundation 的服务器。

**语法**

**express.Publish ( *Target* , *LinkSource* )**

\*express是一个代表 ListObject 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                       |
|--------------|-----------|----------|--------------------------------------------|
| *Target*     | 必选      | Variant  | 包含一个 String 值数组，如“备注”部分所述。 |
| *LinkSource* | 必选      | Boolean  |                                            |

**返回值**

一个字符串，代表 SharePoint 网站上已发布列表的 URL。

**说明**

*Target* 参数包含一个 **String** 元素数组，如下表中所述：

| 元素号 | 内容                    |
|--------|-------------------------|
| 0      | SharePoint 服务器的 URL |
| 1      | ListName（显示名称）    |
| 2      | 列表的说明。可选。      |

如果 **ListObject** 对象目前没有链接到 SharePoint 网站上的列表，在将 *LinkSource* 设置为 **True** 时，将在指定的 SharePoint 网站上创建新的列表。如果 **ListObject** 对象目前已经链接到 SharePoint 网站上的列表，在将 *LinkSource* 参数设置为 **True** 时，将替换现有的链接（只能将列表链接到一个 SharePoint 网站）。如果 **ListObject** 对象目前没有链接，在将 *LinkSource* 设置为 **False** 时，将保留 **ListObject** 对象不进行链接。如果 **ListObject** 对象目前已经链接到 SharePoint 网站，在将 *LinkSource* 设置为 **False** 时，将保持 **ListObject** 对象与当前 SharePoint 网站的链接。

### ListObject.Refresh

从运行 Microsoft SharePoint Foundation 的服务器上检索列表的当前数据和架构。此方法只能被用于链接到 SharePoint 网站的列表。如果 SharePoint 网站不可用，调用此方法将返回错误。

**语法**

**express.Refresh ()**

\*express是一个代表 ListObject 对象的变量。

**说明**

对 **Refresh** 方法的调用不会将更改提交到 ET 工作簿内的列表中。在 ET 中，当调用 **Refresh** 方法时，将放弃列表中未提交的更改。

### ListObject.Resize

**Resize** 方法允许 **ListObject** 对象在新区域内调整大小。不插入或移动单元格。

**语法**

**express.Resize ( *Range* )**

\*express是一个代表 ListObject 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明     |
|---------|-----------|----------|----------|
| *Range* | 必选      | Range    | 新区域。 |

**说明**

对于链接到正在运行 Microsoft SharePoint Foundation 的服务器的表格，可使用此方法调整列表大小，方法是提供一个仅在所包含的行数方面与 **ListObject** 的当前区域不同的 *Range* 参数。如果通过添加或删除列（在 *Range* 参数中）来调整链接到 SharePoint Foundation 的列表的大小，将导致运行时错误。

**示例**

``` JavaScript
/*下例使用了 Resize 方法来调整活动工作簿的 Sheet1 中默认 ListObject 对象的大小。*/
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let objListObj = wrksht.ListObjects.Item(1)
            
   objListObj.Resize (Range("A1:B10"))

}
```

### ListObject.Unlink

从列表中删除与 Microsoft SharePoint Foundation 网站的链接。返回 **Nothing** 。

**语法**

**express.Unlink ()**

\*express是一个代表 ListObject 对象的变量。

**说明**

调用此方法后，列表被取消链接，无法再恢复。

**示例**

``` JavaScript
/*下例从 SharePoint 网站上删除了列表链接。*/
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let objListObj = wrksht.ListObjects.Item(1)
   objListObj.Unlink()
}
```

### ListObject.Unlist

从 **ListObject** 对象删除列表功能。使用此方法后，组成列表的单元格区域将变为常规的数据区域。

**语法**

**express.Unlist ()**

\*express是一个代表 ListObject 对象的变量。

**说明**

运行此方法可保留工作表中的单元格数据、格式和公式。 **“汇总行”** 也保留不变。此方法可删除 Microsoft SharePoint Foundation 网站的所有链接。 **“自动筛选”** 也从列表中删除。

**示例**

``` JavaScript
/*下例将列表功能从工作表上的某个列表中删除。*/
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let objListObj = wrksht.ListObjects.Item(1) 
   objListObj.Unlist()
}
```

## 成员属性

### ListObject.Active

返回一个 **Boolean** 类型的值，该值表示工作表中的 **ListObject** 对象是否是活动的，即活动的单元格是否在 **ListObject** 对象的区域之内。只读 **Boolean** 类型。

**语法**

**express.Active**

\*express是一个代表 ListObject 对象的变量。

**说明**

因为 **ListObject** 对象没有 **Activate** 方法，所以只能通过激活 列表 ?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。） 中某个单元格区域的方式激活 **ListObject** 对象。

**示例**

``` JavaScript
/*下例激活了活动工作簿第一个工作表的默认 ListObject 对象的列表。调用 Range 对象的 Activate 方法时不提供单元格参数，将激活列表的整个区域。*/
function test(){
let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let objList = wrksht.ListObjects.Item(1)
    objList.Range.Activate()

    let MakeListActive = objList.Active
}
```

### ListObject.AlternativeText

返回或设置指定表的描述性（可选）文本字符串。可读/写

**语法**

**express.AlternativeText**

\*express是一个代表 ListObject 对象的变量。

**说明**

**AlternativeText** 属性的值对应于 **“可选文字”** 对话框（右键单击表、指向 **“表”** ，然后单击 **“可选文字”** 即可显示该对话框）中 **“标题”** 框的设置。

### ListObject.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ListObject 对象的变量。

**示例**

``` JavaScript
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### ListObject.AutoFilter

使用“自动筛选”筛选一个列表。只读。

**语法**

**express.AutoFilter**

\*express是一个代表 ListObject 对象的变量。

### ListObject.Comment

返回或设置与列表对象关联的批注。可读/写 **String** 类型。

**语法**

**express.Comment**

\*express是一个代表 ListObject 对象的变量。

**说明**

批注文本不能超过 255 个字符。

### ListObject.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ListObject 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ListObject.DataBodyRange

返回 **Range** 对象，它代表表中值的范围，但不包括标题行。只读。

**语法**

**express.DataBodyRange**

\*express是一个代表 ListObject 对象的变量。

**示例**

``` JavaScript
/*此示例选定列表中的活动数据区域。*/
function test(){
Application.Worksheets.Item("Sheet1").Activate()
ActiveWorksheet.ListObjects.Item(1).DataBodyRange.Select()
}
```

### ListObject.DisplayName

返回或设置指定的 **ListObject** 对象的显示名称。可读/写 **String** 类型。

**语法**

**express.DisplayName**

\*express是一个代表 ListObject 对象的变量。

### ListObject.DisplayRightToLeft

如果指定的 **ListObject** 从右到左而不是从左到右显示，则此属性为 **True** 。如果对象从左到右显示，则此属性为 **False** 。 **Boolean** 类型，只读。

**语法**

**express.DisplayRightToLeft**

\*express是一个代表 ListObject 对象的变量。

### ListObject.HeaderRowRange

返回一个 **Range** 对象，该对象表示 列表 ?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。） 标题行的区域。 **Range** 类型，只读。

**语法**

**express.HeaderRowRange**

\*express是一个代表 ListObject 对象的变量。

**示例**

``` JavaScript
/*下例激活了由活动工作簿第一个工作表中默认 ListObject 对象的 HeaderRowRange 属性指定的区域。*/
function test(){
let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let objList = wrksht.ListObjects.Item(1)
    let objListRng = objList.HeaderRowRange
    objListRng.Activate()
}
```

### ListObject.InsertRowRange

返回一个 **Range** 对象，该对象代表指定的 **ListObject** 对象的 “插入”行 （如果有）。 **Range** 类型，只读。

**语法**

**express.InsertRowRange**

\*express是一个代表 ListObject 对象的变量。

**说明**

如果由于 列表 ?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。） 处于非活动状态而使得“插入”行不可见，则返回 **Nothing** 对象。

**示例**

``` JavaScript
/*下例激活了由活动工作簿第一个工作表内默认 ListObject 对象中 InsertRowRange 指定的区域。*/
function test(){
  let wrksht = Application.ActiveWorkbook.Worksheets.Item(1)
    let objList = wrksht.ListObjects.Item(1)
    let objListRng = objList.InsertRowRange
    if(objListRng){
        ActivateInsertRow = false
    }
    else{
        objListRng.Activate()
        ActivateInsertRow = true
    }
}
```

### ListObject.ListColumns

返回一个 **ListColumns** 集合，该集合代表 **ListObject** 对象中的所有列。只读。

**语法**

**express.ListColumns**

\*express是一个代表 ListObject 对象的变量。

**示例**

``` JavaScript
/*下例显示了通过对 ListColumns 属性的调用所创建的 ListColumns 集合对象中第二列的名称。要让本代码运行，Sheet1 工作表必须包含一个表格。*/
function test(){
let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let objListObj = wrksht.ListObjects.Item(1)
    let objListCols = objListObj.ListColumns
    alert(objListCols.Item(2).Name)
}
```

### ListObject.ListRows

返回一个 **ListRows** 对象，该对象代表 **ListObject** 对象中的所有数据行。只读。

**语法**

**express.ListRows**

\*express是一个代表 ListObject 对象的变量。

**说明**

返回的 **ListRows** 对象不包括标题行、总计行或插入行。

**示例**

``` JavaScript
/*下例删除了由 ListRows 集合中的数字指定的行，该集合是通过调用 ListRows 属性创建的。*/
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let objListObj = wrksht.ListObjects.Item(1)
    let objListRows = objListObj.ListRows

    if((iRowNumber != 0) && (iRowNumber < objListRows.Count - 1)){
        objListRows.Item(iRowNumber).Delete()
    }
}
```

### ListObject.Name

返回或设置一个 **String** 值，它代表 **ListObject** 对象的名称。

**语法**

**express.Name**

\*express是一个代表 ListObject 对象的变量。

**说明**

此名称只用作 **ListObjects** 集合对象的 **Item** 属性的唯一标识符。只能通过对象模型设置此属性。

默认情况下，每个 **ListObject** 对象名称都以“List”开头，后面接着数字（无空格）。如果试图将 **Name** 属性设置为其他 **ListObject** 对象已经使用的名称，将产生运行时错误。

**示例**

``` JavaScript
/*以下示例显示活动工作簿的 sheet1 中的默认 ListObject 对象的名称。*/
function test(){
  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let oListObj = wrksht.ListObjects.Item(1)
   alert (oListObj.Name)  
}
```

### ListObject.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ListObject 对象的变量。

### ListObject.QueryTable

返回 **QueryTable** 对象，它为 **ListObject** 对象提供了一个列表服务器的链接。只读。

**语法**

**express.QueryTable**

\*express是一个代表 ListObject 对象的变量。

**示例**

下例创建一个到 SharePoint 网站的链接，并将名为 ` List1 ` 的 **ListObject** 对象发布到服务器。另外，还创建了对列表对象的 **QueryTable** 对象的引用，并将 **QueryTable** 对象的 **MaintainConnection** 属性设为 **True** ，以便保持服务器与 SharePoint 网站之间的连接。

``` JavaScript
function test(){
let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
let objListObj = wrksht.ListObjects.Item(1)

let arr = []
arr[0] = "0"
arr[1] = "http://myteam/project1"
arr[2] = "1"
arr[3] = "List1"
   
let strSTSConnection = objListObj.Publish(arr, true)
   
let objQryTbl = objListObj.QueryTable
   
objQryTbl.MaintainConnection = true
}
```

### ListObject.Range

返回一个 **Range** 对象，它代表在以上列表中的指定列表对象所应用的区域。

**语法**

**express.Range**

\*express是一个代表 ListObject 对象的变量。

### ListObject.SharePointURL

返回一个 **String** 类型的数值，该数值表示给定 **ListObject** 对象的 SharePoint 列表的 URL。 **String** 类型，只读。

**语法**

**express.SharePointURL**

\*express是一个代表 ListObject 对象的变量。

**说明**

如果列表未与 SharePoint 网站链接，则访问该属性会产生运行时错误。

**示例**

下例为 **Publish** 方法的 *Target* 参数设置了元素，以便将 **ListObject** 对象推送到 SharePoint 网站。此代码示例使用 **SharePointURL** 属性将 URL 分配给数组，使用 **Name** 属性分配列表的名称。然后使用 **Publish** 方法将数组中的信息传递到 SharePoint 网站。

``` JavaScript
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let objListObj = wrksht.ListObjects.Item(1)

    let arr = []
    arr[0] = "0"
    arr[1] = objListObj.SharePointURL
    arr[2] = "1"
    arr[3] = objListObj.Name
   
    let strSTSConnection = objListObj.Publish(arr, true)
}
```

### ListObject.ShowAutoFilter

返回一个 **Boolean** 类型的数值，该数值指示是否显示“自动筛选”。 **Boolean** 类型，可读写。

**语法**

**express.ShowAutoFilter**

\*express是一个代表 ListObject 对象的变量。

**说明**

对于新的 **ListObject** 对象， **ShowAutoFilter** 属性默认为 **True** 。

**示例**

``` JavaScript
/*下例显示了对活动工作簿的 Sheet 1 中默认列表进行的 ShowAutoFilter 属性设置。*/
function test(){
let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
let oListCol = wrksht.ListObjects.Item(1)
Debug.Print (oListCol.ShowAutoFilter)
}
```

### ListObject.ShowHeaders

返回或设置是否显示指定的 **ListObject** 对象的标题信息。可读/写 **Boolean** 类型。

**语法**

**express.ShowHeaders**

\*express是一个代表 ListObject 对象的变量。

### ListObject.ShowTableStyleColumnStripes

返回或设置指定的 **ListObject** 对象是否使用“列条纹”表样式。可读/写 **Boolean** 类型。

**语法**

**express.ShowTableStyleColumnStripes**

\*express是一个代表 ListObject 对象的变量。

### ListObject.ShowTableStyleFirstColumn

返回或设置是否显示指定的 **ListObject** 对象的第一列。可读/写 **Boolean** 类型。

**语法**

**express.ShowTableStyleFirstColumn**

\*express是一个代表 ListObject 对象的变量。

### ListObject.ShowTableStyleLastColumn

返回或设置是否显示指定的 **ListObject** 对象的最后一列。可读/写 **Boolean** 类型。

**语法**

**express.ShowTableStyleLastColumn**

\*express是一个代表 ListObject 对象的变量。

### ListObject.ShowTableStyleRowStripes

返回或设置指定的 **ListObject** 对象是否使用“行条纹”表样式。可读/写 **Boolean** 类型。

**语法**

**express.ShowTableStyleRowStripes**

\*express是一个代表 ListObject 对象的变量。

### ListObject.ShowTotals

获取或设置一个 **Boolean** 类型的数值以指示“总计”行是否可见。 **Boolean** 类型，可读写。

**语法**

**express.ShowTotals**

\*express是一个代表 ListObject 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例显示了对活动工作簿的 Sheet1 中默认列表的 ShowTotals 属性的当前设置。*/
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let oListObj = wrksht.ListObjects.Item(1)
   Debug.Print (oListObj.ShowTotals)
}
```

### ListObject.Sort

获取或设置 **ListObject** 集合的排序列和排序次序。

**语法**

**express.Sort**

\*express是一个代表 ListObject 对象的变量。

### ListObject.SourceType

返回一个 **XlListObjectSourceType** 值，它代表列表的当前源。

**语法**

**express.SourceType**

\*express是一个代表 ListObject 对象的变量。

**示例**

``` JavaScript
/*下面的示例代码返回 XlListObjectSourceType 常量，该常量表示活动工作簿的 Sheet1 上的默认列表的源。*/
function test(){
  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let oListObj = wrksht.ListObjects.Item(1)
   
   Debug.Print (oListObj.SourceType)
}
```

### ListObject.Summary

如果指定区域为分级显示的汇总行或汇总列，则该值为 **True** 。该区域应为一行或一列。 **Variant** 类型，只读。

**语法**

**express.Summary**

\*express是一个代表 ListObject 对象的变量。

**示例**

``` JavaScript
/**/
function test(){
let rows = Application.Worksheets.Item("Sheet1").Rows.Item(4)
if(rows.Summary == true){
    rows.Font.Bold = true
    rows.Font.Italic = true
}
}
```

### ListObject.TableStyle

获取或设置指定的 **ListObject** 对象的表样式。可读/写 **Variant** 类型。

**语法**

**express.TableStyle**

\*express是一个代表 ListObject 对象的变量。

### ListObject.TotalsRowRange

返回一个 **Range** ，代表指定的 **ListObject** 对象的“汇总”行（如果有）。只读。

**语法**

**express.TotalsRowRange**

\*express是一个代表 ListObject 对象的变量。

**示例**

``` JavaScript
/*下面的示例代码返回了活动工作簿的 Sheet1 中默认列表的“汇总”行的地址。该代码可在“汇总”行尚未显示的情况下显示该行。*/
function test(){
let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let objListObj = wrksht.ListObjects.Item(1)
    objListObj.ShowTotals = true
    alert (objListObj.TotalsRowRange.Address())
}
```

### ListObject.XmlMap

返回一个 **XmlMap** 对象，该对象表示用于指定表格的架构映射。只读。

**语法**

**express.XmlMap**

\*express是一个代表 ListObject 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ListObject](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ListRows 对象

## 简介

指定的 ListObject 对象中 ListRow 对象的集合。

## 说明

每一个 ListRow 对象都代表表格中的一行。

使用 ListObject 对象的 ListRows 属性可返回 ListRows 集合。下例给工作簿的第一张工作表的默认 ListObject 对象添加新行。由于未指定位置，因此在表格结束处添加一个新行。

``` JavaScript
let myNewRow = Application.Worksheets.Item(1).ListObjects.Item(1).ListRows.Add()
```

## 方法

| 序号 | 名称                 | 说明                                         |
|:----:|:---------------------|:---------------------------------------------|
|  1   | [Add](#ListRows.Add) | 向指定的 ListObject 所代表的表格中添加新行。 |

## 属性

| 序号 | 名称                                 | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ListRows.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#ListRows.Count)             | 返回一个 Integer 值，它代表集合中对象的数量。                                                                                                                                                                                   |
|  3   | [Creator](#ListRows.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Item](#ListRows.Item)               | 从集合中返回一个对象。                                                                                                                                                                                                          |
|  5   | [Parent](#ListRows.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### ListRows.Add

向指定的 **ListObject** 所代表的表格中添加新行。

**语法**

**express.Add ( *Position* , *AlwaysInsert* )**

\*express是一个代表 ListRows 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                       |
|----------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Position*     | 可选      | Variant  | Integer 类型。指定新行的相对位置。                                                                                                                                                                                                                                                                         |
| *AlwaysInsert* | 可选      | Variant  | Boolean。指定在插入新行时，是否始终移动表格中最后一行下面的单元格中的数据，而不考虑表格下面的行是否为空行。如果为 True，则表格下面的单元格将下移一行。如果为 False，并且表格下面的行为空行，则表格将扩展以占用该空行而不移动它下面的单元格；但如果表格下面的行包含数据，则在插入新行时，这些单元格将下移。 |

**返回值**

一个代表新行的 ListRow 对象。

**说明**

如果不指定 *Position* ，新行将添加到底部。如果不指定 *AlwaysInsert* ，表格下面的单元格将下移一行（与指定 **True** 的作用相同）。

**示例**

``` JavaScript
/*以下示例向工作簿的第一张工作表中的默认 ListObject 对象添加新行。由于未指定位置，新行将添加到列表的底部。*/
let myNewColumn = Application.ActiveWorkbook.Worksheets.Item(1).ListObject.Item(1).ListRows.Add()
```

## 成员属性

### ListRows.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ListRows 对象的变量。

**示例**

``` JavaScript
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### ListRows.Count

返回一个 **Integer** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 ListRows 对象的变量。

### ListRows.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ListRows 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ListRows.Item

从集合中返回一个对象。

**语法**

**express.Item**

\*express是一个代表 ListRows 对象的变量。

### ListRows.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ListRows 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ListRows](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Panes 对象

## 简介

指定的窗口中所有 Pane 对象的集合。

## 说明

Pane 对象只对于工作表和 ET 4.0 宏表存在。

使用 Panes 属性返回 Panes 集合。下例中，如果当前活动窗口包含若干窗格，则冻结这些窗格。

``` JavaScript
if(Application.ActiveWindow.Panes.Count > 1) { 
    Application.ActiveWindow.FreezePanes = true
}
```

使用 Panes ( *index* )（其中 *index* 是窗格索引号）可返回一个 Pane 对象。下例滚动 Sheet1 所在窗口中左上角的窗格。

``` JavaScript
function test(){
    Application.Worksheets.Item("sheet1").Activate()
    Application.Windows.Item(1).Panes.Item(1).LargeScroll(1)
}
```

## 方法

| 序号 | 名称                | 说明                   |
|:----:|:--------------------|:-----------------------|
|  1   | [Item](#Panes.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                              | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Panes.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Panes.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Panes.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Panes.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Panes.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Panes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明         |
|---------|-----------|----------|--------------|
| *Index* | 必选      | Long     | 对象的索引号 |

**示例**

``` JavaScript
function test(){  
  /*此示例拆分第一张工作表所在的窗口，然后滚动窗口左下角的窗格，直至第五行到达此窗格的顶部。*/
  Application.Worksheets.Item(1).Activate()
  Application.ActiveWindow.Split = true
  Application.ActiveWindow.Panes.Item(3).ScrollRow = 5
}
```

## 成员属性

### Panes.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Panes 对象的变量。

**示例**

``` JavaScript
function test(){  
  /*本示例显示一条有关创建 myObject 的应用程序的消息。*/
  if(Application.ActiveWorkbook.Application.Value == "ET") {
      alert("This is an ET Application object.")
  }                             
  else {
      alert("This is not an ET Application object.")
  }
}
```

### Panes.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Panes 对象的变量。

### Panes.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Panes 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Panes.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Panes 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Panes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotCell 对象

## 简介

代表数据透视表中的一个单元格。

## 说明

使用 Range 集合的 PivotCell 属性可返回一个 PivotCell 对象。

返回 PivotCell 对象后，可以使用 ColumnItems 或 RowItems 属性来确定 PivotItems 集合，对应于代表所选编号的列轴或行轴上的项目。下例使用 PivotCell 对象的 ColumnItems 属性来返回一个 PivotItemList 集合。

返回了 PivotCell 对象后，可以使用 PivotCellType 属性来确定某个特定区域是什么单元格类型。下例确定数据透视表中的单元格 A5 是否是数据项目并通知用户。此示例假定活动工作表上存在数据透视表而且单元格 A5 包含在该数据透视表中。如果单元格 A5 不在数据透视表中，该示例会处理运行时错误。

``` JavaScript
/**/
function test(){
try {
        // Determine if cell A5 is a data item in the PivotTable.
        if(Application.Range("A5").PivotCell.PivotCellType == xlPivotCellValue) {
            alert ("The PivotCell at A5 is a data item.")
        }
        else {
            alert ("The PivotCell at A5 is a data item.")
        }
    }
                                    
    catch(exception){
        alert("The chosen cell is not in a PivotTable.")
    }
}
```

此示例确定单元格 B5 的数据项所在的列字段。然后判断列字段标题是否与“Inventory”相匹配，并通知用户。此示例假定数据透视表位于活动工作表上，并且工作表的 B 列包含数据透视表的列字段。

``` JavaScript
/**/
function test(){
  // Determine if there is a match between the item and column field.
    if(Application.Range("B5").PivotCell.ColumnItems.Item(1) == "Inventory") {
        alert ("Item in B5 is a member of the 'Inventory' column field.")
    }
    else {
        alert ("Item in B5 is not a member of the 'Inventory' column field.")
    }
}
```

## 方法

| 序号 | 名称                                      | 说明                                 |
|:----:|:------------------------------------------|:-------------------------------------|
|  1   | [DiscardChange](#PivotCell.DiscardChange) | 放弃对数据透视表中指定单元格的更改。 |

## 属性

| 序号 | 名称                                                        | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#PivotCell.Application)                       | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [CellChanged](#PivotCell.CellChanged)                       | 返回自创建数据透视表以来，或上次执行提交操作以来，数据透视表值单元格是否经过了编辑或重新计算。只读                                                                                                                              |
|  3   | [ColumnItems](#PivotCell.ColumnItems)                       | 返回一个 PivotItemList 集合，该集合对应于表示选定区域的列轴上的项目。                                                                                                                                                           |
|  4   | [Creator](#PivotCell.Creator)                               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [CustomSubtotalFunction](#PivotCell.CustomSubtotalFunction) | 返回 PivotCell 对象的自定义分类汇总函数字段的设置。 XlConsolidationFunction 类型，只读。                                                                                                                                        |
|  6   | [DataField](#PivotCell.DataField)                           | 返回一个 PivotField 对象，该对象与所选数据字段相对应。                                                                                                                                                                          |
|  7   | [DataSourceValue](#PivotCell.DataSourceValue)               | 为数据透视表中经过编辑的单元格，返回上次从数据源检索到的值。只读                                                                                                                                                                |
|  8   | [MDX](#PivotCell.MDX)                                       | 返回一个元组，该元组提供了使用 OLAP 数据源的数据透视表中指定值单元格的完整 MDX 坐标。只读                                                                                                                                       |
|  9   | [Parent](#PivotCell.Parent)                                 | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  10  | [PivotCellType](#PivotCell.PivotCellType)                   | 返回一个 XlPivotCellType 常量，该常量标识单元格所对应的数据透视表实体。只读。                                                                                                                                                   |
|  11  | [PivotColumnLine](#PivotCell.PivotColumnLine)               | 在列上返回特定 PivotCell 对象的 PivotLine 。只读 PivotLine 类型。                                                                                                                                                               |
|  12  | [PivotField](#PivotCell.PivotField)                         | 返回一个 PivotField 对象，它代表指定区域左上角所在的数据透视表字段。                                                                                                                                                            |
|  13  | [PivotItem](#PivotCell.PivotItem)                           | 返回一个 PivotItem 对象，它代表指定区域左上角所在的数据透视表项。                                                                                                                                                               |
|  14  | [PivotRowLine](#PivotCell.PivotRowLine)                     | 在行上返回特定 PivotCell 对象的 PivotLine。只读 PivotLine 类型。                                                                                                                                                                |
|  15  | [PivotTable](#PivotCell.PivotTable)                         | 返回一个 PivotTable 对象，它代表与 PivotCell 相关联的数据透视表。                                                                                                                                                               |
|  16  | [Range](#PivotCell.Range)                                   | 返回一个 Range 对象，它代表指定的 PivotCell 所应用的区域。                                                                                                                                                                      |
|  17  | [RowItems](#PivotCell.RowItems)                             | 返回一个 PivotItemList 集合，该集合对应于分类轴上表示选定单元格的项目。                                                                                                                                                         |

## 成员方法

### PivotCell.DiscardChange

放弃对数据透视表中指定单元格的更改。

**语法**

**express.DiscardChange ()**

\*express是一个代表 PivotCell 对象的变量。

**返回值**

Nothing

**说明**

对于基于 OLAP 数据源的数据透视表中的单元格， **DiscardChange** 方法删除在该指定单元格中输入的值或公式，然后运行数据透视表更新从数据源检索最新值。该指定单元格的对应数据源值将由此方法设置为 **NULL** 。该方法还将从 ET 更改列表中删除对此单元格所做的所有更改，以使这些更改不再包含在对该数据源的 **UPDATE CUBE** 语句中。此方法还将重新创建在事务中发生的所有其他更改的会话状态。这包括执行 **ROLLBACK TRANSACTION** 语句，然后执行 **UPDATE CUBE** 语句应用除该指定单元格之外的所有更改。对于基于非 OLAP 数据源的数据透视表中的单元格，此方法清除编辑过的单元格。

## 成员属性

### PivotCell.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotCell 对象的变量。

**示例**

``` JavaScript
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### PivotCell.CellChanged

返回自创建数据透视表以来，或上次执行提交操作以来，数据透视表值单元格是否经过了编辑或重新计算。只读

**语法**

**express.CellChanged**

\*express是一个代表 PivotCell 对象的变量。

**说明**

table { word-break:break-all; }

**CellChanged** 属性的值默认为 **xlCellNotChanged** 。

对于使用非 OLAP 数据源的数据透视表，此属性的值只能为 **xlCellNotChanged** 或 **xlCellChanged** 。对于未编辑的单元格，其值为 **xlCellNotChanged** ；对于编辑过的单元格，其值为 **xlCellChanged** 。放弃更改将把其值设置为 **xlCellNotChanged** 。

应用并保存更改只适用于使用 OLAP 数据源的数据透视表。下面的列表是 **CellChange** 属性各种可能的状态的说明，它们中适用于使用 OLAP 数据源的数据透视表：

- **xlCellNotChanged** - 自创建数据透视表以来，或者自上次执行保存或放弃更改操作以来，单元格尚未编辑或重新计算过（如果单元格包含公式）。

- **xlCellChanged** - 自创建数据透视表以来，或者自上次执行应用更改或保存更改操作以来，单元格已经过编辑或重新计算，但是尚未应用该更改（尚未对该更改运行 **UPDATE CUBE** 语句）。

- **xlCellChangeApplied** - 自创建数据透视表以来，或者自上次执行应用更改、保存更改或放弃更改操作以来，单元格已经过编辑或重新计算，并且已经应用了该更改（已经对该更改运行 **UPDATE CUBE** 语句）。

下表说明不同的用户操作如何影响使用 OLAP 数据源的数据透视表中 **CellChanged** 属性的设置。

| 用户操作                                                              | 无公式单元格的 CellChanged 属性的设置                            | 有公式单元格的 CellChanged 属性的设置                           |
|-----------------------------------------------------------------------|------------------------------------------------------------------|-----------------------------------------------------------------|
| 在一个或多个单元格中输入值或公式。                                    | 对这些单元格，设置为 **xlCellChanged** 。                        | 对这些单元格，设置为 **xlCellChanged** 。                       |
| 重新计算含公式的一个或多个单元格（手动计算 (F9)，或由 ET 自动计算）。 | 不适用                                                           | 对这些单元格，设置为 **xlCellChanged** 。                       |
| 保存（提交）更改。                                                    | 对所有经过编辑的无公式单元格，设置为 **xlCellNotChanged** 。     | 对所有经过编辑的有公式单元格，设置为 **xlCellChangeApplied** 。 |
| 放弃所有更改。                                                        | 对所有经过编辑的无公式单元格，设置为 **xlCellNotChanged** 。     | 对所有经过编辑的有公式单元格，设置为 **xlCellNotChanged** 。    |
| 放弃单个单元格中的更改。                                              | 只对该单元格设置为 **xlCellNotChanged** 。                       | 只对该单元格设置为 **xlCellNotChanged** 。                      |
| 一次操作清除多个单元格。                                              | 对这些单元格设置为 **xlCellNotChanged** 。                       | 对这些单元格设置为 **xlCellNotChanged** 。                      |
| 清除一个单元格。                                                      | 只对该单元格设置为 **xlCellNotChanged** 。                       | 只对该单元格设置为 **xlCellNotChanged** 。                      |
| 应用值之前，执行撤消操作，将该值改回先前编辑的值。                    | 对所有经过编辑的无公式单元格，设置为 **xlCellChanged** 。        | 对所有经过编辑的有公式单元格，设置为 **xlCellChanged** 。       |
| 应用值之后，执行撤消操作，将该值改回先前编辑的值。                    | 对所有经过编辑的无公式单元格，设置为 **xlCellChangedApplied** 。 | 对所有经过编辑的有公式单元格，设置为 **xlCellChangeApplied** 。 |

### PivotCell.ColumnItems

返回一个 **PivotItemList** 集合，该集合对应于表示选定区域的列轴上的项目。

**语法**

**express.ColumnItems**

\*express是一个代表 PivotCell 对象的变量。

**示例**

本示例判断单元格 B5 中的数据项是否在第一个列字段中的“Inventory”项目之下，并通知用户。本示例假定数据透视表位于活动工作表上，并且第 B 列包含数据透视表的列字段。

``` JavaScript
function test(){
// Determine if there is a match between the item and column field.
if(Application.Range("B5").PivotCell.ColumnItems.Item(1) == "Inventory") {
    MsgBox("Item in B5 is a member of the 'Inventory' column field.")
}
else {
    MsgBox("Item in B5 is not a member of the 'Inventory' column field.")
}
}
```

### PivotCell.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotCell 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotCell.CustomSubtotalFunction

返回 **PivotCell** 对象的自定义分类汇总函数字段的设置。 **XlConsolidationFunction** 类型，只读。

**语法**

**express.CustomSubtotalFunction**

\*express是一个代表 PivotCell 对象的变量。

**说明**

|                                                                             |
|-----------------------------------------------------------------------------|
| **XlConsolidationFunction** 可为以下 **XlConsolidationFunction** 常量之一。 |
| **xlAverage**                                                               |
| **xlCount**                                                                 |
| **xlCountNums**                                                             |
| **xlMax**                                                                   |
| **xlMin**                                                                   |
| **xlProduct**                                                               |
| **xlStDev**                                                                 |
| **xlStDevP**                                                                |
| **xlSum**                                                                   |
| **xlUnknown**                                                               |
| **xlVar**                                                                   |
| **xlVarP**                                                                  |

如果 **PivotCell** 对象类型不是自定义分类汇总，那么 **CustomSubtotalFunction** 属性将返回一个错误。该属性只应用于 非 OLAP 源数据 ?（非 OLAP 源数据：非 OLAP 源数据是数据透视表或数据透视图使用的基本数据，该数据来自 OLAP 数据库之外的源。这些数据源包括关系数据库、 ET 工作表中的清单以及文本文件数据库。） 。

**示例**

``` JavaScript
/*本示例确定单元格 C20 是否包含使用计数合并函数的自定义分类汇总函数，并通知用户。本示例假定数据透视表位于活动工作表上。*/
function test(){
  try{
        // Determine if custom subtotal function is a count function.
        if(Application.Range("C20").PivotCell.CustomSubtotalFunction == xlCount) {
            MsgBox("The custom subtotal function is a Count.")
        }
        else
            alert("The custom subtotal function is not a Count.")
        }
                                    
    catch(exception) {
        alert("The selected cell is not a custom subtotal function.")
    }
}
```

### PivotCell.DataField

返回一个 **PivotField** 对象，该对象与所选数据字段相对应。

**语法**

**express.DataField**

\*express是一个代表 PivotCell 对象的变量。

**说明**

如果 **PivotCell** 对象不是 **XlPivotCellType** 允许的常量： **xlPivotCellTypeDataField** 、 **xlPivotCellTypeSubtotal** 或 **xlPivotCellTypeGrandTotal** 之一，那么该属性将返回一个错误。

**示例**

本示例确定单元格 L10 是否在数据透视表的数据字段中，然后通过通知用户来返回对应于该单元格的数据透视表字段，或处理运行错误。本示例假定数据透视表位于活动工作表上。

``` JavaScript
function test(){
  try{
        alert(Application.Range("L10").PivotCell.DataField)
    }
                                    
    catch(exception) {
        alert("The selected range is not in the data field of the PivotTable.")
    }
}
```

### PivotCell.DataSourceValue

为数据透视表中经过编辑的单元格，返回上次从数据源检索到的值。只读

**语法**

**express.DataSourceValue**

\*express是一个代表 PivotCell 对象的变量。

**说明**

每次编辑数据透视表值区域中的某个单元格时，在编辑发生之前， **DataSourceValue** 属性将保存上次从数据源检索到的值。对于尚未编辑的数据透视表值单元格，或者尚未显式检索数据源值的数据透视表值单元格，此属性将返回 **NULL** 。对于使用 OLAP 数据源的数据透视表， **DataSourceValue** 属性的值将从单独的连接检索，以确保其不包含用户可能已执行的任何回写操作的值。

如果对数据透视表值区域之外的单元格读取其 **DataSourceValue** 属性，那么会生成运行时错误。

### PivotCell.MDX

返回一个元组，该元组提供了使用 OLAP 数据源的数据透视表中指定值单元格的完整 MDX 坐标。只读

**语法**

**express.MDX**

\*express是一个代表 PivotCell 对象的变量。

**说明**

元组中由 **MDX** 属性返回的维度包含行坐标和列坐标，以及报表筛选坐标。对于数据透视表值区域之外的单元格，以及数据透视表之外的单元格，访问此属性将生成运行时错误。对于在报表筛选字段中有多项选择的数据透视表，访问此属性也将生成运行时错误。

### PivotCell.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotCell 对象的变量。

### PivotCell.PivotCellType

返回一个 **XlPivotCellType** 常量，该常量标识单元格所对应的数据透视表实体。只读。

**语法**

**express.PivotCellType**

\*express是一个代表 PivotCell 对象的变量。

**说明**

|                                                                                            |
|--------------------------------------------------------------------------------------------|
| **XlPivotCellType** 可为以下 **XlPivotCellType** 常量之一。                                |
| **xlPivotCellBlankCell?** 数据透视表中的一个结构空白单元格。                               |
| **xlPivotCellCustomSubtotal** 自定义分类汇总的行或列区域中的一个单元格。                   |
| **xlPivotCellDataField?** 一个数据字段标签（不是 “数据”按钮）。                            |
| **xlPivotCellDataPivotField** “数据”按钮。                                                 |
| **xlPivotCellGrandTotal?** 总计行或列区域中的一个单元格。                                  |
| **xlPivotCellPageFieldItem?** 显示页字段的选定项的单元格。                                 |
| **xlPivotCellPivotField?** 字段的按钮（不是 “数据”按钮）。                                 |
| **xlPivotCellPivotItem?** 不是分类汇总、总计、自定义分类汇总或空行的行或列区域中的单元格。 |
| **xlPivotCellSubtotal?** 分类汇总的行或列区域中的单元格。                                  |
| **xlPivotCellValue?** 数据区区域（除空行以外）中的任一单元格。                             |

**示例**

``` JavaScript
/*本示例确定数据透视表中的单元格 A5 是否是一个数据项，并通知用户。本示例假定数据透视表位于活动工作表上。如果单元格 A5 不在数据透视表中，则本示例处理运行错误。*/
function test(){
 try{
        // Determine if cell A5 is a data item in the PivotTable.
        if(Application.Range("A5").PivotCell.PivotCellType == xlPivotCellValue) {
            alert ("The cell at A5 is a data item.")
        }
        else {
        alert("The cell at A5 is not a data item.")
        }
    }
    catch(exception) {
        alert ("The chosen cell is not in a PivotTable." )
    }
}
```

### PivotCell.PivotColumnLine

在列上返回特定 **PivotCell** 对象的 **PivotLine** 。只读 **PivotLine** 类型。

**语法**

**express.PivotColumnLine**

\*express是一个代表 PivotCell 对象的变量。

**说明**

如果 PivotCell 在行上，则 **PivotColumnLine** 属性返回一个运行时错误。

如果 PivotCell 在列上，则 **PivotColumnLine** 属性返回列 **PivotLine** 对象。

如果 PivotCell 在数据区域中，则 **PivotColumnLine** 属性返回对应的列 **PivotLine** 对象。

### PivotCell.PivotField

返回一个 **PivotField** 对象，它代表指定区域左上角所在的数据透视表字段。

**语法**

**express.PivotField**

\*express是一个代表 PivotCell 对象的变量。

**示例**

``` JavaScript
/*此示例显示包含活动单元格的数据透视表字段的名称。*/
function test(){
Worksheets.Item("Sheet1").Activate()
alert("The active cell is in the field " + ActiveCell.PivotField.Name)
}
```

### PivotCell.PivotItem

返回一个 **PivotItem** 对象，它代表指定区域左上角所在的数据透视表项。

**语法**

**express.PivotItem**

\*express是一个代表 PivotCell 对象的变量。

**示例**

``` JavaScript
/*此示例显示包含 Sheet1 中活动单元格的数据透视表数据项的名称.*/
function test(){
Worksheets.Item("Sheet1").Activate()
alert("The active cell is in the field " + ActiveCell.PivotField.Name)
}
```

### PivotCell.PivotRowLine

在行上返回特定 **PivotCell** 对象的 PivotLine。只读 **PivotLine** 类型。

**语法**

**express.PivotRowLine**

\*express是一个代表 PivotCell 对象的变量。

**说明**

如果 PivotCell 在行上，则 **PivotRowLine** 返回行的 **PivotLine** 对象。

如果 PivotCell 在列上，则 **PivotRowLine** 返回一个运行时错误。

如果 PivotCell 在数据区域中，则 **PivotRowLine** 返回对应的行的 **PivotLine** 对象。

### PivotCell.PivotTable

返回一个 **PivotTable** 对象，它代表与 PivotCell 相关联的数据透视表。

**语法**

**express.PivotTable**

\*express是一个代表 PivotCell 对象的变量。

**示例**

此示例将 Sheet1 上的数据透视表的当前页设置为名为“Canada”的页。

``` JavaScript
let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
pvtTable.PivotFields("Country").CurrentPage = "Canada"
```

此示例确定与“Sales”图表相关联的数据透视表，然后将名为“Oregon”的页设置为数据透视表的当前页。

``` JavaScript
let objPT = ActiveSheet.Charts("Sales").PivotLayout.PivotTable
objPT.PivotFields.Item("State").CurrentPageName = "Oregon"
```

### PivotCell.Range

返回一个 **Range** 对象，它代表指定的 PivotCell 所应用的区域。

**语法**

**express.Range**

\*express是一个代表 PivotCell 对象的变量。

**示例**

下例将应用于“Crew”工作表的“自动筛选”的地址保存在变量中。

``` JavaScript
let rAddress = Worksheets.Item("Crew").AutoFilter.Range.Address()
```

此示例滚动工作簿窗口，直至超链接区域位于活动窗口的左上角。

``` JavaScript
Workbooks.Item(1).Activate()
let hr = ActiveSheet.Hyperlinks.Item(1).Range
ActiveWindow.ScrollRow = hr.Row
ActiveWindow.ScrollColumn = hr.Column
```

### PivotCell.RowItems

返回一个 **PivotItemList** 集合，该集合对应于分类轴上表示选定单元格的项目。

**语法**

**express.RowItems**

\*express是一个代表 PivotCell 对象的变量。

**示例**

本示例判断单元格 B5 中的数据项是否在第一个行字段中的“Inventory”项之下，并通知用户。本示例假定数据透视表位于活动工作表上，并且工作表的列 B 包含数据透视表的行项目。

``` JavaScript
function test(){
 // Determine if there is a match between the item and row field.
                                    
    if(Application.Range("B5").PivotCell.RowItems.Item(1) == "Inventory" ) {
        alert("Cell B5 is a member of the 'Inventory' row field.")
    }                                   
    else {
        alert("Cell B5 is not a member of the 'Inventory' row field.")
    }
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotCell](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Ranges 对象

## 简介

由 Range 对象组成的集合。

## 属性

| 序号 | 名称                               | 说明                                                                                                                                                               |
|:----:|:-----------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Ranges.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Count](#Ranges.Count)             | 返回集合中对象的数目。只读 Long 类型。                                                                                                                             |
|  3   | [Creator](#Ranges.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Item](#Ranges.Item)               | 返回一个 Range 对象，该对象代表工作簿中的项的区域。只读。                                                                                                          |
|  5   | [Parent](#Ranges.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |

## 成员属性

### Ranges.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Ranges 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### Ranges.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 Ranges 对象的变量。

### Ranges.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Ranges 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Ranges.Item

返回一个 **Range** 对象，该对象代表工作簿中的项的区域。只读。

**语法**

**express.Item**

\*express是一个代表 Ranges 对象的变量。

### Ranges.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Ranges 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Ranges](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Series 对象

## 简介

代表图表上的系列。

## 说明

Series 对象是 SeriesCollection 集合的成员。

使用 SeriesCollection ( *index* )（其中 *index* 是系列索引号或名称）可返回一个 Series 对象。下例设置 Sheet1 上嵌入式图表一中第一个系列的内部颜色。

系列索引号指明了系列添加到图表中的顺序。 ` SeriesCollection(1) ` 是第一个添加到图表中的系列，而 ` SeriesCollection(SeriesCollection.Count) ` 是最后一个添加到图表中的系列。

``` JavaScript
Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.SeriesCollection(1).Interior.Color = (255, 0, 0)
```

## 方法

| 序号 | 名称                                       | 说明                                                                                                |
|:----:|:-------------------------------------------|:----------------------------------------------------------------------------------------------------|
|  1   | [ApplyDataLabels](#Series.ApplyDataLabels) | 向系列应用数据标签。                                                                                |
|  2   | [AxisGroup](#Series.AxisGroup)             | 返回或设置指定系列的组。可读/写                                                                     |
|  3   | [ClearFormats](#Series.ClearFormats)       | 清除对象的格式设置。                                                                                |
|  4   | [Copy](#Series.Copy)                       | 如果系列具有图片填充，此方法将把图片复制到剪贴板。                                                  |
|  5   | [DataLabels](#Series.DataLabels)           | 返回代表数据系列中单个数据标签（ DataLabel 对象）或所有数据标签的集合（ DataLabels 集合）的对象。   |
|  6   | [Delete](#Series.Delete)                   | 删除对象。                                                                                          |
|  7   | [ErrorBar](#Series.ErrorBar)               | 将误差线应用于系列。 Variant 类型。                                                                 |
|  8   | [Paste](#Series.Paste)                     | 从剪贴板中粘贴图片，将其作为所选系列的标志。                                                        |
|  9   | [Points](#Series.Points)                   | 返回一个对象，该对象表示数据系列中单个数据点（ Point 对象）或所有数据点的集合（ Points 集合）。只读 |
|  10  | [Select](#Series.Select)                   | 选择对象。                                                                                          |
|  11  | [Trendlines](#Series.Trendlines)           | 返回一个对象，该对象表示单个趋势线（ Trendline 对象）或所有趋势线的集合（ Trendlines 集合）。       |

## 属性

| 序号 | 名称                                                             | 说明                                                                                                                                                                                                                                 |
|:----:|:-----------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Series.Application)                               | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。      |
|  2   | [ApplyPictToEnd](#Series.ApplyPictToEnd)                         | 如果图片应用于系列中数据点或所有数据点之后，则为 True 。 Boolean 类型，可读写。                                                                                                                                                      |
|  3   | [ApplyPictToFront](#Series.ApplyPictToFront)                     | 如果图片应用于系列中数据点或所有数据点之前，则为 True 。 Boolean 类型，可读写。                                                                                                                                                      |
|  4   | [ApplyPictToSides](#Series.ApplyPictToSides)                     | 如果将图片应用于系列中某个数据点或所有数据点的旁边，则为 True 。 Boolean 类型，可读写。                                                                                                                                              |
|  5   | [BarShape](#Series.BarShape)                                     | 返回或设置用于三维条形图或柱形图的形状。 XlBarShape 类型，可读写。                                                                                                                                                                   |
|  6   | [BubbleSizes](#Series.BubbleSizes)                               | 返回或设置一个字符串，该字符串引用包含气泡图 x 值、y 值和大小数据的工作表单元格。在返回单元格引用时，它将返回以 A1 样式注释描述单元格的字符串。要为气泡图设置大小数据，必须使用 R1 样式注释。仅适用于气泡图。 Variant 类型，可读写。 |
|  7   | [ChartType](#Series.ChartType)                                   | 返回或设置图表类型。 XlChartType 类型，可读写。                                                                                                                                                                                      |
|  8   | [Creator](#Series.Creator)                                       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                           |
|  9   | [ErrorBars](#Series.ErrorBars)                                   | 返回一个 ErrorBars 对象，该对象表示数据系列中的误差线。只读。                                                                                                                                                                        |
|  10  | [Explosion](#Series.Explosion)                                   | 返回或设置饼图或圆环图的扇区的分离程度值。如果没有分离，则返回 0（零），即扇区的中心与饼图中心重合。 Long 型，可读写。                                                                                                               |
|  11  | [Format](#Series.Format)                                         | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                        |
|  12  | [Formula](#Series.Formula)                                       | 返回或设置一个 String 值，它代表 A1 样式表示法和宏语言中的对象的公式。                                                                                                                                                               |
|  13  | [FormulaLocal](#Series.FormulaLocal)                             | 返回或设置指定对象的公式，使用用户语言 A1 格式引用。 String 型，可读写。                                                                                                                                                             |
|  14  | [FormulaR1C1](#Series.FormulaR1C1)                               | 返回或设置指定对象的公式，使用宏语言 R1C1 格式符号表示。 String 型，可读写。                                                                                                                                                         |
|  15  | [FormulaR1C1Local](#Series.FormulaR1C1Local)                     | 返回或设置指定对象的公式，使用用户语言 R1C1 格式符号表示。 String 型，可读写。                                                                                                                                                       |
|  16  | [Has3DEffect](#Series.Has3DEffect)                               | 如果数据系列有三维外观，则为 True 。可读/写 Boolean 类型。                                                                                                                                                                           |
|  17  | [HasDataLabels](#Series.HasDataLabels)                           | 如果数据系列具有数据标签，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                |
|  18  | [HasErrorBars](#Series.HasErrorBars)                             | 如果序列有错误条时，该值为 True 。该属性对三维图表无效。 Boolean 类型，可读写。                                                                                                                                                      |
|  19  | [HasLeaderLines](#Series.HasLeaderLines)                         | 如果数据系列有引导线，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                    |
|  20  | [InvertColor](#Series.InvertColor)                               | 返回或设置系列中负数据点的填充色。可读/写                                                                                                                                                                                            |
|  21  | [InvertColorIndex](#Series.InvertColorIndex)                     | 返回或设置系列中负数据点的填充色。可读/写                                                                                                                                                                                            |
|  22  | [InvertIfNegative](#Series.InvertIfNegative)                     | 如果指定项与一个负数相对应时 ET 就将其反色，则为 True 。 Boolean 类型，可读写。                                                                                                                                                      |
|  23  | [LeaderLines](#Series.LeaderLines)                               | 返回一个 LeaderLines 对象，该对象表示系列的引导线。只读。                                                                                                                                                                            |
|  24  | [MarkerBackgroundColor](#Series.MarkerBackgroundColor)           | 将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。 Long 型，可读写。                                                                                                                            |
|  25  | [MarkerBackgroundColorIndex](#Series.MarkerBackgroundColorIndex) | 返回或设置数据标志的背景色，表示为当前调色板中的索引或下列 XlColorIndex 常量之一： xlColorIndexAutomatic 或 xlColorIndexNone 。仅适用于折线图、散点图和雷达图。 Long 类型，可读写。                                                  |
|  26  | [MarkerForegroundColor](#Series.MarkerForegroundColor)           | 将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。 Long 型，可读写。                                                                                                                            |
|  27  | [MarkerForegroundColorIndex](#Series.MarkerForegroundColorIndex) | 返回或设置数据标志的前景色，表示为当前调色板中的索引或下列 XlColorIndex 常量之一： xlColorIndexAutomatic 或 xlColorIndexNone 。仅适用于折线图、散点图和雷达图。 Long 类型，可读写。                                                  |
|  28  | [MarkerSize](#Series.MarkerSize)                                 | 返回或设置数据标志的大小，以 磅 （磅：指打印的字符的高度的度量单位。1 磅等于 1/72 英寸，或大约等于 1 厘米的 1/28。） 为单位。可以是 2 到 72 之间的一个值。 Long 型，可读写。                                                         |
|  29  | [MarkerStyle](#Series.MarkerStyle)                               | 返回或设置折线图、散点图或雷达图中数据点或数据系列的数据标志样式。 XlMarkerStyle 类型，可读写。                                                                                                                                      |
|  30  | [Name](#Series.Name)                                             | 返回或设置一个 String 值，它代表对象的名称。                                                                                                                                                                                         |
|  31  | [Parent](#Series.Parent)                                         | 返回指定对象的父对象。只读。                                                                                                                                                                                                         |
|  32  | [PictureType](#Series.PictureType)                               | 返回或设置一个 XlChartPictureType 值，它代表在柱形或条形图上图片的显示方式。                                                                                                                                                         |
|  33  | [PictureUnit2](#Series.PictureUnit2)                             | 如果 PictureType 属性设置为 xlStackScale ，则返回或设置图表上每一个图片的单位。否则忽略此属性。可读/写 Double 类型。                                                                                                                 |
|  34  | [PlotColorIndex](#Series.PlotColorIndex)                         | 返回内部用于将系列格式设置与图表元素关联的索引值。只读                                                                                                                                                                               |
|  35  | [PlotOrder](#Series.PlotOrder)                                   | 返回或设置图表组中选定数据系列的绘制顺序。 Long 类型，可读写。                                                                                                                                                                       |
|  36  | [Shadow](#Series.Shadow)                                         | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                                    |
|  37  | [Smooth](#Series.Smooth)                                         | 如果折线图或散点图有平滑线，则为 True 。仅适用于折线图和散点图。可读写。                                                                                                                                                             |
|  38  | [Type](#Series.Type)                                             | 返回或设置一个代表系列类型的 Long 值。                                                                                                                                                                                               |
|  39  | [Values](#Series.Values)                                         | 返回或设置一个 Variant 值，它代表系列中所有值的集合。                                                                                                                                                                                |
|  40  | [XValues](#Series.XValues)                                       | 返回或设置图表系列中 x 值的数组。 XValues 属性可设置为工作表区域或数值数组，但不能为二者的组合。 Variant 类型，可读写。                                                                                                              |

## 成员方法

### Series.ApplyDataLabels

向系列应用数据标签。

**语法**

**express.ApplyDataLabels ( *Type* , *LegendKey* , *AutoText* , *HasLeaderLines* , *ShowSeriesName* , *ShowCategoryName* , *ShowValue* , *ShowPercentage* , *ShowBubbleSize* , *Separator* )**

\*express是一个代表 Series 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型         | 说明                                                         |
|--------------------|-----------|------------------|--------------------------------------------------------------|
| *Type*             | 可选      | XlDataLabelsType | 要应用的数据标签的类型。                                     |
| *LegendKey*        | 可选      | Variant          | 如果为 True，则在数据点旁边显示图例项标示。默认值为 False。  |
| *AutoText*         | 可选      | Variant          | 如果对象根据内容自动生成相应的文字，则为 True。              |
| *HasLeaderLines*   | 可选      | Variant          | 对于 Chart 和 Series 对象，如果数据系列有引导线，则为 True。 |
| *ShowSeriesName*   | 可选      | Variant          | 传递布尔值以启用或禁用数据标签的系列名称。                   |
| *ShowCategoryName* | 可选      | Variant          | 传递布尔值以启用或禁用数据标签的分类名称。                   |
| *ShowValue*        | 可选      | Variant          | 传递布尔值以启用或禁用数据标签的值。                         |
| *ShowPercentage*   | 可选      | Variant          | 传递布尔值以启用或禁用数据标签的百分比。                     |
| *ShowBubbleSize*   | 可选      | Variant          | 传递布尔值以启用或禁用数据标签的气泡大小。                   |
| *Separator*        | 可选      | Variant          | 数据标签的分隔符。                                           |

**说明**

|                                                                                              |
|----------------------------------------------------------------------------------------------|
| **XlDataLabelsType** 可为以下 **XlDataLabelsType** 常量之一。                                |
| **xlDataLabelsShowBubbleSizes** 。数据标签的气泡尺寸。                                       |
| **xlDataLabelsShowLabelAndPercent** 。占总数的百分比及数据点所属的分类。仅用于饼图和圆环图。 |
| **xlDataLabelsShowPercent** 。占总数的百分比。仅用于饼图和圆环图。                           |
| **xlDataLabelsShowLabel** 。数据点所属的分类。                                               |
| **xlDataLabelsShowNone** 。无数据标签。                                                      |
| **xlDataLabelsShowValue** 。默认。数据点的值（若未指定此参数，则假定为此值）。               |

**示例**

``` JavaScript
//此示例对 Chart1 上的第一个数据系列应用分类标签。
Application.Charts.Item("Chart1").SeriesCollection().Item(1).ApplyDataLabels(xlDataLabelsShowLabel)
```

### Series.AxisGroup

返回或设置指定系列的组。可读/写

**语法**

**express.AxisGroup ()**

\*express是一个代表 Series 对象的变量。

**返回值**

XlAxisGroup

### Series.ClearFormats

清除对象的格式设置。

**语法**

**express.ClearFormats ()**

\*express是一个代表 Series 对象的变量。

### Series.Copy

如果系列具有图片填充，此方法将把图片复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 Series 对象的变量。

### Series.DataLabels

返回代表数据系列中单个数据标签（ **DataLabel** 对象）或所有数据标签的集合（ **DataLabels** 集合）的对象。

**语法**

**express.DataLabels ()**

\*express是一个代表 Series 对象的变量。

**说明**

如果指定数据标签的“值” 选项处于打开状态，则返回的集合中每个数据点最多包含一个数据标签。可分别设置系列中数据点的数据标签开关选项。

如果指定系列处于面积图中，并且数据标签的“系列名称” 选项处于打开状态，则返回的集合中仅包含单个数据标签，即面积系列的标签。

**示例**

``` JavaScript
//本示例对 Chart1 的第一个数据系列的数据标签进行设置，显示其关键字段的值。本示例假定在运行时这些值是可见的。
function test()
{
    let series = Application.Charts.Item("Chart1").SeriesCollection().Item(1)
    series.HasDataLabels = true
    let bel= series.DataLabels()
    bel.ShowLegendKey = true
}
```

### Series.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Series 对象的变量。

### Series.ErrorBar

将误差线应用于系列。 **Variant** 类型。

**语法**

**express.ErrorBar ( *Direction* , *Include* , *Type* , *Amount* , *MinusValues* )**

\*express是一个代表 Series 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型            | 说明                                                       |
|---------------|-----------|---------------------|------------------------------------------------------------|
| *Direction*   | 必选      | XlErrorBarDirection | 误差线方向。                                               |
| *Include*     | 必选      | XlErrorBarInclude   | 要包含的误差线部分。                                       |
| *Type*        | 必选      | XlErrorBarType      | 误差线类型。                                               |
| *Amount*      | 可选      | Variant             | 误差量。当 Type 为 xlErrorBarTypeCustom 时只用于正误差量。 |
| *MinusValues* | 可选      | Variant             | 当 Type 为 xlErrorBarTypeCustom 时，为负误差量。           |

**返回值**

Variant

**示例**

``` JavaScript
//本示例显示 Chart1 中的第一个数据系列在 Y 方向上的标准误差线。误差线是正负两个方向的。本示例应在二维折线图上运行。
Application.Charts.Item("Chart1").SeriesCollection().Item(1).ErrorBar(xlY,xlErrorBarIncludeBoth,xlErrorBarTypeStError)
```

### Series.Paste

从剪贴板中粘贴图片，将其作为所选系列的标志。

**语法**

**express.Paste ()**

\*express是一个代表 Series 对象的变量。

**说明**

此方法可用于柱形图、条形图、折线图或雷达图，并且将 **MarkerStyle** 属性设为 **xlMarkerStylePicture** 。

**示例**

``` JavaScript
//本示例将剪贴板中的图片粘贴到 Chart1 上的第一个系列中。
Application.Charts.Item("Chart1").SeriesCollection().Item(1).Paste()
```

### Series.Points

返回一个对象，该对象表示数据系列中单个数据点（ **Point** 对象）或所有数据点的集合（ **Points** 集合）。只读

**语法**

**express.Points ( *Index* )**

\*express是一个代表 Series 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 可选      | Variant  | 数据点的名称或编号。 |

**返回值**

Object

**示例**

``` JavaScript
//本示例对 Chart1 中第一个数据系列的第一个数据点应用数据标签。
Application.Charts.Item("Chart1").SeriesCollection().Item(1).Points(1).ApplyDataLabels()
```

### Series.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 Series 对象的变量。

**返回值**

Variant

### Series.Trendlines

返回一个对象，该对象表示单个趋势线（ **Trendline** 对象）或所有趋势线的集合（ **Trendlines** 集合）。

**语法**

**express.Trendlines ( *Index* )**

\*express是一个代表 Series 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 可选      | Variant  | 趋势线的名称或编号。 |

**返回值**

Object

**示例**

``` JavaScript
//本示例向 Chart1 中第一个数据系列添加线性趋势线。
Application.Charts.Item("Chart1").SeriesCollection().Item(1).Trendlines.Add(xlLinear)
```

## 成员属性

### Series.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//本示例显示一条有关创建 myObject 的应用程序的消息。
function test()
{
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") 
    {
        alert("This is an ET Application object.")
    }
    else 
    {
        alert("This is not an ET Application object.")
    }
}
```

### Series.ApplyPictToEnd

如果图片应用于系列中数据点或所有数据点之后，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ApplyPictToEnd**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//此示例将图片应用于第一个数据系列中所有数据点之后。图片必须已应用于该系列（设置此属性将更改图片的方向）。
Application.Charts.Item(1).SeriesCollection(1).ApplyPictToEnd = true
```

### Series.ApplyPictToFront

如果图片应用于系列中数据点或所有数据点之前，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ApplyPictToFront**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//此示例将图片应用于第一个数据系列中所有数据点之前。图片必须已应用于该系列（设置此属性将更改图片的方向）。
Application.Charts.Item(1).SeriesCollection(1).ApplyPictToFront = true
```

### Series.ApplyPictToSides

如果将图片应用于系列中某个数据点或所有数据点的旁边，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ApplyPictToSides**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//此示例将图片应用于第一个数据系列中所有数据点的旁边。图片必须已应用于该系列（设置此属性将更改图片的方向）。
Application.Charts.Item(1).SeriesCollection(1).ApplyPictToSides = true
```

### Series.BarShape

返回或设置用于三维条形图或柱形图的形状。 **XlBarShape** 类型，可读写。

**语法**

**express.BarShape**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//本示例设置用于第一个图表的第一个数据系列的形状。
Application.Charts.Item(1).SeriesCollection(1).BarShape = xlConeToPoint
```

### Series.BubbleSizes

返回或设置一个字符串，该字符串引用包含气泡图 x 值、y 值和大小数据的工作表单元格。在返回单元格引用时，它将返回以 A1 样式注释描述单元格的字符串。要为气泡图设置大小数据，必须使用 R1 样式注释。仅适用于气泡图。 **Variant** 类型，可读写。

**语法**

**express.BubbleSizes**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//本示例显示单元格引用，该单元格中包含气泡图 x 值、y 值和大小数据。
alert(Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).BubbleSizes)

//此示例说明如何使用 R1 样式标住设置该属性。
Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).BubbleSizes = "=Sheet1!r1c5:r5c5"
```

### Series.ChartType

返回或设置图表类型。 **XlChartType** 类型，可读写。

**语法**

**express.ChartType**

\*express是一个代表 Series 对象的变量。

**说明**

某些图表类型对数据透视图报表无效。

**示例**

### Series.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Series 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Series.ErrorBars

返回一个 **ErrorBars** 对象，该对象表示数据系列中的误差线。只读。

**语法**

**express.ErrorBars**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//本示例对 Chart1 中第一个系列的误差线颜色进行设置。本示例应在其第一个系列有误差线的二维折线图上运行。
function test()
{
    let Ser = Application.Charts.Item("Chart1").SeriesCollection(1)
    Ser.ErrorBars.Border.ColorIndex = 8
}
```

### Series.Explosion

返回或设置饼图或圆环图的扇区的分离程度值。如果没有分离，则返回 0（零），即扇区的中心与饼图中心重合。 **Long** 型，可读写。

**语法**

**express.Explosion**

\*express是一个代表 Series 对象的变量。

### Series.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 Series 对象的变量。

### Series.Formula

返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

**语法**

**express.Formula**

\*express是一个代表 Series 对象的变量。

**说明**

此属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此 **Formula** 属性返回一个空字符串。如果单元格包含公式， **Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

### Series.FormulaLocal

返回或设置指定对象的公式，使用用户语言 A1 格式引用。 **String** 型，可读写。

**语法**

**express.FormulaLocal**

\*express是一个代表 Series 对象的变量。

**说明**

如果指定单元格包含常量，此属性返回的就是该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式，此属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的格式的值或公式设为日期类型，ET 将检查该单元格的格式是否符合某个日期或时间数组格式，如果不符合，将采用默认的短日期数字格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

对多重单元格区域设置公式，则该区域中所有单元格都用此公式填充。

### Series.FormulaR1C1

返回或设置指定对象的公式，使用宏语言 R1C1 格式符号表示。 **String** 型，可读写。

**语法**

**express.FormulaR1C1**

\*express是一个代表 Series 对象的变量。

**说明**

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式，此属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的格式的值或公式设为日期类型，ET 将检查该单元格的格式是否符合某个日期或时间数组格式，如果不符合，将采用默认的短日期数字格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

对多重单元格区域设置公式，则该区域中所有单元格都用此公式填充。

### Series.FormulaR1C1Local

返回或设置指定对象的公式，使用用户语言 R1C1 格式符号表示。 **String** 型，可读写。

**语法**

**express.FormulaR1C1Local**

\*express是一个代表 Series 对象的变量。

**说明**

如果指定单元格包含常量，此属性返回的就是该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式，此属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的格式的值或公式设为日期类型，ET 将检查该单元格的格式是否符合某个日期或时间数组格式，如果不符合，将采用默认的短日期数字格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

对多重单元格区域设置公式，则该区域中所有单元格都用此公式填充。

### Series.Has3DEffect

如果数据系列有三维外观，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.Has3DEffect**

\*express是一个代表 Series 对象的变量。

**说明**

该属性仅适用于气泡图。

### Series.HasDataLabels

如果数据系列具有数据标签，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HasDataLabels**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//本示例显示 Chart1 中第三个数据系列的数据标签。
function test()
{
    let Ser = Application.Charts.Item("Chart1").SeriesCollection(3)
    Ser.HasDataLabels = true
    Ser.ApplyDataLabels(xlValue)
}
```

### Series.HasErrorBars

如果序列有错误条时，该值为 **True** 。该属性对三维图表无效。 **Boolean** 类型，可读写。

**语法**

**express.HasErrorBars**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//本示例删除 Chart1 中第一个数据系列的误差线。本示例应在其第一个系列有误差线的二维折线图上运行。
Application.Charts.Item("Chart1").SeriesCollection(1).HasErrorBars = false
```

### Series.HasLeaderLines

如果数据系列有引导线，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HasLeaderLines**

\*express是一个代表 Series 对象的变量。

**说明**

该属性仅应用于饼图。

**示例**

``` JavaScript
//本示例向饼图上的数据系列 1 添加数据标签和蓝色引导线。如果看不到引导线，则本示例代码将失败。这种情况下，可以手动将数据标签之一拖出饼图以便让引导线显示。
function test()
{
    let Ser = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    Ser.HasDataLabels = true
    Ser.DataLabels.Position = xlLabelPositionBestFit
    Ser.HasLeaderLines = true
    Ser.LeaderLines.Border.ColorIndex = 5
}
```

### Series.InvertColor

返回或设置系列中负数据点的填充色。可读/写

**语法**

**express.InvertColor**

\*express是一个代表 Series 对象的变量。

**说明**

**InvertColor** 属性可用来设置负数据点的填充色，颜色可设置为特定的数值、十六进制值、八进制值或 RGB 颜色值。若要将该值设置为 RGB 值，请使用 Visual Basic **RGB** 函数。如果不使用 **InvertColor** 属性，还可以使用 **InvertColorIndex** 属性，它使用当前调色板上更简单的整数值集合。

为了使 **InvertColor** 属性生效，还必须将 **Series** 对象的 **InvertIfNegative** 属性设置为 **True** 。

**示例**

``` JavaScript
//下面的代码示例将“Chart 2”第一个系列中的负数据点的填充色设置为洋红。
function test()
{
    Application.ActiveSheet.ChartObjects("Chart 2").Activate()
    Application.ActiveChart.SeriesCollection(1).InvertIfNegative = true
    Application.ActiveChart.SeriesCollection(1).InvertColor = 255, 0, 255
}
```

### Series.InvertColorIndex

返回或设置系列中负数据点的填充色。可读/写

**语法**

**express.InvertColorIndex**

\*express是一个代表 Series 对象的变量。

**说明**

**InvertColorIndex** 属性可用来将负数据点的填充色设置为 0 到 56 之间的颜色索引值。有关颜色索引值的详细信息，请参阅使用 ColorIndex 属性将颜色添加到 ET工作表。如果不使用 **InvertColorIndex** 属性，还可以使用 **InvertColor** 属性，该属性可用来将颜色值作为特定的数值、十六进制值、八进制值和 RGB 颜色值来设置颜色。

为了使 **InvertColorIndex** 属性生效，还必须将 **Series** 对象的 **InvertIfNegative** 属性设置为 **True** 。

**示例**

``` JavaScript
//下面的代码示例将“Chart 2”第一个系列中的负数据点的填充色设置为洋红。
function test()
{
    Application.ActiveSheet.ChartObjects("Chart 2").Activate()
    Application.ActiveChart.SeriesCollection(1).InvertIfNegative = true
    Application.ActiveChart.SeriesCollection(1).InvertColorIndex = 7
}
```

### Series.InvertIfNegative

如果指定项与一个负数相对应时 ET 就将其反色，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.InvertIfNegative**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//此示例反转 Chart1 中第一个数据系列的负值的图案。此示例须在二维柱形图上运行。
Application.Charts.Item("Chart1").SeriesCollection(1).InvertIfNegative = true
```

### Series.LeaderLines

返回一个 **LeaderLines** 对象，该对象表示系列的引导线。只读。

**语法**

**express.LeaderLines**

\*express是一个代表 Series 对象的变量。

**说明**

该属性仅应用于饼图。

**示例**

``` JavaScript
//本示例向饼图上的数据系列 1 添加数据标签和蓝色引导线。如果看不到引导线，则本示例代码将失败。这种情况下，可以手动将数据标签之一拖出饼图以便让引导线显示。
function test()
{
    let Ser = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    Ser.HasDataLabels = true
    Ser.DataLabels.Position = xlLabelPositionBestFit
    Ser.HasLeaderLines = true
    Ser.LeaderLines.Border.ColorIndex = 5
}
```

### Series.MarkerBackgroundColor

将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。 **Long** 型，可读写。

**语法**

**express.MarkerBackgroundColor**

\*express是一个代表 Series 对象的变量。

### Series.MarkerBackgroundColorIndex

返回或设置数据标志的背景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一： **xlColorIndexAutomatic** 或 **xlColorIndexNone** 。仅适用于折线图、散点图和雷达图。 **Long** 类型，可读写。

**语法**

**express.MarkerBackgroundColorIndex**

\*express是一个代表 Series 对象的变量。

### Series.MarkerForegroundColor

将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。 **Long** 型，可读写。

**语法**

**express.MarkerForegroundColor**

\*express是一个代表 Series 对象的变量。

### Series.MarkerForegroundColorIndex

返回或设置数据标志的前景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一： **xlColorIndexAutomatic** 或 **xlColorIndexNone** 。仅适用于折线图、散点图和雷达图。 **Long** 类型，可读写。

**语法**

**express.MarkerForegroundColorIndex**

\*express是一个代表 Series 对象的变量。

### Series.MarkerSize

返回或设置数据标志的大小，以 磅 （磅：指打印的字符的高度的度量单位。1 磅等于 1/72 英寸，或大约等于 1 厘米的 1/28。） 为单位。可以是 2 到 72 之间的一个值。 **Long** 型，可读写。

**语法**

**express.MarkerSize**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//此示例设置系列一中所有数据标志的尺寸。
Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).MarkerSize = 10
```

### Series.MarkerStyle

返回或设置折线图、散点图或雷达图中数据点或数据系列的数据标志样式。 **XlMarkerStyle** 类型，可读写。

**语法**

**express.MarkerStyle**

\*express是一个代表 Series 对象的变量。

**说明**

|                                                         |
|---------------------------------------------------------|
| **XlMarkerStyle** 可为以下 **XlMarkerStyle** 常量之一。 |
| **xlMarkerStyleAutomatic** 。自动设置标记               |
| **xlMarkerStyleCircle** 。圆形标记                      |
| **xlMarkerStyleDash** 。长条形标记                      |
| **xlMarkerStyleDiamond** 。菱形标记                     |
| **xlMarkerStyleDot** 。短条形标记                       |
| **xlMarkerStyleNone** 。无标记                          |
| **xlMarkerStylePicture** 。图片标记                     |
| **xlMarkerStylePlus** 。带加号的方形标记                |
| **xlMarkerStyleSquare** 。方形标记                      |
| **xlMarkerStyleStar** 。带星号的方形标记                |
| **xlMarkerStyleTriangle** 。三角形标记                  |
| **xlMarkerStyleX** 。带 X 记号的方形标记                |

**示例**

``` JavaScript
//本示例设置 Chart1 中第一个数据系列的标记样式。本示例应在二维折线图上运行。
Application.Charts.Item("Chart1").SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
```

### Series.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Series 对象的变量。

**说明**

您可以使用 R1C1 符号进行引用，例如“=Sheet1!R1C1”。

### Series.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Series 对象的变量。

### Series.PictureType

返回或设置一个 **XlChartPictureType** 值，它代表在柱形或条形图上图片的显示方式。

**语法**

**express.PictureType**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//本示例设置 Chart1 中的系列一伸展图片。本示例应在带图片数据标志的二维柱形图上运行。
Application.Charts.Item("Chart1").SeriesCollection(1).PictureType = xlStretch
```

### Series.PictureUnit2

如果 **PictureType** 属性设置为 **xlStackScale** ，则返回或设置图表上每一个图片的单位。否则忽略此属性。可读/写 **Double** 类型。

**语法**

**express.PictureUnit2**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//此示例将 Chart1 中的系列一设为堆积图片，设置每一图片代表五个单位。此示例应在带图片数据标志的二维柱形图上运行。
function test()
{
    let Ser = Application.Charts.Item("Chart1").SeriesCollection(1)
    Ser.PictureType = xlScale
    Ser.PictureUnit2 = 5
}
```

### Series.PlotColorIndex

返回内部用于将系列格式设置与图表元素关联的索引值。只读

**语法**

**express.PlotColorIndex**

\*express是一个代表 Series 对象的变量。

**说明**

此属性由 ET 内部使用，并非用于从代码中调用。

### Series.PlotOrder

返回或设置图表组中选定数据系列的绘制顺序。 **Long** 类型，可读写。

**语法**

**express.PlotOrder**

\*express是一个代表 Series 对象的变量。

**说明**

只能设置图表组内的绘制顺序（在图表中包含若干图表类型的情况下，不能设置整个图表的绘制顺序）。图表组指具有相同图表类型和子类型的数据系列的集合。

更改某个系列的绘制顺序将导致同一图表组内其他数据系列的绘制顺序的必要调整。

**示例**

``` JavaScript
//本示例使 Chart1 中的数据系列二在绘制时第三个出现。本示例应在至少包含三个数据系列的二维柱形图上运行。
Application.Charts.Item("Chart1").ChartGroups(1).SeriesCollection(2).PlotOrder = 3
```

### Series.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 Series 对象的变量。

### Series.Smooth

如果折线图或散点图有平滑线，则为 **True** 。仅适用于折线图和散点图。可读写。

**语法**

**express.Smooth**

\*express是一个代表 Series 对象的变量。

**示例**

``` JavaScript
//此示例使 Chart1 中的系列一具有平滑线。此示例应在二维折线图上运行。
Application.Charts.Item("Chart1").SeriesCollection(1).Smooth = true
```

### Series.Type

返回或设置一个代表系列类型的 Long 值。

**语法**

**express.Type**

\*express是一个代表 Series 对象的变量。

**说明**

|                          |
|--------------------------|
| 以下常量可以用于此属性。 |
| **xlColumn**             |
| **xlBar**                |
| **xl3DBar**              |
| **xlLine**               |
| **xlPie**                |
| **xlXYScatter**          |
| **xlArea**               |
| **xl3DArea**             |
| **xlDoughnut**           |
| **xlRadar**              |
| **xl3DSurface**          |
| **xl3DColumn**           |
| **xl3DBar**              |

### Series.Values

返回或设置一个 **Variant** 值，它代表系列中所有值的集合。

**语法**

**express.Values**

\*express是一个代表 Series 对象的变量。

**说明**

此属性值可以是工作表中的某一区域或常量值的数组，但不能是两者的组合。有关详细信息，请参阅此示例。

**示例**

``` JavaScript
//此示例设置区域的系列值。
Application.Charts.Item("Chart1").SeriesCollection(1).Values = Application.Worksheets.Item("Sheet1").Range("C5:T5")

//若要为每个单独的数据点赋以常量值，则必须使用数组。
Application.Charts.Item("Chart1").SeriesCollection(1).Values = [1, 3, 5, 7, 11, 13, 17, 19]
```

### Series.XValues

返回或设置图表系列中 x 值的数组。 **XValues** 属性可设置为工作表区域或数值数组，但不能为二者的组合。 **Variant** 类型，可读写。

**语法**

**express.XValues**

\*express是一个代表 Series 对象的变量。

**说明**

对于数据透视图，该属性为只读。

**示例**

``` JavaScript
//本示例将 Chart1 中系列一的 x 值设置为 Sheet1 上 B1:B5 区域的值。
Application.Charts.Item("Chart1").SeriesCollection(1).XValues = Application.Worksheets.Item("Sheet1").Range("B1:B5")

//本示例使用数组为 Chart1 中系列一的各点逐一赋值。
Application.Charts.Item("Chart1").SeriesCollection(1).XValues = [5.0, 6.3, 12.6, 28, 50]
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Series](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SortFields 对象

## 简介

SortFields 集合是 SortField 对象的集合。开发人员可以使用该集合存储工作簿、列表和自动筛选的排序状态。

## 说明

该对象具有添加、计数、排序和删除 SortField 对象的属性。

``` JavaScript
function test(){
    Application.ActiveSheet.Sort.SortFields.Add(Range("A1"),null,xlDescending) 
    Application.ActiveSheet.Sort.SortFields.Add(Range("B1"),null,xlDescending)  
}
```

## 方法

| 序号 | 名称                       | 说明                                                                    |
|:----:|:---------------------------|:------------------------------------------------------------------------|
|  1   | [Add](#SortFields.Add)     | 创建新的排序字段，并返回一个 SortFields 对象。返回 SortField值          |
|  2   | [Clear](#SortFields.Clear) | 清除所有 Sort Fields 对象。                                             |
|  3   | [Item](#SortFields.Item)   | 返回一个 SortField 对象，该对象代表可以存储在工作簿中的项的集合。只读。 |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                               |
|:----:|:---------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#SortFields.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Count](#SortFields.Count)             | 返回集合中对象的数目。只读 Long 类型。                                                                                                                             |
|  3   | [Creator](#SortFields.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Parent](#SortFields.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |

## 成员方法

### SortFields.Add

创建新的排序字段，并返回一个 **SortFields** 对象。返回 SortField值

**语法**

**express.Add ( *Key* , *SortOn* , *Order* , *CustomOrder* , *DataOption* )**

\*express是一个代表 SortFields 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                           |
|---------------|-----------|----------|--------------------------------|
| *Key*         | 必选      | Range    | 指定用于排序的键值。           |
| *SortOn*      | 可选      | Variant  | 要进行排序的字段。             |
| *Order*       | 可选      | Variant  | 指定排序次序。                 |
| *CustomOrder* | 可选      | Variant  | 指定是否应使用自定义排序次序。 |
| *DataOption*  | 可选      | Variant  | 指定数据选项。                 |

### SortFields.Clear

清除所有 **Sort Fields** 对象。

**语法**

**express.Clear ()**

\*express是一个代表 SortFields 对象的变量。

### SortFields.Item

返回一个 **SortField** 对象，该对象代表可以存储在工作簿中的项的集合。只读。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SortFields 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明               |
|---------|-----------|----------|--------------------|
| *Index* | 必选      | Variant  | 排序字段的索引值。 |

## 成员属性

### SortFields.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 SortFields 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### SortFields.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 SortFields 对象的变量。

### SortFields.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 SortFields 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### SortFields.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SortFields 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/SortFields](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TableStyleElements 对象

## 简介

代表表格样式元素。

## 说明

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，标题行、最后一列或总计行是表格的元素，表格样式可以规定标题行的填充色为蓝色，最后一列为红色。

表格中表格样式元素的格式设置可在适用于该元素的表格样式中指定。 XlTableStyleElementType 枚举包含可供使用的表格样式元素的类型。

## 方法

| 序号 | 名称                             | 说明                   |
|:----:|:---------------------------------|:-----------------------|
|  1   | [Item](#TableStyleElements.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                           | 说明                                                                                                                                                               |
|:----:|:-----------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#TableStyleElements.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Count](#TableStyleElements.Count)             | 返回集合中对象的数目。只读 Long 类型。                                                                                                                             |
|  3   | [Creator](#TableStyleElements.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Parent](#TableStyleElements.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |

## 成员方法

### TableStyleElements.Item

从集合中返回一个对象。

**语法**

**express.Item ( *index* )**

\*express是一个代表 TableStyleElements 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型                | 说明         |
|---------|-----------|-------------------------|--------------|
| *index* | 必选      | XlTableStyleElementType | 表样式元素。 |

**返回值**

TableStyleElement

## 成员属性

### TableStyleElements.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 TableStyleElements 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### TableStyleElements.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 TableStyleElements 对象的变量。

### TableStyleElements.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 TableStyleElements 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### TableStyleElements.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 TableStyleElements 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/TableStyleElements](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Trendline 对象

## 简介

代表图表上的趋势线。

## 说明

趋势线显示系列中数据的趋势或方向。 Trendline 对象是 Trendlines 集合的成员。 Trendlines 集合包含某一个系列的所有 Trendline 对象。

使用 Trendlines ( *index* )（其中 *index* 是趋势线索引号）可返回一个 Trendline 对象。

索引号指出趋势线添加到系列中的顺序。 ` Trendlines(1) ` 是第一个添加到系列中的趋势线，而 ` Trendlines(Trendlines.Count) ` 是最后一个。

下例更改工作表一上嵌入式图表一中第一个系列的趋势线类型。如果该系列没有趋势线，则此示例会失败。

``` JavaScript
function test() {
    Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1).Type = Application.Enum.xlMovingAvg
}
```

## 方法

| 序号 | 名称                                    | 说明                 |
|:----:|:----------------------------------------|:---------------------|
|  1   | [ClearFormats](#Trendline.ClearFormats) | 清除对象的格式设置。 |
|  2   | [Delete](#Trendline.Delete)             | 删除对象。           |
|  3   | [Select](#Trendline.Select)             | 选择对象。           |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Trendline.Application)         | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Backward2](#Trendline.Backward2)             | 返回或设置趋势线向后延伸的周期数目（或散点图中的单位个数）。可读/写 Double 类型。                                                                                                                                               |
|  3   | [Border](#Trendline.Border)                   | 返回一个 Border 对象，它代表对象的边框。                                                                                                                                                                                        |
|  4   | [Creator](#Trendline.Creator)                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [DataLabel](#Trendline.DataLabel)             | 返回一个 DataLabel 对象，它代表与趋势线相关的数据标签。只读。                                                                                                                                                                   |
|  6   | [DisplayEquation](#Trendline.DisplayEquation) | 如果显示图表中的趋势线公式，则该值为 True （其数据标签与 R-平方值相同）。将该属性设为 True 可自动显示数据标签。 Boolean 类型，可读写。                                                                                          |
|  7   | [DisplayRSquared](#Trendline.DisplayRSquared) | 如果显示图表中趋势线的 R-平方值，则该值为 True （其数据标签与公式的相同）。将该属性设为 True 可自动显示数据标签。 Boolean 类型，可读写。                                                                                        |
|  8   | [Format](#Trendline.Format)                   | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  9   | [Forward2](#Trendline.Forward2)               | 返回或设置趋势线向前延伸的周期数目（或散点图中的单位个数）。可读/写 Double 类型。                                                                                                                                               |
|  10  | [Index](#Trendline.Index)                     | 返回 Long 值，它代表对象在其同类对象所组成的集合内的索引号。                                                                                                                                                                    |
|  11  | [Intercept](#Trendline.Intercept)             | 返回或设置趋势线与数值轴的交叉点。可读/写 Double 类型。                                                                                                                                                                         |
|  12  | [InterceptIsAuto](#Trendline.InterceptIsAuto) | 如果趋势线与数值轴的交叉点由回归分析自动确定，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                       |
|  13  | [Name](#Trendline.Name)                       | 返回或设置一个 String 值，它代表对象的名称。                                                                                                                                                                                    |
|  14  | [NameIsAuto](#Trendline.NameIsAuto)           | 如果 ET 自动确定趋势线的名称，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                       |
|  15  | [Order](#Trendline.Order)                     | 返回或设置一个 Long 值，它代表趋势线类型为 xlPolynomial 时趋势线的顺序号（大于 1 的整数）。                                                                                                                                     |
|  16  | [Parent](#Trendline.Parent)                   | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  17  | [Period](#Trendline.Period)                   | 返回或设置移动平均趋势线的周期。可以是从 2 至 255 的值。可读/写 Long 类型。                                                                                                                                                     |
|  18  | [Type](#Trendline.Type)                       | 返回或设置一个 XlTrendlineType 值，它代表趋势线的类型。                                                                                                                                                                         |

## 成员方法

### Trendline.ClearFormats

清除对象的格式设置。

**语法**

**express.ClearFormats ()**

\*express是一个代表 Trendline 对象的变量。

**返回值**

Variant

### Trendline.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Trendline 对象的变量。

**返回值**

Variant

### Trendline.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 Trendline 对象的变量。

**返回值**

Variant

## 成员属性

### Trendline.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Trendline 对象的变量。

**示例**

本示例显示一条有关创建 ` myObject ` 的应用程序的消息。

``` JavaScript
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### Trendline.Backward2

返回或设置趋势线向后延伸的周期数目（或散点图中的单位个数）。可读/写 **Double** 类型。

**语法**

**express.Backward2**

\*express是一个代表 Trendline 对象的变量。

**示例**

本示例对 Chart1 中的趋势线向前和向后延伸的单位数进行设置。本示例应在包含单个带趋势线系列的二维柱形图上运行。

``` JavaScript
function test() {
    let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)
    rng.Forward2 = 5
    rng.Backward2 = 0.5
}
```

### Trendline.Border

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

\*express是一个代表 Trendline 对象的变量。

### Trendline.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Trendline 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Trendline.DataLabel

返回一个 **DataLabel** 对象，它代表与趋势线相关的数据标签。只读。

**语法**

**express.DataLabel**

\*express是一个代表 Trendline 对象的变量。

**说明**

若要打开趋势线的数据标签，需要将 **DisplayEquation** 属性和 **DisplayRSquared** 属性设置为 **True** 。

### Trendline.DisplayEquation

如果显示图表中的趋势线公式，则该值为 **True** （其数据标签与 R-平方值相同）。将该属性设为 **True** 可自动显示数据标签。 **Boolean** 类型，可读写。

**语法**

**express.DisplayEquation**

\*express是一个代表 Trendline 对象的变量。

**示例**

本示例显示 Chart1 中趋势线一的 R-平方值和公式。本示例应在其第一个系列有趋势线的二维柱形图上运行。

``` JavaScript
function test() {
    let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)
    rng.DisplayRSquared = true
    rng.DisplayEquation = true
}
```

### Trendline.DisplayRSquared

如果显示图表中趋势线的 R-平方值，则该值为 **True** （其数据标签与公式的相同）。将该属性设为 **True** 可自动显示数据标签。 **Boolean** 类型，可读写。

**语法**

**express.DisplayRSquared**

\*express是一个代表 Trendline 对象的变量。

**示例**

本示例显示 Chart1 中趋势线一的 R-平方值和公式。本示例应在其第一个系列有趋势线的二维柱形图上运行。

``` JavaScript
function test() {
    let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)
    rng.DisplayRSquared = true
    rng.DisplayEquation = true
}
```

### Trendline.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 Trendline 对象的变量。

### Trendline.Forward2

返回或设置趋势线向前延伸的周期数目（或散点图中的单位个数）。可读/写 **Double** 类型。

**语法**

**express.Forward2**

\*express是一个代表 Trendline 对象的变量。

**示例**

本示例对 Chart1 中的趋势线向前和向后延伸的单位数进行设置。本示例应在包含单个带趋势线系列的二维柱形图上运行。

``` JavaScript
function test() {
    let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)
    rng.Forward2 = 5
    rng.Backward2 = 0.5
}
```

### Trendline.Index

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 Trendline 对象的变量。

### Trendline.Intercept

返回或设置趋势线与数值轴的交叉点。可读/写 **Double** 类型。

**语法**

**express.Intercept**

\*express是一个代表 Trendline 对象的变量。

**说明**

设置该属性会将 **InterceptIsAuto** 属性设置为 **False** 。

**示例**

本示例设置 Chart1 的第一条趋势线与数值轴相交于 5。本示例应在包含单个带趋势线系列的二维柱形图上运行。

``` JavaScript
function test() {
    Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1).Intercept = 5
}
```

### Trendline.InterceptIsAuto

如果趋势线与数值轴的交叉点由回归分析自动确定，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.InterceptIsAuto**

\*express是一个代表 Trendline 对象的变量。

**说明**

设置 **Intercept** 属性会将该属性设置为 **False** 。

**示例**

本示例设置 ET 自动判断 Chart1 的趋势线与数值坐标轴的交点。本示例应在包含单个带趋势线系列的二维柱形图上运行。

``` JavaScript
function test() {
    Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1).InterceptIsAuto = true
}
```

### Trendline.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Trendline 对象的变量。

### Trendline.NameIsAuto

如果 ET 自动确定趋势线的名称，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.NameIsAuto**

\*express是一个代表 Trendline 对象的变量。

**示例**

本示例设置 ET 自动确定 Chart1 的趋势线一的名称。本示例应在包含单个带趋势线系列的二维柱形图上运行。

``` JavaScript
function test() {
    Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1).NameIsAuto = true
}
```

### Trendline.Order

返回或设置一个 **Long** 值，它代表趋势线类型为 **xlPolynomial** 时趋势线的顺序号（大于 1 的整数）。

**语法**

**express.Order**

\*express是一个代表 Trendline 对象的变量。

### Trendline.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Trendline 对象的变量。

### Trendline.Period

返回或设置移动平均趋势线的周期。可以是从 2 至 255 的值。可读/写 **Long** 类型。

**语法**

**express.Period**

\*express是一个代表 Trendline 对象的变量。

**示例**

本示例设置 Chart1 上移动平均趋势线的周期。本示例应在二维柱形图上运行，该柱形图中只有一个包含 10 个数据点的系列，并且有移动平均趋势线。

``` JavaScript
function test() {
    let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)
    if (rng.Type == Application.Enum.xlMovingAvg) {
        rng.Period = 5
    }
}
```

### Trendline.Type

返回或设置一个 **XlTrendlineType** 值，它代表趋势线的类型。

**语法**

**express.Type**

\*express是一个代表 Trendline 对象的变量。

**示例**

本示例更改第一个系列（第一张工作表的第一个嵌入图表中）的趋势线类型。如果该系列没有趋势线，则本示例无效。

``` JavaScript
function test() {
    Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1).Type = Application.Enum.xlMovingAvg
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Trendline](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Trendlines 对象

## 简介

指定的数据系列中所有 Trendline 对象的集合。

## 说明

每一个 Trendline 对象都代表图表中的趋势线。趋势线显示系列中数据的趋势或方向。

## 方法

| 序号 | 名称                     | 说明                   |
|:----:|:-------------------------|:-----------------------|
|  1   | [Add](#Trendlines.Add)   | 创建新的趋势线。       |
|  2   | [Item](#Trendlines.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Trendlines.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |

## 成员方法

### Trendlines.Add

创建新的趋势线。

**语法**

**express.Add ( *Type* , *Order* , *Period* , *Forward* , *Backward* , *Intercept* , *DisplayEquation* , *DisplayRSquared* , *Name* )**

\*express是一个代表 Trendlines 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型        | 说明                                                                                                 |
|-------------------|-----------|-----------------|------------------------------------------------------------------------------------------------------|
| *Type*            | 可选      | XlTrendlineType | 趋势线类型。                                                                                         |
| *Order*           | 可选      | Variant         | Variant。如果 Type 为 xlPolynomial。趋势线顺序号。必须为 2 到 6 之间的整数（包括 2 和 6）。          |
| *Period*          | 可选      | Variant         | 如果 Type 为 xlMovingAvg。趋势线周期。必须为大于 1，而小于正添加趋势线的数据系列中数据点个数的整数。 |
| *Forward*         | 可选      | Variant         | 趋势线向前延伸的周期数目（或散点图中的单位个数）。                                                   |
| *Backward*        | 可选      | Variant         | 趋势线向后延伸的周期数目（或散点图中的单位个数）。                                                   |
| *Intercept*       | 可选      | Variant         | 趋势线的截距。如果省略此参数，则由回归分析自动设置截距。                                             |
| *DisplayEquation* | 可选      | Variant         | 如果为 True，则在图表中显示趋势线的公式（与 R 平方值显示在同一数据标签中）。默认值为 False。         |
| *DisplayRSquared* | 可选      | Variant         | 如果为 True，则在图表中显示趋势线的 R 平方值（与公式显示在同一数据标签中）。默认值为 False。         |
| *Name*            | 可选      | Variant         | 作为文本的趋势线名称。如果省略此参数，则由 ET 生成名称。                                             |

**返回值**

一个代表新趋势线的 Trendline 对象。

**示例**

本示例在 Chart1 中新建线性趋势线。

``` JavaScript
Application.ActiveWorkbook.Charts.Item("Chart1").SeriesCollection(1).Trendlines().Add()
```

### Trendlines.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Trendlines 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *Index* | 可选      | Variant  | 对象的索引号。 |

**返回值**

包含在集合中的一个 Trendline 对象。

**示例**

``` JavaScript
function test()
{
  let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines().Item(1)
  rng.Forward = 5
  rng.Backward = 0.5
}
```

## 成员属性

### Trendlines.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Trendlines 对象的变量。

**示例**

本示例显示一条有关创建 ` myObject ` 的应用程序的消息。

``` JavaScript
function TEST()
{
   let myObject = ActiveWorkbook
   if(myObject.Application.Value == "ET"){
       MsgBox("This is an ET Application object.")
   }
   else{
       MsgBox("This is not an ET Application object.")
   }
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Trendlines](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Worksheets 对象

## 简介

指定的或活动工作簿中所有 Worksheet 对象的集合。每个 Worksheet 对象都代表一个工作表。

## 说明

Worksheet 对象也是 Sheets 集合的成员。 Sheets 集合包含工作簿中所有的工作表（图表工作表和工作表）。

使用 Worksheets 属性可返回 Worksheets 集合。下例将所有工作表移到工作簿尾部。

``` JavaScript
Application.Worksheets.Move(Application.Worksheets.Item(Application.Worksheets.Count))
```

使用 Add 方法可创建一个新工作表并将它添加到集合。下例将两个新工作表添加到活动工作簿的工作表一之前。

``` JavaScript
Application.Worksheets.Add(Application.Worksheets.Item(1), undefined, 2)
```

使用 Worksheets ( *index* )（其中 *index* 是工作表索引号或名称）可返回一个 Worksheet 对象。下例隐藏活动工作簿中的工作表一。

``` JavaScript
Application.Worksheets.Item(1).Visible = false
```

## 方法

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#Worksheets.Add">Add</a></td>
<td style="text-align: left;" data-valian="middle"><p>新建工作表、图表或宏表。新建的工作表将成为活动工作表。</p>
<p>返回 一个 Object 值，它代表新的工作表、图表或宏表。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#Worksheets.Copy">Copy</a></td>
<td style="text-align: left;" data-valian="middle"><p>将工作表复制到工作簿的另一位置。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#Worksheets.Delete">Delete</a></td>
<td style="text-align: left;" data-valian="middle"><p>删除对象</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#Worksheets.FillAcrossSheets">FillAcrossSheets</a></td>
<td style="text-align: left;" data-valian="middle"><p>将单元格区域复制到集合中所有其他工作表的同一位置。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#Worksheets.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>从集合中返回一个对象。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#Worksheets.Move">Move</a></td>
<td style="text-align: left;" data-valian="middle"><p>将工作表移到工作簿中的其他位置。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#Worksheets.PrintOut">PrintOut</a></td>
<td style="text-align: left;" data-valian="middle"><p>打印对象。</p>
<p>返回 Variant 值</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#Worksheets.PrintPreview">PrintPreview</a></td>
<td style="text-align: left;" data-valian="middle"><p>按对象打印后的外观效果显示对象的预览。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#Worksheets.Select">Select</a></td>
<td style="text-align: left;" data-valian="middle"><p>选择对象。</p></td>
</tr>
</tbody>
</table>

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Worksheets.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Worksheets.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Worksheets.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [HPageBreaks](#Worksheets.HPageBreaks) | 返回一个 HPageBreaks 集合，它代表工作表上的水平分页符。只读。                                                                                                                                                                   |
|  5   | [Parent](#Worksheets.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  6   | [VPageBreaks](#Worksheets.VPageBreaks) | 返回一个 VPageBreaks 集合，它代表工作表上的垂直分页符。只读。                                                                                                                                                                   |
|  7   | [Visible](#Worksheets.Visible)         | 返回或设置一个 Variant 值，它确定对象是否可见。                                                                                                                                                                                 |

## 成员方法

### Worksheets.Add

新建工作表、图表或宏表。新建的工作表将成为活动工作表。

返回 一个 Object 值，它代表新的工作表、图表或宏表。

**语法**

**express.Add ( *Before* , *After* , *Count* , *Type* )**

\*express是一个代表 Worksheets 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                        |
|----------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Before* | 可选      | Variant  | 指定工作表的对象，新建的工作表将置于此工作表之前。                                                                                                                                          |
| *After*  | 可选      | Variant  | 指定工作表的对象，新建的工作表将置于此工作表之后。                                                                                                                                          |
| *Count*  | 可选      | Variant  | 要添加的工作表数。默认值为 1。                                                                                                                                                              |
| *Type*   | 可选      | Variant  | 指定工作表类型。可以为下列 XlSheetType 常量之一：xlWorksheet、xlChart、xlExcel4MacroSheet 或 xlExcel4IntlMacroSheet。如果基于现有模板插入工作表，则指定该模板的路径。默认值为 xlWorksheet。 |

**说明**

如果同时省略 *Before* 和 *After* ，则新工作表插入到活动工作表之前。

### Worksheets.Copy

将工作表复制到工作簿的另一位置。

**语法**

**express.Copy ( *Before* , *After* )**

\*express是一个代表 Worksheets 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                        |
|----------|-----------|----------|-----------------------------------------------------------------------------|
| *Before* | 可选      | Variant  | 将要在其之前放置所复制工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选      | Variant  | 将要在其之后放置所复制工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After* ，则 ET 将新建一个工作簿，其中包含复制的工作表。

**示例**

``` JavaScript
/*此示例复制工作表 Sheet1，并将其放置在工作表 Sheet3 之后。*/
Application.Worksheets.Item("Sheet1").Copy(null, Application.Worksheets.Item("Sheet3"))
```

### Worksheets.Delete

删除对象

**语法**

**express.Delete ()**

\*express是一个代表 Worksheets 对象的变量。

### Worksheets.FillAcrossSheets

将单元格区域复制到集合中所有其他工作表的同一位置。

**语法**

**express.FillAcrossSheets ( *Range* , *Type* )**

\*express是一个代表 Worksheets 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型   | 说明                                                                       |
|---------|-----------|------------|----------------------------------------------------------------------------|
| *Range* | 必选      | Range      | 要填充到集合中所有工作表上的单元格区域。该区域必须来自集合中的某个工作表。 |
| *Type*  | 可选      | XlFillWith | 指定如何复制区域。                                                         |

**示例**

``` JavaScript
function test(){
  /*本示例用工作表 Sheet1 上 A1:C5 单元格区域的内容填充工作表 Sheet5 和 Sheet7 上的相同区域。*/
  let x = ["Sheet1", "Sheet5", "Sheet7"]
  Application.Sheets.Item(x).FillAcrossSheets(Application.Worksheets.Item("Sheet1").Range("A1:C5"))
}
```

### Worksheets.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Worksheets 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**说明**

有关返回集合中单个成员的详细信息，请参阅 <span style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;background-color:#FFFFFF;">返回集合中的对象</span> 。

**示例**

``` JavaScript
/*Item 是集合的默认成员*/
Application.ActiveWorkbook.Worksheets.Item(1)
```

### Worksheets.Move

将工作表移到工作簿中的其他位置。

**语法**

**express.Move ( *Before* , *After* )**

\*express是一个代表 Worksheets 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                  |
|----------|-----------|----------|-----------------------------------------------------------------------|
| *Before* | 可选      | Variant  | 在其之前放置移动工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选      | Variant  | 在其之后放置移动工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After* ，ET 将新建一个工作簿，其中包含所移动的工作表。

**示例**

``` JavaScript
/*此示例将当前活动工作簿的 Sheet1 移到 Sheet3 之后。*/
Application.Worksheets.Item("Sheet1").Move(null, Application.Worksheets.Item("Sheet3"))
```

### Worksheets.PrintOut

打印对象。

返回 Variant 值

**语法**

**express.PrintOut ( *From* , *To* , *Copies* , *Preview* , *ActivePrinter* , *PrintToFile* , *Collate* , *PrToFileName* , *IgnorePrintAreas* )**

\*express是一个代表 Worksheets 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                                                                                              |
|--------------------|-----------|----------|---------------------------------------------------------------------------------------------------|
| *From*             | 可选      | Variant  | 打印的开始页号。如果省略此参数，则从起始位置开始打印。                                            |
| *To*               | 可选      | Variant  | 打印的终止页号。如果省略此参数，则打印至最后一页。                                                |
| *Copies*           | 可选      | Variant  | 打印份数。如果省略此参数，则只打印一份。                                                          |
| *Preview*          | 可选      | Variant  | 如果为 True，ET 将在打印对象之前调用打印预览。如果为 False（或省略该参数），则立即打印对象。      |
| *ActivePrinter*    | 可选      | Variant  | 设置活动打印机的名称。                                                                            |
| *PrintToFile*      | 可选      | Variant  | 如果为 True，则打印到文件。如果没有指定 PrToFileName，ET 将提示用户输入要使用的输出文件的文件名。 |
| *Collate*          | 可选      | Variant  | 如果为 True，则逐份打印多个副本。                                                                 |
| *PrToFileName*     | 可选      | Variant  | 如果PrintToFile 设为 True，则该参数指定要打印到的文件名。                                         |
| *IgnorePrintAreas* | 可选      | Variant  | 如果为 True，则忽略打印区域并打印整个对象。                                                       |

**说明**

*From* 和 *To* 所描述的“页”指的是要打印的页，并非指定工作表或工作簿中的全部页。

**示例**

``` JavaScript
/*此示例打印当前活动工作表。*/
Application.ActiveSheet.PrintOut()
```

### Worksheets.PrintPreview

按对象打印后的外观效果显示对象的预览。

**语法**

**express.PrintPreview ( *EnableChanges* )**

\*express是一个代表 Worksheets 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                          |
|-----------------|-----------|----------|-------------------------------------------------------------------------------|
| *EnableChanges* | 可选      | Variant  | 传递 Boolean 值，以指定用户是否可更改边距和打印预览中可用的其他页面设置选项。 |

**示例**

``` JavaScript
/*此示例在打印预览中显示工作表 Sheet1。*/
Application.Worksheets.Item("Sheet1").PrintPreview()
```

### Worksheets.Select

选择对象。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 Worksheets 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                                              |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

## 成员属性

### Worksheets.Application

如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Worksheets 对象的变量。

**示例**

``` JavaScript
function test(){
  /*本示例显示一条有关创建 myObject 的应用程序的消息。*/
  let Wsheet = Application.ActiveWorkbook
  if(Wsheet.Application.Value == "ET"){
     alert("This is an ET Application object.")
  }
  else {
     alert("This is not an ET Application object.")
  }
}
```

### Worksheets.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Worksheets 对象的变量。

### Worksheets.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Worksheets 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Worksheets.HPageBreaks

返回一个 HPageBreaks 集合，它代表工作表上的水平分页符。只读。

**语法**

**express.HPageBreaks**

\*express是一个代表 Worksheets 对象的变量。

**说明**

每张工作表最多有 1026 个水平分页符。

**示例**

``` JavaScript
function test(){
  /*此示例显示全屏水平分页符和打印区水平分页符的个数。*/
  let cFull = 0
  let cPartial = 0
  for(let i = 1; i <= Application.Worksheets.Item(1).HPageBreaks.Count; i++){
      if(Application.Worksheets.Item(1).HPageBreaks.Extent == xlPageBreakFull){
          cFull++
      }
      else{
          cPartial++
      }
  }
  alert(cFull + " full-screen page breaks, " + cPartial + " print-area page breaks")
}
```

### Worksheets.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Worksheets 对象的变量。

### Worksheets.VPageBreaks

返回一个 VPageBreaks 集合，它代表工作表上的垂直分页符。只读。

**语法**

**express.VPageBreaks**

\*express是一个代表 Worksheets 对象的变量。

**示例**

``` JavaScript
function test(){
  /*此示例显示全屏垂直分页符和打印区域垂直分页符的总数。*/
  let cFull = 0
  let cPartial = 0
  for(let i = 1; i <= Application.Worksheets.Item(1).VPageBreaks.Count; i++){
      if(Application.Worksheets.Item(1).VPageBreaks.Extent == xlPageBreakFull){
          cFull++
      }
      else{
          cPartial++
      }
  }
  alert(cFull + " full-screen page breaks, " + cPartial + " print-area page breaks")
}
```

### Worksheets.Visible

返回或设置一个 **Variant** 值，它确定对象是否可见。

**语法**

**express.Visible**

\*express是一个代表 Worksheets 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Worksheets](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# WorksheetView 对象

## 简介

指定的或活动工作簿中所有 WorksheetView 对象的集合。

## 说明

通过使用 DisplayFormulas、DisplayGridlines 和 DisplayHeadings 等属性控制应用程序或工作簿级视图的外观和风格。

## 属性

| 序号 | 名称                                                | 说明                                                                                                                                                               |
|:----:|:----------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#WorksheetView.Application)           | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [DisplayFormulas](#WorksheetView.DisplayFormulas)   | 返回或设置在当前工作表视图中是显示还是隐藏公式。可读/写 Boolean 类型。                                                                                             |
|  3   | [DisplayGridlines](#WorksheetView.DisplayGridlines) | 如果显示网格线，则为 True 。可读/写 Boolean 类型。                                                                                                                 |
|  4   | [DisplayHeadings](#WorksheetView.DisplayHeadings)   | 如果同时显示行标题和列标题，则为 True ；如果未显示标题，则为 False 。可读/写 Boolean 类型。                                                                        |
|  5   | [DisplayOutline](#WorksheetView.DisplayOutline)     | 如果显示分级显示符号，则为 True 。可读/写 Boolean 类型。                                                                                                           |
|  6   | [DisplayZeros](#WorksheetView.DisplayZeros)         | 如果显示零值，则为 True 。可读/写 Boolean 类型。                                                                                                                   |
|  7   | [Sheet](#WorksheetView.Sheet)                       | 返回指定 WorksheetView 对象的工作表名称。只读。                                                                                                                    |

## 成员属性

### WorksheetView.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 WorksheetView 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### WorksheetView.DisplayFormulas

返回或设置在当前工作表视图中是显示还是隐藏公式。可读/写 **Boolean** 类型。

**语法**

**express.DisplayFormulas**

\*express是一个代表 WorksheetView 对象的变量。

### WorksheetView.DisplayGridlines

如果显示网格线，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.DisplayGridlines**

\*express是一个代表 WorksheetView 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

此属性仅影响显示的网格线。使用 **PrintGridlines** 属性可以控制网格线的打印。

### WorksheetView.DisplayHeadings

如果同时显示行标题和列标题，则为 **True** ；如果未显示标题，则为 **False** 。可读/写 **Boolean** 类型。

**语法**

**express.DisplayHeadings**

\*express是一个代表 WorksheetView 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

此属性仅影响显示的标题。使用 **PrintHeadings** 属性可以控制标题的打印。

### WorksheetView.DisplayOutline

如果显示分级显示符号，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.DisplayOutline**

\*express是一个代表 WorksheetView 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

### WorksheetView.DisplayZeros

如果显示零值，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.DisplayZeros**

\*express是一个代表 WorksheetView 对象的变量。

**说明**

该属性仅适用于工作表和宏工作表。

### WorksheetView.Sheet

返回指定 **WorksheetView** 对象的工作表名称。只读。

**语法**

**express.Sheet**

\*express是一个代表 WorksheetView 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/WorksheetView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Cells 对象

## 简介

由表格列、表格行、选定内容或区域中的 Cell 对象组成的集合。

## 方法

| 序号 | 名称                                        | 说明                                                               |
|:----:|:--------------------------------------------|:-------------------------------------------------------------------|
|  1   | [Add](#Cells.Add)                           | 返回一个 Cell 对象，该对象代表添加到表格中的单元格。               |
|  2   | [AutoFit](#Cells.AutoFit)                   | 改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。 |
|  3   | [Delete](#Cells.Delete)                     | 删除一个或多个表格单元格并可选择控制如何移动剩余的单元格。         |
|  4   | [DistributeHeight](#Cells.DistributeHeight) | 将指定单元格调整为等高。                                           |
|  5   | [DistributeWidth](#Cells.DistributeWidth)   | 将指定单元格调整为等宽。                                           |
|  6   | [Item](#Cells.Item)                         | 返回集合中的单个 Cell 对象。                                       |
|  7   | [Merge](#Cells.Merge)                       | 将指定的多个表格单元格相互合并。合并后的结果是一个表格单元格。     |
|  8   | [SetHeight](#Cells.SetHeight)               | 设定表格单元格的高度。                                             |
|  9   | [SetWidth](#Cells.SetWidth)                 | 设置表格的列或单元格的宽度。                                       |
|  10  | [Split](#Cells.Split)                       | 拆分表格单元格范围。                                               |

## 属性

| 序号 | 名称                                            | 说明                                                                                          |
|:----:|:------------------------------------------------|:----------------------------------------------------------------------------------------------|
|  1   | [Application](#Cells.Application)               | 返回一个 Application 对象，该对象代表 WPS 应用程序。                                          |
|  2   | [Borders](#Cells.Borders)                       | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                         |
|  3   | [Count](#Cells.Count)                           | 返回 Cells 集合中的项目数。 Number 类型，只读。                                               |
|  4   | [Creator](#Cells.Creator)                       | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。                  |
|  5   | [Height](#Cells.Height)                         | 返回或设置指定表格单元格的高度。可读/写 Single 类型。                                         |
|  6   | [HeightRule](#Cells.HeightRule)                 | 返回或设置一个 WdRowHeightRule 常量，该常量代表确定指定单元格高度的规则。可读写。             |
|  7   | [NestingLevel](#Cells.NestingLevel)             | 返回指定单元格的嵌套层。 Long 类型，只读。                                                    |
|  8   | [Parent](#Cells.Parent)                         | 返回一个 Object 类型的值，该值代表指定 Cells 对象的父对象。                                   |
|  9   | [PreferredWidth](#Cells.PreferredWidth)         | 返回或设置指定单元格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。 Single 类型，可读写。  |
|  10  | [PreferredWidthType](#Cells.PreferredWidthType) | 返回或设置用于指定单元格的宽度的首选度量单位。 WdPreferredWidthType 类型，只读。              |
|  11  | [Shading](#Cells.Shading)                       | 返回一个 Shading 对象，该对象指明指定对象的底纹格式。                                         |
|  12  | [VerticalAlignment](#Cells.VerticalAlignment)   | 返回或设置表格中一个或多个单元格中文字的垂直对齐方式。 WdCellVerticalAlignment 类型，可读写。 |
|  13  | [Width](#Cells.Width)                           | 返回或设置表格单元格的宽度（以磅为单位）。 Number 类型，可读写。                              |

## 成员方法

### Cells.Add

返回一个 **Cell** 对象，该对象代表添加到表格中的单元格。

**语法**

**express.Add ( *BeforeCell* )**

\*express是一个代表 Cells 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                           |
|--------------|-----------|----------|----------------------------------------------------------------|
| *BeforeCell* | 可选      | Variant  | 指定一个 Cell 对象，该对象代表将紧靠新单元格右侧显示的单元格。 |

**返回值**

Cell

### Cells.AutoFit

改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。

**语法**

**express.AutoFit ()**

\*express是一个代表 Cells 对象的变量。

**说明**

如果表格的宽度已等于从左边界到右边界的距离，则此方法无效。

**示例**

``` JavaScript
/*本示例在新文档中创建一个 3x3 表格，然后调整所有列的宽度，使之与文本的宽度相称。*/
function test()
{
let docNew = Documents.Add()
let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)
tableNew.Cell(1,1).Range.InsertAfter("First cell")
tableNew.Cell(1,2).Range.InsertAfter("This is cell (1,2)")
tableNew.Cell(1,3).Range.InsertAfter("(1,3)")
tableNew.Columns.AutoFit()
}
```

### Cells.Delete

删除一个或多个表格单元格并可选择控制如何移动剩余的单元格。

**语法**

**express.Delete ( *ShiftCells* )**

\*express是一个代表 Cells 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                  |
|--------------|-----------|----------|-------------------------------------------------------------------------------------------------------|
| *ShiftCells* | 可选      | Variant  | 剩余单元格移动的方向。可以是任意 WdDeleteCells 常量。如果省略，最后删除的单元格的右侧单元格向左移动。 |

### Cells.DistributeHeight

将指定单元格调整为等高。

**语法**

**express.DistributeHeight ()**

\*express是一个代表 Cells 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档的第一张表格的高度调整为相等。*/
Application.ActiveDocument.Tables.Item(1).Rows.DistributeHeight()
```

``` JavaScript
/*本示例将第一张表格的前三行高度调整为相等。*/
function test() {
    if (Application.ActiveDocument.Tables.Count >= 1) {
        let rngTemp = Application.ActiveDocument.Range(Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Range.Start, Application.ActiveDocument.Tables.Item(1).Rows.Item(3).Range.End)
        rngTemp.Rows.DistributeHeight()
    }
}
```

### Cells.DistributeWidth

将指定单元格调整为等宽。

**语法**

**express.DistributeWidth ()**

\*express是一个代表 Cells 对象的变量。

**示例**

``` JavaScript
/* 调整选中单元格，使宽度一致 */
Application.Selection.Cells.DistributeWidth()
```

### Cells.Item

返回集合中的单个 **Cell** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Cells 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

**返回值**

Cell

### Cells.Merge

将指定的多个表格单元格相互合并。合并后的结果是一个表格单元格。

**语法**

**express.Merge ()**

\*express是一个代表 Cells 对象的变量。

**示例**

``` JavaScript
/*以下示例将所选内容第一行中的多个单元格合并为一个单元格，然后为该行应用底纹。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        let myrow = Application.Selection.Rows.Item(1)
        myrow.Cells.Merge()
        myrow.Shading.Texture = wdTexture10Percent
    }
}
```

### Cells.SetHeight

设定表格单元格的高度。

**语法**

**express.SetHeight ( *RowHeight* , *HeightRule* )**

\*express是一个代表 Cells 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型        | 说明                           |
|--------------|-----------|-----------------|--------------------------------|
| *RowHeight*  | 必选      | Variant         | 一行或多行的高度，以磅为单位。 |
| *HeightRule* | 必选      | WdRowHeightRule | 确定指定单元格高度的方法。     |

**说明**

设置 **Cells** 对象的 **SetHeight** 属性将自动设置整行的高度。

**示例**

``` JavaScript
/*以下示例将所选单元格的行高设置为至少 18 磅。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        Application.Selection.Cells.SetHeight(18, wdRowHeightAtLeast)
    }
    else {
        alert("The insertion point is not in a table.")
    }
}
```

### Cells.SetWidth

设置表格的列或单元格的宽度。

**语法**

**express.SetWidth ( *ColumnWidth* , *RulerStyle* )**

\*express是一个代表 Cells 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型     | 说明                            |
|---------------|-----------|--------------|---------------------------------|
| *ColumnWidth* | 必选      | Single       | 指定列的宽度，以磅为单位。      |
| *RulerStyle*  | 必选      | WdRulerStyle | 控制 WPS 调整单元格宽度的方式。 |

**说明**

上述 **WdRulerStyle** 行为适用于左对齐的表格。 **WdRulerStyle** 行为用于居中和右对齐的表格时可能无法预测，因此 **SetWidth** 方法应谨慎使用。

**示例**

``` JavaScript
/*以下示例在新文档中创建一张表格，并将第二行第一个单元格的宽度设置为 1.5 英寸。此示例保持表格中其他单元格的宽度。*/
function test() {
    let newDoc = Application.Documents.Add()
    let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)
    myTable.Cell(2, 1).SetWidth(InchesToPoints(1.5), wdAdjustNone)
}
```

``` JavaScript
/*以下示例将包含插入点的单元格的宽度设置为 36 磅。此示例还缩小第一列的宽度以保持表格的右边缘位置。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        Application.Selection.Cells.Item(1).SetWidth(36, wdAdjustFirstColumn)
    }
    else {
        alert("The insertion point is not in a table.")
    }
}
```

### Cells.Split

拆分表格单元格范围。

**语法**

**express.Split ( *NumRows* , *NumColumns* , *MergeBeforeSplit* )**

\*express是一个代表 Cells 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                                                        |
|--------------------|-----------|----------|-------------------------------------------------------------|
| *NumRows*          | 可选      | Variant  | 单元格或单元格组拆分成的行数。                              |
| *NumColumns*       | 可选      | Variant  | 单元格或单元格组拆分成的列数。                              |
| *MergeBeforeSplit* | 可选      | Variant  | 如果该参数值为 True，则在拆分多个单元格前会将它们相互合并。 |

**示例**

``` JavaScript
/*以下示例首先将选定单元格合并为一个单元格，然后将该单元格拆分为同一行中的三个单元格。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        Application.Selection.Cells.Split(1, 3, true)
    }
}
```

## 成员属性

### Cells.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Cells 对象的变量。

### Cells.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Cells 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例对活动文档中第一个表格首行中的单元格应用内部和外部边框。*/
function test() {
    let objTable = Application.ActiveDocument.Tables.Item(1)
    let Bor = objTable.Rows.Item(1).Cells.Borders
    Bor.InsideLineStyle = wdLineStyleSingle
    Bor.OutsideLineStyle = wdLineStyleDouble
}
```

### Cells.Count

返回 **Cells** 集合中的项目数。 **Number** 类型，只读。

**语法**

**express.Count**

\*express是一个代表 Cells 对象的变量。

### Cells.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Cells 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

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

### Cells.Height

返回或设置指定表格单元格的高度。可读/写 **Single** 类型。

**语法**

**express.Height**

\*express是一个代表 Cells 对象的变量。

**说明**

如果指定行的 **HeightRule** 属性为 **wdRowHeightAuto** ，则 **Height** 返回 **wdUndefined** ；如果设置 **Height** 属性，系统会将 **HeightRule** 设置为 **wdRowHeightAtLeast** 。

### Cells.HeightRule

返回或设置一个 **WdRowHeightRule** 常量，该常量代表确定指定单元格高度的规则。可读写。

**语法**

**express.HeightRule**

\*express是一个代表 Cells 对象的变量。

**说明**

设置 **Cells** 集合的 **HeightRule** 属性将自动设置整行的高度。

### Cells.NestingLevel

返回指定单元格的嵌套层。 **Long** 类型，只读。

**语法**

**express.NestingLevel**

\*express是一个代表 Cells 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

**示例**

``` JavaScript
/*本示例新建一个文档，创建一个三层嵌套表格，并在每个表格的第一个单元格中填入该表格所在的嵌套层数。*/
function test() {
    Application.Documents.Add()
    Application.ActiveDocument.Tables.Add(Selection.Range, 3, 3, wdWord9TableBehavior, wdAutoFitContent)
    let Rng1 = Application.ActiveDocument.Tables.Item(1).Range
    Rng1.Copy()
    Rng1.Cells.Item(1).Range.Text = Rng1.Cells.NestingLevel
    Rng1.Cells.Item(5).Range.PasteAsNestedTable()
    let Rng2 = Rng1.Cells.Item(5).Tables.Item(1).Range
    Rng2.Cells.Item(1).Range.Text = Rng2.Cells.NestingLevel
    Rng2.Cells.Item(5).Range.PasteAsNestedTable()
    let Rng3 = Rng2.Cells.Item(5).Tables.Item(1).Range
    Rng3.Cells.Item(1).Range.Text = Rng3.Cells.NestingLevel
}
```

### Cells.Parent

返回一个 **Object** 类型的值，该值代表指定 **Cells** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Cells 对象的变量。

### Cells.PreferredWidth

返回或设置指定单元格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。 **Single** 类型，可读写。

**语法**

**express.PreferredWidth**

\*express是一个代表 Cells 对象的变量。

**说明**

如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPoints** ，则 **PreferredWidth** 属性返回或设置宽度（以磅为单位）。如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPercent** ，则 **PreferredWidth** 属性返回或设置宽度（以窗口宽度的百分比表示）。

### Cells.PreferredWidthType

返回或设置用于指定单元格的宽度的首选度量单位。 **WdPreferredWidthType** 类型，只读。

**语法**

**express.PreferredWidthType**

\*express是一个代表 Cells 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为以窗口宽度的百分比来表示宽度，然后将文档中第一张表格的宽度设置为窗口宽度的 50%。*/
function test() {
    let rng = Application.ActiveDocument.Tables.Item(1)
    rng.PreferredWidthType = wdPreferredWidthPercent
    rng.PreferredWidth = 50
}
```

### Cells.Shading

返回一个 **Shading** 对象，该对象指明指定对象的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Cells 对象的变量。

**说明**

**示例**

``` JavaScript
/*本示例对表格 1 的第一行应用水平线纹理。*/function test() {    if (Application.ActiveDocument.Tables.Count >= 1) {        Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Shading.Texture = wdTextureHorizontal    }}
```

### Cells.VerticalAlignment

返回或设置表格中一个或多个单元格中文字的垂直对齐方式。 **WdCellVerticalAlignment** 类型，可读写。

**语法**

**express.VerticalAlignment**

\*express是一个代表 Cells 对象的变量。

**示例**

``` JavaScript
/*本示例在一个新文档中创建一个 3x3 表格，并为表格中的每个单元格分配连续的单元格号。然后将第一行的高度设置为 20 磅，并在单元格的顶端垂直对齐文本。*/
function test() {
    let newDoc = Application.Documents.Add()
    let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)
    let i = 1
    for (let j = 1; j <= myTable.Range.Cells.Count; j++) {
        myTable.Range.Cells.Item(j).Range.InsertAfter("Cell " + j)
        i++
    }
    let rng = myTable.Rows.Item(1)
    rng.Height = 20
    rng.Cells.VerticalAlignment = wdAlignVerticalTop
}
```

### Cells.Width

返回或设置表格单元格的宽度（以磅为单位）。 **Number** 类型，可读写。

**语法**

**express.Width**

\*express是一个代表 Cells 对象的变量。

**示例**

``` JavaScript
/* 获取选区单元格的宽度 */
alert(Application.Selection.Cells.Width);
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Cells](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DocumentField 对象

## 简介

代表文档中的一个公文域。

## 方法

| 序号 | 名称                            | 说明           |
|:----:|:--------------------------------|:---------------|
|  1   | [Delete](#DocumentField.Delete) | 删除指定的域。 |
|  2   | [Select](#DocumentField.Select) | 选择指定的域。 |

## 属性

| 序号 | 名称                                      | 说明                                                                          |
|:----:|:------------------------------------------|:------------------------------------------------------------------------------|
|  1   | [Application](#DocumentField.Application) | 返回一个代表 WPS 应用程序的 [Application](Application) [](Application) 对象。 |
|  2   | [Creator](#DocumentField.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。  |
|  3   | [Hidden](#DocumentField.Hidden)           | 如果为 True ，则隐藏文档中的公文域。 Boolean 类型，可读写。                   |
|  4   | [Name](#DocumentField.Name)               | 返回或设置指定对象的名称。可读/写 String 类型。                               |
|  5   | [Parent](#DocumentField.Parent)           | 返回一个 Object 类型值，该值代表指定 DocumentField 对象的父对象。             |
|  6   | [PrintOut](#DocumentField.PrintOut)       | 该公文域对象是否可以打印 。                                                   |
|  7   | [Range](#DocumentField.Range)             | 返回指定对象的区域 Range。                                                    |
|  8   | [ReadOnly](#DocumentField.ReadOnly)       | 如果为True，代表该公文域值只读。 Boolean 类型， 可读/写                       |
|  9   | [StoryType](#DocumentField.StoryType)     | 返回一个Long类型值，该值代表此公文域的类型。只读。                            |
|  10  | [Value](#DocumentField.Value)             | 返回一个字符串，该字符串为此公文域中的内容。 String 类型，可读/写。           |

## 成员方法

### DocumentField.Delete

删除指定的域。

**语法**

**express.Delete ( *DeleteWithContent* )**

\*express是一个代表 DocumentField 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明                              |
|---------------------|-----------|----------|-----------------------------------|
| *DeleteWithContent* | 可选      | Boolean  | 是否删除该域中的内容，默认为false |

### DocumentField.Select

选择指定的域。

**语法**

**express.Select ()**

\*express是一个代表 DocumentField 对象的变量。

## 成员属性

### DocumentField.Application

返回一个代表 WPS 应用程序的 **[Application](Application)** [](Application) 对象。

**语法**

**express.Application**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 DocumentField 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### DocumentField.Hidden

如果为 **True** ，则隐藏文档中的公文域。 **Boolean** 类型，可读写。

**语法**

**express.Hidden**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.Name

返回或设置指定对象的名称。可读/写 **String** 类型。

**语法**

**express.Name**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.Parent

返回一个 **Object** 类型值，该值代表指定 **DocumentField** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.PrintOut

该公文域对象是否可以打印 。

**语法**

**express.PrintOut**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.Range

返回指定对象的区域 Range。

**语法**

**express.Range**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.ReadOnly

如果为True，代表该公文域值只读。 **Boolean** 类型， 可读/写

**语法**

**express.ReadOnly**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.StoryType

返回一个Long类型值，该值代表此公文域的类型。只读。

**语法**

**express.StoryType**

\*express是一个代表 DocumentField 对象的变量。

### DocumentField.Value

返回一个字符串，该字符串为此公文域中的内容。 **String** 类型，可读/写。

**语法**

**express.Value**

\*express是一个代表 DocumentField 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/DocumentField](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Editor 对象

## 简介

表示已被分配特定权限可编辑文档部分的单个用户。

## 说明

可授予权限的用户包括单独的参与者以及为“文档工作区”站点定义的用户组。

分配给范围和所选内容的权限仅在文档受到保护之后生效。使用 Editors 集合和 Editor 对象可将特定权限分配到文档的节。然后使用 Protect 方法来保护该文档。

使用 Add 集合的 Editors 方法可向指定的用户或组授予权限以修改文档中的范围或所选内容。以下示例为当前用户分配编辑权限以修改活动的所选内容。

``` JavaScript
let objEditor = Selection.Editors.Add(wdEditorCurrent)
```

## 方法

| 序号 | 名称                           | 说明                                       |
|:----:|:-------------------------------|:-------------------------------------------|
|  1   | [Delete](#Editor.Delete)       | 删除指定的 Editor 对象。                   |
|  2   | [DeleteAll](#Editor.DeleteAll) | 删除特定用户在文档中的所有编辑权限。       |
|  3   | [SelectAll](#Editor.SelectAll) | 选择文档中由单个用户插入或编辑的所有形状。 |

## 属性

| 序号 | 名称                               | 说明                                                                                          |
|:----:|:-----------------------------------|:----------------------------------------------------------------------------------------------|
|  1   | [Application](#Editor.Application) | Visual Basic 的 CreateObject 和 GetObject 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。 |
|  2   | [Creator](#Editor.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                  |
|  3   | [ID](#Editor.ID)                   | 将父文档另存为网页时，返回或设置指定对象的标识标签。只读 String 类型。                        |
|  4   | [Name](#Editor.Name)               | 返回或设置指定对象的名称。只读 String 类型。                                                  |
|  5   | [NextRange](#Editor.NextRange)     | 返回一个 Range 对象，该对象代表用户有权修改的下一个区域。                                     |
|  6   | [Parent](#Editor.Parent)           | 返回一个 Object 类型值，该值代表指定 Editor 对象的父对象。                                    |
|  7   | [Range](#Editor.Range)             | 返回一个 Range 对象，该对象代表指定对象所包含的文档部分。                                     |

## 成员方法

### Editor.Delete

删除指定的 **Editor** 对象。

**语法**

**express.Delete ()**

\*express是一个代表 Editor 对象的变量。

### Editor.DeleteAll

删除特定用户在文档中的所有编辑权限。

**语法**

**express.DeleteAll ()**

\*express是一个代表 Editor 对象的变量。

**示例**

``` JavaScript
/*以下示例删除第一个用户在活动文档 Editors 集合中的所有编辑权限。*/
function test()
{
let objEditor = Selection.Editors.Item(1)
objEditor.DeleteAll()
}
```

### Editor.SelectAll

选择文档中由单个用户插入或编辑的所有形状。

**语法**

**express.SelectAll ()**

\*express是一个代表 Editor 对象的变量。

**说明**

此方法不能选择 **InlineShape** 对象。

## 成员属性

### Editor.Application

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

**语法**

**express.Application**

\*express是一个代表 Editor 对象的变量。

### Editor.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Editor 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Editor.ID

将父文档另存为网页时，返回或设置指定对象的标识标签。只读 **String** 类型。

**语法**

**express.ID**

\*express是一个代表 Editor 对象的变量。

### Editor.Name

返回或设置指定对象的名称。只读 **String** 类型。

**语法**

**express.Name**

\*express是一个代表 Editor 对象的变量。

### Editor.NextRange

返回一个 **Range** 对象，该对象代表用户有权修改的下一个区域。

**语法**

**express.NextRange**

\*express是一个代表 Editor 对象的变量。

**说明**

还可使用 **Range** 对象的 **GoToEditableRange** 方法返回用户有权修改的下一区域。

**示例**

``` JavaScript
/*以下示例返回活动选定内容中第一个编辑者的下一区域。*/
function test()
{
let objEditor = Selection.Editors.Item(1)
let objRange = objEditor.NextRange
}
```

### Editor.Parent

返回一个 **Object** 类型值，该值代表指定 **Editor** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Editor 对象的变量。

### Editor.Range

返回一个 **Range** 对象，该对象代表指定对象所包含的文档部分。

**语法**

**express.Range**

\*express是一个代表 Editor 对象的变量。

**说明**

有关从文档返回范围或从形状集合返回形状范围的信息，请参阅 **Range** 方法。

**示例**

``` JavaScript
/*以下示例为活动文档中的第一段应用“标题 1”样式。*/
ActiveDocument.Paragraphs.Item(1).Range.Style = wdStyleHeading1
```

``` JavaScript
/*以下示例复制表格 1 中的第一行。*/
function test()
{
if(ActiveDocument.Tables.Count >= 1) {
    ActiveDocument.Tables.Item(1).Rows.Item(1).Range.Copy()
}
}
```

``` JavaScript
/*以下示例更改文档中第一个批注的文本。*/
function test()
{
let comments = ActiveDocument.Comments.Item(1).Range
comments.Delete()
comments.InsertBefore("new comment text")
}
```

``` JavaScript
/*以下示例在第一节的结尾插入文本。*/
function test()
{
let myRange = ActiveDocument.Sections.Item(1).Range
myRange.MoveEnd(wdCharacter, -1)
myRange.Collapse(wdCollapseEnd)
myRange.InsertParagraphAfter()
myRange.InsertAfter("End of section")
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Editor](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Hyperlink 对象

## 简介

代表一个超链接。 Hyperlink 对象是 Hyperlinks 集合的一个成员。

## 说明

使用 Hyperlink 属性可返回一个与形状相关联的 Hyperlink 对象（一个形状只能有一个超链接）。以下示例激活与活动文档中第一个形状相关联的超链接。

``` JavaScript
Application.ActiveDocument.Shapes.Item(1).Hyperlink.Follow()
```

使用 Hyperlinks ( *Index* ) 可从文档、范围或所选内容返回单个 Hyperlink 对象，其中 *Index* 为索引号。以下示例激活所选内容中的第一个超链接。

``` JavaScript
function test() {
    if(Application.Selection.Hyperlinks.Count >= 1){
       Application.Selection.Hyperlinks.Item(1).Follow()
    }
}
```

## 方法

| 序号 | 名称                                              | 说明                                                                                                                                          |
|:----:|:--------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddToFavorites](#Hyperlink.AddToFavorites)       | 创建文档或超链接的快捷方式，并将其添加到“收藏夹”。                                                                                            |
|  2   | [CreateNewDocument](#Hyperlink.CreateNewDocument) | 新建一篇链接到指定超链接的文档。                                                                                                              |
|  3   | [Delete](#Hyperlink.Delete)                       | 删除指定的超链接。                                                                                                                            |
|  4   | [Follow](#Hyperlink.Follow)                       | 显示一个与指定的 Hyperlink 对象相关的缓存文档（如果该文档已下载）。否则，此方法将解析该超链接，下载目标文档，并在相应的应用程序中显示此文档。 |

## 属性

| 序号 | 名称                                              | 说明                                                                                 |
|:----:|:--------------------------------------------------|:-------------------------------------------------------------------------------------|
|  1   | [Address](#Hyperlink.Address)                     | 返回或设置指定超链接的地址（例如，文件名或 URL）。 String 类型，可读写。             |
|  2   | [Application](#Hyperlink.Application)             | 返回一个代表 WPS 应用程序的 Application 对象。                                       |
|  3   | [Creator](#Hyperlink.Creator)                     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。         |
|  4   | [EmailSubject](#Hyperlink.EmailSubject)           | 返回或设置指定超链接的主题行的文本字符串。 String 类型，可读写。                     |
|  5   | [ExtraInfoRequired](#Hyperlink.ExtraInfoRequired) | 如果在解析指定的超链接时需要额外信息，则该属性值为 True 。 Boolean 类型，只读。      |
|  6   | [Name](#Hyperlink.Name)                           | 返回指定对象的名称。 String 类型，只读。                                             |
|  7   | [Parent](#Hyperlink.Parent)                       | 返回一个 Object 类型值，该值代表指定 Hyperlink 对象的父对象。                        |
|  8   | [Range](#Hyperlink.Range)                         | 返回一个 Range 对象，该对象代表超链接中包含的文档部分。                              |
|  9   | [ScreenTip](#Hyperlink.ScreenTip)                 | 返回或设置当鼠标指针置于指定超链接之上时显示的“屏幕提示”文字。 String 类型，可读写。 |
|  10  | [Shape](#Hyperlink.Shape)                         | 返回一个指定超链接或图表顶点的 Shape 对象。                                          |
|  11  | [SubAddress](#Hyperlink.SubAddress)               | 返回或设置指定超链接目标中命名的位置。 String 类型，可读写。                         |
|  12  | [Target](#Hyperlink.Target)                       | 返回或设置在其中加载超链接的图文框或窗口的名称。可读写 String 类型。                 |
|  13  | [TextToDisplay](#Hyperlink.TextToDisplay)         | 返回或设置文档中指定超链接的可视文本。 String 类型，可读写。                         |
|  14  | [Type](#Hyperlink.Type)                           | 返回超链接的类型。只读 MsoHyperlinkType 类型。                                       |

## 成员方法

### Hyperlink.AddToFavorites

创建文档或超链接的快捷方式，并将其添加到“收藏夹”。

**语法**

**express.AddToFavorites ()**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

``` JavaScript
/*本示例为活动文档的所有超链接创建快捷方式，并将其添加到“收藏夹”文件夹。*/
function test() {
  for(let i = 1; i <= Application.ActiveDocument.Hyperlinks.Count; i++){
      Application.ActiveDocument.Hyperlinks.Item(i).AddToFavorites()
  }
}

/*本示例为 Sales.doc 创建快捷方式，并将其添加至“收藏夹”文件夹。如果 Sales.doc 还未打开，本示例将从 C:\Documents 文件夹打开该文档。*/
function test() {
  let isOpen
  for(let i = 1; i <= Application.Documents.Count; i++){
      if(Application.Documents.Item(i).Name == "sales.doc"){
         isOpen = true
      }
  }
  if(isOpen != true){
      Application.Documents.Open("C:\\Documents\\Sales.doc")
      Application.Documents.Item("Sales.doc").AddToFavorites()
  }
}
```

### Hyperlink.CreateNewDocument

新建一篇链接到指定超链接的文档。

**语法**

**express.CreateNewDocument ( *FileName* , *EditNow* , *Overwrite* )**

\*express是一个代表 Hyperlink 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                                |
|-------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------|
| *FileName*  | 必选      | String   | 指定文档的文件名。                                                                                                                                  |
| *EditNow*   | 必选      | Boolean  | 如果为 True，则立即在相关编辑环境中打开指定文档。默认值为 True。                                                                                    |
| *Overwrite* | 必选      | Boolean  | 如果为 True，则覆盖相同文件夹中任何现有的同名文件。如果为 False，则保留所有现有的同名文件，并且参数 FileName 将指定一个新的文件名。默认值为 False。 |

**示例**

``` JavaScript
/*本示例基于第一篇文档中新的超链接新建一篇文档，然后在 WPS 中加载该新文档以进行编辑。该文档名为“Overview.doc”，它将覆盖 \\Server1\Annual 文件夹中的任何同名文件。*/
function test() {
  let objHyper = Application.Documents.Item(1).Hyperlinks.Add(Selection.Range, "\\Server1\\Annual\\Overview.doc")
  objHyper.CreateNewDocument("\\Server1\\Annual\\Overview.doc", true, true)
}
```

### Hyperlink.Delete

删除指定的超链接。

**语法**

**express.Delete ()**

\*express是一个代表 Hyperlink 对象的变量。

### Hyperlink.Follow

显示一个与指定的 **Hyperlink** 对象相关的缓存文档（如果该文档已下载）。否则，此方法将解析该超链接，下载目标文档，并在相应的应用程序中显示此文档。

**语法**

**express.Follow ( *NewWindow* , *AddHistory* , *ExtraInfo* , *Method* , *HeaderInfo* )**

\*express是一个代表 Hyperlink 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                  |
|--------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *NewWindow*  | 必选      | Variant  | 如果该属性值为 True，则在新窗口中显示目标文档。默认值为 False。                                                                                                                                                                       |
| *AddHistory* | 必选      | Variant  | 该参数为以后使用而保留。                                                                                                                                                                                                              |
| *ExtraInfo*  | 必选      | Variant  | 指定用于解析超链接的 HTTP 附加信息的字符串或字节数组。例如，您可以使用 ExtraInfo 来指定图像映射的坐标、窗体的内容或 FAT 文件的名称。根据 Method 的值来决定是发布还是附加该字符串。可使用 ExtraInfoRequired 属性确定是否需要额外信息。 |
| *Method*     | 必选      | Variant  | 指定 HTTP 附加信息的处理方式。可以是任何 MsoExtraInfoMethod 常量。                                                                                                                                                                    |
| *HeaderInfo* | 必选      | Variant  | 指定 HTTP 请求标题信息的字符串。默认值为一个空字符串。可以使用下列语法将几个标题行合并成一个字符串："string1" & vbCr & "string2"。指定的字符串自动转换为 ANSI 字符。注意，HeaderInfo 参数可能会覆盖默认的 HTTP 标题字段。             |

**说明**

如果超链接使用文件协议，该方法不是下载而是打开该文档。

**示例**

``` JavaScript
/*本示例显示 Home.doc 中的第一个超链接。*/
Application.Documents.Item("Home.doc").Hyperlinks.Item(1).Follow()

/*本示例首先插入指向 www.msn.com 的超链接，然后显示此超链接。*/
function test() {
  Application.Selection.Collapse(wdCollapseEnd)
  Application.Selection.InsertAfter("MSN ")
  Application.Selection.Previous()
  let hypTemp = Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "http://www.msn.com")
      hypTemp.Follow(false, true)
}
```

## 成员属性

### Hyperlink.Address

返回或设置指定超链接的地址（例如，文件名或 URL）。 **String** 类型，可读写。

**语法**

**express.Address**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

如果不存在与某对象关联的超链接，则设置 **Address** 属性将返回一个错误。在此情况下，请将 **Add** 方法用于 **Hyperlinks** 集合来添加超链接。以下示例说明了具体的操作方法。

``` JavaScript
Application.ActiveDocument.Hyperlinks.Add(Selection.Range,"http://www.wps.cn")
```

**示例**

``` JavaScript
/*本示例可实现的功能是：将超链接添加到活动文档的选定内容，设置地址然后在消息框中显示该地址。*/
let aHLink = Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "http://forms")
alert("The hyperlink goes to " + aHLink.Address)

/*如果活动文档包含超链接，则本示例会在文档的末尾插入超链接目标的列表。*/
let myRange = Application.ActiveDocument.Range(ActiveDocument.Content.End - 1)
let Count = 0
for(let i = 1; i <= Application.ActiveDocument.Hyperlinks.Count; i++){
    Count = Count + 1
    myRange.InsertAfter("Hyperlink #" + Count + "\t")
    myRange.InsertAfter(Application.ActiveDocument.Hyperlinks.Item(i).Address)
    myRange.InsertParagraphAfter()
}
```

### Hyperlink.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 Hyperlink 对象的变量。

### Hyperlink.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Hyperlink.EmailSubject

返回或设置指定超链接的主题行的文本字符串。 **String** 类型，可读写。

**语法**

**express.EmailSubject**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

该主题行附加到超链接的 Internet 地址或 URL 中。该属性通常用于电子邮件超链接。其值优先于在相同 **Hyperlink** 对象的 Address 属性中指定的其他任何电子邮件主题。

**示例**

``` JavaScript
/*本示例实现的功能是：检查活动文档中的电子邮件超链接。如果发现某个超链接有空主题行，则为其添加主题“NewProducts”。*/
function test() {
  for(let i = 1; i <= Application.ActiveDocument.Hyperlinks.Count; i++){
      if(Application.ActiveDocument.Hyperlinks.Item(i).Address.indexOf("mailto*") == 0 && Application.ActiveDocument.Hyperlinks.Item(i).Address == ActiveDocument.Hyperlinks.Item(i).EmailSubject){
          Application.ActiveDocument.Hyperlinks.Item(i).EmailSubject = "NewProducts"
      }
  }
}
```

### Hyperlink.ExtraInfoRequired

如果在解析指定的超链接时需要额外信息，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.ExtraInfoRequired**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

可以将 *ExtraInfo* 参数与 **Follow** 或 **FollowHyperlink** 方法结合使用，以便指定其他信息。例如，可以使用 *ExtraInfo* 指定图像映射的坐标、窗体的内容或 FAT 文件名。

**示例**

``` JavaScript
/*本示例首先插入指向 www.msn.com 的超链接，如果不需要额外信息，则跟踪此超链接。*/
function test() {
  Application.Selection.Collapse(wdCollapseEnd)
  Application.Selection.InsertAfter("MSN ")
  Application.Selection.Previous()
  
  let hypTemp = Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "http://www.msn.com")
  if(hypTemp.ExtraInfoRequired == false){
     hypTemp.Follow()
  }
}
```

### Hyperlink.Name

返回指定对象的名称。 **String** 类型，只读。

**语法**

**express.Name**

\*express是一个代表 Hyperlink 对象的变量。

### Hyperlink.Parent

返回一个 **Object** 类型值，该值代表指定 **Hyperlink** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Hyperlink 对象的变量。

### Hyperlink.Range

返回一个 **Range** 对象，该对象代表超链接中包含的文档部分。

**语法**

**express.Range**

\*express是一个代表 Hyperlink 对象的变量。

### Hyperlink.ScreenTip

返回或设置当鼠标指针置于指定超链接之上时显示的“屏幕提示”文字。 **String** 类型，可读写。

**语法**

**express.ScreenTip**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

``` JavaScript
/*本示例设置活动文档中第一个超链接的“屏幕提示”文字。*/
Application.ActiveDocument.Hyperlinks.Item(1).ScreenTip = "Home"
```

### Hyperlink.Shape

返回一个指定超链接或图表顶点的 **Shape** 对象。

**语法**

**express.Shape**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

如果超链接不是由形状表示的，则会产生错误。

**示例**

``` JavaScript
/*以下示例改变形状的填充颜色，该形状代表活动文档中的第一个超链接。为使本示例能运行，超链接必须由形状表示。*/
Application.ActiveDocument.Hyperlinks.Item(1).Shape.Fill.ForeColor.RGB = (255, 255, 0)
```

### Hyperlink.SubAddress

返回或设置指定超链接目标中命名的位置。 **String** 类型，可读写。

**语法**

**express.SubAddress**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

命名的位置可以是 WPS 文档中的书签、ET 工作表中命名的单元格或单元格引用、Access 数据库中命名的对象，或者是 WPP 演示文稿中幻灯片的序号。

**示例**

``` JavaScript
/*本例显示所选超链接的子地址。*/
function test() {
  if(Application.Selection.Range.Hyperlinks.Count >= 1){
      alert(Application.Selection.Range.Hyperlinks.Item(1).SubAddress)
  }
}

/*本例为活动文档中的选定内容添加超链接，设置超链接的目标和子地址，然后将其显示在消息框中。*/
function test() {
  let SCut = Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "C:\\My Documents\\Other.doc", "temp")
  alert("The hyperlink goes to " + SCut.SubAddress)
}
```

### Hyperlink.Target

返回或设置在其中加载超链接的图文框或窗口的名称。可读写 **String** 类型。

**语法**

**express.Target**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

``` JavaScript
/*以下示例设置在一个新的浏览器窗口中打开指定的超链接。*/
Application.ActiveDocument.Hyperlinks.Item(1).Target = "_blank"

/*以下示例设置在名为“left”的图文框中打开指定的超链接。*/
Application.ActiveDocument.Hyperlinks.Item(1).Target = "left"
```

### Hyperlink.TextToDisplay

返回或设置文档中指定超链接的可视文本。 **String** 类型，可读写。

**语法**

**express.TextToDisplay**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

``` JavaScript
/*本示例设置活动文档中第一个超链接的显示文本。*/
Application.ActiveDocument.Hyperlinks.Item(1).TextToDisplay = "Follow this link for more information..."
```

### Hyperlink.Type

返回超链接的类型。只读 **MsoHyperlinkType** 类型。

**语法**

**express.Type**

\*express是一个代表 Hyperlink 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Hyperlink](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Sentences 对象

## 简介

Range 对象的集合，这些对象代表选定内容、区域或文档中的所有句子。不存在 Sentence 对象。

## 方法

| 序号 | 名称                    | 说明                                     |
|:----:|:------------------------|:-----------------------------------------|
|  1   | [Item](#Sentences.Item) | 返回集合中的单个 Range 对象。返回Range值 |

## 属性

| 序号 | 名称                                  | 说明                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------|
|  1   | [Application](#Sentences.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                  |
|  2   | [Count](#Sentences.Count)             | 返回一个 Long 类型的值，该值代表集合中句子的数量。只读。                        |
|  3   | [Creator](#Sentences.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。    |
|  4   | [First](#Sentences.First)             | 返回一个 Range 对象，该对象代表文档、区域或选定内容中的句子集合中的第一个句子。 |
|  5   | [Last](#Sentences.Last)               | 返回一个 Range 对象，该对象代表文档、所选内容或范围中的最后一句。               |
|  6   | [Parent](#Sentences.Parent)           | 返回一个 Object 类型值，该值代表指定 Sentences 对象的父对象。                   |

## 成员方法

### Sentences.Item

返回集合中的单个 **Range** 对象。返回Range值

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Sentences 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

## 成员属性

### Sentences.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Sentences 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Sentences.Count

返回一个 **Long** 类型的值，该值代表集合中句子的数量。只读。

**语法**

**express.Count**

\*express是一个代表 Sentences 对象的变量。

### Sentences.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Sentences 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Sentences.First

返回一个 Range 对象，该对象代表文档、区域或选定内容中的句子集合中的第一个句子。

**语法**

**express.First**

\*express是一个代表 Sentences 对象的变量。

**示例**

``` JavaScript
返回一个 Range 对象，该对象代表文档、区域或选定内容中的句子集合中的第一个句子。
```

### Sentences.Last

返回一个 **Range** 对象，该对象代表文档、所选内容或范围中的最后一句。

**语法**

**express.Last**

\*express是一个代表 Sentences 对象的变量。

### Sentences.Parent

返回一个 **Object** 类型值，该值代表指定 **Sentences** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Sentences 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Sentences](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ShapeRange 对象

## 简介

代表一个形状范围，即某个文档中的一组形状。一个形状范围可以只包含一个形状，也可以包含该文档中的所有形状。

## 方法

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Align">Align</a></td>
<td style="text-align: left;" data-valian="middle"><p>对齐指定形状范围中的形状。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Apply">Apply</a></td>
<td style="text-align: left;" data-valian="middle"><p>应用于使用 PickUp 方法复制的特定图形格式。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.CanvasCropBottom">CanvasCropBottom</a></td>
<td style="text-align: left;" data-valian="middle"><p>从绘图画布的底部裁剪一定百分比的 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">绘图画布</span> 高度。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.CanvasCropLeft">CanvasCropLeft</a></td>
<td style="text-align: left;" data-valian="middle"><p>从绘图画布左侧裁剪一定百分比的 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">绘图画布</span> 宽度。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.CanvasCropRight">CanvasCropRight</a></td>
<td style="text-align: left;" data-valian="middle"><p>从绘图画布右侧裁剪一定百分比的 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">绘图画布</span> 宽度。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.CanvasCropTop">CanvasCropTop</a></td>
<td style="text-align: left;" data-valian="middle"><p>从绘图画布顶部裁剪一定百分比的 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">绘图画布</span> 高度。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ConvertToInlineShape">ConvertToInlineShape</a></td>
<td style="text-align: left;" data-valian="middle"><p>将文档绘图层的指定形状转换为文字层的内嵌形状。只能转换代表图片、OLE 对象或 ActiveX 控件的形状。</p>
<p>返回 InlineShape 值</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Delete">Delete</a></td>
<td style="text-align: left;" data-valian="middle"><p>删除指定区域的形状。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Distribute">Distribute</a></td>
<td style="text-align: left;" data-valian="middle"><p>在指定的形状范围内均匀分布形状。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Duplicate">Duplicate</a></td>
<td style="text-align: left;" data-valian="middle"><p>创建指定的 ShapeRange 对象的副本，以标准的偏移将新图形区域从原图形添加至 Shapes 集合，然后返回 Shape 对象。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Flip">Flip</a></td>
<td style="text-align: left;" data-valian="middle"><p>水平或垂直翻转一个图形。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Group">Group</a></td>
<td style="text-align: left;" data-valian="middle"><p>组合指定区域中的图形并将组合图形作为单个 Shape 对象返回。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.IncrementLeft">IncrementLeft</a></td>
<td style="text-align: left;" data-valian="middle"><p>将指定形状水平移动指定的磅数。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.IncrementRotation">IncrementRotation</a></td>
<td style="text-align: left;" data-valian="middle"><p>使指定的形状绕 Z 轴旋转指定的角度。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.IncrementTop">IncrementTop</a></td>
<td style="text-align: left;" data-valian="middle"><p>以指定磅数垂直移动指定形状。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">16</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回集合中的单个 Shape 对象。返回Shape值</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">17</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.PickUp">PickUp</a></td>
<td style="text-align: left;" data-valian="middle"><p>复制指定形状的格式。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">18</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ScaleHeight">ScaleHeight</a></td>
<td style="text-align: left;" data-valian="middle"><p>按指定的比例缩放形状范围的高度。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">19</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ScaleWidth">ScaleWidth</a></td>
<td style="text-align: left;" data-valian="middle"><p>按指定比例调整形状的宽度。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">20</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Select">Select</a></td>
<td style="text-align: left;" data-valian="middle"><p>选择指定的形状范围。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">21</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.SetShapesDefaultProperties">SetShapesDefaultProperties</a></td>
<td style="text-align: left;" data-valian="middle"><p>将文档中默认形状的格式应用于指定的形状范围。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">22</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Ungroup">Ungroup</a></td>
<td style="text-align: left;" data-valian="middle"><p>取消指定形状范围中所有组合形状的组合，分解指定形状或形状范围中图片和 OLE 对象的组合，将取消组合后的形状以单个 ShapeRange 对象的形式返回。</p>
<p>返回 ShapeRange值</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">23</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ZOrder">ZOrder</a></td>
<td style="text-align: left;" data-valian="middle"><p>将集合中指定的形状区域移动到其他形状的前面或后面（也就是说，更改形状区域在 Z 顺序中的位置）。</p></td>
</tr>
</tbody>
</table>

## 属性

| 序号 | 名称                                                                 | 说明                                                                                                                                                                                                                                                                                     |
|:----:|:---------------------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Adjustments](#ShapeRange.Adjustments)                               | 返回一个 Adjustments 对象, 该对象包含所有对指定 ShapeRange 对象（代表自选图形或艺术字）进行调整操作的调整值。只读。                                                                                                                                                                      |
|  2   | [AlternativeText](#ShapeRange.AlternativeText)                       | 返回或设置与网页的图形相关联的 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">可选文字</span> 。 String 类型，可读写。                                                                                                     |
|  3   | [Anchor](#ShapeRange.Anchor)                                         | 返回一个 Range 对象，该对象代表指定图形区域的锁定范围。只读。                                                                                                                                                                                                                            |
|  4   | [Application](#ShapeRange.Application)                               | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                                                                                                                                           |
|  5   | [AutoShapeType](#ShapeRange.AutoShapeType)                           | 返回或设置指定的 ShapeRange 对象的图形类型，该对象不是代表线条或任意多边形，而是代表自选图形。 MsoAutoShapeType 类型，可读写。                                                                                                                                                           |
|  6   | [BackgroundStyle](#ShapeRange.BackgroundStyle)                       | 设置或返回指定形状范围中形状的背景样式。可读/写 MsoBackgroundStyleIndex 。                                                                                                                                                                                                               |
|  7   | [Callout](#ShapeRange.Callout)                                       | 返回 CalloutFormat 对象，该对象包含指定图形的标注格式属性。只读。                                                                                                                                                                                                                        |
|  8   | [CanvasItems](#ShapeRange.CanvasItems)                               | 返回一个 CanvasShapes 对象，该对象代表绘图画布上图形的集合。                                                                                                                                                                                                                             |
|  9   | [Child](#ShapeRange.Child)                                           | 如果图形是子图形或位于图形区域的所有图形都是同一父图形的子图形，则该属性值为 True 。 MsoTriState 类型，只读。                                                                                                                                                                            |
|  10  | [Count](#ShapeRange.Count)                                           | 返回一个 Long 类型的值，该值代表集合中图形的数量。只读。                                                                                                                                                                                                                                 |
|  11  | [Creator](#ShapeRange.Creator)                                       | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                                                                                                                                             |
|  12  | [Fill](#ShapeRange.Fill)                                             | 返回一个 FillFormat 对象，该对象包含指定图形的填充格式属性。只读。                                                                                                                                                                                                                       |
|  13  | [Glow](#ShapeRange.Glow)                                             | 返回一个 GlowFormat 对象，该对象代表形状区域的发光格式。只读。                                                                                                                                                                                                                           |
|  14  | [GroupItems](#ShapeRange.GroupItems)                                 | 返回一个 GroupShapes 对象，该对象代表指定图形组中的单个图形。只读。                                                                                                                                                                                                                      |
|  15  | [Height](#ShapeRange.Height)                                         | 返回或设置指定图形区域的高度。 Single 类型，可读写。                                                                                                                                                                                                                                     |
|  16  | [HeightRelative](#ShapeRange.HeightRelative)                         | 返回或设置一个 Single 类型的值，该值代表将形状区域大小调整到的目标形状的百分比。可读写。                                                                                                                                                                                                 |
|  17  | [HorizontalFlip](#ShapeRange.HorizontalFlip)                         | 表示该形状范围已进行水平翻转。只读 MsoTriState 类型。                                                                                                                                                                                                                                    |
|  18  | [Hyperlink](#ShapeRange.Hyperlink)                                   | 返回一个 Hyperlink 对象，该对象代表与指定 ShapeRange 对象相关联的超链接。只读。                                                                                                                                                                                                          |
|  19  | [ID](#ShapeRange.ID)                                                 | 返回形状范围的标识类型。只读 Long 类型。                                                                                                                                                                                                                                                 |
|  20  | [LayoutInCell](#ShapeRange.LayoutInCell)                             | 返回一个 Long 类型的值，该值代表表格中的形状是显示在表格内部还是表格外部。                                                                                                                                                                                                               |
|  21  | [Left](#ShapeRange.Left)                                             | 返回或设置一个 Single 类型的值，该值代表指定形状范围的水平位置，以磅为单位。也可以是任何有效的 WdShapePosition 常量。可读写。                                                                                                                                                            |
|  22  | [LeftRelative](#ShapeRange.LeftRelative)                             | 返回或设置一个 Single 类型的值，该值代表形状区域左侧的相对位置。可读写。                                                                                                                                                                                                                 |
|  23  | [Line](#ShapeRange.Line)                                             | 返回一个 LineFormat 对象，该对象包含指定形状范围的线条格式属性。只读。                                                                                                                                                                                                                   |
|  24  | [LockAnchor](#ShapeRange.LockAnchor)                                 | 如果指定 ShapeRange 对象的锁定标记锁定到锁定范围，则该属性值为 True 。可读写 Long 类型。                                                                                                                                                                                                 |
|  25  | [LockAspectRatio](#ShapeRange.LockAspectRatio)                       | 如果在调整指定形状的大小时保留其最初比例，则该属性值为 MsoTrue ；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 MsoFalse 。可读写 MsoTriState 类型。                                                                                                                           |
|  26  | [Name](#ShapeRange.Name)                                             | 返回或设置指定对象的名称。 String 类型，可读写。                                                                                                                                                                                                                                         |
|  27  | [Nodes](#ShapeRange.Nodes)                                           | 返回一个 ShapeNodes 集合，该集合代表指定形状的几何描述。                                                                                                                                                                                                                                 |
|  28  | [Parent](#ShapeRange.Parent)                                         | 返回一个 Object 类型值，该值代表指定 ShapeRange 对象的父对象。                                                                                                                                                                                                                           |
|  29  | [ParentGroup](#ShapeRange.ParentGroup)                               | 返回一个 Shape 对象，该对象代表形状范围的通用父形状。                                                                                                                                                                                                                                    |
|  30  | [PictureFormat](#ShapeRange.PictureFormat)                           | 返回一个 PictureFormat 对象，该对象包含指定形状范围的图片格式属性。只读。                                                                                                                                                                                                                |
|  31  | [Reflection](#ShapeRange.Reflection)                                 | 返回一个 ReflectionFormat 对象，该对象代表形状区域的反射格式。只读。                                                                                                                                                                                                                     |
|  32  | [RelativeHorizontalPosition](#ShapeRange.RelativeHorizontalPosition) | 指定形状范围的相对水平位置。可读写 WdRelativeHorizontalPosition 类型。                                                                                                                                                                                                                   |
|  33  | [RelativeHorizontalSize](#ShapeRange.RelativeHorizontalSize)         | 返回或设置一个 WdRelativeHorizontalSize 常量，该常量代表形状区域相对的对象。可读写。                                                                                                                                                                                                     |
|  34  | [RelativeVerticalPosition](#ShapeRange.RelativeVerticalPosition)     | 指定形状范围的相对垂直位置。可读写 WdRelativeHorizontalPosition 类型。                                                                                                                                                                                                                   |
|  35  | [RelativeVerticalSize](#ShapeRange.RelativeVerticalSize)             | 返回或设置一个 WdRelativeVerticalSize 常量，该常量代表形状区域相对的对象。可读写。                                                                                                                                                                                                       |
|  36  | [Rotation](#ShapeRange.Rotation)                                     | 返回或设置指定形状绕 Z 轴旋转的度数。可读写 Single 类型。                                                                                                                                                                                                                                |
|  37  | [Shadow](#ShapeRange.Shadow)                                         | 返回一个 ShadowFormat 对象，该对象代表指定形状的阴影格式。                                                                                                                                                                                                                               |
|  38  | [ShapeStyle](#ShapeRange.ShapeStyle)                                 | 设置或返回指定形状范围中形状的形状样式。可读/写 MsoShapeStyleIndex。                                                                                                                                                                                                                     |
|  39  | [SoftEdge](#ShapeRange.SoftEdge)                                     | 返回一个 SoftEdgeFormat 对象，该对象代表形状区域的软边缘格式。只读。                                                                                                                                                                                                                     |
|  40  | [TextEffect](#ShapeRange.TextEffect)                                 | 返回一个 TextEffectFormat 对象，该对象包含指定形状的文本效果格式属性。只读。                                                                                                                                                                                                             |
|  41  | [TextFrame](#ShapeRange.TextFrame)                                   | 返回一个 TextFrame 对象，该对象包含指定形状范围的文字。                                                                                                                                                                                                                                  |
|  42  | [TextFrame2](#ShapeRange.TextFrame2)                                 | 返回一个 TextFrame2 对象，包含指定形状区域的文本。只读。                                                                                                                                                                                                                                 |
|  43  | [ThreeD](#ShapeRange.ThreeD)                                         | 返回一个 ThreeDFormat 对象，该对象包含指定形状范围的三维格式属性。只读。                                                                                                                                                                                                                 |
|  44  | [Title](#ShapeRange.Title)                                           | 返回或设置 String 类型值，该值包含指定形状范围中形状的标题。可读/写。                                                                                                                                                                                                                    |
|  45  | [Top](#ShapeRange.Top)                                               | 返回或设置指定形状或形状范围的垂直位置（以磅为单位）。可读写 Single 类型。                                                                                                                                                                                                               |
|  46  | [TopRelative](#ShapeRange.TopRelative)                               | 返回或设置一个 Single 类型的值，该值代表形状区域顶部的相对位置。可读写。                                                                                                                                                                                                                 |
|  47  | [Type](#ShapeRange.Type)                                             | 返回形状类型。只读 MsoShapeType 类型。                                                                                                                                                                                                                                                   |
|  48  | [VerticalFlip](#ShapeRange.VerticalFlip)                             | 如果指定形状围绕垂直轴进行翻转，则该属性值为 True 。 MsoTriState 类型，只读。                                                                                                                                                                                                            |
|  49  | [Vertices](#ShapeRange.Vertices)                                     | 该属性以一系列 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">坐标对</span> 的形式返回指定任意多边形图形顶点（和贝赛尔曲线的控点）的坐标。可将该属性返回的数组用作 AddCurve 或 AddPolyLine 方法的参数。只读 Variant 类型。 |
|  50  | [Visible](#ShapeRange.Visible)                                       | 如果指定对象或应用于该对象的格式是可见的，则该属性值为 True 。 MsoTriState 类型，可读写。                                                                                                                                                                                                |
|  51  | [Width](#ShapeRange.Width)                                           | 返回或设置范围内形状的宽度（以磅为单位）。可读写 Long 类型。                                                                                                                                                                                                                             |
|  52  | [WidthRelative](#ShapeRange.WidthRelative)                           | 返回或设置一个 Single 类型的值，该值代表形状区域的相对宽度。可读写。                                                                                                                                                                                                                     |
|  53  | [WrapFormat](#ShapeRange.WrapFormat)                                 | 返回一个 WrapFormat 对象，该对象包含在指定的形状范围四周文字环绕的属性。只读。                                                                                                                                                                                                           |
|  54  | [ZOrderPosition](#ShapeRange.ZOrderPosition)                         | 返回一个 Long 类型的值，该值代表指定的形状在 Z 顺序中的位置。只读。                                                                                                                                                                                                                      |

## 成员方法

### ShapeRange.Align

对齐指定形状范围中的形状。

**语法**

**express.Align ( *Align* , *RelativeTo* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型    | 说明                                                                                        |
|--------------|-----------|-------------|---------------------------------------------------------------------------------------------|
| *Align*      | 必选      | MsoAlignCmd | 指定特定形状范围中形状的对齐方式。                                                          |
| *RelativeTo* | 必选      | Long        | 如果该参数值为 True，则相对于文档边缘对齐形状。如果该参数值为 False，则相对于彼此对齐形状。 |

**示例**

``` JavaScript
/*以下示例将活动文档中选定范围内所有形状的左边缘与该范围最左边的形状的左边缘对齐。*/
function test()
{
    let myShapeRange = Application.Selection.ShapeRange
    myShapeRange.Align(msoAlignLefts, false)
}  
```

### ShapeRange.Apply

应用于使用 **PickUp** 方法复制的特定图形格式。

**语法**

**express.Apply ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

如果以前没有使用 **PickUp** 方法复制 **ShapeRange** 对象的格式，则使用 **Apply** 方法会导致错误。

### ShapeRange.CanvasCropBottom

从绘图画布的底部裁剪一定百分比的 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">绘图画布</span> 高度。

**语法**

**express.CanvasCropBottom ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                       |
|-------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 裁剪后所需保留的绘图画布高度的百分比数值。输入 0.9 作为增量从底部裁剪绘图画布高度的百分之十。输入 0.1 从底部裁剪绘图画布高度的百分之九十。 |

**示例**

``` JavaScript
/*假定活动文档中的第一个图形为绘图画布的情况下，则本示例从活动文档中第一块绘图画布的底部裁剪绘图画布高度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在文档中添加绘图画布。*/
function test()
{
    let shpCanvas = Application.ActiveDocument.Shapes.Item(1)
    shpCanvas.CanvasCropBottom(0.75)
}
```

### ShapeRange.CanvasCropLeft

从绘图画布左侧裁剪一定百分比的 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">绘图画布</span> 宽度。

**语法**

**express.CanvasCropLeft ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                               |
|-------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 裁剪绘图画布后所需保留的绘图画布宽度的百分比数值。输入 0.9 作为增量从左侧裁剪绘图画布宽度的百分之十。输入 0.1 从左侧裁剪绘图画布宽度的百分之九十。 |

**示例**

``` JavaScript
/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从活动文档中第一块绘图画布的左侧裁剪绘图画布宽度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在文档中添加绘图画布。*/
function test()
{
    let shpCanvas = Application.ActiveDocument.Shapes.Item(1)
    shpCanvas.CanvasCropLeft(0.75)
}
```

### ShapeRange.CanvasCropRight

从绘图画布右侧裁剪一定百分比的 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">绘图画布</span> 宽度。

**语法**

**express.CanvasCropRight ( *Incremet* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                                                               |
|------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------|
| *Incremet* | 必选      | Single   | 裁剪绘图画布后所需保留的绘图画布宽度的百分比数值。输入 0.9 作为增量从右侧裁剪绘图画布宽度的百分之十。输入 0.1 从右侧裁剪绘图画布宽度的百分之九十。 |

**示例**

``` JavaScript
/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从右侧裁剪活动文档中第一块绘图画布宽度的百分之二十五。如果不是这种情况，需要使用 AddCanvas 方法在文档中添加绘图画布。*/
function test(){
    let shpCanvas = Application.ActiveDocument.Shapes.Item(1)
    shpCanvas.CanvasCropRight(0.75)
}
```

### ShapeRange.CanvasCropTop

从绘图画布顶部裁剪一定百分比的 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">绘图画布</span> 高度。

**语法**

**express.CanvasCropTop ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                               |
|-------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 裁剪绘图画布后所需保留的绘图画布高度的百分比数值。输入 0.9 作为增量从顶部裁剪绘图画布高度的百分之十。输入 0.1 从顶部裁剪绘图画布高度的百分之九十。 |

**示例**

``` JavaScript
/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从活动文档中第一块绘图画布的顶部裁剪绘图画布高度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在活动文档中添加绘图画布。*/
function test()
{
    let shpCanvas = Application.ActiveDocument.Shapes.Item(1)
    shpCanvas.CanvasCropTop(0.75)
}
```

### ShapeRange.ConvertToInlineShape

将文档绘图层的指定形状转换为文字层的内嵌形状。只能转换代表图片、OLE 对象或 ActiveX 控件的形状。

返回 InlineShape **值**

**语法**

**express.ConvertToInlineShape ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

支持附加文字的形状不能转换为内嵌形状。对于这样的形状，请使用 **ConvertToFrame** 方法。

如果对包含多个形状的 **ShapeRange** 对象使用此方法，将导致出错。

**示例**

``` JavaScript
/*本示例将 MyDoc.doc 中的每张图片转换为内嵌形状。*/

function test()
{
  for(let i = 1; i <= Application.Documents.Item("MyDoc.doc").Shapes.Count; i++){
    if(Application.Documents.Item("MyDoc.doc").Shapes.Item(i).Type == msoPicture){
        Application.Documents.Item("MyDoc.doc").Shapes.Item(i).ConvertToInlineShape()
    }
  }
}
```

### ShapeRange.Delete

删除指定区域的形状。

**语法**

**express.Delete ()**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Distribute

在指定的形状范围内均匀分布形状。

**语法**

**express.Distribute ( *Distribute* , *RelativeTo* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型         | 说明                                                                                                                |
|--------------|-----------|------------------|---------------------------------------------------------------------------------------------------------------------|
| *Distribute* | 必选      | MsoDistributeCmd | 指定图形是横向分布还是纵向分布。                                                                                    |
| *RelativeTo* | 必选      | Long             | 如果为 True，则将图形均布在页面的整个横向或者纵向空间上。如果为 False，则图形在其原来占有的横向或者纵向空间内分布。 |

**说明**

可选择纵向分布或横向分布，也可选择将各图形分布于整个页面还是仅限于原先占据的空间。

**示例**

``` JavaScript
/*本示例可实现的功能为：定义一个图形区域，该区域包含活动文档中所有的自选图形，再将各图形水平分布在此区域内。*/
function test()
{
  let docShapes = Application.ActiveDocument.Shapes
  let numShapes = docShapes.Count
  let autoShpArray = []
  let numAutoShapes
  let asRange
  if(numShapes > 1){ 
    numAutoShapes = 0
    for(let i = 1; i <= numShapes; i++) {
        if(docShapes.Item(i).Type == msoAutoShape) {
            numAutoShapes++
            autoShpArray[numAutoShapes] = docShapes.Item(i).Name
        }
    }
    if(numAutoShapes > 1) {
        asRange = docShapes.Range(autoShpArray)      
        asRange.Distribute(msoDistributeHorizontally, false)
    }
  }
}
```

### ShapeRange.Duplicate

创建指定的 **ShapeRange** 对象的副本，以标准的偏移将新图形区域从原图形添加至 **Shapes** 集合，然后返回 **Shape** 对象。

**语法**

**express.Duplicate ()**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*本示例创建活动文档第一个图形的副本，然后改变新图形的填充格式。*/
function test()
{
    let newShape = Application.ActiveDocument.Shapes.Item(1).Duplicate()
    newShape.Fill.PresetGradient(msoGradientVertical, 1, msoGradientGold)
}
```

### ShapeRange.Flip

水平或垂直翻转一个图形。

**语法**

**express.Flip ( *FlipCmd* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型   | 说明       |
|-----------|-----------|------------|------------|
| *FlipCmd* | 必选      | MsoFlipCmd | 翻转方向。 |

**示例**

``` JavaScript
/*本示例首先向活动文档添加一个三角形，然后复制此三角形，再垂直翻转复制的三角形并将其设为红色。*/
function test(){
    let shapes = Application.ActiveDocument.Shapes.AddShape(msoShapeRightTriangle, 150, 150, 50, 50).Duplicate()
    shapes.Fill.ForeColor.RGB = (255, 0, 0)
    shapes.Flip(msoFlipVertical)
}
```

### ShapeRange.Group

组合指定区域中的图形并将组合图形作为单个 **Shape** 对象返回。

**语法**

**express.Group ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

由于一组形状作为单个形状处理，所以创建和分解形状组将改变 **Shapes** 集合中的项目数，而且由于影响集合中的项目，还会改变部分项目的索引号。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加两个图形，并组合这两个新图形，接着设置该组的填充格式并对该组进行旋转，然后将该组置于绘图层的后面。*/
function test()
{
    let myDocument = Application.ActiveDocument.Shapes
    myDocument.AddShape(msoShapeCan, 50, 10, 100, 200).Name = "shpOne"
    myDocument.AddShape(msoShapeCube, 150, 250, 100, 200).Name = "shpTwo"
    let newmyDocument = myDocument.Range(Array("shpOne", "shpTwo")).Group()
    newmyDocument.Fill.PresetTextured(msoTextureBlueTissuePaper)
    newmyDocument.Rotation = 45
    newmyDocument.ZOrder(msoSendToBack)
}
```

### ShapeRange.IncrementLeft

将指定形状水平移动指定的磅数。

**语法**

**express.IncrementLeft ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                         |
|-------------|-----------|----------|------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定形状水平移动的距离，以磅为单位。为正值时将形状右移；为负值时将形状左移。 |

### ShapeRange.IncrementRotation

使指定的形状绕 Z 轴旋转指定的角度。

**语法**

**express.IncrementRotation ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                               |
|-------------|-----------|----------|------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定形状的水平旋转量，以度为单位。为正值时顺时针旋转形状，为负值时逆时针旋转形状。 |

**说明**

使用 **Rotation** 属性可以设置形状的绝对旋转角度。如果要将三维形状绕 X 轴或 Y 轴旋转，请使用 **ThreeDFormat** 的 **IncrementRotationX** 或 **IncrementRotationY** 方法。

### ShapeRange.IncrementTop

以指定磅数垂直移动指定形状。

**语法**

**express.IncrementTop ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                         |
|-------------|-----------|----------|------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定对象的垂直移动距离，以磅为单位。为正值时将形状下移；为负值时将形状上移。 |

### ShapeRange.Item

返回集合中的单个 **Shape** 对象。返回Shape值

**语法**

**express.Item ( *Index* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

### ShapeRange.PickUp

复制指定形状的格式。

**语法**

**express.PickUp ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

用 Apply 方法可将复制的格式应用于另一个形状。

**示例**

``` JavaScript
/*以下示例首先复制 myDocument 上第一个形状的格式，然后将复制的形状格式应用于第二个形状。*/
Application.ActiveDocument.Shapes.Item(1).PickUp()
Application.ActiveDocument.Shapes.Item(2).Apply()
```

### ShapeRange.ScaleHeight

按指定的比例缩放形状范围的高度。

**语法**

**express.ScaleHeight ( *Factor* , *RelativeToOriginalSize* , *Scale* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型     | 说明                                                                                                                                                        |
|--------------------------|-----------|--------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Factor*                 | 必选      | Single       | 指定形状调整后的高度与当前或原始高度的比例。例如，要将一个矩形放大百分之五十，请将此参数指定为 1.5。                                                        |
| *RelativeToOriginalSize* | 必选      | MsoTriState  | 如果该参数值为 True，则相对于原始大小缩放形状。如果该参数值为 False，则相对于当前大小缩放形状。仅当指定的形状为图片或 OLE 对象时，才能将此参数指定为 True。 |
| *Scale*                  | 可选      | MsoScaleFrom | 在缩放形状时，形状中位置不变的部分。                                                                                                                        |

**说明**

对于图片和 OLE 对象，您可以说明是相对于原始大小还是相对于当前大小缩放形状。对于图片和 OLE 对象以外的形状总是相对于当前高度缩放。

### ShapeRange.ScaleWidth

按指定比例调整形状的宽度。

**语法**

**express.ScaleWidth ( *Factor* , *RelativeToOriginalSize* , *Scale* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型     | 说明                                                                                                                                                        |
|--------------------------|-----------|--------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Factor*                 | 必选      | Single       | 指定形状调整后的宽度与当前或原始宽度的比例。例如，要将一个矩形放大百分之五十，请将此参数指定为 1.5。                                                        |
| *RelativeToOriginalSize* | 必选      | MsoTriState  | 如果该参数值为 True，则相对于原始大小缩放形状。如果该参数值为 False，则相对于当前大小缩放形状。仅当指定的形状为图片或 OLE 对象时，才能将此参数指定为 True。 |
| *Scale*                  | 可选      | MsoScaleFrom | 在缩放形状时，形状中位置不变的部分。                                                                                                                        |

**说明**

对于图片和 OLE 对象，您可以说明是相对于原始大小还是相对于当前大小缩放形状的范围。图片和 OLE 对象以外的形状总是相对于当前宽度缩放。

### ShapeRange.Select

选择指定的形状范围。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | 在添加形状时，如果该参数值为 True，则替换所选内容。如果该参数值为 False，则将新形状添加到所选内容。 |

### ShapeRange.SetShapesDefaultProperties

将文档中默认形状的格式应用于指定的形状范围。

**语法**

**express.SetShapesDefaultProperties ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

新形状将继承默认形状的许多属性。

### ShapeRange.Ungroup

取消指定形状范围中所有组合形状的组合，分解指定形状或形状范围中图片和 OLE 对象的组合，将取消组合后的形状以单个 **ShapeRange** 对象的形式返回。

返回 ShapeRange值

**语法**

**express.Ungroup ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

由于将组合形状作为单个对象来处理，对形状进行组合和取消组合将会改变 **Shapes** 集合中项目的数量，并且改变集合中取消组合的形状后面的项目的索引号。

### ShapeRange.ZOrder

将集合中指定的形状区域移动到其他形状的前面或后面（也就是说，更改形状区域在 Z 顺序中的位置）。

**语法**

**express.ZOrder ( *ZOrderCmd* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型     | 说明                                             |
|-------------|-----------|--------------|--------------------------------------------------|
| *ZOrderCmd* | 必选      | MsoZOrderCmd | 指定相对于其他形状将指定的形状区域移动到的位置。 |

**说明**

使用 ZOrderPosition 属性确定形状区域在 Z 顺序中的当前位置。

## 成员属性

### ShapeRange.Adjustments

返回一个 **Adjustments** 对象, 该对象包含所有对指定 **ShapeRange** 对象（代表自选图形或艺术字）进行调整操作的调整值。只读。

**语法**

**express.Adjustments**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.AlternativeText

返回或设置与网页的图形相关联的 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">可选文字</span> 。 **String** 类型，可读写。

**语法**

**express.AlternativeText**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：为活动窗口中选定的图形设定可选文字。选定图形是一幅野鸭的图*/
Application.ActiveWindow.Selection.ShapeRange.AlternativeText = "This is a mallard duck."
```

### ShapeRange.Anchor

返回一个 **Range** 对象，该对象代表指定图形区域的锁定范围。只读。

**语法**

**express.Anchor**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

如果对包含多个图形的 **ShapeRange** 对象使用该属性，将导致出错。

所有 **Shape** 对象都锁定于某一个文字区域，但可将其置于锁定标记所在页的任意位置。如果创建图形时指定了锁定范围，则锁定标记位于包含该锁定范围的第一段落的段首。如果没有指定锁定范围，则将自动选择锁定范围，图形参照页面的左边框和上边框进行定位。

图形总是与锁定标记处在同一页上。如果图形的 **LockAnchor** 属性为 **True** ，则不能在页面上拖动锁定标记。

### ShapeRange.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### ShapeRange.AutoShapeType

返回或设置指定的 **ShapeRange** 对象的图形类型，该对象不是代表线条或任意多边形，而是代表自选图形。 **MsoAutoShapeType** 类型，可读写。

**语法**

**express.AutoShapeType**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

改变一个形状的类型时，该形状保留其大小、颜色和其他属性。

### ShapeRange.BackgroundStyle

设置或返回指定形状范围中形状的背景样式。可读/写 MsoBackgroundStyleIndex 。

**语法**

**express.BackgroundStyle**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Callout

返回 **CalloutFormat** 对象，该对象包含指定图形的标注格式属性。只读。

**语法**

**express.Callout**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

该属性应用于代表标注的 **ShapeRange** 对象。

**示例**

``` JavaScript
/*该示例在 myDocument 中添加一个椭圆和一个指向椭圆的标注。标注文本不带有边框，但有一条强调线将文本和标注线隔开。*/
function test()
{
    let myShapes = Application.ActiveDocument.Shapes
    myShapes.AddShape(msoShapeOval, 180, 200, 280, 130)

    let newCallout = myShapes.AddCallout(msoCalloutTwo, 420, 170, 170, 40)
    newCallout.TextFrame.TextRange.Text = "My oval"
    newCallout.Callout.Accent = true
    newCallout.Callout.Border = false
}
```

### ShapeRange.CanvasItems

返回一个 **CanvasShapes** 对象，该对象代表绘图画布上图形的集合。

**语法**

**express.CanvasItems**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Child

如果图形是子图形或位于图形区域的所有图形都是同一父图形的子图形，则该属性值为 **True** 。 **MsoTriState** 类型，只读。

**语法**

**express.Child**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Count

返回一个 **Long** 类型的值，该值代表集合中图形的数量。只读。

**语法**

**express.Count**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### ShapeRange.Fill

返回一个 **FillFormat** 对象，该对象包含指定图形的填充格式属性。只读。

**语法**

**express.Fill**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Glow

返回一个 **GlowFormat** 对象，该对象代表形状区域的发光格式。只读。

**语法**

**express.Glow**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.GroupItems

返回一个 **GroupShapes** 对象，该对象代表指定图形组中的单个图形。只读。

**语法**

**express.GroupItems**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

该属性应用于代表组合图形的 **ShapeRange** 对象。使用 **GroupShapes** 对象的 **Item** 方法可从图形组中返回单个图形。

### ShapeRange.Height

返回或设置指定图形区域的高度。 **Single** 类型，可读写。

**语法**

**express.Height**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.HeightRelative

返回或设置一个 **Single** 类型的值，该值代表将形状区域大小调整到的目标形状的百分比。可读写。

**语法**

**express.HeightRelative**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

将此属性与 RelativeVerticalSize 属性一起使用。当设置为 **wdShapeSizeRelativeNone** (-999999)（参见 **WdShapeSizeRelative** 枚举）时，应忽略此属性，因为形状不能使用百分比大小。高度是由 Height 属性单独确定的。

### ShapeRange.HorizontalFlip

表示该形状范围已进行水平翻转。只读 **MsoTriState** 类型。

**语法**

**express.HorizontalFlip**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Hyperlink

返回一个 **Hyperlink** 对象，该对象代表与指定 **ShapeRange** 对象相关联的超链接。只读。

**语法**

**express.Hyperlink**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*如果不存在与指定形状区域相关联的超链接，则会发生错误。在此情况下，请将 Add 方法用于 Hyperlinks 集合来为指定形状区域添加超链接。以下示例说明了具体的操作方法。*/
Application.ActiveDocument.Hyperlinks.Add(Selection.ShapeRange.Item(1), "http://www.microsoft.com")
```

### ShapeRange.ID

返回形状范围的标识类型。只读 **Long** 类型。

**语法**

**express.ID**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.LayoutInCell

返回一个 **Long** 类型的值，该值代表表格中的形状是显示在表格内部还是表格外部。

**语法**

**express.LayoutInCell**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

**LayoutInCell** 属性对应于图片格式的 **“高级版式”** 对话框中的 **“表格单元格中的版式”** 选项。如果为 **True** ，则表示指定的图片显示在表格内部。如果为 **False** ，则表示指定的图片显示在表格外部。

| 注释                                                                                                                                             |
|--------------------------------------------------------------------------------------------------------------------------------------------------|
| 仅当 **WrapFormat** 对象的 **Type** 属性设置为除 **wdWrapTypeInline** 或 **wdWrapTypeNone** 之外的值时，对 **LayoutInCell** 属性的设置才会生效。 |

### ShapeRange.Left

返回或设置一个 **Single** 类型的值，该值代表指定形状范围的水平位置，以磅为单位。也可以是任何有效的 WdShapePosition 常量。可读写。

**语法**

**express.Left**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

形状的位置是根据形状边框的左上角到形状锁定标记的相对距离进行计算的。 **RelativeHorizontalPosition** 属性控制锁定标记是沿字符、文本栏、页边距还是页面边缘放置。

对于包含多个形状的 **ShapeRange** 对象， **Left** 属性设置每个形状的水平位置。

**示例**

``` JavaScript
/*以下示例将活动文档中第一、第二个形状的水平位置设置为距离文本栏的左边缘 1 英寸。*/

function test()
{
  let shapes = Application.ActiveDocument.Shapes.Range([1,2])
  shapes.RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
  shapes.Left = InchesToPoints(1)
}
```

### ShapeRange.LeftRelative

返回或设置一个 **Single** 类型的值，该值代表形状区域左侧的相对位置。可读写。

**语法**

**express.LeftRelative**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

将此属性与 RelativeHorizontalPosition 属性一起使用。当设置为 **[wdShapePositionRelativeNone](wdShapePositionRelativeNone)** [](wdShapePositionRelativeNone) (-999999)（参见 WdShapePositionRelative 枚举）时，应忽略此属性，因为形状不能使用百分比放置。水平位置是由 Left 属性单独确定的。

### ShapeRange.Line

返回一个 **LineFormat** 对象，该对象包含指定形状范围的线条格式属性。只读。

**语法**

**express.Line**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

对于线条来说， **LineFormat** 对象代表线条本身；而对于带有边框的形状范围来说， **LineFormat** 对象代表边框。

### ShapeRange.LockAnchor

如果指定 **ShapeRange** 对象的锁定标记锁定到锁定范围，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.LockAnchor**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

如果形状范围有一个锁定的标记，则不能通过拖动来移动形状的锁定标记。锁定标记不会随形状移动而移动。

虽然 **ShapeRange** 对象已锁定到某一文字范围，但是您可以将该对象放在页面上的任何位置。如果形状范围锁定到包含锁定范围的第一段的开头，则形状总是与其锁定标记位于同一页上。

### ShapeRange.LockAspectRatio

如果在调整指定形状的大小时保留其最初比例，则该属性值为 **MsoTrue** ；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 **MsoFalse** 。可读写 **MsoTriState** 类型。

**语法**

**express.LockAspectRatio**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Name

返回或设置指定对象的名称。 **String** 类型，可读写。

**语法**

**express.Name**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Nodes

返回一个 **ShapeNodes** 集合，该集合代表指定形状的几何描述。

**语法**

**express.Nodes**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Parent

返回一个 **Object** 类型值，该值代表指定 **ShapeRange** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.ParentGroup

返回一个 **Shape** 对象，该对象代表形状范围的通用父形状。

**语法**

**express.ParentGroup**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.PictureFormat

返回一个 **PictureFormat** 对象，该对象包含指定形状范围的图片格式属性。只读。

**语法**

**express.PictureFormat**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

适用于代表图片或 OLE 对象的 **ShapeRange** 对象。

### ShapeRange.Reflection

返回一个 **ReflectionFormat** 对象，该对象代表形状区域的反射格式。只读。

**语法**

**express.Reflection**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.RelativeHorizontalPosition

指定形状范围的相对水平位置。可读写 WdRelativeHorizontalPosition 类型。

**语法**

**express.RelativeHorizontalPosition**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例重新放置选定形状对象。*/


function test()
{
  let shaperange = Application.Selection.ShapeRange
  shaperange.Left = InchesToPoints(0.6)
  shaperange.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
  shaperange.Top = InchesToPoints(1)
  shaperange.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
}
```

### ShapeRange.RelativeHorizontalSize

返回或设置一个 WdRelativeHorizontalSize 常量，该常量代表形状区域相对的对象。可读写。

**语法**

**express.RelativeHorizontalSize**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

此属性可与 [WidthRelative](WidthRelative) 属性一起使用。

### ShapeRange.RelativeVerticalPosition

指定形状范围的相对垂直位置。可读写 WdRelativeHorizontalPosition 类型。

**语法**

**express.RelativeVerticalPosition**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例重新放置选定形状对象。*/

function test()
{
let shaperange = Application.Selection.ShapeRange
shaperange.Left = InchesToPoints(0.6)
shaperange.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
shaperange.Top = InchesToPoints(1)
shaperange.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
}
```

### ShapeRange.RelativeVerticalSize

返回或设置一个 **WdRelativeVerticalSize** 常量，该常量代表形状区域相对的对象。可读写。

**语法**

**express.RelativeVerticalSize**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

此属性可与 HeightRelative 属性一起使用。

### ShapeRange.Rotation

返回或设置指定形状绕 Z 轴旋转的度数。可读写 **Single** 类型。

**语法**

**express.Rotation**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

正值表示顺时针旋转；负值表示逆时针旋转。要设置三维形状绕 X 轴或 Y 轴的旋转，请使用 **ThreeDFormat** 对象的 **RotationX** 属性或 **RotationY** 属性。

### ShapeRange.Shadow

返回一个 **ShadowFormat** 对象，该对象代表指定形状的阴影格式。

**语法**

**express.Shadow**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.ShapeStyle

设置或返回指定形状范围中形状的形状样式。可读/写 MsoShapeStyleIndex。

**语法**

**express.ShapeStyle**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.SoftEdge

返回一个 **SoftEdgeFormat** 对象，该对象代表形状区域的软边缘格式。只读。

**语法**

**express.SoftEdge**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.TextEffect

返回一个 **TextEffectFormat** 对象，该对象包含指定形状的文本效果格式属性。只读。

**语法**

**express.TextEffect**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

适用于代表艺术字的 **ShapeRange** 对象。

### ShapeRange.TextFrame

返回一个 **TextFrame** 对象，该对象包含指定形状范围的文字。

**语法**

**express.TextFrame**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.TextFrame2

返回一个 **TextFrame2** 对象，包含指定形状区域的文本。只读。

**语法**

**express.TextFrame2**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.ThreeD

返回一个 **ThreeDFormat** 对象，该对象包含指定形状范围的三维格式属性。只读。

**语法**

**express.ThreeD**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Title

返回或设置 **String** 类型值，该值包含指定形状范围中形状的标题。可读/写。

**语法**

**express.Title**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

使用 **Title** 属性可为形状提供可选文字标题。该属性将标题文本添加到 WPS 2015 中 **“设置形状格式”** 对话框的 **“可选文字”** 窗格上的 **“标题”** 文本框中。

| 注释                                                                                                                       |
|----------------------------------------------------------------------------------------------------------------------------|
| Web 浏览器在加载表格的过程中或表格丢失时显示可选文字。Web 搜索引擎利用可选文字帮助查找网页。可选文字也可用来帮助残障人士。 |

**示例**

``` JavaScript
/*下面的代码示例向活动文档中第一个和第三个形状中添加可选文字标题。*/
Application.ActiveDocument.Shapes.Range([1, 3]).Title = "Part of a shape array."
```

### ShapeRange.Top

返回或设置指定形状或形状范围的垂直位置（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.Top**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

形状的位置是根据形状边框的左上角到形状锁定标记的相对距离进行计算的。 **RelativeVerticalPosition** 属性控制形状的锁定标记是沿行、段落、页边距还是页面边缘放置。

对于包含多个形状的 **ShapeRange** 对象， **Top** 属性设置每个形状的垂直位置。

**示例**

``` JavaScript
/*以下示例将活动文档中第一个和第二个形状的垂直位置设置为距页面顶部 1 英寸。*/

function test()
{
  let shapes = Application.ActiveDocument.Shapes.Range(Array(1, 2))
  shapes.RelativeVerticalPosition = wdRelativeVerticalPositionPage
  shapes.Top = InchesToPoints(1)
}
```

### ShapeRange.TopRelative

返回或设置一个 **Single** 类型的值，该值代表形状区域顶部的相对位置。可读写。

**语法**

**express.TopRelative**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

将此属性与 RelativeVerticalPosition 属性一起使用。当设置为 wdShapePositionRelativeNone (-999999)（参见 WdShapePositionRelative 枚举）时，应忽略此属性，因为形状不能使用百分比放置。垂直位置是由 Top 属性单独确定的。

### ShapeRange.Type

返回形状类型。只读 **MsoShapeType** 类型。

**语法**

**express.Type**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.VerticalFlip

如果指定形状围绕垂直轴进行翻转，则该属性值为 **True** 。 **MsoTriState** 类型，只读。

**语法**

**express.VerticalFlip**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例将 myDocument 中所有进行过水平翻转或垂直翻转的形状还原至初始状态。*/

function test()
{
  let s
  for(let i =1; i <= Application.ActiveDocument.Range().ShapeRange.Count; i++) {
      s = Application.ActiveDocument.Range().ShapeRange.Item(i)
      if(s.HorizontalFlip) {
          s.Flip(msoFlipHorizontal)
      }
      if(s.VerticalFlip) { 
          s.Flip(msoFlipVertical)
      }
  }
}
```

### ShapeRange.Vertices

该属性以一系列 <span class="glossary" style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;">坐标对</span> 的形式返回指定任意多边形图形顶点（和贝赛尔曲线的控点）的坐标。可将该属性返回的数组用作 **AddCurve** 或 **AddPolyLine** 方法的参数。只读 **Variant** 类型。

**语法**

**express.Vertices**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

下表显示 **Vertices** 属性如何将数组 *vertArray()* 的值与三角形的顶点坐标相关联。

``` JavaScript
vertArray(1, 1)vertArray(1, 2)vertArray(2, 1)vertArray(2, 2)vertArray(3, 1)vertArray(3, 2)
```

### ShapeRange.Visible

如果指定对象或应用于该对象的格式是可见的，则该属性值为 **True** 。 **MsoTriState** 类型，可读写。

**语法**

**express.Visible**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Width

返回或设置范围内形状的宽度（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Width**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.WidthRelative

返回或设置一个 **Single** 类型的值，该值代表形状区域的相对宽度。可读写。

**语法**

**express.WidthRelative**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

将此属性与 RelativeHorizontalSize 属性一起使用。当设置为 wdShapeSizeRelativeNone (-999999)（参见 WdShapeSizeRelative 枚举）时，应忽略此属性，因为形状不能使用百分比大小。宽度是由 Width 属性单独确定的。

### ShapeRange.WrapFormat

返回一个 **WrapFormat** 对象，该对象包含在指定的形状范围四周文字环绕的属性。只读。

**语法**

**express.WrapFormat**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.ZOrderPosition

返回一个 **Long** 类型的值，该值代表指定的形状在 Z 顺序中的位置。只读。

**语法**

**express.ZOrderPosition**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

` Shapes(1) ` 返回 Z 顺序中的最后一个形状，而 ` Shapes(Shapes.Count) ` 返回 Z 顺序中的第一个形状。该属性为只读。要设置形状在 Z 顺序中的位置，请使用 **ZOrder** 方法。

形状在 Z 顺序中的位置与 Shapes 集合中形状的索引号相对应。例如，如果在 myDocument 上有四个形状，则表达式 ` myDocument.Shapes(1) ` 返回 Z 顺序中的最后一个形状，而表达式 ` myDocument.Shapes(4) ` 返回 Z 顺序中的第一个形状。

无论何时将新形状添加到集合中，默认情况下该形状都将添加到 Z 顺序的最前端。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/ShapeRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Shapes 对象

## 简介

Shape 对象的集合，这些对象代表某个文档中的所有形状，或某个文档的所有页眉和页脚中的所有形状。每个 Shape 对象代表绘制层的一个对象，如自选图形、任意多边形、OLE 对象或图片。

## 方法

| 序号 | 名称                                   | 说明                                                                                                        |
|:----:|:---------------------------------------|:------------------------------------------------------------------------------------------------------------|
|  1   | [AddCanvas](#Shapes.AddCanvas)         | 在文档中添加 绘图画布 。返回代表该画布的 Shape 对象并将其添加到 Shapes 集合。                               |
|  2   | [AddChart](#Shapes.AddChart)           | 将指定类型的图表作为形状插入到活动文档中，并且在ET 中打开包含 WPS 用于创建该图表的默认数据的工作表。        |
|  3   | [AddCurve](#Shapes.AddCurve)           | 返回一个 Shape 对象，该对象代表绘图画布上的贝赛尔曲线。                                                     |
|  4   | [AddLabel](#Shapes.AddLabel)           | 在绘图画布上添加一个文本标签。                                                                              |
|  5   | [AddLine](#Shapes.AddLine)             | 在绘图画布上添加一条直线。                                                                                  |
|  6   | [AddOLEControl](#Shapes.AddOLEControl) | 创建一个 ActiveX 控件（以前称为 OLE 控件）。返回代表新 ActiveX 控件的 InlineShape 对象。                    |
|  7   | [AddOLEObject](#Shapes.AddOLEObject)   | 创建一个 OLE 对象。返回代表新 OLE 对象的 InlineShape 对象。                                                 |
|  8   | [AddPicture](#Shapes.AddPicture)       | 在绘图画布上添加一幅图片。返回一个 Shape 对象，该对象代表图片并将其添加至 CanvasShapes 集合。               |
|  9   | [AddPolyline](#Shapes.AddPolyline)     | 在绘图画布上添加一个开放或封闭的多边形。返回一个代表该多边形的 Shape 对象，并将其添加到 CanvasShapes 集合。 |
|  10  | [AddShape](#Shapes.AddShape)           | 向某文档中添加一个自选图形。返回代表该自选图形的 Shape 对象并将其添加到 Shapes 集合。                       |
|  11  | [AddSmartArt](#Shapes.AddSmartArt)     | 将指定 SmartArt 图形插入活动文档。                                                                          |
|  12  | [AddTextbox](#Shapes.AddTextbox)       | 在绘图画布上添加一个文本框。                                                                                |
|  13  | [AddTextEffect](#Shapes.AddTextEffect) | 在绘图画布上添加一个“艺术字”图形。返回一个 Shape 对象，该对象代表“艺术字”，并将其添加至 CanvasShapes 集合。 |
|  14  | [BuildFreeform](#Shapes.BuildFreeform) | 建立任意多边形对象。                                                                                        |
|  15  | [Item](#Shapes.Item)                   | 返回集合中的单个 Shape 对象。                                                                               |
|  16  | [Range](#Shapes.Range)                 | 返回代表范围内的形状的 ShapeRange 对象。                                                                    |
|  17  | [SelectAll](#Shapes.SelectAll)         | 选择形状集合中的所有形状。                                                                                  |

## 属性

| 序号 | 名称                               | 说明                                                                         |
|:----:|:-----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [AddCallout](#Shapes.AddCallout)   | 在绘图画布上添加无框线的标注。                                               |
|  2   | [Application](#Shapes.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  3   | [Count](#Shapes.Count)             | 返回一个 Long 类型的值，该值代表集合中图形的数量。只读。                     |
|  4   | [Creator](#Shapes.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  5   | [Parent](#Shapes.Parent)           | 返回一个 Object 类型值，该值代表指定 Shapes 对象的父对象。                   |

## 成员方法

### Shapes.AddCanvas

在文档中添加 绘图画布 。返回代表该画布的 **Shape** 对象并将其添加到 **Shapes** 集合。

**语法**

**express.AddCanvas ( *Left* , *Top* , *Width* , *Height* , *Anchor* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                                                                                                         |
|----------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Left*   | 必选      | Single   | 画布左侧边缘相对于锁定标记的位置，以磅为单位。                                                                                                                               |
| *Top*    | 必选      | Single   | 画布上部边缘相对于锁定标记的位置，以磅为单位。                                                                                                                               |
| *Width*  | 必选      | Single   | 画布的宽度，以磅为单位。                                                                                                                                                     |
| *Height* | 必选      | Single   | 画布的高度，以磅为单位。                                                                                                                                                     |
| *Anchor* | 可选      | Variant  | 代表画布所绑定的文本的 Range 对象。如果指定 Anchor，则锁定标记将出现在锁定区域第一段的开头。如果省略该参数，将自动选定锁定区域，而画布将相对于页面的上边缘和左边缘进行定位。 |

**返回值**

Shape

**示例**

``` JavaScript
/*以下示例在新文档中添加画布，然后在画布上添加两个图形，并设置填充和线条属性。*/
function AddInlineCanvas(){
    //Add a drawing canvas to the new document
    let shpCanvas = Documents.Add().Shapes.AddCanvas(150, 150, 70, 70)
    shpCanvas.WrapFormat.Type = wdWrapInline
    //Add shapes to drawing canvas
    let canvasitems = shpCanvas.CanvasItems
    canvasitems.AddShape(msoShapeHeart, 10, 10, 50, 60)
    canvasitems.AddLine(0, 0, 70, 70)
    let shpcanvas = shpCanvas
    shpcanvas.CanvasItems.Item(1).Fill.ForeColor.RGB = 255, 0, 0
    shpcanvas.CanvasItems.Item(2).Line.EndArrowheadStyle = msoArrowheadTriangle
}
```

### Shapes.AddChart

将指定类型的图表作为形状插入到活动文档中，并且在ET 中打开包含 WPS 用于创建该图表的默认数据的工作表。

**语法**

**express.AddChart ( *Type* , *Left* , *Top* , *Width* , *Height* , *Anchor* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型    | 说明                                                                                                                                                                     |
|----------|-----------|-------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Type*   | 可选      | XlChartType | 指定要创建的图表的类型。                                                                                                                                                 |
| *Left*   | 可选      | Variant     | 图表左边缘相对于定位标记的位置（以磅为单位）。                                                                                                                           |
| *Top*    | 可选      | Variant     | 图表上边缘相对于定位标记的位置（以磅为单位）。                                                                                                                           |
| *Width*  | 可选      | Variant     | 图表的宽度（以磅为单位）。                                                                                                                                               |
| *Height* | 可选      | Variant     | 图表的高度（以磅为单位）。                                                                                                                                               |
| *Anchor* | 可选      | Variant     | 代表图表所绑定文本的 Range 对象。如果指定 Anchor，则定位标记将出现在定位区域中第一段的开头。如果省略该参数，则自动选择定位区域，并且相对于页面的上边缘和左边缘放置图表。 |

**返回值**

Shape

**示例**

``` JavaScript
/*创建一个新的三维柱形图，然后将其添加到活动文档中。*/
ActiveDocument.Shapes.AddChart(xl3DColumn)
```

### Shapes.AddCurve

返回一个 **Shape** 对象，该对象代表绘图画布上的贝赛尔曲线。

**语法**

**express.AddCurve ( *SafeArrayOfPoints* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                                                                   |
|---------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *SafeArrayOfPoints* | 必选      | Variant  | 指定该曲线的顶点和控制点的坐标对?（坐标对：一对值，表示两维数组中存储的点的 x 和 y 坐标，该数组中包含许多点的坐标。）数组。指定的第一个点是起始顶点，随后指定的两个点是第一段贝塞尔曲线的控制点。如果该曲线还有其他段，则每段都应该指定一个顶点和两个控制点。指定的最后一个点是该曲线的终止顶点。请注意，须始终指定 3n + 1 个点，此处 n 为曲线的段数。 |

**返回值**

Shape

**说明**

------------------------------------------------------------------------

**示例**

``` JavaScript
/*本示例在新绘图画布上添加一条贝赛尔曲线。*/
function CanvasBezier(){
    let sngArray = [[0,0],[0,0],[0,0],[0,0],[0,0],[0,0],[0,0]]
    let docNew = Documents.Add()
    //Create a new drawing canvas
    let shpCanvas = docNew.Shapes.AddCanvas(100, 100, 300, 50)
    sngArray[0][0] = 0
    sngArray[0][1] = 0
    sngArray[1][0] = 50
    sngArray[1][1] = 50
    sngArray[2][0] = 100
    sngArray[2][1] = 0
    sngArray[3][0] = 150
    sngArray[3][1] = 50
    sngArray[4][0] = 200
    sngArray[4][1] = 0
    sngArray[5][0] = 250
    sngArray[5][1] = 50
    sngArray[6][0] = 300
    sngArray[6][1] = 0
    //Add Bezier curve to drawing canvas
    shpCanvas.CanvasItems.AddCurve(sngArray)
}
```

### Shapes.AddLabel

在绘图画布上添加一个文本标签。

**语法**

**express.AddLabel ( *Orientation* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型           | 说明                                                 |
|---------------|-----------|--------------------|------------------------------------------------------|
| *Orientation* | 必选      | MsoTextOrientation | 文本的方向。                                         |
| *Left*        | 必选      | Single             | 标签左边缘相对于绘图画布左边缘的位置（以磅为单位）。 |
| *Top*         | 必选      | Single             | 标签上边缘相对于绘图画布上边缘的位置（以磅为单位）。 |
| *Width*       | 必选      | Single             | 标签的宽度（以磅为单位）。                           |
| *Height*      | 必选      | Single             | 标签的高度（以磅为单位）。                           |

**返回值**

Shapes

**示例**

``` JavaScript
/*本示例在活动文档的新绘图画布上添加一个蓝色的文本标签，该标签带有文字“Hello World”。*/
function NewCanvasTextLabel(){
    //Add a drawing canvas to the active document
    let shpCanvas = ActiveDocument.Shapes.AddCanvas(100, 75, 150, 200)
    
    //Add a label to the drawing canvas
    let shpLabel = shpCanvas.CanvasItems.AddLabel(msoTextOrientationHorizontal, 15, 15, 100, 100)
    
    //Fill the label textbox with a color,
    //add text to the label and format it
    let fill = shpLabel.Fill
        fill.BackColor.RGB = 0, 0, 192
        //Make the fill visible
        fill.Visible = msoTrue
    let newtextrange = shpLabel.TextFrame.TextRange
        newtexttange.Text = "Hello World."
        newtexttange.Bold = true
        newtexttange.Font.Name = "Tahoma"
}
```

### Shapes.AddLine

在绘图画布上添加一条直线。

**语法**

**express.AddLine ( *BeginX* , *BeginY* , *EndX* , *EndY* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                             |
|----------|-----------|----------|--------------------------------------------------|
| *BeginX* | 必选      | Single   | 直线起点相对于绘图画布的水平位置（以磅为单位）。 |
| *BeginY* | 必选      | Single   | 直线起点相对于绘图画布的垂直位置（以磅为单位）。 |
| *EndX*   | 必选      | Single   | 直线终点相对于绘图画布的水平位置（以磅为单位）。 |
| *EndY*   | 必选      | Single   | 直线终点相对于绘图画布的垂直位置（以磅为单位）。 |

**返回值**

Shape

**说明**

若要创建带箭头的直线，请使用 **Line** 属性来设置线条格式。

**示例**

``` JavaScript
/*本示例在新绘图画布上添加一条紫色带箭头的直线。*/
function NewCanvasLine(){
    //Add new drawing canvas to the active document
    let shpCanvas = ActiveDocument.Shapes.AddCanvas(100, 75, 150, 200)

    //Add a line to the drawing canvas
    let shpLine = shpCanvas.CanvasItems.AddLine(25, 25, 150, 150)

    //Add an arrow to the line and sets the color to purple
    let line = shpLine.Line
        line.BeginArrowheadStyle = msoArrowheadDiamond
        line.BeginArrowheadWidth = msoArrowheadWide
        line.ForeColor.RGB = 150, 0, 255
}
```

### Shapes.AddOLEControl

创建一个 ActiveX 控件（以前称为 OLE 控件）。返回代表新 ActiveX 控件的 **InlineShape** 对象。

**语法**

**express.AddOLEControl ( *ClassType* , *Range* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                   |
|-------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------|
| *ClassType* | 可选      | Variant  | 用于要创建的 ActiveX 控件的编程标识符。                                                                                                |
| *Range*     | 可选      | Variant  | 一个范围，ActiveX 控件将放置在该范围的文本中。如果该范围未折叠，则 ActiveX 控件将替换该范围。如果省略该参数，则自动放置 ActiveX 控件。 |

**返回值**

InlineShape

**说明**

在 WPS 中，将 ActiveX 控件表示为 **Shape** 对象或 **InlineShape** 对象。要修改 ActiveX 控件的属性，可以对指定的形状或内嵌形状使用 **OLEFormat** 对象的 **Object** 属性。

有关可用的 ActiveX 控件类类型的信息，请参阅 OLE 编程标识符。

### Shapes.AddOLEObject

创建一个 OLE 对象。返回代表新 OLE 对象的 **InlineShape** 对象。

**语法**

**express.AddOLEObject ( *ClassType* , *FileName* , *LinkToFile* , *DisplayAsIcon* , *IconFileName* , *IconIndex* , *IconLabel* , *Range* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                       |
|-----------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ClassType*     | 可选      | Variant  | 用于激活指定 OLE 对象的应用程序的名称。                                                                                                                                                                                                                                    |
| *FileName*      | 可选      | Variant  | 创建对象的文件。如果省略该参数，则使用当前文件夹。必须指定对象的 ClassType 或 FileName 参数，但不能同时指定两者。                                                                                                                                                          |
| *LinkToFile*    | 可选      | Variant  | 如果该参数值为 True，则将 OLE 对象链接到创建该对象的文件。如果该参数值为 False，则 OLE 对象成为该文件的独立副本。如果该参数值为 ClassType 指定了值，则 LinkToFile 参数必须为 False。默认值为 False。                                                                       |
| *DisplayAsIcon* | 可选      | Variant  | 如果该参数值为 True，则将 OLE 对象显示为图标。默认值为 False。                                                                                                                                                                                                             |
| *IconFileName*  | 可选      | Variant  | 包含将要显示的图标的文件。                                                                                                                                                                                                                                                 |
| *IconIndex*     | 可选      | Variant  | IconFileName 内的图标的索引号。当选中“显示为图标”复选框时，指定文件中图标的顺序对应于“更改图标”对话框中图标显示的顺序。文件中的第一个图标的索引号为 0（零）。如果 IconFileName 中不存在指定索引号的图标，则使用索引号为 1 的图标（文件中的第二个图标）。默认值为 0（零）。 |
| *IconLabel*     | 可选      | Variant  | 显示在图标下面的标签（题注）。                                                                                                                                                                                                                                             |
| *Range*         | 可选      | Variant  | 一个范围，OLE 对象将放置在该范围的文本中。除非范围是折叠的，否则 OLE 对象将替换该范围。如果省略此参数，则自动放置对象。                                                                                                                                                    |

**示例**

``` JavaScript
/*以下示例在活动文档中添加一个新的浮动位图映射。该位图链接到另一个文件。*/
ActiveDocument.Shapes.AddOLEObject(null,"c:\\my documents\\MyDrawing.bmp", true)
```

### Shapes.AddPicture

在绘图画布上添加一幅图片。返回一个 **Shape** 对象，该对象代表图片并将其添加至 **CanvasShapes** 集合。

**语法**

**express.AddPicture ( *FileName* , *LinkToFile* , *SaveWithDocument* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                                                                                                                    |
|--------------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------|
| *FileName*         | 必选      | String   | 图片的路径和文件名。                                                                                                    |
| *LinkToFile*       | 可选      | Variant  | 如果该参数值为 True ，则将图片链接到创建它的文件。如果该参数值为 False ，则将图片作为该文件的独立副本。默认值为 False。 |
| *SaveWithDocument* | 可选      | Variant  | 如果该参数值为 True ，则将链接的图片与文档一起保存。默认值为 False 。                                                   |
| *Left*             | 可选      | Variant  | 新图片的左边缘相对于绘图画布的位置，以磅为单位。                                                                        |
| *Top*              | 可选      | Variant  | 新图片的上边缘相对于绘图画布的位置，以磅为单位。                                                                        |
| *Width*            | 可选      | Variant  | 图片的宽度，以磅为单位。                                                                                                |
| *Height*           | 可选      | Variant  | 图片的高度，以磅为单位。                                                                                                |

**示例**

``` JavaScript
/*以下示例在活动文档中新创建的绘图画布上添加一幅图片*/
function NewCanvasPicture(){
    //Add a drawing canvas to the active document
    let shpCanvas = Application.ActiveDocument.Shapes.AddCanvas(100, 75, 200, 300)

    //Add a graphic to the drawing canvas
    shpCanvas.CanvasItems.AddPicture("C:\\Program Files\\Microsoft Office\\" + "Office\\Bitmaps\\Styles\\stone.bmp", false, true)
}
```

### Shapes.AddPolyline

在绘图画布上添加一个开放或封闭的多边形。返回一个代表该多边形的 **Shape** 对象，并将其添加到 **CanvasShapes** 集合。

**语法**

**express.AddPolyline ( *SafeArrayOfPoints* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明                                                                                                                      |
|---------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------|
| *SafeArrayOfPoints* | 必选      | Variant  | 指定连续线段图形的顶点的坐标对?（坐标对：一对值，表示两维数组中存储的点的 x 和 y 坐标，该数组中包含许多点的坐标。）数组。 |

**示例**

``` JavaScript
/*本示例在新绘图画布上创建一个 V 形开放多边形。*/
function NewCanvasPolyline(){
    let sngArray = [[0,0],[0,0],[0,0]]
    //Creates a new document and adds a drawing canvas
    let docNew = Documents.Add()
    let shpCanvas = docNew.Shapes.AddCanvas(100, 75, 200, 300)

    //Sets the coordinates of the array
    sngArray[0][0] = 100
    sngArray[0][1] = 75
    sngArray[1][0] = 150
    sngArray[1][1] = 100
    sngArray[2][0] = 100
    sngArray[2][1] = 125
    
    //Adds a V-shaped open polyline to the drawing canvas
    shpCanvas.CanvasItems.AddPolyline(sngArray)
}
```

### Shapes.AddShape

向某文档中添加一个自选图形。返回代表该自选图形的 **Shape** 对象并将其添加到 **Shapes** 集合。

**语法**

**express.AddShape ( *Type* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                 |
|----------|-----------|----------|------------------------------------------------------|
| *Type*   | 必选      | Long     | 要返回的图形类型。可以是任何 MsoAutoShapeType 常量。 |
| *Left*   | 必选      | Single   | 自选图形左边缘的位置（以磅为单位）。                 |
| *Top*    | 必选      | Single   | 自选图形上边缘的位置（以磅为单位）。                 |
| *Width*  | 必选      | Single   | 自选图形的宽度（以磅为单位）。                       |
| *Height* | 必选      | Single   | 自选图形的高度（以磅为单位）。                       |

**说明**

要改变所添加的自选图形类型，可设置其 **AutoShapeType** 属性。

**示例**

``` JavaScript
/*向活动文档中添加一个矩形*/
Application.ActiveDocument.Shapes.AddShape(1,100, 100, 100, 100)
```

### Shapes.AddSmartArt

将指定 SmartArt 图形插入活动文档。

**语法**

**express.AddSmartArt ( *Layout* , *Left* , *Top* , *Width* , *Height* , *Anchor* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型           | 说明                                                                                                                                                                                         |
|----------|-----------|--------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Layout* | 必选      | \[SMARTARTLAYOUT\] | 一个指定 SmartArt 图形的布局的 SmartArtLayout 对象。                                                                                                                                         |
| *Left*   | 可选      | Variant            | 从幻灯片左边缘到 SmartArt 图形左边缘之间的距离（以磅为单位）。                                                                                                                               |
| *Top*    | 可选      | Variant            | 从幻灯片上边缘到 SmartArt 图形上边缘之间的距离（以磅为单位）。                                                                                                                               |
| *Width*  | 可选      | Variant            | SmartArt 图形的宽度。                                                                                                                                                                        |
| *Height* | 可选      | Variant            | SmartArt 图形的高度。                                                                                                                                                                        |
| *Anchor* | 可选      | Variant            | 一个代表 SmartArt 所绑定的文本的 Range 对象。如果指定 Anchor，则定位标记将出现在定位区域第一段的开头。如果省略该参数，将自动选定定位区域，并且相对于页面的上边缘和左边缘放置 SmartArt 图形。 |

**返回值**

Shape

### Shapes.AddTextbox

在绘图画布上添加一个文本框。

**语法**

**express.AddTextbox ( *Orientation* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型           | 说明                                                                             |
|---------------|-----------|--------------------|----------------------------------------------------------------------------------|
| *Orientation* | 必选      | MsoTextOrientation | 文本的方向。由于选择或安装的语言支持（例如美国英语）不同，有些常量可能无法使用。 |
| *Left*        | 必选      | Single             | 文本框左边缘的位置（以磅为单位）。                                               |
| *Top*         | 必选      | Single             | 文本框上边缘的位置（以磅为单位）。                                               |
| *Width*       | 必选      | Single             | 文本框的宽度（以磅为单位）。                                                     |
| *Height*      | 必选      | Single             | 文本框的高度（以磅为单位）。                                                     |

**示例**

``` JavaScript
/*本示例在新文档中的绘图画布上添加一个文本框。*/
function NewCanvasTextbox(){
    //Create a new document and add a drawing canvas
    let docNew = Application.Documents.Add()
    let shpCanvas = docNew.Shapes.AddCanvas(100, 75, 150, 200)

    //Add a text box to the drawing canvas
    shpCanvas.CanvasItems.AddTextbox(msoTextOrientationHorizontal, 1, 1, 100, 100)
}
```

### Shapes.AddTextEffect

在绘图画布上添加一个“艺术字”图形。返回一个 **Shape** 对象，该对象代表“艺术字”，并将其添加至 **CanvasShapes** 集合。

**语法**

**express.AddTextEffect ( *PresetTextEffect* , *Text* , *FontName* , *FontSize* , *FontBold* , *FontItalic* , *Left* , *Top* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型            | 说明                                                                                                                 |
|--------------------|-----------|---------------------|----------------------------------------------------------------------------------------------------------------------|
| *PresetTextEffect* | 必选      | MsoPresetTextEffect | 预设的文本效果。MsoPresetTextEffect 常量的值对应于“‘艺术字’库”对话框中所列的格式（按从上到下、从左到右的顺序编号）。 |
| *Text*             | 必选      | String              | 艺术字中的文字。                                                                                                     |
| *FontName*         | 必选      | String              | 艺术字中所用字体的名称。                                                                                             |
| *FontSize*         | 必选      | Single              | 艺术字中所用字体的大小（以磅为单位）。                                                                               |
| *FontBold*         | 必选      | MsoTriState         | 如果为 MsoTrue，则对艺术字字体应用加粗格式。                                                                         |
| *FontItalic*       | 必选      | MsoTriState         | 如果为 MsoTrue，则对艺术字字体应用倾斜格式。                                                                         |
| *Left*             | 必选      | Single              | 艺术字形状左边缘相对于绘图画布左边缘的位置（以磅为单位）。                                                           |
| *Top*              | 必选      | Single              | 艺术字形状上边缘相对于绘图画布上边缘的位置（以磅为单位）。                                                           |

**说明**

向文档添加艺术字对象时，该对象的高度和宽度将自动根据所指定的文字的大小和数量来设置。

**示例**

``` JavaScript
/*本示例在新文档中添加一个绘图画布，并在绘图画布中插入包含文字“Hello, World”的“艺术字”图形。*/
function test(){
    //Create a new document and add a drawing canvas
    let docNew = Documents.Add()
    let shpCanvas = docNew.Shapes.AddCanvas(100, 100, 150, 50)

    //Add WordArt shape to the drawing canvas
    shpCanvas.CanvasItems.AddTextEffect(msoTextEffect20, "Hello, World", "Tahoma", 15, msoTrue, msoFalse, 120, 120)
}
```

### Shapes.BuildFreeform

建立任意多边形对象。

**语法**

**express.BuildFreeform ( *EditingType* , *X1* , *Y1* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型       | 说明                                                       |
|---------------|-----------|----------------|------------------------------------------------------------|
| *EditingType* | 必选      | MsoEditingType | 第一个顶点的编辑属性。                                     |
| *X1*          | 必选      | Single         | 任意多边形第一个顶点相对于文档左边缘的位置（以磅为单位）。 |
| *Y1*          | 必选      | Single         | 任意多边形第一个顶点相对于文档上边缘的位置（以磅为单位）。 |

**返回值**

FreeformBuilder

**说明**

使用 **AddNodes** 方法将段添加到任意多边形。向该多边形添加至少一个段后，可使用 **ConvertToShape** 方法将 **FreeformBuilder** 对象转换为具有在 **FreeformBuilder** 对象中定义的几何形状描述的 **Shape** 对象。

**示例**

``` JavaScript
/*本示例将一个具有五个顶点的任意多边形添加到活动文档中。*/
function test()
{
let docActive = ActiveDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    docActive.AddNodes(msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300)
    docActive.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 200)
    docActive.AddNodes(msoSegmentLine, msoEditingAuto, 480, 400)
    docActive.AddNodes(msoSegmentLine, msoEditingAuto, 360, 200)
    docActive.ConvertToShape()
}
```

### Shapes.Item

返回集合中的单个 **Shape** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

### Shapes.Range

返回代表范围内的形状的 **ShapeRange** 对象。

**语法**

**express.Range ( *Index* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                         |
|---------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 指定要在指定区域内包含哪些形状。可以是指定 Shapes 集合内某个形状的索引号的整数，也可以是指定形状名称的字符串，或者是包含整数或字符串的数组。 |

**返回值**

ShapeRange

**说明**

**Shape** 对象总是与定位的区域位于同一页上。

| 注释                                                                                                                                              |
|---------------------------------------------------------------------------------------------------------------------------------------------------|
| 对 **Shape** 对象进行的大部分操作也适用于只包含一个形状的 **ShapeRange** 对象。有些操作如果用于包含多个形状的 **ShapeRange** 对象，则会产生错误。 |

**示例**

``` JavaScript
/*以下示例将活动文档中第一个形状的填充前景色设置为紫色。*/
function ShRange(){
    let range = ActiveDocument.Shapes.Range(1).Fill
        range.ForeColor.RGB = 255, 0, 255
        range.Visible = msoTrue
}
```

``` JavaScript
/*以下示例为活动文档中变量形状应用阴影。*/
function ShpRange2(strShpName){
    ActiveDocument.Shapes.Range(strShpName).Shadow.Type = msoShadow6
}
```

``` JavaScript
/*要调用前一个子程序，请在标准代码模块中输入以下代码。*/
function CallShpRange2(){
    let shpArrow = ActiveDocument.Shapes.AddShape(msoShapeLeftArrow, 200, 400, 50, 75)
    shpArrow.Name = "myShape"
    let strName = shpArrow.Name
    ShpRange2(strName)
}
```

``` JavaScript
/*以下示例选择活动文档的第一和第三个形状。*/
function SelectShapeRange(){
    ActiveDocument.Shapes.Range(Array(1, 3)).Select()
}
```

``` JavaScript
/*以下示例选择并删除活动文档的第一个形状中的形状。本示例假定第一个形状是画布形状。*/
function CanvasShapeRange(){
    let rngCanvasShapes = ActiveDocument.Shapes.Item(1).CanvasItems.Range(1)
        rngCanvasShapes.Select()
        rngCanvasShapes.Delete()
}
```

### Shapes.SelectAll

选择形状集合中的所有形状。

**语法**

**express.SelectAll ()**

\*express是一个代表 Shapes 对象的变量。

**说明**

此方法不能选择 **InlineShape** 对象。不能使用此方法选择多个画布。

**示例**

``` JavaScript
/*以下示例选择活动文档中的所有形状。*/
function SelectAllShapes(){
    ActiveDocument.Shapes.SelectAll()
}
```

``` JavaScript
/*以下示例选择活动文档的页眉和页脚中的所有形状，并给每一个形状添加红色阴影。*/
function SelectAllHeaderShapes(){
    let selectall = ActiveDocument.ActiveWindow
        selectall.View.Type = wdPrintView
        selectall.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Shapes.SelectAll()
    let shadow = Selection.ShapeRange.Shadow
        shadow.Type = msoShadow10
        shadow.ForeColor.RGB = 220, 0, 0
}
```

## 成员属性

### Shapes.AddCallout

在绘图画布上添加无框线的标注。

**语法**

**express.AddCallout**

\*express是一个代表 Shapes 对象的变量。

**说明**

使用 **AddShape** 方法可以插入多种不同的标注，例如气球和云朵。

**示例**

``` JavaScript
/*本示例在新创建的绘图画布上添加标注。*/
function NewCanvasCallout(){
    //Add drawing canvas to the active document
    let shpCanvas = ActiveDocument.Shapes.AddCanvas(150, 150, 200, 300)

    //Add callout to the drawing canvas
    shpCanvas.CanvasItems.AddCallout(msoCalloutTwo, 100, 40, 150, 75)
}
```

### Shapes.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Shapes 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Shapes.Count

返回一个 **Long** 类型的值，该值代表集合中图形的数量。只读。

**语法**

**express.Count**

\*express是一个代表 Shapes 对象的变量。

### Shapes.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Shapes 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Shapes.Parent

返回一个 **Object** 类型值，该值代表指定 **Shapes** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Shapes 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Shapes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TextInput 对象

## 简介

代表一个文本型窗体域。

## 说明

使用 FormFields ( *Index* ) 可返回一个 FormField 对象，其中 *Index* 是与文本型窗体域或索引号相关联的书签名称。使用 TextInput 属性和 FormField 对象可返回一个 TextInput 对象。以下示例删除活动文档中名为“Text1”的文本型窗体域的内容。

``` JavaScript
ActiveDocument.FormFields.Item("Text1").TextInput.Clear()
```

索引号代表窗体域在 FormFields 集合中的位置。以下示例检查活动文档中第一个窗体域的类型。如果该窗体域是文本型窗体域，则该示例将“Mission Critical”设置为该域的值。

``` JavaScript
function test()
{
if(ActiveDocument.FormFields.Item(1).Type == wdFieldFormTextInput){
    ActiveDocument.FormFields.Item(1).Result = "Mission Critical"
}
}
```

以下示例确定 *ffield* 变量在设置为默认的文本以前，是否在活动文档中代表一个有效的文本型窗体域。

``` JavaScript
function test()
{
let ffield = ActiveDocument.FormFields.Item(1).TextInput

if(ffield.Valid == true){
    ffield.Default = "Type your name here"
}
else{
    MsgBox("First field is not a text box")
}
}
```

使用 Add 方法和 FormFields 对象可添加一个文本型窗体域。以下示例在活动文档的开始处添加文本型窗体域，然后将该窗体域的名称设置为“FirstName”。

``` JavaScript
function test()
{
let ffield = ActiveDocument.FormFields.Add(ActiveDocument.Range(0, 0), wdFieldFormTextInput)
ffield.Name = "FirstName"
}
```

## 方法

| 序号 | 名称                            | 说明                             |
|:----:|:--------------------------------|:---------------------------------|
|  1   | [Clear](#TextInput.Clear)       | 删除指定的文字型窗体域中的文字。 |
|  2   | [EditType](#TextInput.EditType) | 为指定的文字型窗体域设置选项。   |

## 属性

| 序号 | 名称                                  | 说明                                                                                |
|:----:|:--------------------------------------|:------------------------------------------------------------------------------------|
|  1   | [Application](#TextInput.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                      |
|  2   | [Creator](#TextInput.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。        |
|  3   | [Default](#TextInput.Default)         | 返回或设置代表默认文本框内容的文本。 String 类型，可读写。                          |
|  4   | [Format](#TextInput.Format)           | 返回指定文本框的文字格式。 String 类型，只读。                                      |
|  5   | [Parent](#TextInput.Parent)           | 返回一个 Object 类型值，该值代表指定 TextInput 对象的父对象。                       |
|  6   | [Type](#TextInput.Type)               | 返回文字型窗体域的类型。只读 WdTextFormFieldType 类型。                             |
|  7   | [Valid](#TextInput.Valid)             | 如果指定的窗体域对象为有效的复选框窗体域，则该属性值为 True 。 Boolean 类型，只读。 |
|  8   | [Width](#TextInput.Width)             | 返回或设置指定文本输入域的宽度（以磅为单位）。可读写 Long 类型。                    |

## 成员方法

### TextInput.Clear

删除指定的文字型窗体域中的文字。

**语法**

**express.Clear ()**

\*express是一个代表 TextInput 对象的变量。

**示例**

``` JavaScript
/*此示例保护文档中的窗体，并当第一个窗体域是文字型窗体域时，删除其中的文本。*/
function test()
{
ActiveDocument.Protect(wdAllowOnlyFormFields, true)

if(ActiveDocument.FormFields.Item(1).Type == wdFieldFormTextInput){
    ActiveDocument.FormFields.Item(1).TextInput.Clear()
}

}
```

### TextInput.EditType

为指定的文字型窗体域设置选项。

**语法**

**express.EditType ( *Type* , *Default* , *Format* , *Enabled* )**

\*express是一个代表 TextInput 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型            | 说明                                                                                                                                                               |
|-----------|-----------|---------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Type*    | 必选      | WdTextFormFieldType | 文本框的类型。                                                                                                                                                     |
| *Default* | 可选      | Variant             | 出现在文本框中的默认文字。                                                                                                                                         |
| *Format*  | 可选      | Variant             | 设置文本、数字或日期格式的格式字符串（例如，“0.00”、“Title Case”或“M/d/yy”）。有关格式的更多示例，请查看“文字型窗体域选项”对话框中指定文字型窗体域类型的格式列表。 |
| *Enabled* | 可选      | Variant             | 如果为 True，则可以向窗体域中输入文字。                                                                                                                            |

**示例**

``` JavaScript
/*本示例在活动文档的开头添加一个名为“Today”的文字型窗体域。EditType 方法将窗体域的类型设置为 wdCurrentDateText，并将日期格式设置为“M/d/yy”。*/
function test()
{
let DateText = ActiveDocument.FormFields.Add(ActiveDocument.Range(0, 0), wdFieldFormTextInput)

DateText.Name = "Today"
DateText.TextInput.EditType(wdCurrentDateText, null, "M/d/yy", false)
}
```

## 成员属性

### TextInput.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 TextInput 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### TextInput.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 TextInput 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### TextInput.Default

返回或设置代表默认文本框内容的文本。 **String** 类型，可读写。

**语法**

**express.Default**

\*express是一个代表 TextInput 对象的变量。

### TextInput.Format

返回指定文本框的文字格式。 **String** 类型，只读。

**语法**

**express.Format**

\*express是一个代表 TextInput 对象的变量。

**示例**

``` JavaScript
/*本示例显示活动文档中第一个域中的文本格式。*/
function test()
{
if(ActiveDocument.FormFields.Item(1).Type == wdFieldFormTextInput){
    MsgBox(ActiveDocument.FormFields.Item(1).TextInput.Format)
}
else{
    MsgBox("First field is not a text form field")
}
}
```

### TextInput.Parent

返回一个 **Object** 类型值，该值代表指定 **TextInput** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 TextInput 对象的变量。

### TextInput.Type

返回文字型窗体域的类型。只读 **WdTextFormFieldType** 类型。

**语法**

**express.Type**

\*express是一个代表 TextInput 对象的变量。

### TextInput.Valid

如果指定的窗体域对象为有效的复选框窗体域，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.Valid**

\*express是一个代表 TextInput 对象的变量。

**示例**

``` JavaScript
/*本示例确定活动文档中第一个窗体域是否为文字型窗体域。如果 Valid 属性值为 True，则该文字型窗体域的内容更改为“Hello”。 */
function test()
{
if(ActiveDocument.FormFields.Item(1).TextInput.Valid == true){
    ActiveDocument.FormFields.Item(1).Result = "Hello"
}
}
```

### TextInput.Width

返回或设置指定文本输入域的宽度（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Width**

\*express是一个代表 TextInput 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/TextInput](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# XMLChildNodeSuggestions 对象

## 简介

代表 XMLChildNodeSuggestion 对象的集合，这些对象代表根据架构可能是指定元素的有效子元素的元素。该集合为只读。 XMLChildNodeSuggestions 集合中的每个 XMLChildNodeSuggestion 对象都是所允许的、可能的 XML 元素列表中的一项，该列表位于 “XML 结构” 任务窗格的底部。

## 方法

| 序号 | 名称                                  | 说明                                           |
|:----:|:--------------------------------------|:-----------------------------------------------|
|  1   | [Item](#XMLChildNodeSuggestions.Item) | 返回集合中的单个 XMLChildNodeSuggestion 对象。 |

## 属性

| 序号 | 名称                                                | 说明                                                                         |
|:----:|:----------------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#XMLChildNodeSuggestions.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#XMLChildNodeSuggestions.Count)             | 返回一个 Long 类型的值，该值代表集合中子元素建议的数量。只读。               |
|  3   | [Creator](#XMLChildNodeSuggestions.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#XMLChildNodeSuggestions.Parent)           | 返回一个 Object 类型值，该值代表指定 XMLChildNodeSuggestions 对象的父对象。  |

## 成员方法

### XMLChildNodeSuggestions.Item

返回集合中的单个 **XMLChildNodeSuggestion** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 XMLChildNodeSuggestions 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

**返回值**

XMLChildNodeSuggestion

## 成员属性

### XMLChildNodeSuggestions.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 XMLChildNodeSuggestions 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### XMLChildNodeSuggestions.Count

返回一个 **Long** 类型的值，该值代表集合中子元素建议的数量。只读。

**语法**

**express.Count**

\*express是一个代表 XMLChildNodeSuggestions 对象的变量。

### XMLChildNodeSuggestions.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 XMLChildNodeSuggestions 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### XMLChildNodeSuggestions.Parent

返回一个 **Object** 类型值，该值代表指定 **XMLChildNodeSuggestions** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 XMLChildNodeSuggestions 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/XMLChildNodeSuggestions](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# XSLTransforms 对象

## 简介

XSLTransform 对象的集合，这些对象代表特定 XML 命名空间的所有可扩展样式表语言转换 (XSLT)。

## 方法

| 序号 | 名称                        | 说明                                                                                                                                            |
|:----:|:----------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Add](#XSLTransforms.Add)   | 返回一个 XSLTransform 对象，该对象代表添加到指定架构的 XSLT 集合的“可扩展样式表语言转换”(Extensible Stylesheet Language Transformation, XSLT)。 |
|  2   | [Item](#XSLTransforms.Item) | 返回集合中的 XSLTransform 对象。                                                                                                                |

## 属性

| 序号 | 名称                                      | 说明                                                                         |
|:----:|:------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#XSLTransforms.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#XSLTransforms.Count)             | 返回一个 Long 类型的值，该值代表集合中 XSLTransform 的数量。只读。           |
|  3   | [Creator](#XSLTransforms.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#XSLTransforms.Parent)           | 返回一个 Object 类型值，该值代表指定 XSLTransforms 对象的父对象。            |

## 成员方法

### XSLTransforms.Add

返回一个 **XSLTransform** 对象，该对象代表添加到指定架构的 XSLT 集合的“可扩展样式表语言转换”(Extensible Stylesheet Language Transformation, XSLT)。

**语法**

**express.Add ( *Location* , *Alias* , *InstallForAllUsers* )**

\*express是一个代表 XSLTransforms 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                                                          |
|----------------------|-----------|----------|-------------------------------------------------------------------------------|
| *Location*           | 必选      | String   | XSLT 的路径和文件名。可以为本地文件路径、网络路径或 Internet 地址。           |
| *Alias*              | 可选      | String   | 架构库中显示的 XSLT 名称。                                                    |
| *InstallForAllUsers* | 可选      | Boolean  | 如果登录到计算机上的所有用户都能访问和使用新架构，则为 True。默认值为 False。 |

**返回值**

XSLTransform

**示例**

``` JavaScript
/*以下代码将一个架构添加到架构库中，然后将 XSLT 添加到该新添加的架构中。*/
function AddXSLT() {
    let objSchema = Application.XMLNamespaces.Item("SimpleSample")
    let objTransform = objSchema.XSLTransforms.Add("c:\\schemas\\simplesample.xsl")
}
```

### XSLTransforms.Item

返回集合中的 **XSLTransform** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 XSLTransforms 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

**返回值**

XSLTransform

## 成员属性

### XSLTransforms.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 XSLTransforms 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### XSLTransforms.Count

返回一个 **Long** 类型的值，该值代表集合中 XSLTransform 的数量。只读。

**语法**

**express.Count**

\*express是一个代表 XSLTransforms 对象的变量。

### XSLTransforms.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 XSLTransforms 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### XSLTransforms.Parent

返回一个 **Object** 类型值，该值代表指定 **XSLTransforms** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 XSLTransforms 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/XSLTransforms](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
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
# Font 对象

## 简介

包含对象的字体属性（如字体名称、字号、颜色等）。

## 属性

| 序号 | 名称                             | 说明                                                                     |
|:----:|:---------------------------------|:-------------------------------------------------------------------------|
|  1   | [AllCaps](#Font.AllCaps)         | 如果字体格式为全部字母大写，则该属性值为 True 。 Long 类型，可读写。     |
|  2   | [Bold](#Font.Bold)               | 如果字体格式为加粗，则该属性值为 True 。 Long 类型，可读写。             |
|  3   | [Color](#Font.Color)             |                                                                          |
|  4   | [Italic](#Font.Italic)           | 如果将字体或范围设置为倾斜格式，则该属性值为 True 。 Long 类型，可读写。 |
|  5   | [Name](#Font.Name)               | 返回或设置指定对象的名称。 String 类型，可读写。                         |
|  6   | [Shadow](#Font.Shadow)           | 如果将指定字体设置为阴影格式，则该属性值为 True 。可读写 Long 类型。     |
|  7   | [Size](#Font.Size)               | 返回或设置字号（以磅为单位）。 Single 类型，可读写。                     |
|  8   | [Subscript](#Font.Subscript)     | 如果字体格式为下标，则该属性值为 True 。 Long 类型，可读写。             |
|  9   | [Superscript](#Font.Superscript) | 如果字体格式为上标，则该属性值为 True 。 Long 类型，可读写。             |
|  10  | [Underline](#Font.Underline)     | 返回或设置应用于字体的下划线类型。可读写 WdUnderline 类型。              |

## 成员属性

### Font.AllCaps

如果字体格式为全部字母大写，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.AllCaps**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdU** **ndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。该属性可设置为 **True** 、 **False** 或 **wdToggle** （与当前设置相反）。

如果设置 **AllCaps** 为 **True** ，则 **SmallCaps** 属性的值为 **False** ，反之亦然。

**示例**

``` JavaScript
/*本示例检查活动文档第三段的文本格式是否设置为全部字母大写。*/
function test() {
  if(Application.ActiveDocument.Paragraphs.Item(3).Range.Font.AllCaps == true){ 
      alert("Text is all caps.")
  }
  else{
      alert("Text is not all caps.")
  }
}

/*本示例将选定文本的字体格式设置为全部字母大写。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal){
      Application.Selection.Font.AllCaps = true
  }
  else{
      alert("You need to select some text.")
  }
}
```

### Font.Bold

如果字体格式为加粗，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Bold**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

**示例**

``` JavaScript
/*本示例实现的功能是：如果选定内容中有一部分字体为加粗格式，则将全部选定内容的格式都设为加粗格式。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal) {
      if(Application.Selection.Font.Bold == wdUndefined) {
          Application.Selection.Font.Bold = true
      }
  }
  else{
      alert("You need to select some text.")
  }
}
```

### Font.Color

**语法**

**express.Color**

\*express是一个代表 Font 对象的变量。

### Font.Italic

如果将字体或范围设置为倾斜格式，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Italic**

\*express是一个代表 Font 对象的变量。

**说明**

该属性返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

**示例**

``` JavaScript
/*以下示例检查所选内容中的倾斜格式并删除所选内容中存在的倾斜格式。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal){
      let mySel = Application.Selection.Font.Italic

      if(mySel == wdUndefined || mySel == true){
          alert("there is italic text in selection. " + "Click OK to remove.")
          Application.Selection.Font.Italic = false
      }
      else{
          alert("No italic text in the selection.")
      }
  }
  else{
      alert("You need to select some text.")
  }
}
```

### Font.Name

返回或设置指定对象的名称。 **String** 类型，可读写。

**语法**

**express.Name**

\*express是一个代表 Font 对象的变量。

**说明**

本示例将选定内容设置为 Arial 字体的加粗格式。

**示例**

``` JavaScript
本示例将选定内容设置为 Arial 字体的加粗格式。
function test() {
  Application.Selection.Font.Name = "Arial"
  Application.Selection.Font.Bold = true
}
```

### Font.Shadow

如果将指定字体设置为阴影格式，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.Shadow**

\*express是一个代表 Font 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例对所选内容应用阴影和加粗格式。/*
function test() {
  if(Application.Selection.Type == wdSelectionNormal){
      Application.Selection.Font.Shadow = true
      Application.Selection.Font.Bold = true
  }
  else {
      alert("You need to select some text.")
  }
}
```

### Font.Size

返回或设置字号（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.Size**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/*以下示例插入文本并将该文本的第七个单词的字号设置为 20 磅。*/
function test() {
  Application.Selection.Collapse(wdCollapseEnd)
  let rng = Application.Selection.Range
  rng.Font.Reset()
  rng.InsertBefore( "This is a demonstration of font size.")
  rng.Words.Item(7).Font.Size = 20
}

/*以下示例确定所选文本的字号。*/
function test() {
  let mySel = Application.Selection.Font.Size
  if(mySel == wdUndefined){
      alert("there is a mix of font sizes in the selection.")
  }
  else {
      alert(mySel + " points")
  }
}
```

### Font.Subscript

如果字体格式为下标，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Subscript**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

如果将 **Subscript** 属性设为 **True** ，则 **Superscript** 属性会设为 **False** ，反之亦然。

**示例**

``` JavaScript
/*本示例在活动文档的开头插入文本，并将第十个字符设为下标。*/
function test() {
  let myRange = Application.ActiveDocument.Range(0, 0)
  myRange.InsertAfter("Water = H20")
  myRange.Characters.Item(10).Font.Subscript = true
}

/*本示例检查所选文本的下标格式。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal){
      let mySel = Application.Selection.Font.Subscript
      if(mySel == wdUndefined || mySel == true){
          alert("Subscript text exists in the selection.")
      }
      else{
          alert("No subscript text in the selection.")
      }
  }
  else{
      alert("You need to select some text.")
  }
}
```

### Font.Superscript

如果字体格式为上标，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Superscript**

\*express是一个代表 Font 对象的变量。

**说明**

返回 **True** 、 **False** 或 **wdUndefined** （当返回值既可为 **True** ，也可为 **False** 时取该值）。可设置为 **True** 、 **False** 或 **wdToggle** 。

如果将 **Superscript** 属性设为 **True** ，则 **Subscript** 属性会设为 **False** ，反之亦然。

**示例**

``` JavaScript
/*本示例在活动文档的开头插入文本，并将第四个单词的两个字符设为上标。*/
function test() {
  let myRange = Application.ActiveDocument.Range(0, 0)
  myRange.InsertAfter("Superscript in the 4th word.")
  Application.ActiveDocument.Range(20, 22).Font.Superscript = true
}

/*本示例将所选文本设为上标。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal){
      Application.Selection.Font.Superscript = true
  }
  else{
      alert("You need to select some text.")
  }
}
```

### Font.Underline

返回或设置应用于字体的下划线类型。可读写 **WdUnderline** 类型。

**语法**

**express.Underline**

\*express是一个代表 Font 对象的变量。

**示例**

``` JavaScript
/*以下示例给所选内容添加单下划线。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal){
      Application.Selection.Font.Underline = wdUnderlineSingle
  }
  else{
      alert("You need to select some text.")
  }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Font](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# HeadersFooters 对象

## 简介

包含指定幻灯片、备注页、讲义或母版上的所有 HeaderFooter 对象。

## 说明

每个 HeaderFooter 对象代表一个页眉、页脚、日期和时间或幻灯片编号。

## 方法

| 序号 | 名称                           | 说明                                                   |
|:----:|:-------------------------------|:-------------------------------------------------------|
|  1   | [Clear](#HeadersFooters.Clear) | 从标尺中清除指定的制表位并将其从 TabStops 集合中删除。 |

## 属性

| 序号 | 名称                                                       | 说明                                                                                                                     |
|:----:|:-----------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#HeadersFooters.Application)                 | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                  |
|  2   | [DateAndTime](#HeadersFooters.DateAndTime)                 | 返回一个代表日期和时间项的 HeaderFooter 对象，该项将显示在幻灯片的左下角，或者备注页、讲义或大纲的右上角。只读。         |
|  3   | [DisplayOnTitleSlide](#HeadersFooters.DisplayOnTitleSlide) | 决定标题幻灯片上是否显示页脚、日期和时间以及幻灯片编号。可读/写 MsoTriState 类型。适用于幻灯片母版。                     |
|  4   | [Footer](#HeadersFooters.Footer)                           | 返回一个 HeaderFooter 对象，该对象代表显示于幻灯片底部或者备注页、讲义或大纲的左下角的页脚。只读。                       |
|  5   | [Header](#HeadersFooters.Header)                           | 返回一个 HeaderFooter 对象，该对象代表显示于幻灯片顶部或者备注页、讲义或大纲的左下角的页眉。只读。                       |
|  6   | [Parent](#HeadersFooters.Parent)                           | 返回指定对象的父对象。                                                                                                   |
|  7   | [SlideNumber](#HeadersFooters.SlideNumber)                 | 返回一个 HeaderFooter 对象，该对象代表位于幻灯片右下角的幻灯片编号，或者位于备注页、打印的讲义或大纲右下角的页码。只读。 |

## 成员方法

### HeadersFooters.Clear

从标尺中清除指定的制表位并将其从 **TabStops** 集合中删除。

**语法**

**express.Clear ()**

\*express是一个代表 HeadersFooters 对象的变量。

**示例**

``` JavaScript
/*本示例清除当前演示文稿中第一张幻灯片第二个形状中的文本的所有制表位。*/
function test() {
    let ts = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).TextFrame.Ruler.TabStops
    for (let i = ts.count; i >= 1; i--) {
        ts.Item(i).Clear()
    }
}
```

## 成员属性

### HeadersFooters.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 HeadersFooters 对象的变量。

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

### HeadersFooters.DateAndTime

返回一个代表日期和时间项的 **HeaderFooter** 对象，该项将显示在幻灯片的左下角，或者备注页、讲义或大纲的右上角。只读。

**语法**

**express.DateAndTime**

\*express是一个代表 HeadersFooters 对象的变量。

**示例**

``` JavaScript
/*本示例为活动演示文稿的母版幻灯片设置日期和时间格式。该设定将对所有基于该幻灯片母版的幻灯片有效。*/
function test(){
　　　　let da = Application.ActivePresentation.SlideMaster.HeadersFooters.DateAndTime
　　　　da.Format = ppDateTimeMdyy
　　　　da.UseFormat = true
}
```

### HeadersFooters.DisplayOnTitleSlide

决定标题幻灯片上是否显示页脚、日期和时间以及幻灯片编号。可读/写 **MsoTriState** 类型。适用于幻灯片母版。

**语法**

**express.DisplayOnTitleSlide**

\*express是一个代表 HeadersFooters 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                                 |
|-------------------------------------------------------------------------------|
| msoCTrue                                                                      |
| msoFalse 在除标题幻灯片之外的所有幻灯片上显示页脚、日期和时间以及幻灯片编号。 |
| msoTriStateMixed                                                              |
| msoTriStateToggle                                                             |
| msoTrue 在标题幻灯片上显示页脚、日期和时间以及幻灯片编号。                    |

**示例**

本示例设置标题幻灯片上不显示页脚、日期和时间及幻灯片编号。

``` JavaScript
Application.ActivePresentation.SlideMaster.HeadersFooters.DisplayOnTitleSlide = msoFalse
```

### HeadersFooters.Footer

返回一个 **HeaderFooter** 对象，该对象代表显示于幻灯片底部或者备注页、讲义或大纲的左下角的页脚。只读。

**语法**

**express.Footer**

\*express是一个代表 HeadersFooters 对象的变量。

**示例**

``` JavaScript
/*本示例为当前演示文稿的幻灯片母版设置脚注的文本，并设定在标题幻灯片上显示脚注、日期和时间及幻灯片编号。*/
function test(){
　　　　let foot = Application.ActivePresentation.SlideMaster.HeadersFooters
　　　　foot.Footer.Text = "Introduction"
　　　　foot.DisplayOnTitleSlide = true
}
```

### HeadersFooters.Header

返回一个 **HeaderFooter** 对象，该对象代表显示于幻灯片顶部或者备注页、讲义或大纲的左下角的页眉。只读。

**语法**

**express.Header**

\*express是一个代表 HeadersFooters 对象的变量。

**示例**

``` JavaScript
/*本示例为当前演示文稿的讲义母版设置页眉文本。以大纲或讲义方式打印演示文稿时，该文本将显示于页的左上角。*/
function test(){
　　　　let myHandHF = Application.ActivePresentation.HandoutMaster.HeadersFooters
　　　　myHandHF.Header.Text = "Third Quarter Report"
}
```

### HeadersFooters.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 HeadersFooters 对象的变量。

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

### HeadersFooters.SlideNumber

返回一个 **HeaderFooter** 对象，该对象代表位于幻灯片右下角的幻灯片编号，或者位于备注页、打印的讲义或大纲右下角的页码。只读。

**语法**

**express.SlideNumber**

\*express是一个代表 HeadersFooters 对象的变量。

**示例**

``` JavaScript
/*本示例隐藏当前演示文稿第二张幻灯片的幻灯片编号（如果该编号当前为可见），或显示该幻灯片编号（如果该编号当前为隐藏）。*/
function test() {
    let slides = Application.ActivePresentation.Slides.Item(2).HeadersFooters.SlideNumber
    if (slides.Visible) {
        slides.Visible = false
    }
    else {
        slides.Visible = true
    }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/HeadersFooters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# LineFormat 对象

## 简介

该对象代表线条和箭头格式。对于线条， LineFormat 对象包含关于线条本身的格式信息；对于带有边框的图形，本对象包含该图形的边框的格式信息。

## 属性

| 序号 | 名称                                                     | 说明                                                                                                                                                          |
|:----:|:---------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#LineFormat.Application)                   | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                       |
|  2   | [BackColor](#LineFormat.BackColor)                       | 返回或设置 ColorFormat 对象，该对象代表指定填充对象或图案线条的背景色。可读/写。                                                                              |
|  3   | [BeginArrowheadLength](#LineFormat.BeginArrowheadLength) | 返回或设置指定线条始端的箭头长度。可读/写 MsoArrowheadLength 类型。                                                                                           |
|  4   | [BeginArrowheadStyle](#LineFormat.BeginArrowheadStyle)   | 返回或设置指定线条始端的箭头样式。可读/写 MsoArrowheadStyle 类型。                                                                                            |
|  5   | [BeginArrowheadWidth](#LineFormat.BeginArrowheadWidth)   | 返回或设置指定线条始端的箭头宽度。可读/写 MsoArrowheadWidth 类型。                                                                                            |
|  6   | [Creator](#LineFormat.Creator)                           | 返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。 |
|  7   | [DashStyle](#LineFormat.DashStyle)                       | 返回或设置指定线条的虚线线型。可读/写 MsoLineDashStyle 类型。                                                                                                 |
|  8   | [EndArrowheadLength](#LineFormat.EndArrowheadLength)     | 返回或设置指定线条末端箭头的长度。可读/写 MsoArrowheadLength 类型。                                                                                           |
|  9   | [EndArrowheadStyle](#LineFormat.EndArrowheadStyle)       | 返回或设置指定线条末端箭头的样式。可读/写 MsoArrowheadStyle 类型。                                                                                            |
|  10  | [EndArrowheadWidth](#LineFormat.EndArrowheadWidth)       | 返回或设置指定线条末端箭头的宽度。可读/写 MsoArrowheadWidth 类型。                                                                                            |
|  11  | [ForeColor](#LineFormat.ForeColor)                       | 返回或设置 ColorFormat 对象，该对象表示填充、线条或阴影的前景色。可读/写。                                                                                    |
|  12  | [InsetPen](#LineFormat.InsetPen)                         | 该属性值为 MsoTrue 时，在指定形状的内部绘制线条。可读/写 MsoTriState 类型。                                                                                   |
|  13  | [Parent](#LineFormat.Parent)                             | 返回指定对象的父对象。                                                                                                                                        |
|  14  | [Pattern](#LineFormat.Pattern)                           | 设置或返回一个值，该值代表应用于指定线条的图案。 MsoPatternType 类型，可读写。                                                                                |
|  15  | [Style](#LineFormat.Style)                               | 返回或设置线条样式。可读/写 MsoLineStyle 类型。                                                                                                               |
|  16  | [Transparency](#LineFormat.Transparency)                 | 该属性返回或设置指定填充、阴影或线条的透明度，可取 0.0（不透明）至 1.0（透明）之间的数值。 Single 类型，可读/写。                                             |
|  17  | [Visible](#LineFormat.Visible)                           | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 MsoTriState 类型。                                                                                |
|  18  | [Weight](#LineFormat.Weight)                             | 以磅为单位返回或设置指定线条的粗细。可读/写 Single 类型。                                                                                                     |

## 成员属性

### LineFormat.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 LineFormat 对象的变量。

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

### LineFormat.BackColor

返回或设置 **ColorFormat** 对象，该对象代表指定填充对象或图案线条的背景色。可读/写。

**语法**

**express.BackColor**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变。*/
function test(){
　　　　let myDocument = ActivePresentation.Slides.Item(1)
　　　　let fill = myDocument.Shapes.AddShape(msoShapeRectangle,90, 90, 90, 50).Fill
　　　　fill.ForeColor.RGB = (128, 0, 0)
　　　　fill.BackColor.RGB = (170, 170, 170)
　　　　fill.TwoColorGradient(msoGradientHorizontal, 1)
}
```

``` JavaScript
/*以下示例将一条带图案的线条添加至 myDocument。*/
function test() {
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let line = myDocument.Shapes.AddLine(10, 100, 250, 0).Line
    line.Weight = 6
    line.ForeColor.RGB = (0, 0, 255)
    line.BackColor.RGB = (128, 0, 0)
    line.Pattern = msoPatternDarkDownwardDiagonal
}
```

### LineFormat.BeginArrowheadLength

返回或设置指定线条始端的箭头长度。可读/写 **MsoArrowheadLength** 类型。

**语法**

**express.BeginArrowheadLength**

\*express是一个代表 LineFormat 对象的变量。

**说明**

| MsoArrowheadLength 可以是下列 MsoArrowheadLength 常量之一。 |
|-------------------------------------------------------------|
| msoArrowheadLengthMedium                                    |
| msoArrowheadLengthMixed                                     |
| msoArrowheadLong                                            |
| msoArrowheadShort                                           |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
　　　　line.BeginArrowheadLength = msoArrowheadShort
　　　　line.BeginArrowheadStyle = msoArrowheadOval
　　　　line.BeginArrowheadWidth = msoArrowheadNarrow
　　　　line.EndArrowheadLength = msoArrowheadLong
　　　　line.EndArrowheadStyle = msoArrowheadTriangle
　　　　line.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.BeginArrowheadStyle

返回或设置指定线条始端的箭头样式。可读/写 **MsoArrowheadStyle** 类型。

**语法**

**express.BeginArrowheadStyle**

\*express是一个代表 LineFormat 对象的变量。

**说明**

| MsoArrowheadStyle 可以是下列 MsoArrowheadStyle 常量之一。 |
|-----------------------------------------------------------|
| msoArrowheadDiamond                                       |
| msoArrowheadNone                                          |
| msoArrowheadOpen                                          |
| msoArrowheadOval                                          |
| msoArrowheadStealth                                       |
| msoArrowheadStyleMixed                                    |
| msoArrowheadTriangle                                      |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
　　　　line.BeginArrowheadLength = msoArrowheadShort
　　　　line.BeginArrowheadStyle = msoArrowheadOval
　　　　line.BeginArrowheadWidth = msoArrowheadNarrow
　　　　line.EndArrowheadLength = msoArrowheadLong
　　　　line.EndArrowheadStyle = msoArrowheadTriangle
　　　　line.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.BeginArrowheadWidth

返回或设置指定线条始端的箭头宽度。可读/写 **MsoArrowheadWidth** 类型。

**语法**

**express.BeginArrowheadWidth**

\*express是一个代表 LineFormat 对象的变量。

**说明**

| MsoArrowheadWidth 可以是下列 MsoArrowheadWidth 常量之一。 |
|-----------------------------------------------------------|
| msoArrowheadNarrow                                        |
| msoArrowheadWide                                          |
| msoArrowheadWidthMedium                                   |
| msoArrowheadWidthMixed                                    |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
　　　　line.BeginArrowheadLength = msoArrowheadShort
　　　　line.BeginArrowheadStyle = msoArrowheadOval
　　　　line.BeginArrowheadWidth = msoArrowheadNarrow
　　　　line.EndArrowheadLength = msoArrowheadLong
　　　　line.EndArrowheadStyle = msoArrowheadTriangle
　　　　line.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.Creator

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

\*express是一个代表 LineFormat 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序

**示例**

``` JavaScript
/*以下示例显示关于 myObject 创建者的消息。*/
function test(){
　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
　　　　if (myObject.Creator == "0x50575054") {
　　　　    alert("This is aWPP object")
　　　　}
　　　　else {
　　　　    alert("This is not aWPP object")
　　　　}
}
```

### LineFormat.DashStyle

返回或设置指定线条的虚线线型。可读/写 **MsoLineDashStyle** 类型。

**语法**

**express.DashStyle**

\*express是一个代表 LineFormat 对象的变量。

**说明**

| MsoLineDashStyle 可以是下列 MsoLineDashStyle 常量之一。 |
|---------------------------------------------------------|
| msoLineDash                                             |
| msoLineDashDot                                          |
| msoLineDashDotDot                                       |
| msoLineDashStyleMixed                                   |
| msoLineLongDash                                         |
| msoLineLongDashDot                                      |
| msoLineRoundDot                                         |
| msoLineSolid                                            |
| msoLineSquareDot                                        |

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加一条蓝色虚线。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let line = myDocument.Shapes.AddLine(10, 10, 250, 250).Line
　　　　line.DashStyle = msoLineDashDotDot
　　　　line.ForeColor.RGB = (50, 0, 128)
}
```

### LineFormat.EndArrowheadLength

返回或设置指定线条末端箭头的长度。可读/写 **MsoArrowheadLength** 类型。

**语法**

**express.EndArrowheadLength**

\*express是一个代表 LineFormat 对象的变量。

**说明**

| MsoArrowheadLength 可以是下列 MsoArrowheadLength 常量之一。 |
|-------------------------------------------------------------|
| msoArrowheadLengthMedium                                    |
| msoArrowheadLengthMixed                                     |
| msoArrowheadLong                                            |
| msoArrowheadShort                                           |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
　　　　line.BeginArrowheadLength = msoArrowheadShort
　　　　line.BeginArrowheadStyle = msoArrowheadOval
　　　　line.BeginArrowheadWidth = msoArrowheadNarrow
　　　　line.EndArrowheadLength = msoArrowheadLong
　　　　line.EndArrowheadStyle = msoArrowheadTriangle
　　　　line.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.EndArrowheadStyle

返回或设置指定线条末端箭头的样式。可读/写 **MsoArrowheadStyle** 类型。

**语法**

**express.EndArrowheadStyle**

\*express是一个代表 LineFormat 对象的变量。

**说明**

| MsoArrowheadStyle 可以是下列 MsoArrowheadStyle 常量之一。 |
|-----------------------------------------------------------|
| msoArrowheadDiamond                                       |
| msoArrowheadNone                                          |
| msoArrowheadOpen                                          |
| msoArrowheadOval                                          |
| msoArrowheadStealth                                       |
| msoArrowheadStyleMixed                                    |
| msoArrowheadTriangle                                      |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
　　　　line.BeginArrowheadLength = msoArrowheadShort
　　　　line.BeginArrowheadStyle = msoArrowheadOval
　　　　line.BeginArrowheadWidth = msoArrowheadNarrow
　　　　line.EndArrowheadLength = msoArrowheadLong
　　　　line.EndArrowheadStyle = msoArrowheadTriangle
　　　　line.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.EndArrowheadWidth

返回或设置指定线条末端箭头的宽度。可读/写 **MsoArrowheadWidth** 类型。

**语法**

**express.EndArrowheadWidth**

\*express是一个代表 LineFormat 对象的变量。

**说明**

| MsoArrowheadWidth 可以是下列 MsoArrowheadWidth 常量之一。 |
|-----------------------------------------------------------|
| msoArrowheadNarrow                                        |
| msoArrowheadWide                                          |
| msoArrowheadWidthMedium                                   |
| msoArrowheadWidthMixed                                    |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加一条线。在该线条的起点处有一个短而窄的椭圆，在线条的终点处有一个长而宽的三角形。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line
　　　　line.BeginArrowheadLength = msoArrowheadShort
　　　　line.BeginArrowheadStyle = msoArrowheadOval
　　　　line.BeginArrowheadWidth = msoArrowheadNarrow
　　　　line.EndArrowheadLength = msoArrowheadLong
　　　　line.EndArrowheadStyle = msoArrowheadTriangle
　　　　line.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.ForeColor

返回或设置 **ColorFormat** 对象，该对象表示填充、线条或阴影的前景色。可读/写。

**语法**

**express.ForeColor**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let fil = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill
　　　　fil.ForeColor.RGB = 128, 0, 0
　　　　fil.BackColor.RGB = 170, 170, 170
　　　　fil.TwoColorGradient(msoGradientHorizontal, 1)
}
```

``` JavaScript
/*以下示例将一条带图案的线条添加至 myDocument。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let lin = myDocument.Shapes.AddLine(10, 100, 250, 0).Line
　　　　lin.Weight = 6
　　　　lin.ForeColor.RGB = 0, 0, 255
　　　　lin.BackColor.RGB = 128, 0, 0
　　　　lin.Pattern = msoPatternDarkDownwardDiagonal
}
```

### LineFormat.InsetPen

该属性值为 **MsoTrue** 时，在指定形状的内部绘制线条。可读/写 **MsoTriState** 类型。

**语法**

**express.InsetPen**

\*express是一个代表 LineFormat 对象的变量。

**说明**

如果此属性试图在不支持插入笔绘图的任何 WPS Office自选图形中设置插入笔绘图，将发生错误。

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue 不适用于此属性。                     |
| msoFalse 默认值。不启用插入笔。               |
| msoTriStateMixed 不适用于此属性。             |
| msoTriStateToggle 不适用于此属性。            |
| msoTrue 启用插入笔。                          |

**示例**

下面的代码行在形状中启用插入笔。本示例假设当前演示文稿的第一张幻灯片中包含形状并且该形状支持插入笔绘图。

``` JavaScript
function DrawLinesInsideShape(){
    Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Line.InsetPen = msoTrue
}
```

### LineFormat.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 LineFormat 对象的变量。

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

### LineFormat.Pattern

设置或返回一个值，该值代表应用于指定线条的图案。 **MsoPatternType** 类型，可读写。

**语法**

**express.Pattern**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加带有图案的线条。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let lin = myDocument.Shapes.AddLine(10, 100, 250, 0).Line
　　　　lin.Weight = 6
　　　　lin.ForeColor.RGB = 0, 0, 255
　　　　lin.BackColor.RGB = 128, 0, 0
　　　　lin.Pattern = msoPatternDarkDownwardDiagonal
}
```

### LineFormat.Style

返回或设置线条样式。可读/写 **MsoLineStyle** 类型。

**语法**

**express.Style**

\*express是一个代表 LineFormat 对象的变量。

**说明**

| MsoLineStyle 可以是下列 MsoLineStyle 常量之一。 |
|-------------------------------------------------|
| msoLineSingle                                   |
| msoLineStyleMixed                               |
| msoLineThickBetweenThin                         |
| msoLineThickThin                                |
| msoLineThinThick                                |
| msoLineThinThin                                 |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加粗的蓝色复合线。该复合线由一根粗线及其两侧的各一细线组成。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let lin = myDocument.Shapes.AddLine(10, 10, 250, 250).Line
　　　　lin.Style = msoLineThickBetweenThin
　　　　lin.Weight = 8
　　　　lin.ForeColor.RGB = 0, 0, 255
}
```

### LineFormat.Transparency

该属性返回或设置指定填充、阴影或线条的透明度，可取 0.0（不透明）至 1.0（透明）之间的数值。 **Single** 类型，可读/写。

**语法**

**express.Transparency**

\*express是一个代表 LineFormat 对象的变量。

**说明**

此属性的值仅影响均一颜色填充和线条的外观。对图案线条及渐变、图案、图片和纹理填充的外观无效。

**示例**

``` JavaScript
/*本例将 myDocument 中第三个形状的阴影设置为半透明的红色。如果形状原来无阴影，以下示例为它添加阴影。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let sha = myDocument.Shapes.Item(3).Shadow
　　　　sha.Visible = true
　　　　sha.ForeColor.RGB = 255, 0, 0
　　　　sha.Transparency = 0.5
}
```

### LineFormat.Visible

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 **MsoTriState** 类型。

**语法**

**express.Visible**

\*express是一个代表 LineFormat 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue：指定对象或对象格式是可见的。         |

**示例**

``` JavaScript
/*以下示例为当前演示文稿中第一张幻灯片的第三个形状设置阴影的水平偏移量和垂直偏移量，使其向右偏移 5 磅并向上偏移 3 磅。如果该形状原来没有阴影，本例将会为其添上阴影。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let sha = myDocument.Shapes.Item(3).Shadow
　　　　sha.Visible = msoTrue
　　　　sha.OffsetX = 5
　　　　sha.OffsetY = -3
}
```

### LineFormat.Weight

以磅为单位返回或设置指定线条的粗细。可读/写 **Single** 类型。

**语法**

**express.Weight**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/*本示例将一条两磅粗的绿色虚线添加到 myDocument。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let lin = myDocument.Shapes.AddLine(10, 10, 250, 250).Line
　　　　lin.DashStyle = msoLineDashDotDot
　　　　lin.ForeColor.RGB = 0, 255, 255
　　　　lin.Weight = 2
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/LineFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PrintRange 对象

## 简介

代表要打印的连续幻灯片或页的范围。

## 说明

PrintRange 对象是 PrintRanges 集合的成员。 PrintRanges 集合包含已为指定演示文稿定义的所有打印范围。

可以在 PrintRanges 集合中设置独立于 RangeType 设置的打印区域。这些打印区域在包含它们的演示文稿加载时始终有效。 RangeType 属性设为 ppPrintSlideRange 时，应用 PrintRanges 集合中的区域。

## 方法

| 序号 | 名称                         | 说明             |
|:----:|:-----------------------------|:-----------------|
|  1   | [Delete](#PrintRange.Delete) | 删除指定的对象。 |

## 属性

| 序号 | 名称                                   | 说明                                                            |
|:----:|:---------------------------------------|:----------------------------------------------------------------|
|  1   | [Application](#PrintRange.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。         |
|  2   | [End](#PrintRange.End)                 | 返回指定打印范围中最后一张幻灯片的编号。只读 Long 类型。        |
|  3   | [Parent](#PrintRange.Parent)           | 返回指定对象的父对象。                                          |
|  4   | [Start](#PrintRange.Start)             | 返回要打印的幻灯片范围内第一张幻灯片的编号。只读 Integer 类型。 |

## 成员方法

### PrintRange.Delete

删除指定的对象。

**语法**

**express.Delete ()**

\*express是一个代表 PrintRange 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

## 成员属性

### PrintRange.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 PrintRange 对象的变量。

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

### PrintRange.End

返回指定打印范围中最后一张幻灯片的编号。只读 **Long** 类型。

**语法**

**express.End**

\*express是一个代表 PrintRange 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条信息，该信息指示当前演示文稿中第一个打印范围的起始和结束幻灯片的编号。*/
function test() {
    let print = Application.ActivePresentation.PrintOptions.Ranges
    if (print.Count > 0) {
        let item = print.Item(1)
        alert("Print range 1 starts on slide " + item.Start + " and ends on slide " + item.End)
    }
}
```

### PrintRange.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 PrintRange 对象的变量。

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

### PrintRange.Start

返回要打印的幻灯片范围内第一张幻灯片的编号。只读 **Integer** 类型。

**语法**

**express.Start**

\*express是一个代表 PrintRange 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条信息，该信息指示当前演示文稿中第一个打印范围的起始和结束幻灯片的编号。*/
function test() {
    let print = Application.ActivePresentation.PrintOptions.Ranges
    if (print.Count > 0) {
        let item = print.Item(1)
        alert("Print range 1 starts on slide " + item.Start + " and ends on slide " + item.End)
    }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/PrintRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomXMLPrefixMapping 对象

## 简介

代表一个命名空间前缀。

## 说明

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

以下示例通过向 CustomXMLPrefixMapping 集合添加命名空间和前缀，创建一个 CustomXMLPrefixMapping 对象。

``` JavaScript
let objNamespace = Application.ActivePresentation.CustomXMLParts.Item(1).NamespaceManager.AddNamespace("xs", "urn:invoice:namespace")
```

## 属性

| 序号 | 名称                                                 | 说明                                                                               |
|:----:|:-----------------------------------------------------|:-----------------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLPrefixMapping.Application)   | 获取一个代表 CustomXMLPrefixMapping 对象的容器应用程序的 Application 对象。只读。  |
|  2   | [Creator](#CustomXMLPrefixMapping.Creator)           | 获取一个 32 位整数，指示创建 CustomXMLPrefixMapping 对象时所使用的应用程序。只读。 |
|  3   | [NamespaceURI](#CustomXMLPrefixMapping.NamespaceURI) | 获取 CustomXMLPrefixMapping 对象的命名空间的唯一地址标识符。只读。                 |
|  4   | [Parent](#CustomXMLPrefixMapping.Parent)             | 获取 CustomXMLPrefixMapping 对象的 Parent 对象。只读。                             |
|  5   | [Prefix](#CustomXMLPrefixMapping.Prefix)             | 获取 CustomXMLPrefixMapping 对象的前缀。只读。                                     |

## 成员属性

### CustomXMLPrefixMapping.Application

获取一个代表 **CustomXMLPrefixMapping** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLPrefixMapping 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMapping.Creator

获取一个 32 位整数，指示创建 **CustomXMLPrefixMapping** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLPrefixMapping 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMapping.NamespaceURI

获取 **CustomXMLPrefixMapping** 对象的命名空间的唯一地址标识符。只读。

**语法**

**express.NamespaceURI**

\*express是一个代表 CustomXMLPrefixMapping 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMapping.Parent

获取 **CustomXMLPrefixMapping** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLPrefixMapping 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMapping.Prefix

获取 **CustomXMLPrefixMapping** 对象的前缀。只读。

**语法**

**express.Prefix**

\*express是一个代表 CustomXMLPrefixMapping 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLPrefixMapping](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomXMLValidationErrors 对象

## 简介

代表 CustomXMLValidationError 对象的集合。

## 方法

| 序号 | 名称                                  | 说明                                                                                       |
|:----:|:--------------------------------------|:-------------------------------------------------------------------------------------------|
|  1   | [Add](#CustomXMLValidationErrors.Add) | 将包含 XML 验证错误的 CustomXMLValidationError 对象添加到 CustomXMLValidationErrors 集合。 |

## 属性

| 序号 | 名称                                                  | 说明                                                                                  |
|:----:|:------------------------------------------------------|:--------------------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLValidationErrors.Application) | 获取一个代表 CustomXMLValidationErrors 对象的容器应用程序的 Application 对象。只读。  |
|  2   | [Count](#CustomXMLValidationErrors.Count)             | 获取一个 Long 类型的值，指示 CustomXMLValidationErrors 集合中的项数。只读。           |
|  3   | [Creator](#CustomXMLValidationErrors.Creator)         | 获取一个 32 位整数，指示创建 CustomXMLValidationErrors 对象时所使用的应用程序。只读。 |
|  4   | [Item](#CustomXMLValidationErrors.Item)               | 获取 CustomXMLValidationErrors 集合中的一个 CustomXMLValidationError 对象。只读。     |
|  5   | [Parent](#CustomXMLValidationErrors.Parent)           | 获取 CustomXMLValidationErrors 对象的 Parent 对象。只读。                             |

## 成员方法

### CustomXMLValidationErrors.Add

将包含 XML 验证错误的 **CustomXMLValidationError** 对象添加到 **CustomXMLValidationErrors** 集合。

**语法**

**express.Add ( *Node* , *ErrorName* , *ErrorText* , *ClearedOnUpdate* )**

\*express是一个代表 CustomXMLValidationErrors 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型      | 说明                                                                         |
|-------------------|-----------|---------------|------------------------------------------------------------------------------|
| *Node*            | 必选      | CustomXMLNode | 代表其中发生了错误的节点。                                                   |
| *ErrorName*       | 必选      | String        | 包含错误的名称。                                                             |
| *ErrorText*       | 可选      | String        | 包含描述性错误文本。                                                         |
| *ClearedOnUpdate* | 可选      | Boolean       | 指定在更正并更新了 XML 后是否要从 CustomXMLValidationErrors 集合中清除错误。 |

**示例**

``` JavaScript
//以下示例将一个错误消息添加到集合。
function test()
{
    let coreNode = Application.ActivePresentation.CustomXMLParts.Item(1).SelectSingleNode("/ns1:coreProperties[1]")
    Application.ActivePresentation.CustomXMLParts.Item(1).Errors.Add(coreNode , "ValidationError", "To add content to this stream, it must be valid, well-formed XML.", true)
}
```

## 成员属性

### CustomXMLValidationErrors.Application

获取一个代表 **CustomXMLValidationErrors** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLValidationErrors 对象的变量。

### CustomXMLValidationErrors.Count

获取一个 **Long** 类型的值，指示 **CustomXMLValidationErrors** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 CustomXMLValidationErrors 对象的变量。

### CustomXMLValidationErrors.Creator

获取一个 32 位整数，指示创建 **CustomXMLValidationErrors** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLValidationErrors 对象的变量。

### CustomXMLValidationErrors.Item

获取 **CustomXMLValidationErrors** 集合中的一个 **CustomXMLValidationError** 对象。只读。

**语法**

**express.Item**

\*express是一个代表 CustomXMLValidationErrors 对象的变量。

### CustomXMLValidationErrors.Parent

获取 **CustomXMLValidationErrors** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLValidationErrors 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLValidationErrors](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ODSOFilters 对象

## 简介

代表应用于附加到邮件合并发布内容的数据源的所有筛选器。 ODSOFilters 对象由 ODSOFilter 对象组成。

## 说明

## 方法

| 序号 | 名称                          | 说明                                            |
|:----:|:------------------------------|:------------------------------------------------|
|  1   | [Add](#ODSOFilters.Add)       | 在 ODSOFilters 集合中新添一个筛选。             |
|  2   | [Delete](#ODSOFilters.Delete) | 从 ODSOFilters 集合中删除筛选对象。             |
|  3   | [Item](#ODSOFilters.Item)     | 代表 ODSOFilters 集合中的一个 ODSOFilter 对象。 |
|  4   | [Parent](#ODSOFilters.Parent) | 获取 ODSOFilters 对象的 Parent 对象。只读。     |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                               |
|:----:|:----------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ODSOFilters.Application) | 获取一个 Application 对象，代表 ODSOFilters 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Count](#ODSOFilters.Count)             | 获取一个 Long 类型的值，指示 ODSOFilters 集合中的项数。只读。                                                                      |
|  3   | [Creator](#ODSOFilters.Creator)         | 获取一个 32 位整数，指示创建 ODSOFilters 对象时所使用的应用程序。只读。                                                            |

## 成员方法

### ODSOFilters.Add

在 **ODSOFilters** 集合中新添一个筛选。

**语法**

**express.Add ( *Column* , *Comparison* , *Conjunction* , *bstrCompareTo* , *DeferUpdate* )**

\*express是一个代表 ODSOFilters 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型             | 说明                                                                                                                   |
|-----------------|-----------|----------------------|------------------------------------------------------------------------------------------------------------------------|
| *Column*        | 必选      | String               | 数据源中的表名称。                                                                                                     |
| *Comparison*    | 必选      | MsoFilterComparison  | 筛选表中数据的方式。                                                                                                   |
| *Conjunction*   | 必选      | MsoFilterConjunction | 确定此筛选与 ODSOFilters 对象中其他筛选的关联方式。                                                                    |
| *bstrCompareTo* | 可选      | String               | 如果 Comparison 参数不是 msoFilterComparisonIsBlank 或 msoFilterComparisonIsNotBlank，则表中的数据与该字符串进行比较。 |
| *DeferUpdate*   | 可选      | Boolean              | 指定是否延迟对筛选器的更新。默认值为 False。                                                                           |

### ODSOFilters.Delete

从 **ODSOFilters** 集合中删除筛选对象。

**语法**

**express.Delete ( *Index* , *DeferUpdate* )**

\*express是一个代表 ODSOFilters 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                         |
|---------------|-----------|----------|----------------------------------------------|
| *Index*       | 必选      | Long     | 要删除的筛选的编号。                         |
| *DeferUpdate* | 可选      | Boolean  | 指定是否延迟对筛选器的更新。默认值为 False。 |

### ODSOFilters.Item

代表 **ODSOFilters** 集合中的一个 **ODSOFilter** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 ODSOFilters 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明         |
|---------|-----------|----------|--------------|
| *Index* | 必选      | Long     | 项目的编号。 |

**示例**

### ODSOFilters.Parent

获取 **ODSOFilters** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent ()**

\*express是一个代表 ODSOFilters 对象的变量。

## 成员属性

### ODSOFilters.Application

获取一个 **Application** 对象，代表 **ODSOFilters** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 ODSOFilters 对象的变量。

### ODSOFilters.Count

获取一个 **Long** 类型的值，指示 **ODSOFilters** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 ODSOFilters 对象的变量。

### ODSOFilters.Creator

获取一个 32 位整数，指示创建 **ODSOFilters** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 ODSOFilters 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/ODSOFilters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PickerProperties 对象

## 简介

PickerProperty 对象的集合。

## 说明

每个 PickerProperty 对象都是一个名称 (ID)/值对，用于将选项值传递给 PickerDialog 对象。可以通过 PickerDialog 对象的 Properties 属性获取 PickerProperties 集合对象。

``` JavaScript
//下面的代码设置选取器对话框的属性，然后显示选取器对话框。
function test()
{
   // Configure the Picker Dialog properties.
   let objPickerDialog = Application.PickerDialog
   objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
   objPickerDialog.Title = "Sample Picker Dialog"
   let objPickerProperties = objPickerDialog.Properties
   let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)
   let objPickerExistingResults = objPickerDialog.CreatePickerResults()
   let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")
    
   // Show the Picker Dialog and get the results.
   let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult)
}
```

## 方法

| 序号 | 名称                               | 说明 |
|:----:|:-----------------------------------|:-----|
|  1   | [Add](#PickerProperties.Add)       |      |
|  2   | [Remove](#PickerProperties.Remove) |      |

## 属性

| 序号 | 名称                                         | 说明                                                                             |
|:----:|:---------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [Application](#PickerProperties.Application) | 获取一个 Application 对象，该对象代表 PickerProperties 对象的容器应用程序。只读  |
|  2   | [Count](#PickerProperties.Count)             | 检索包含在 PickerProperties 集合中的 PickerProperty 对象数的计数。只读           |
|  3   | [Creator](#PickerProperties.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 PickerProperties 对象的应用程序。只读 |
|  4   | [Item](#PickerProperties.Item)               | 检索位于指定索引处的 PickerProperty 对象。只读                                   |

## 成员方法

### PickerProperties.Add

**语法**

**express.Add ( *Id* , *Value* , *Type* )**

\*express是一个代表 PickerProperties 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型       | 说明           |
|---------|-----------|----------------|----------------|
| *Id*    | 必选      | String         | 属性的项名称。 |
| *Value* | 必选      | Boolean        | 属性的值。     |
| *Type*  | 必选      | MsoPickerField | 属性的类型。   |

**返回值**

PickerProperty

**示例**

``` JavaScript
//下面的代码设置 PickerDialog 对象的各种属性。
function test()
{
   // Configure the Picker Dialog properties.
   let objPickerDialog = Application.PickerDialog
   objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
   objPickerDialog.Title = "Sample Picker Dialog"
   let objPickerProperties = objPickerDialog.Properties
   let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldtypeText)
}
```

### PickerProperties.Remove

**语法**

**express.Remove ( *Id* )**

\*express是一个代表 PickerProperties 对象的变量。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                               |
|------|-----------|----------|------------------------------------|
| *Id* | 必选      | String   | 要删除的 PickerProperty 的标识符。 |

## 成员属性

### PickerProperties.Application

获取一个 **Application** 对象，该对象代表 **PickerProperties** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 PickerProperties 对象的变量。

### PickerProperties.Count

检索包含在 **PickerProperties** 集合中的 **PickerProperty** 对象数的计数。只读

**语法**

**express.Count**

\*express是一个代表 PickerProperties 对象的变量。

### PickerProperties.Creator

获取一个 32 位整数，该整数指示在其中创建了 **PickerProperties** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 PickerProperties 对象的变量。

### PickerProperties.Item

检索位于指定索引处的 **PickerProperty** 对象。只读

**语法**

**express.Item**

\*express是一个代表 PickerProperties 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/PickerProperties](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SmartArtLayout 对象

## 简介

代表 Smart Art 图表。

## 说明

选项包括基本列表、图片题注列表、垂直项目符号列表等。

``` JavaScript
/*下面的代码更改 WPP 中的 Smart Art 图表的图表样式。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Layout = Application.SmartArtLayouts.Item(1)
```

## 属性

| 序号 | 名称                                       | 说明                                                                             |
|:----:|:-------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtLayout.Application) | 获取一个 Application 对象，该对象代表 SmartArtLayout 对象的容器应用程序。只读。  |
|  2   | [Category](#SmartArtLayout.Category)       | 检索与 SmartArt 布局关联的主类别名称。只读。                                     |
|  3   | [Creator](#SmartArtLayout.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtLayout 对象的应用程序。只读。 |
|  4   | [Description](#SmartArtLayout.Description) | 检索 SmartArt 布局的说明。只读。                                                 |
|  5   | [Id](#SmartArtLayout.Id)                   | 检索关联的 SmartArt 布局的唯一 Id。只读。                                        |
|  6   | [Name](#SmartArtLayout.Name)               | 检索 SmartArt 布局的字符串名称。只读。                                           |
|  7   | [Parent](#SmartArtLayout.Parent)           | 返回调用对象。只读。                                                             |

## 成员属性

### SmartArtLayout.Application

获取一个 **Application** 对象，该对象代表 **SmartArtLayout** 对象的容器应用程序。只读。

**语法**

**express.Application**

\*express是一个代表 SmartArtLayout 对象的变量。

### SmartArtLayout.Category

检索与 SmartArt 布局关联的主类别名称。只读。

**语法**

**express.Category**

\*express是一个代表 SmartArtLayout 对象的变量。

### SmartArtLayout.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtLayout** 对象的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SmartArtLayout 对象的变量。

### SmartArtLayout.Description

检索 SmartArt 布局的说明。只读。

**语法**

**express.Description**

\*express是一个代表 SmartArtLayout 对象的变量。

### SmartArtLayout.Id

检索关联的 SmartArt 布局的唯一 Id。只读。

**语法**

**express.Id**

\*express是一个代表 SmartArtLayout 对象的变量。

**说明**

与此属性关联的 ID 区分大小写。

### SmartArtLayout.Name

检索 SmartArt 布局的字符串名称。只读。

**语法**

**express.Name**

\*express是一个代表 SmartArtLayout 对象的变量。

### SmartArtLayout.Parent

返回调用对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SmartArtLayout 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtLayout](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
