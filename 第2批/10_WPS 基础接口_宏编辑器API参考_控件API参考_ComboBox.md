# ComboBox 对象

## 简介

复合框由一个文本输入控件和一个下拉菜单组成，用户可以从预先定义的列表中的选择一个选项，同时也可以直接通过文本框输入

## 说明

复合框由一个文本输入控件和一个下拉菜单组成，用户可以从预先定义的列表中的选择一个选项，同时也可以直接通过文本框输入

## 方法

| 序号 | 名称                               | 说明                                                 |
|:----:|:-----------------------------------|:-----------------------------------------------------|
|  1   | [AddItem](#ComboBox.AddItem)       | 将新项目添加到指定组合框控件显示的值列表中。         |
|  2   | [Clear](#ComboBox.Clear)           | 清除列表框中的所有选项                               |
|  3   | [DropDown](#ComboBox.DropDown)     | 您可以使用 Dropdown 方法强制指定组合框中的列表下拉。 |
|  4   | [Move](#ComboBox.Move)             | 将对象移动到参数指定的位置                           |
|  5   | [RemoveItem](#ComboBox.RemoveItem) | 从指定组合框控件显示的值列表中删除项目。             |

## 属性

| 序号 | 名称                                     | 说明                                                                      |
|:----:|:-----------------------------------------|:--------------------------------------------------------------------------|
|  1   | [AutoSize](#ComboBox.AutoSize)           | 指定该对象是否根据内容自动设置它的大小                                    |
|  2   | [BackColor](#ComboBox.BackColor)         | 获取或设置指定对象的内部颜色                                              |
|  3   | [BackStyle](#ComboBox.BackStyle)         | 您可以使用该属性指定控件是否透明                                          |
|  4   | [BorderColor](#ComboBox.BorderColor)     | 您可以通过该属性指定控件边框的颜色                                        |
|  5   | [BorderStyle](#ComboBox.BorderStyle)     | 指定控件边框的显示方式                                                    |
|  6   | [BoundColumn](#ComboBox.BoundColumn)     | 您可以使用该属性指定是否将列表框里第一个选项的文本显示出来                |
|  7   | [ColumnCount](#ComboBox.ColumnCount)     | 您可以使用 ColumnCount 属性来指定在列表框或组合框的列表框部分中显示的列数 |
|  8   | [ColumnHeads](#ComboBox.ColumnHeads)     | 您可以使用 ColumnHeads 属性指定列表框是否显示单行列标题。                 |
|  9   | [ColumnWidths](#ComboBox.ColumnWidths)   | 您可以使用 ColumnWidths 属性来指定多列组合框中每列的宽度。                |
|  10  | [ControlSource](#ComboBox.ControlSource) | 您可以使用该属性指定控件中显示的数据。                                    |
|  11  | [Enabled](#ComboBox.Enabled)             | 您可以通过该属性设置或返回是否启用该控件                                  |
|  12  | [Font](#ComboBox.Font)                   | 指定对象的字体，只读                                                      |
|  13  | [ForeColor](#ComboBox.ForeColor)         | 您可以通过该属性指定控件中文本的颜色                                      |
|  14  | [Height](#ComboBox.Height)               | 获取或设置指定对象的高度                                                  |
|  15  | [HeightPolicy](#ComboBox.HeightPolicy)   | 获取或设置指定对像的高度策略                                              |
|  16  | [Left](#ComboBox.Left)                   | 您可以通过该属性指定控件在对象或报表中的位置                              |
|  17  | [ListRows](#ComboBox.ListRows)           | 您可以使用 ListRows 属性设置在组合框的列表框部分中显示的最大行数。        |
|  18  | [ListWidth](#ComboBox.ListWidth)         | 您可以使用 ListWidth 属性设置组合框的列表框部分的宽度。                   |
|  19  | [MaxLength](#ComboBox.MaxLength)         | 指定用户可以在文本框或组合框中输入的最大字符数                            |
|  20  | [Name](#ComboBox.Name)                   | 您可以通过该属性指定或确定表示对象名称的字符串，只读                      |
|  21  | [RowSource](#ComboBox.RowSource)         | 您可以使用 RowSource 属性指定如何向指定的对象提供数据。                   |
|  22  | [TabIndex](#ComboBox.TabIndex)           | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置                 |
|  23  | [TabStop](#ComboBox.TabStop)             | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上                     |
|  24  | [Text](#ComboBox.Text)                   | 您可以通过该属性设置或返回文本框中包含的文本值                            |
|  25  | [TextAlign](#ComboBox.TextAlign)         | 该属性用于指定对象中文本的对齐方式                                        |
|  26  | [Top](#ComboBox.Top)                     | 您可以通过该属性指定控件在对象或报表中的位置                              |
|  27  | [TopIndex](#ComboBox.TopIndex)           | 设置并返回出现在列表最顶部位置的项目。                                    |
|  28  | [Value](#ComboBox.Value)                 | 确定或指定文本框中的文本                                                  |
|  29  | [Visible](#ComboBox.Visible)             | 返回或设置该对象是否可见                                                  |
|  30  | [Width](#ComboBox.Width)                 | 获取或设置指定对象的宽度                                                  |
|  31  | [WidthPolicy](#ComboBox.WidthPolicy)     | 获取或设置指定对象的宽度策略                                              |

## 成员方法

### ComboBox.AddItem

将新项目添加到指定组合框控件显示的值列表中。

**语法**

**express.AddItem ( *Content* , *Index* )**

\*express是一个代表 ComboBox 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                   |
|-----------|-----------|----------|------------------------|
| *Content* | 必选      | String   | 列表框里显示的文本内容 |
| *Index*   | 必选      | Number   | 该选项在列表框里的位置 |

**说明**

将新项目添加到指定组合框控件显示的值列表中。

**示例**

``` JavaScript
/*将test0设置为ComboBox1控件列表框中的第一个选项*/
function func1()
{
    UserForm1.ComboBox1.AddItem("test0",0);
}
```

### ComboBox.Clear

清除列表框中的所有选项

**语法**

**express.Clear ()**

\*express是一个代表 ComboBox 对象的变量。

**说明**

清除列表框中的所有选项

**示例**

``` JavaScript
/*清除ComboBox1中列表框中的所有选项*/
function func1()
{
    UserForm1.ComboBox1.Clear()
}
```

### ComboBox.DropDown

您可以使用 Dropdown 方法强制指定组合框中的列表下拉。

**语法**

**express.DropDown ()**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以使用 Dropdown 方法强制指定组合框中的列表下拉。

**示例**

``` JavaScript
/*强制指定ComboBox1中的列表下拉*/
function func1()
{
    UserForm1.ComboBox1.DropDown()
}
```

### ComboBox.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 ComboBox 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                               |
|----------|-----------|----------|------------------------------------|
| *Left*   | 必选      | Number   | 对象移动后相对于窗口的左边缘的距离 |
| *Top*    | 可选      | Number   | 对象移动后相对于窗口的上边缘的距离 |
| *Width*  | 可选      | Number   | 对象移动后的预期宽度               |
| *Height* | 可选      | Number   | 对象移动后的预期长度               |

**说明**

将对象移动到参数指定的位置

**示例**

``` JavaScript
/*将ComboBox1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.ComboBox1.Move(100,100,100,100)
}
```

### ComboBox.RemoveItem

从指定组合框控件显示的值列表中删除项目。

**语法**

**express.RemoveItem ( *index* )**

\*express是一个代表 ComboBox 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                            |
|---------|-----------|----------|---------------------------------|
| *index* | 可选      | Number   | 删除列表框中index指定的下标的值 |

**说明**

从指定组合框控件显示的值列表中删除项目。

**示例**

``` JavaScript
/*删除ComboBox1中下标为3的值*/
function func1()
{
    UserForm1.ComboBox1.RemoveItem(3)
}
```

## 成员属性

### ComboBox.AutoSize

指定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 ComboBox 对象的变量。

**说明**

指定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置ComboBox1控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.ComboBox1.AutoSize = true
}
```

### ComboBox.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 ComboBox 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将ComboBox1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.ComboBox1.BackColor = 0x665898
}
```

### ComboBox.BackStyle

您可以使用该属性指定控件是否透明

**语法**

**express.BackStyle**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以使用该属性指定控件是否透明

**示例**

``` JavaScript
/*将ComboBox1设置为透明*/
function func1()
{
    UserForm1.ComboBox1.BackStyle = 0
}
```

### ComboBox.BorderColor

您可以通过该属性指定控件边框的颜色

**语法**

**express.BorderColor**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以通过该属性指定控件边框的颜色

**示例**

``` JavaScript
/*将ComboBox1的边框设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.ComboBox1.BorderColor = 0x665898
}
```

### ComboBox.BorderStyle

指定控件边框的显示方式

**语法**

**express.BorderStyle**

\*express是一个代表 ComboBox 对象的变量。

**说明**

指定控件边框的显示方式，可以设置为以下值：

| 设置   | 值  | 描述     |
|--------|-----|----------|
| None   | 0   | 无边框   |
| Single | 1   | 实线边框 |

**示例**

``` JavaScript
/*将ComboBox1的边框显示方式设置为实线*/
function func1()
{
    UserForm1.ComboBox1.BorderStyle = 1
}
```

### ComboBox.BoundColumn

您可以使用该属性指定是否将列表框里第一个选项的文本显示出来

**语法**

**express.BoundColumn**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以使用该属性指定是否将列表框里第一个选项的文本显示出来，可以设置为以下值：

| 设置   | 值  | 描述                   |
|--------|-----|------------------------|
| 显示   | 0   | 显示第一个选项的文本   |
| 不显示 | 1   | 不显示第一个选项的文本 |

**示例**

``` JavaScript
/*显示ComboBox1列表框里第一个选项的值*/
function func1()
{
    UserForm1.ComboBox1.BoundColumn = 0
}
```

### ComboBox.ColumnCount

您可以使用 ColumnCount 属性来指定在列表框或组合框的列表框部分中显示的列数

**语法**

**express.ColumnCount**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以使用 ColumnCount 属性来指定在列表框或组合框的列表框部分中显示的列数

**示例**

``` JavaScript
/*将ComboBox1的列表框的列数设置为3*/
function func1()
{
    UserForm1.ComboBox1.ColumnCount = 3
}
```

### ComboBox.ColumnHeads

您可以使用 ColumnHeads 属性指定列表框是否显示单行列标题。

**语法**

**express.ColumnHeads**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以使用 ColumnHeads 属性指定列表框是否显示单行列标题。可以设置为以下值：

| 设置 | 值    | 描述         |
|------|-------|--------------|
| Yes  | true  | 显示列标题   |
| No   | false | 不显示列标题 |

**示例**

``` JavaScript
/*不显示ComboBox1列表框的列标题*/
function func1()
{
    UserForm1.ComboBox1.ColumnHeads = false
}
```

### ComboBox.ColumnWidths

您可以使用 ColumnWidths 属性来指定多列组合框中每列的宽度。

**语法**

**express.ColumnWidths**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以使用 ColumnWidths 属性来指定多列组合框中每列的宽度。

**示例**

``` JavaScript
/*设置ComboBox1列表框中每列的宽度为5*/
function func1()
{
    UserForm1.ComboBox1.ColumnWidths = 5
}
```

### ComboBox.ControlSource

您可以使用该属性指定控件中显示的数据。

**语法**

**express.ControlSource**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以使用该属性指定控件中显示的数据。您可以显示和编辑绑定到表的数据、查询的数据或SQL语句中的字段的数据。您还可以显示表达式的结果

### ComboBox.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.ComboBox1.Enabled = true
}
```

### ComboBox.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 ComboBox 对象的变量。

**说明**

指定对象的字体，只读

### ComboBox.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将ComboBox1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.ComboBox1.ForeColor = 0x658978
}
```

### ComboBox.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 ComboBox 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将ComboBox1控件的高度设置为100*/
function func1()
{
    UserForm1.ComboBox1.Height = 100
}
```

### ComboBox.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 ComboBox 对象的变量。

**说明**

获取或设置指定对像的高度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将ComboBox1的高度设置为固定*/
function func1()
{
    UserForm1.ComboBox1.HeightPolicy = 2
}
```

### ComboBox.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ComboBox1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.ComboBox1.Left = 100
}
```

### ComboBox.ListRows

您可以使用 ListRows 属性设置在组合框的列表框部分中显示的最大行数。

**语法**

**express.ListRows**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以使用 ListRows 属性设置在组合框的列表框部分中显示的最大行数。

**示例**

``` JavaScript
/*将ComboBox1中的列表框的显示的最大行数设为8*/
function func1()
{
    UserForm1.ComboBox1.ListRows = 8
}
```

### ComboBox.ListWidth

您可以使用 ListWidth 属性设置组合框的列表框部分的宽度。

**语法**

**express.ListWidth**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以使用 ListWidth 属性设置组合框的列表框部分的宽度。

**示例**

``` JavaScript
/*将ComboBox1的列表框部分的宽度设为100*/
function func1()
{
    UserForm1.ComboBox1.ListWidth = 100
}
```

### ComboBox.MaxLength

指定用户可以在文本框或组合框中输入的最大字符数

**语法**

**express.MaxLength**

\*express是一个代表 ComboBox 对象的变量。

**说明**

指定用户可以在文本框或组合框中输入的最大字符数

**示例**

``` JavaScript
/*将ComboBox1控件中允许输入的最大字符数设置为32767*/
function func1()
{
    UserForm1.ComboBox1.MaxLength = 32767
}
```

### ComboBox.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### ComboBox.RowSource

您可以使用 RowSource 属性指定如何向指定的对象提供数据。

**语法**

**express.RowSource**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以使用 RowSource 属性指定如何向指定的对象提供数据。

### ComboBox.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将ComboBox1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.ComboBox1.TabIndex = 2
}
```

### ComboBox.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将ComboBox1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.ComboBox1.TabStop = true
}
```

### ComboBox.Text

您可以通过该属性设置或返回文本框中包含的文本值

**语法**

**express.Text**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以通过该属性设置或返回文本框中包含的文本值

**示例**

``` JavaScript
/*将ComboBox1控件中的文本设置为test test*/
function func1()
{
    UserForm1.ComboBox1.Text = "test test"
}
```

### ComboBox.TextAlign

该属性用于指定对象中文本的对齐方式

**语法**

**express.TextAlign**

\*express是一个代表 ComboBox 对象的变量。

**说明**

该属性用于指定对象中文本的对齐方式，可设置为以下值：

**示例**

``` JavaScript
/*将ComboBox1中的文本设置为居中对齐*/
function func1()
{
    UserForm1.ComboBox1.TextAlign = 2
}
```

### ComboBox.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 ComboBox 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ComboBox1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.ComboBox1.Top = 100
}
```

### ComboBox.TopIndex

设置并返回出现在列表最顶部位置的项目。

**语法**

**express.TopIndex**

\*express是一个代表 ComboBox 对象的变量。

**说明**

设置并返回出现在列表最顶部位置的项目。

**示例**

``` JavaScript
/*将ComboBox1的下标为3的选项设置到列表框的顶部*/
function func1()
{
    UserForm1.ComboBox1.TopIndex = 3
}
```

### ComboBox.Value

确定或指定文本框中的文本

**语法**

**express.Value**

\*express是一个代表 ComboBox 对象的变量。

**说明**

确定或指定文本框中的文本

**示例**

``` JavaScript
/*将ComboBox1控件中的文本指定为test test*/
function func1()
{
    UserForm1.ComboBox1.Value = "test test"
}
```

### ComboBox.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 ComboBox 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将ComboBox1设置为可见*/
function func1()
{
    UserForm1.ComboBox1.Visible = true
}
```

### ComboBox.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 ComboBox 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将ComboBox1的宽度设置为100*/
function func1()
{
    UserForm1.ComboBox1.Width = 100
}
```

### ComboBox.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 ComboBox 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将ComboBox1的宽度设置为固定*/
function func1()
{
    UserForm1.ComboBox1.WidthPolicy = 2
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/ComboBox](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
