# ListBox 对象

## 简介

用来显示用户可以选择的项目列表。如果不能一次显示全部项目的话，列表可以滚动。

## 说明

用来显示用户可以选择的项目列表。如果不能一次显示全部项目的话，列表可以滚动。

## 方法

| 序号 | 名称                              | 说明                                         |
|:----:|:----------------------------------|:---------------------------------------------|
|  1   | [AddItem](#ListBox.AddItem)       | 将新项目添加到指定组合框控件显示的值列表中。 |
|  2   | [Clear](#ListBox.Clear)           | 清除列表框中的所有选项                       |
|  3   | [Move](#ListBox.Move)             | 将对象移动到参数指定的位置                   |
|  4   | [RemoveItem](#ListBox.RemoveItem) | 从指定组合框控件显示的值列表中删除项目       |

## 属性

| 序号 | 名称                                    | 说明                                                                      |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------|
|  1   | [BackColor](#ListBox.BackColor)         | 获取或设置指定对象的内部颜色                                              |
|  2   | [BorderColor](#ListBox.BorderColor)     | 您可以通过该属性指定控件边框的颜色                                        |
|  3   | [BorderStyle](#ListBox.BorderStyle)     | 指定控件边框的显示方式                                                    |
|  4   | [BoundColumn](#ListBox.BoundColumn)     | 您可以使用该属性指定是否将列表框里第一个选项的文本显示出来                |
|  5   | [ColumnCount](#ListBox.ColumnCount)     | 您可以使用 ColumnCount 属性来指定在列表框或组合框的列表框部分中显示的列数 |
|  6   | [ColumnHeads](#ListBox.ColumnHeads)     | 您可以使用 ColumnHeads 属性指定列表框是否显示单行列标题                   |
|  7   | [ColumnWidths](#ListBox.ColumnWidths)   | 您可以使用 ColumnWidths 属性来指定多列组合框中每列的宽度                  |
|  8   | [ControlSource](#ListBox.ControlSource) | 您可以使用该属性指定控件中显示的数据                                      |
|  9   | [Enabled](#ListBox.Enabled)             | 您可以通过该属性设置或返回是否启用该控件                                  |
|  10  | [Font](#ListBox.Font)                   | 指定对象的字体，只读                                                      |
|  11  | [ForeColor](#ListBox.ForeColor)         | 您可以通过该属性指定控件中文本的颜色                                      |
|  12  | [Height](#ListBox.Height)               | 获取或设置指定对象的高度                                                  |
|  13  | [HeightPolicy](#ListBox.HeightPolicy)   | 获取或设置指定对像的高度策略                                              |
|  14  | [Left](#ListBox.Left)                   | 您可以通过该属性指定控件在对象或报表中的位置                              |
|  15  | [Name](#ListBox.Name)                   | 您可以通过该属性指定或确定表示对象名称的字符串，只读                      |
|  16  | [RowSource](#ListBox.RowSource)         | 您可以使用 RowSource 属性指定如何向指定的对象提供数据                     |
|  17  | [TabIndex](#ListBox.TabIndex)           | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置                 |
|  18  | [TabStop](#ListBox.TabStop)             | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上                     |
|  19  | [Text](#ListBox.Text)                   | 您可以通过该属性设置或返回文本框中包含的文本值                            |
|  20  | [TextAlign](#ListBox.TextAlign)         | 该属性用于指定对象中文本的对齐方式                                        |
|  21  | [Top](#ListBox.Top)                     | 您可以通过该属性指定控件在对象或报表中的位置                              |
|  22  | [TopIndex](#ListBox.TopIndex)           | 设置并返回出现在列表最顶部位置的项目。                                    |
|  23  | [Value](#ListBox.Value)                 | 确定或指定列表框中的哪个值或选项被选中。                                  |
|  24  | [Visible](#ListBox.Visible)             | 返回或设置该对象是否可见                                                  |
|  25  | [Width](#ListBox.Width)                 | 获取或设置指定对象的宽度                                                  |
|  26  | [WidthPolicy](#ListBox.WidthPolicy)     | 获取或设置指定对象的宽度策略                                              |

## 成员方法

### ListBox.AddItem

将新项目添加到指定组合框控件显示的值列表中。

**语法**

**express.AddItem ( *Content* , *Index* )**

\*express是一个代表 ListBox 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                   |
|-----------|-----------|----------|------------------------|
| *Content* | 必选      | String   | 列表框里显示的文本内容 |
| *Index*   | 必选      | Number   | 该选项在列表框里的位置 |

**说明**

将新项目添加到指定组合框控件显示的值列表中。

**示例**

``` JavaScript
/*将test0设置为ListBox1控件列表框中的第一个选项*/
function func1()
{
    UserForm1.ListBox1.AddItem("test0",0);
}
```

### ListBox.Clear

清除列表框中的所有选项

**语法**

**express.Clear ()**

\*express是一个代表 ListBox 对象的变量。

**说明**

清除列表框中的所有选项

**示例**

``` JavaScript
/*清除ListBox1中列表框中的所有选项*/
function func1()
{
    UserForm1.ListBox1.Clear()
}
```

### ListBox.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 ListBox 对象的变量。

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
/*将ListBox1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.ListBox1.Move(100,100,100,100)
}
```

### ListBox.RemoveItem

从指定组合框控件显示的值列表中删除项目

**语法**

**express.RemoveItem ( *Index* )**

\*express是一个代表 ListBox 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                            |
|---------|-----------|----------|---------------------------------|
| *Index* | 必选      | Number   | 删除列表框中index指定的下标的值 |

**说明**

从指定组合框控件显示的值列表中删除项目

**示例**

``` JavaScript
/*删除ListBox1中下标为3的值*/
function func1()
{
    UserForm1.ListBox1.RemoveItem(3)
}
```

## 成员属性

### ListBox.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 ListBox 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将ListBox1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.ListBox1.BackColor = 0x665898
}
```

### ListBox.BorderColor

您可以通过该属性指定控件边框的颜色

**语法**

**express.BorderColor**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以通过该属性指定控件边框的颜色

**示例**

``` JavaScript
/*将ListBox1的边框设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.ListBox1.BorderColor = 0x665898
}
```

### ListBox.BorderStyle

指定控件边框的显示方式

**语法**

**express.BorderStyle**

\*express是一个代表 ListBox 对象的变量。

**说明**

指定控件边框的显示方式，可以设置为以下值：

| 设置   | 值  | 描述     |
|--------|-----|----------|
| None   | 0   | 无边框   |
| Single | 1   | 实线边框 |

**示例**

``` JavaScript
/*将ListBox1的边框显示方式设置为实线*/
function func1()
{
    UserForm1.ListBox1.BorderStyle = 1
}
```

### ListBox.BoundColumn

您可以使用该属性指定是否将列表框里第一个选项的文本显示出来

**语法**

**express.BoundColumn**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以使用该属性指定是否将列表框里第一个选项的文本显示出来，可以设置为以下值：

| 设置   | 值  | 描述                 |
|--------|-----|----------------------|
| 显示   | 0   | 显示第一个选项的值   |
| 不显示 | 1   | 不显示第一个选项的值 |

**示例**

``` JavaScript
/*显示ListBox1列表框里第一个选项的值*/
function func1()
{
    UserForm1.ListBox1.BoundColumn = 0
}
```

### ListBox.ColumnCount

您可以使用 ColumnCount 属性来指定在列表框或组合框的列表框部分中显示的列数

**语法**

**express.ColumnCount**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以使用 ColumnCount 属性来指定在列表框或组合框的列表框部分中显示的列数

**示例**

``` JavaScript
/*将ListBox1的列表框的列数设置为3*/
function func1()
{
    UserForm1.ListBox1.ColumnCount = 3
}
```

### ListBox.ColumnHeads

您可以使用 ColumnHeads 属性指定列表框是否显示单行列标题

**语法**

**express.ColumnHeads**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以使用 ColumnHeads 属性指定列表框是否显示单行列标题。可以设置为以下值：

| 设置 | 值    | 描述         |
|------|-------|--------------|
| Yes  | true  | 显示列标题   |
| No   | false | 不显示列标题 |

**示例**

``` JavaScript
/*不显示ListBox1列表框的列标题*/
function func1()
{
    UserForm1.ListBox1.ColumnHeads = false
}
```

### ListBox.ColumnWidths

您可以使用 ColumnWidths 属性来指定多列组合框中每列的宽度

**语法**

**express.ColumnWidths**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以使用 ColumnWidths 属性来指定多列组合框中每列的宽度

**示例**

``` JavaScript
/*设置ListBox1列表框中每列的宽度为5*/
function func1()
{
    UserForm1.ListBox1.ColumnWidths = 5
}
```

### ListBox.ControlSource

您可以使用该属性指定控件中显示的数据

**语法**

**express.ControlSource**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以使用该属性指定控件中显示的数据。您可以显示和编辑绑定到表的数据、查询的数据或SQL语句中的字段的数据。您还可以显示表达式的结果

### ListBox.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.ListBox1.Enabled = true
}
```

### ListBox.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 ListBox 对象的变量。

**说明**

指定对象的字体，只读

### ListBox.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将ListBox1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.ListBox1.ForeColor = 0x658978
}
```

### ListBox.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 ListBox 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将ListBox1控件的高度设置为100*/
function func1()
{
    UserForm1.ListBox1.Height = 100
}
```

### ListBox.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 ListBox 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将ListBox1的高度设置为固定*/
function func1()
{
    UserForm1.ListBox1.HeightPolicy = 2
}
```

### ListBox.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ListBox1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.ListBox1.Left = 100
}
```

### ListBox.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### ListBox.RowSource

您可以使用 RowSource 属性指定如何向指定的对象提供数据

**语法**

**express.RowSource**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以使用 RowSource 属性指定如何向指定的对象提供数据

### ListBox.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将ListBox1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.ListBox1.TabIndex = 2
}
```

### ListBox.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将ListBox1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.ListBox1.TabStop = true
}
```

### ListBox.Text

您可以通过该属性设置或返回文本框中包含的文本值

**语法**

**express.Text**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以通过该属性设置或返回文本框中包含的文本值

**示例**

``` JavaScript
/*将ListBox1控件中的文本设置为test test*/
function func1()
{
    UserForm1.ListBox1.Text = "test test"
}
```

### ListBox.TextAlign

该属性用于指定对象中文本的对齐方式

**语法**

**express.TextAlign**

\*express是一个代表 ListBox 对象的变量。

**说明**

该属性用于指定对象中文本的对齐方式，包含以下值：

| 设置   | 值  | 描述                     |
|--------|-----|--------------------------|
| Left   | 1   | 文本数字和日期都左对齐   |
| Center | 2   | 文本数字和日期都居中对齐 |
| Right  | 3   | 文本数字和日期都右对齐   |

**示例**

``` JavaScript
/*将ListBox1中的文本设置为居中对齐*/
function func1()
{
    UserForm1.ListBox1.TextAlign = 2
}
```

### ListBox.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 ListBox 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ListBox1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.ListBox1.Top = 100
}
```

### ListBox.TopIndex

设置并返回出现在列表最顶部位置的项目。

**语法**

**express.TopIndex**

\*express是一个代表 ListBox 对象的变量。

**说明**

设置并返回出现在列表最顶部位置的项目。

**示例**

``` JavaScript
/*将ListBox1的下标为3的选项设置到列表框的顶部*/
function func1()
{
    UserForm1.ListBox1.TopIndex = 3
}
```

### ListBox.Value

确定或指定列表框中的哪个值或选项被选中。

**语法**

**express.Value**

\*express是一个代表 ListBox 对象的变量。

**说明**

确定或指定列表框中的哪个值或选项被选中。

**示例**

``` JavaScript
/*将ListBox1控件中的选项指定为test test*/
function func1()
{
    UserForm1.ListBox1.Value = "test test"
}
```

### ListBox.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 ListBox 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将ListBox1设置为可见*/
function func1()
{
    UserForm1.ListBox1.Visible = true
}
```

### ListBox.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 ListBox 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将ListBox1的宽度设置为100*/
function func1()
{
    UserForm1.ListBox1.Width = 100
}
```

### ListBox.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 ListBox 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将ListBox1的宽度设置为固定*/
function func1()
{
    UserForm1.ListBox1.WidthPolicy = 2
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/ListBox](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
