# CheckBox 对象

## 简介

可以显示出多重选择，用户可以选择一个或者多个选项，如果选中多个选项则显示出多重选择。

## 说明

可以显示出多重选择，用户可以选择一个或者多个选项，如果选中多个选项则显示出多重选择。

## 方法

| 序号 | 名称                   | 说明                       |
|:----:|:-----------------------|:---------------------------|
|  1   | [Move](#CheckBox.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                     | 说明                                                      |
|:----:|:-----------------------------------------|:----------------------------------------------------------|
|  1   | [AutoSize](#CheckBox.AutoSize)           | 指定该对象是否根据内容自动设置它的大小                    |
|  2   | [BackColor](#CheckBox.BackColor)         | 获取或设置指定对象的内部颜色                              |
|  3   | [BackStyle](#CheckBox.BackStyle)         | 您可以使用该属性指定控件是否透明                          |
|  4   | [Caption](#CheckBox.Caption)             | 获取或设置出现在控件中的文本                              |
|  5   | [ControlSource](#CheckBox.ControlSource) | 您可以使用该属性指定控件中显示的数据                      |
|  6   | [Enabled](#CheckBox.Enabled)             | 您可以通过该属性设置或返回是否启用该控件                  |
|  7   | [Font](#CheckBox.Font)                   | 指定对象的字体，只读                                      |
|  8   | [ForeColor](#CheckBox.ForeColor)         | 您可以通过该属性指定控件中文本的颜色                      |
|  9   | [Height](#CheckBox.Height)               | 获取或设置指定对象的高度                                  |
|  10  | [HeightPolicy](#CheckBox.HeightPolicy)   | 获取或设置指定对像的高度策略                              |
|  11  | [Left](#CheckBox.Left)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  12  | [Name](#CheckBox.Name)                   | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  13  | [TabIndex](#CheckBox.TabIndex)           | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  14  | [TabStop](#CheckBox.TabStop)             | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  15  | [Top](#CheckBox.Top)                     | 您可以通过该属性指定控件在对象或报表中的位置              |
|  16  | [Value](#CheckBox.Value)                 | 确定或指定是否选中指定的复选框                            |
|  17  | [Visible](#CheckBox.Visible)             | 返回或设置该对象是否可见                                  |
|  18  | [Width](#CheckBox.Width)                 | 获取或设置指定对象的宽度                                  |
|  19  | [WidthPolicy](#CheckBox.WidthPolicy)     | 获取或设置指定对象的宽度策略                              |
|  20  | [WordWrap](#CheckBox.WordWrap)           | 指定控件的内容是否在行尾自动换行                          |

## 成员方法

### CheckBox.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 CheckBox 对象的变量。

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
/*将CheckBox1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.CheckBox1.Move(100,100,100,100)
}
```

## 成员属性

### CheckBox.AutoSize

指定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 CheckBox 对象的变量。

**说明**

指定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置CheckBox1控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.CheckBox1.AutoSize = true
}
```

### CheckBox.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 CheckBox 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将CheckBox1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.CheckBox1.BackColor = 0x665898
}
```

### CheckBox.BackStyle

您可以使用该属性指定控件是否透明

**语法**

**express.BackStyle**

\*express是一个代表 CheckBox 对象的变量。

**说明**

您可以使用该属性指定控件是否透明

**示例**

``` JavaScript
/*将CheckBox1设置为透明*/
function func1()
{
    UserForm1.CheckBox1.BackStyle = 0
}
```

### CheckBox.Caption

获取或设置出现在控件中的文本

**语法**

**express.Caption**

\*express是一个代表 CheckBox 对象的变量。

**说明**

获取或设置出现在控件中的文本

**示例**

``` JavaScript
/*将CheckBox1中的文本设置为"CheckBox1"*/
function func1()
{
    UserForm1.CheckBox1.Caption = "CheckBox1"
}
```

### CheckBox.ControlSource

您可以使用该属性指定控件中显示的数据

**语法**

**express.ControlSource**

\*express是一个代表 CheckBox 对象的变量。

**说明**

您可以使用该属性指定控件中显示的数据。您可以显示和编辑绑定到表的数据、查询的数据或SQL语句中的字段的数据。您还可以显示表达式的结果

### CheckBox.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 CheckBox 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.CheckBox1.Enabled = true
}
```

### CheckBox.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 CheckBox 对象的变量。

**说明**

指定对象的字体，只读

### CheckBox.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 CheckBox 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将CheckBox1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.CheckBox1.ForeColor = 0x658978
}
```

### CheckBox.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 CheckBox 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将CheckBox1控件的高度设置为100*/
function func1()
{
    UserForm1.CheckBox1.Height = 100
}
```

### CheckBox.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 CheckBox 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将CheckBox1的高度设置为固定*/
function func1()
{
    UserForm1.CheckBox1.HeightPolicy = 2
}
```

### CheckBox.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 CheckBox 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将CheckBox1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.CheckBox1.Left = 100
}
```

### CheckBox.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 CheckBox 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### CheckBox.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 CheckBox 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将CheckBox1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.CheckBox1.TabIndex = 2
}
```

### CheckBox.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 CheckBox 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将CheckBox1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.CheckBox1.TabStop = true
}
```

### CheckBox.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 CheckBox 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将CheckBox1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.CheckBox1.Top = 100
}
```

### CheckBox.Value

确定或指定是否选中指定的复选框

**语法**

**express.Value**

\*express是一个代表 CheckBox 对象的变量。

**说明**

确定或指定是否选中指定的复选框。设为True为选定该复选框，默认为False

**示例**

``` JavaScript
/*将CheckBox1控件中的选项指定为test test*/
function func1()
{
    UserForm1.CheckBox1.Value = "test test"
}
```

### CheckBox.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 CheckBox 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将CheckBox1设置为可见*/
function func1()
{
    UserForm1.CheckBox1.Visible = true
}
```

### CheckBox.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 CheckBox 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将CheckBox1的宽度设置为100*/
function func1()
{
    UserForm1.CheckBox1.Width = 100
}
```

### CheckBox.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 CheckBox 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 属性     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将CheckBox1的宽度设置为固定*/
function func1()
{
    UserForm1.CheckBox1.WidthPolicy = 2
}
```

### CheckBox.WordWrap

指定控件的内容是否在行尾自动换行

**语法**

**express.WordWrap**

\*express是一个代表 CheckBox 对象的变量。

**说明**

指定控件的内容是否在行尾自动换行

**示例**

``` JavaScript
/*将CheckBox1设置为控件的内容在行尾自动换行*/
function func1()
{
    UserForm1.CheckBox1.WordWrap = true
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/CheckBox](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
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
# OptionButton 对象

## 简介

可以显示出多重选择，用户只能选择一个。

## 说明

可以显示出多重选择，用户只能选择一个。

## 方法

| 序号 | 名称                       | 说明                       |
|:----:|:---------------------------|:---------------------------|
|  1   | [Move](#OptionButton.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                         | 说明                                                      |
|:----:|:---------------------------------------------|:----------------------------------------------------------|
|  1   | [AutoSize](#OptionButton.AutoSize)           | 指定该对象是否根据内容自动设置它的大小                    |
|  2   | [BackColor](#OptionButton.BackColor)         | 获取或设置指定对象的内部颜色                              |
|  3   | [BackStyle](#OptionButton.BackStyle)         | 您可以使用该属性指定控件是否透明                          |
|  4   | [Caption](#OptionButton.Caption)             | 获取或设置出现在控件中的文本                              |
|  5   | [ControlSource](#OptionButton.ControlSource) | 您可以使用该属性指定控件中显示的数据                      |
|  6   | [Enabled](#OptionButton.Enabled)             | 您可以通过该属性设置或返回是否启用该控件                  |
|  7   | [Font](#OptionButton.Font)                   | 指定对象的字体，只读                                      |
|  8   | [ForeColor](#OptionButton.ForeColor)         | 您可以通过该属性指定控件中文本的颜色                      |
|  9   | [GroupName](#OptionButton.GroupName)         | 您可以通过设置组名，创建一组互斥的OptionButton            |
|  10  | [Height](#OptionButton.Height)               | 获取或设置指定对象的高度                                  |
|  11  | [HeightPolicy](#OptionButton.HeightPolicy)   | 获取或设置指定对像的高度策略                              |
|  12  | [Left](#OptionButton.Left)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  13  | [Name](#OptionButton.Name)                   | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  14  | [TabIndex](#OptionButton.TabIndex)           | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  15  | [TabStop](#OptionButton.TabStop)             | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  16  | [Top](#OptionButton.Top)                     | 您可以通过该属性指定控件在对象或报表中的位置              |
|  17  | [Visible](#OptionButton.Visible)             | 返回或设置该对象是否可见                                  |
|  18  | [Width](#OptionButton.Width)                 | 获取或设置指定对象的宽度                                  |
|  19  | [WidthPolicy](#OptionButton.WidthPolicy)     | 获取或设置指定对象的宽度策略                              |
|  20  | [WordWrap](#OptionButton.WordWrap)           | 指定控件的内容是否在行尾自动换行                          |

## 成员方法

### OptionButton.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 OptionButton 对象的变量。

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
/*将OptionButton1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.OptionButton1.Move(100,100,100,100)
}
```

## 成员属性

### OptionButton.AutoSize

指定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 OptionButton 对象的变量。

**说明**

指定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置OptionButton1控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.OptionButton1.AutoSize = true
}
```

### OptionButton.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将OptionButton1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.OptionButton1.BackColor = 0x665898
}
```

### OptionButton.BackStyle

您可以使用该属性指定控件是否透明

**语法**

**express.BackStyle**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以使用该属性指定控件是否透明

**示例**

``` JavaScript
/*将OptionButton1设置为透明*/
function func1()
{
    UserForm1.OptionButton1.BackStyle = 0
}
```

### OptionButton.Caption

获取或设置出现在控件中的文本

**语法**

**express.Caption**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置出现在控件中的文本

**示例**

``` JavaScript
/*将OptionButton1中的文本设置为"OptionButton1"*/
function func1()
{
    UserForm1.OptionButton1.Caption = "OptionButton1"
}
```

### OptionButton.ControlSource

您可以使用该属性指定控件中显示的数据

**语法**

**express.ControlSource**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以使用该属性指定控件中显示的数据。您可以显示和编辑绑定到表的数据、查询的数据或SQL语句中的字段的数据。您还可以显示表达式的结果

### OptionButton.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.OptionButton1.Enabled = true
}
```

### OptionButton.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 OptionButton 对象的变量。

**说明**

指定对象的字体，只读

### OptionButton.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将OptionButton1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.OptionButton1.ForeColor = 0x658978
}
```

### OptionButton.GroupName

您可以通过设置组名，创建一组互斥的OptionButton

**语法**

**express.GroupName**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过设置组名，创建一组互斥的OptionButton。同一组里的OptionButton将采用相同的设置，点击组内的任何一个选项都会将其他选项置为false。不同容器下相同的组名会被视为不同组（每个容器一个组）。

**示例**

``` JavaScript
/*将OptionButton1的组名设置为group one*/
function func1()
{
    UserForm1.OptionButton1.GroupName = "group one"
}
```

### OptionButton.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将OptionButton1控件的高度设置为100*/
function func1()
{
    UserForm1.OptionButton1.Height = 100
}
```

### OptionButton.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将OptionButton1的高度设置为固定*/
function func1()
{
    UserForm1.OptionButton1.HeightPolicy = 2
}
```

### OptionButton.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将OptionButton1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.OptionButton1.Left = 100
}
```

### OptionButton.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### OptionButton.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将OptionButton1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.OptionButton1.TabIndex = 2
}
```

### OptionButton.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将OptionButton1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.OptionButton1.TabStop = true
}
```

### OptionButton.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 OptionButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将OptionButton1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.OptionButton1.Top = 100
}
```

### OptionButton.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 OptionButton 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将OptionButton1设置为可见*/
function func1()
{
    UserForm1.OptionButton1.Visible = true
}
```

### OptionButton.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将OptionButton1的宽度设置为100*/
function func1()
{
    UserForm1.OptionButton1.Width = 100
}
```

### OptionButton.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 OptionButton 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将OptionButton1的宽度设置为固定*/
function func1()
{
    UserForm1.OptionButton1.WidthPolicy = 2
}
```

### OptionButton.WordWrap

指定控件的内容是否在行尾自动换行

**语法**

**express.WordWrap**

\*express是一个代表 OptionButton 对象的变量。

**说明**

指定控件的内容是否在行尾自动换行

**示例**

``` JavaScript
/*将OptionButton1设置为控件的内容在行尾自动换行*/
function func1()
{
    UserForm1.OptionButton1.WordWrap = true
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/OptionButton](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ScrollBar 对象

## 简介

提供在长列表工程或大量信息中快速浏览的图形工具，以比例方式指示出当前位置，或是做为一个输入设备，或速度及数量的指示器。

## 说明

提供在长列表工程或大量信息中快速浏览的图形工具，以比例方式指示出当前位置，或是做为一个输入设备，或速度及数量的指示器。

## 方法

| 序号 | 名称                    | 说明                       |
|:----:|:------------------------|:---------------------------|
|  1   | [Move](#ScrollBar.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                      | 说明                                                      |
|:----:|:------------------------------------------|:----------------------------------------------------------|
|  1   | [AutoSize](#ScrollBar.AutoSize)           | 定该对象是否根据内容自动设置它的大小                      |
|  2   | [BackColor](#ScrollBar.BackColor)         | 获取或设置指定对象的内部颜色                              |
|  3   | [ControlSource](#ScrollBar.ControlSource) | 您可以使用该属性指定控件中显示的数据                      |
|  4   | [Enabled](#ScrollBar.Enabled)             | 您可以通过该属性设置或返回是否启用该控件                  |
|  5   | [Font](#ScrollBar.Font)                   | 指定对象的字体，只读                                      |
|  6   | [ForeColor](#ScrollBar.ForeColor)         | 您可以通过该属性指定控件中文本的颜色                      |
|  7   | [Height](#ScrollBar.Height)               | 获取或设置指定对象的高度                                  |
|  8   | [HeightPolicy](#ScrollBar.HeightPolicy)   | 获取或设置指定对像的高度策略                              |
|  9   | [LargeChange](#ScrollBar.LargeChange)     | 指定当用户在滚动框和滚动箭头之间单击时发生的移动量。      |
|  10  | [Left](#ScrollBar.Left)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  11  | [Max](#ScrollBar.Max)                     | 为该控件的Value属性指定最大可接受值。                     |
|  12  | [Min](#ScrollBar.Min)                     | 为该控件的Value属性指定最小可接受值。                     |
|  13  | [Name](#ScrollBar.Name)                   | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  14  | [Orientation](#ScrollBar.Orientation)     | 指定该控件是垂直定向还是水平定向                          |
|  15  | [SmallChange](#ScrollBar.SmallChange)     | 指定当用户单击滚动箭头时发生的移动量。                    |
|  16  | [TabIndex](#ScrollBar.TabIndex)           | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  17  | [TabStop](#ScrollBar.TabStop)             | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  18  | [Top](#ScrollBar.Top)                     | 您可以通过该属性指定控件在对象或报表中的位置              |
|  19  | [Value](#ScrollBar.Value)                 | 您可以通过该属性指定一个在Max和Min之间的值                |
|  20  | [Visible](#ScrollBar.Visible)             | 返回或设置该对象是否可见                                  |
|  21  | [Width](#ScrollBar.Width)                 | 获取或设置指定对象的宽度                                  |
|  22  | [WidthPolicy](#ScrollBar.WidthPolicy)     | 获取或设置指定对象的宽度策略                              |

## 成员方法

### ScrollBar.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 ScrollBar 对象的变量。

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
/*将ScrollBar1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.ScrollBar1.Move(100,100,100,100)
}
```

## 成员属性

### ScrollBar.AutoSize

定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置ScrollBar1控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.ScrollBar1.AutoSize = true
}
```

### ScrollBar.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将ScrollBar1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.ScrollBar1.BackColor = 0x665898
}
```

### ScrollBar.ControlSource

您可以使用该属性指定控件中显示的数据

**语法**

**express.ControlSource**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以使用该属性指定控件中显示的数据。您可以显示和编辑绑定到表的数据、查询的数据或SQL语句中的字段的数据。您还可以显示表达式的结果

### ScrollBar.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.ScrollBar1.Enabled = true
}
```

### ScrollBar.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

指定对象的字体，只读

### ScrollBar.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将ScrollBar1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.ScrollBar1.ForeColor = 0x658978
}
```

### ScrollBar.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将ScrollBar1控件的高度设置为100*/
function func1()
{
    UserForm1.ScrollBar1.Height = 100
}
```

### ScrollBar.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将ScrollBar1的高度设置为固定*/
function func1()
{
    UserForm1.ScrollBar1.HeightPolicy = 2
}
```

### ScrollBar.LargeChange

指定当用户在滚动框和滚动箭头之间单击时发生的移动量。

**语法**

**express.LargeChange**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

指定当用户在滚动框和滚动箭头之间单击时发生的移动量。

**示例**

``` JavaScript
/*将ScrollBar1在用户点击滚动框和滚动箭头时发生的移动量设为200*/
function func1()
{
    UserForm1.ScrollBar1.LargeChange = 200
}
```

### ScrollBar.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ScrollBar1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.ScrollBar1.Left = 100
}
```

### ScrollBar.Max

为该控件的Value属性指定最大可接受值。

**语法**

**express.Max**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

为该控件的Value属性指定最大可接受值。

**示例**

``` JavaScript
/*将ScrollBar1的Value属性的最大可接受值设为2000*/
function func1()
{
    UserForm1.ScrollBar1.Max = 2000
}
```

### ScrollBar.Min

为该控件的Value属性指定最小可接受值。

**语法**

**express.Min**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

为该控件的Value属性指定最小可接受值。

**示例**

``` JavaScript
/*将ScrollBar1的Value属性的最小可接受值设为0*/
function func1()
{
    UserForm1.ScrollBar1.Mix = 0
}
```

### ScrollBar.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### ScrollBar.Orientation

指定该控件是垂直定向还是水平定向

**语法**

**express.Orientation**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

指定该控件是垂直定向还是水平定向，可以设置为以下值：

| 设置       | 值  | 描述                                 |
|------------|-----|--------------------------------------|
| Auto       | -1  | 根据控件的尺寸自动确定方向（默认）。 |
| Vertical   | 0   | 控件垂直呈现。                       |
| Horizontal | 1   | 控件水平呈现。                       |

**示例**

``` JavaScript
/*将ScrollBar1垂直呈现*/
function func1()
{
    UserForm1.ScrollBar1.Orientation = 0
}
```

### ScrollBar.SmallChange

指定当用户单击滚动箭头时发生的移动量。

**语法**

**express.SmallChange**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

指定当用户单击滚动箭头时发生的移动量。

**示例**

``` JavaScript
/*将ScrollBar1在用户点击滚动箭头时发生的移动量设为100*/
function func1()
{
    UserForm1.ScrollBar1.LargeChange = 100
}
```

### ScrollBar.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将ScrollBar1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.ScrollBar1.TabIndex = 2
}
```

### ScrollBar.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将ScrollBar1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.ScrollBar1.TabStop = true
}
```

### ScrollBar.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ScrollBar1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.ScrollBar1.Top = 100
}
```

### ScrollBar.Value

您可以通过该属性指定一个在Max和Min之间的值

**语法**

**express.Value**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

您可以通过该属性指定一个在Max和Min之间的值

**示例**

``` JavaScript
/*将ScrollBar1的Valued的值为500*/
function func1()
{
    UserForm1.ScrollBar1.Value = 500
}
```

### ScrollBar.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将ScrollBar1设置为可见*/
function func1()
{
    UserForm1.ScrollBar1.Visible = true
}
```

### ScrollBar.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将ScrollBar1的宽度设置为100*/
function func1()
{
    UserForm1.ScrollBar1.Width = 100
}
```

### ScrollBar.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 ScrollBar 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将ScrollBar1的宽度设置为固定*/
function func1()
{
    UserForm1.ScrollBar1.WidthPolicy = 2
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/ScrollBar](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Axes 对象

## 简介

指定图表中所有 Axis 对象的集合。

## 方法

| 序号 | 名称               | 说明                               |
|:----:|:-------------------|:-----------------------------------|
|  1   | [Item](#Axes.Item) | 从 Axes 集合中返回一个 Axis 对象。 |

## 属性

| 序号 | 名称                             | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Axes.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Axes.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Axes.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Axes.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Axes.Item

从 **Axes** 集合中返回一个 **Axis** 对象。

**语法**

**express.Item ( *Type* , *AxisGroup* )**

\*express是一个代表 Axes 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型    | 说明         |
|-------------|-----------|-------------|--------------|
| *Type*      | 必选      | XlAxisType  | 坐标轴类型。 |
| *AxisGroup* | 可选      | XlAxisGroup | 坐标轴。     |

**说明**

**示例**

``` JavaScript
/*本示例为 Chart1 中分类轴设置标题文本。*/
function test()
{
  let myChart = Application.Charts.Item("Chart1").Axes.Item(xlCategory)
  myChart.HasTitle = true
  myChart.AxisTitle.Caption = "1994"
}
```

## 成员属性

### Axes.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Axes 对象的变量。

### Axes.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Axes 对象的变量。

### Axes.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Axes 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Axes.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Axes 对象的变量。

**示例**

``` JavaScript
/*本示例显示包含 myAxis 的图表的名称。*/
function test(){
    let myAxis = Application.Charts.Item(1).Axes(xlValue)
    alert(myAxis.Parent.Name)
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Axes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Characters 对象

## 简介

代表包含文本的对象中的字符。

## 说明

使用 Characters 对象可修改包含在全文本字符串中的任意字符序列。

使用 Characters ( *start* , *length* )（其中 *start* 为起始字符号，而 *length* 为要返回的字符个数）返回 Characters 对象。

``` JavaScript
/*下例向单元格 B1 中添加文本，并将第二个单词设置为加粗。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("B1").Value = "New Title"
    Application.Worksheets.Item("Sheet1").Range("B1").Characters(5, 5).Font.Bold = true
}
```

仅当需要更改对象中文本的一部分而不影响其余部分时，才有必要使用 Characters 方法（如果对象不支持格式文本，则不能使用 Characters 方法对文本中的一部分单独设置格式）。要同时更改所有文本，通常可以对该对象直接应用某一适当的方法或属性。

``` JavaScript
/*下例将单元格 A5 的内容设置为倾斜。*/
Application.Worksheets.Item("Sheet1").Range("A5").Font.Italic = true
```

## 方法

| 序号 | 名称                         | 说明                             |
|:----:|:-----------------------------|:---------------------------------|
|  1   | [Delete](#Characters.Delete) | 删除对象。                       |
|  2   | [Insert](#Characters.Insert) | 在所选的字符前面插入一个字符串。 |

## 属性

| 序号 | 名称                                                 | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Characters.Application)               | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Caption](#Characters.Caption)                       | 返回一个 String 值，它代表此字符区域的文本。                                                                                                                                                                                    |
|  3   | [Count](#Characters.Count)                           | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  4   | [Creator](#Characters.Creator)                       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [Font](#Characters.Font)                             | 返回一个 Font 对象，它代表指定对象的字体。                                                                                                                                                                                      |
|  6   | [Parent](#Characters.Parent)                         | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  7   | [PhoneticCharacters](#Characters.PhoneticCharacters) | 返回或设置指定 Characters 对象中的拼音文本。String 类型，可读写。                                                                                                                                                               |
|  8   | [Text](#Characters.Text)                             | 返回或设置指定对象中的文本。 String 型，可读写。                                                                                                                                                                                |

## 成员方法

### Characters.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Characters 对象的变量。

**说明**

返回值 Variant

### Characters.Insert

在所选的字符前面插入一个字符串。

**语法**

**express.Insert ( *String* )**

\*express是一个代表 Characters 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明             |
|----------|-----------|----------|------------------|
| *String* | 必选      | String   | 要插入的字符串。 |

**说明**

返回值 Variant

## 成员属性

### Characters.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Characters 对象的变量。

### Characters.Caption

返回一个 **String** 值，它代表此字符区域的文本。

**语法**

**express.Caption**

\*express是一个代表 Characters 对象的变量。

### Characters.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Characters 对象的变量。

**示例**

``` JavaScript
/*此示例将 A1 单元格的最后一个字符设为上标字符。*/
function test(){
    let n = Application.Worksheets.Item("Sheet1").Range("A1").Characters.Count
    Application.Worksheets.Item("Sheet1").Range("A1").Characters(n, 1).Font.Superscript = true
}
```

### Characters.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Characters 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Characters.Font

返回一个 **Font** 对象，它代表指定对象的字体。

**语法**

**express.Font**

\*express是一个代表 Characters 对象的变量。

### Characters.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Characters 对象的变量。

### Characters.PhoneticCharacters

返回或设置指定 **Characters** 对象中的拼音文本。String 类型，可读写。

**语法**

**express.PhoneticCharacters**

\*express是一个代表 Characters 对象的变量。

**说明**

要在单元格中添加拼音信息，请不要使用该属性，应使用 **Phonetics** 集合的 **Add** 方法；另外，应使用 **Phonetic** 对象的 **Text** 属性返回或设置单元格中的拼音文本字符串。

**示例**

``` JavaScript
/*本示例将活动单元格中从文本起始位置开始的第四个字符替换。*/
ActiveCell.Characters(1,3).PhoneticCharacters = "替换字符"
```

### Characters.Text

返回或设置指定对象中的文本。 **String** 型，可读写。

**语法**

**express.Text**

\*express是一个代表 Characters 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Characters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ChartObject 对象

## 简介

代表工作表上的嵌入图表。

## 说明

ChartObject 对象充当 Chart 对象的容器。 ChartObject 对象的属性和方法控制工作表上嵌入图表的外观和大小。 ChartObject 对象是 ChartObjects 集合的成员。 ChartObjects 集合包含单一工作表上的所有嵌入图表。

使用 ChartObjects ( *index* )（其中 *index* 是嵌入图表的索引号或名称）可以返回单个 ChartObject 对象。

示例 以下示例设置名为“Sheet1”的工作表上嵌入图表 Chart 1 中的图表区图案。

``` JavaScript
Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format.Fill.Pattern = msoPatternLightDownwardDiagonal 
```

当选定嵌入图表时，其名称显示在“名称”框中。使用 Name 属性可设置或返回 ChartObject 对象的名称。以下示例对工作表“Sheet1”上的嵌入图表“Chart 1”使用了圆角。

``` JavaScript
Application.Worksheets.Item("Sheet1").ChartObjects("Chart 1").RoundedCorners = true
```

## 方法

| 序号 | 名称                                      | 说明                                 |
|:----:|:------------------------------------------|:-------------------------------------|
|  1   | [Activate](#ChartObject.Activate)         | 使当前图表成为活动图表。             |
|  2   | [BringToFront](#ChartObject.BringToFront) | 将对象放到 z-次序前面。              |
|  3   | [Copy](#ChartObject.Copy)                 | 将对象复制到剪贴板。                 |
|  4   | [CopyPicture](#ChartObject.CopyPicture)   | 将所选对象作为图片复制到剪贴板。     |
|  5   | [Cut](#ChartObject.Cut)                   | 将对象剪切到剪贴板。                 |
|  6   | [Delete](#ChartObject.Delete)             | 删除对象。                           |
|  7   | [Duplicate](#ChartObject.Duplicate)       | 复制对象，并返回对新复制对象的引用。 |
|  8   | [Select](#ChartObject.Select)             | 选择对象。                           |
|  9   | [SendToBack](#ChartObject.SendToBack)     | 将对象放到 z-次序的后面。            |

## 属性

| 序号 | 名称                                                  | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ChartObject.Application)               | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [BottomRightCell](#ChartObject.BottomRightCell)       | 返回一个 Range 对象，它代表位于该对象右下角下方的单元格。只读。                                                                                                                                                                 |
|  3   | [Chart](#ChartObject.Chart)                           | 返回一个 Chart 对象，该对象代表对象中包含的图表。只读。                                                                                                                                                                         |
|  4   | [Creator](#ChartObject.Creator)                       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [Height](#ChartObject.Height)                         | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |
|  6   | [Index](#ChartObject.Index)                           | 返回 Long 值，它代表对象在其同类对象所组成的集合内的索引号。                                                                                                                                                                    |
|  7   | [Left](#ChartObject.Left)                             | 返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                       |
|  8   | [Locked](#ChartObject.Locked)                         | 返回或设置一个 Boolean 值，它指明对象是否已被锁定。                                                                                                                                                                             |
|  9   | [Name](#ChartObject.Name)                             | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  10  | [Parent](#ChartObject.Parent)                         | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  11  | [Placement](#ChartObject.Placement)                   | 返回或设置一个包含 XlPlacement 常量的 Variant 值，它代表对象附加到它所在的单元格的方式。                                                                                                                                        |
|  12  | [PrintObject](#ChartObject.PrintObject)               | 如果打印文档时也打印指定对象，则为 True 。 Boolean 类型，可读写。                                                                                                                                                               |
|  13  | [ProtectChartObject](#ChartObject.ProtectChartObject) | 如果不能通过用户界面对嵌入图表框架执行移动、调整大小或删除操作，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                     |
|  14  | [RoundedCorners](#ChartObject.RoundedCorners)         | 如果嵌入式图表使用圆角，则为 True 。 Boolean 类型，可读写。                                                                                                                                                                     |
|  15  | [Shadow](#ChartObject.Shadow)                         | 返回或设置 Boolean 值，它确定字体是否是阴影字体或对象是否带有阴影。                                                                                                                                                             |
|  16  | [ShapeRange](#ChartObject.ShapeRange)                 | 返回一个 ShapeRange 对象，它代表指定的一个或多个对象。只读。                                                                                                                                                                    |
|  17  | [Top](#ChartObject.Top)                               | 返回或设置一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                      |
|  18  | [TopLeftCell](#ChartObject.TopLeftCell)               | 返回一个 Range 对象，它代表位于指定对象左上角下方的单元格。只读。                                                                                                                                                               |
|  19  | [Visible](#ChartObject.Visible)                       | 返回或设置一个 Boolean 值，它确定对象是否可见。可读写。                                                                                                                                                                         |
|  20  | [Width](#ChartObject.Width)                           | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |
|  21  | [ZOrder](#ChartObject.ZOrder)                         | 返回指定对象的 z-次序位置。 Long 型，只读。                                                                                                                                                                                     |

## 成员方法

### ChartObject.Activate

使当前图表成为活动图表。

**语法**

**express.Activate ()**

\*express是一个代表 ChartObject 对象的变量。

### ChartObject.BringToFront

将对象放到 z-次序前面。

**语法**

**express.BringToFront ()**

\*express是一个代表 ChartObject 对象的变量。

### ChartObject.Copy

将对象复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 ChartObject 对象的变量。

### ChartObject.CopyPicture

将所选对象作为图片复制到剪贴板。

**语法**

**express.CopyPicture ( *Appearance* , *Format* )**

\*express是一个代表 ChartObject 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型            | 说明                                    |
|--------------|-----------|---------------------|-----------------------------------------|
| *Appearance* | 可选      | XlPictureAppearance | 指定图片的复制方式。默认值为 xlScreen。 |
| *Format*     | 可选      | XlCopyPictureFormat | 图片的格式。默认值为 xlPicture。        |

### ChartObject.Cut

将对象剪切到剪贴板。

**语法**

**express.Cut ()**

\*express是一个代表 ChartObject 对象的变量。

**说明**

只可剪切嵌入式图表。

### ChartObject.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 ChartObject 对象的变量。

### ChartObject.Duplicate

复制对象，并返回对新复制对象的引用。

**语法**

**express.Duplicate ()**

\*express是一个代表 ChartObject 对象的变量。

**示例**

``` JavaScript
/*此示例复制工作表 Sheet1 上嵌入的第一个图表，然后选择新复制的图表。*/
function test()
{
    let dChart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Duplicate()
    dChart.Select()
}
```

### ChartObject.Select

选择对象。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 ChartObject 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                            |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | 如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

### ChartObject.SendToBack

将对象放到 z-次序的后面。

**语法**

**express.SendToBack ()**

\*express是一个代表 ChartObject 对象的变量。

**示例**

``` JavaScript
/*此示例将工作表 Sheet1 上嵌入的第一个图表放到 z-次序的后面。*/
Application.Worksheets.Item("Sheet1").ChartObjects(1).SendToBack()
```

## 成员属性

### ChartObject.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ChartObject 对象的变量。

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

### ChartObject.BottomRightCell

返回一个 **Range** 对象，它代表位于该对象右下角下方的单元格。只读。

**语法**

**express.BottomRightCell**

\*express是一个代表 ChartObject 对象的变量。

**示例**

``` JavaScript
/*此示例显示 Sheet1 中嵌入的第一个图表右下角单元格的地址。*/
alert("The bottom right corner is over cell " +
    Application.Worksheets.Item("Sheet1").ChartObjects(1).BottomRightCell.Address())
```

### ChartObject.Chart

返回一个 **Chart** 对象，该对象代表对象中包含的图表。只读。

**语法**

**express.Chart**

\*express是一个代表 ChartObject 对象的变量。

**示例**

``` JavaScript
/*本示例为 Sheet1 上嵌入的第一个图表添加标题。*/
function test()
{
    let myChart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    myChart.HasTitle = true
    myChart.ChartTitle.Text = "1995 Rainfall Totals by Month"
}
```

### ChartObject.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ChartObject 对象的变量。

### ChartObject.Height

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 ChartObject 对象的变量。

### ChartObject.Index

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 ChartObject 对象的变量。

### ChartObject.Left

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 ChartObject 对象的变量。

**示例**

``` JavaScript
/*此示例使嵌入图表的左边界与 B 列的左边界对齐。*/
function test()
{
    let mySheet = Application.Worksheets.Item("Sheet1")
    mySheet.ChartObjects(1).Left = mySheet.Columns.Item("B").Left
}
```

### ChartObject.Locked

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

\*express是一个代表 ChartObject 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True** ；如果在工作表处于受保护状态时仍能修改对象，则返回 **False** 。

### ChartObject.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 ChartObject 对象的变量。

**说明**

此属性对图表对象（嵌入式图表）而言是只读的。

### ChartObject.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ChartObject 对象的变量。

### ChartObject.Placement

返回或设置一个包含 **XlPlacement** 常量的 **Variant** 值，它代表对象附加到它所在的单元格的方式。

**语法**

**express.Placement**

\*express是一个代表 ChartObject 对象的变量。

**示例**

``` JavaScript
/*本示例设置工作表 Sheet1 上嵌入的第一个图表为可自由浮动（既不随下方单元格移动，也不随其改变大小）。*/
Application.Worksheets.Item("Sheet1").ChartObjects(1).Placement = xlFreeFloating
```

### ChartObject.PrintObject

如果打印文档时也打印指定对象，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintObject**

\*express是一个代表 ChartObject 对象的变量。

**示例**

``` JavaScript
/*此示例设置当打印工作表 sheet1 时连同嵌入的第一个图表一起打印。*/
Application.Worksheets.Item("Sheet1").ChartObjects(1).PrintObject = true
```

### ChartObject.ProtectChartObject

如果不能通过用户界面对嵌入图表框架执行移动、调整大小或删除操作，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ProtectChartObject**

\*express是一个代表 ChartObject 对象的变量。

**说明**

将该属性设置为 **True** 后将无法防止通过对象模型修改嵌入图表框架。

**示例**

``` JavaScript
/*本示例保护第一张工作表上嵌入的第一个图表。*/
Application.Worksheets.Item(1).ChartObjects(1).ProtectChartObject = true
```

### ChartObject.RoundedCorners

如果嵌入式图表使用圆角，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.RoundedCorners**

\*express是一个代表 ChartObject 对象的变量。

**示例**

``` JavaScript
/*此示例为 Sheet1 的第一个嵌入式图表添加圆角。*/
Application.Worksheets.Item("Sheet1").ChartObjects(1).RoundedCorners = true
```

### ChartObject.Shadow

返回或设置 **Boolean** 值，它确定字体是否是阴影字体或对象是否带有阴影。

**语法**

**express.Shadow**

\*express是一个代表 ChartObject 对象的变量。

**说明**

对于 **Font** 对象，该属性在 Microsoft Windows 中无效，但保留其值（可被设置或返回）。

### ChartObject.ShapeRange

返回一个 **ShapeRange** 对象，它代表指定的一个或多个对象。只读。

**语法**

**express.ShapeRange**

\*express是一个代表 ChartObject 对象的变量。

**示例**

``` JavaScript
/*此示例创建一个形状区域，该区域代表第一张工作表上的嵌入图表。*/
let sr = Application.Worksheets.Item(1).ChartObjects.ShapeRange
```

### ChartObject.Top

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 ChartObject 对象的变量。

### ChartObject.TopLeftCell

返回一个 **Range** 对象，它代表位于指定对象左上角下方的单元格。只读。

**语法**

**express.TopLeftCell**

\*express是一个代表 ChartObject 对象的变量。

**示例**

``` JavaScript
/*此示例显示 Sheet1 上嵌入的第一张图表左上角下面的单元格的地址。*/
function test()
{
    alert("The top left corner is over cell " +
    Application.Worksheets.Item("Sheet1").ChartObjects(1).TopLeftCell.Address())
}
```

### ChartObject.Visible

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 ChartObject 对象的变量。

### ChartObject.Width

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 ChartObject 对象的变量。

**示例**

``` JavaScript
/*此示例设置嵌入图表的宽度。*/
Application.Worksheets.Item("Sheet1").ChartObjects(1).Width = 360
```

### ChartObject.ZOrder

返回指定对象的 z-次序位置。 **Long** 型，只读。

**语法**

**express.ZOrder**

\*express是一个代表 ChartObject 对象的变量。

**说明**

在任何对象集合中，z-次序尾端的对象为 *collection* (1)，z-次序前端的对象为 *collection* ( *collection* . **Count** )。例如，如果活动工作表中有嵌入图表，z-次序尾端的图表为 ` ActiveSheet.ChartObjects(1) ` ，z-次序前端的图表为 ` ActiveSheet.ChartObjects(ActiveSheet.ChartObjects.Count) ` 。

**示例**

``` JavaScript
/*此示例显示 Sheet1 上嵌入的第一张图表的 z-次序位置。*/
function test()
{
    alert("The chart's z-order position is " +
    Application.Worksheets.Item("Sheet1").ChartObjects(1).ZOrder)
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ChartObject](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Charts 对象

## 简介

指定的或活动工作簿中所有图表工作表的集合。

## 方法

| 序号 | 名称                                 | 说明                                   |
|:----:|:-------------------------------------|:---------------------------------------|
|  1   | [Add](#Charts.Add)                   | 新建图表工作表，并返回 Chart 对象。    |
|  2   | [Copy](#Charts.Copy)                 | 将工作表复制到工作簿的另一位置。       |
|  3   | [Delete](#Charts.Delete)             | 删除对象。                             |
|  4   | [Item](#Charts.Item)                 | 从集合中返回一个对象。                 |
|  5   | [Move](#Charts.Move)                 | 将图表移到工作簿的另一位置。           |
|  6   | [PrintOut](#Charts.PrintOut)         | 打印对象。                             |
|  7   | [PrintPreview](#Charts.PrintPreview) | 按对象打印后的外观效果显示对象的预览。 |
|  8   | [Select](#Charts.Select)             | 选择对象。                             |

## 属性

| 序号 | 名称                               | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Charts.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Charts.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Charts.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [HPageBreaks](#Charts.HPageBreaks) | 返回一个 HPageBreaks 集合，它代表图表上的水平分页符。只读。                                                                                                                                                                     |
|  5   | [Parent](#Charts.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  6   | [VPageBreaks](#Charts.VPageBreaks) | 返回一个 VPageBreaks 集合，它代表工作表上的垂直分页符。只读。                                                                                                                                                                   |
|  7   | [Visible](#Charts.Visible)         | 返回或设置一个 Variant 值，它确定对象是否可见。                                                                                                                                                                                 |

## 成员方法

### Charts.Add

新建图表工作表，并返回 **Chart** 对象。

**语法**

**express.Add ( *Before* , *After* , *Count* , *NewLayout* )**

\*express是一个代表 Charts 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                            |
|-------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------|
| *Before*    | 可选      | Variant  | 指定工作表的对象，新建的工作表将置于此工作表之前。                                                              |
| *After*     | 可选      | Variant  | 指定工作表的对象，新建的工作表将置于此工作表之后。                                                              |
| *Count*     | 可选      | Variant  | 要添加的工作表数。默认值为 1。                                                                                  |
| *NewLayout* | 可选      | Variant  | 如果 NewLayout 为 True， 则通过使用新的动态格式规则插入图表 (Title 为打开，并且 Legend 仅在有多个数据系列时) 。 |

**说明**

如果 *Before* 和 *After* 两者都被省略，新建的图表工作表将插入到活动工作表之前。

**示例**

``` JavaScript
/*本示例创建空白图表工作表，并将其插入到最后一张工作表之前。*/
Application.ActiveWorkbook.Charts.Add(Worksheets.Item(Worksheets.Count))
```

### Charts.Copy

将工作表复制到工作簿的另一位置。

**语法**

**express.Copy ( *Before* , *After* )**

\*express是一个代表 Charts 对象的变量。

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
Application.Worksheets.Item("Sheet1").Copy(null, Worksheets.Item("Sheet3"))
```

### Charts.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Charts 对象的变量。

### Charts.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Charts 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**示例**

``` JavaScript
/*此示例对 Chart1 中的趋势线向前和向后延伸的单位数进行设置。此示例应在包含单个带趋势线系列的二维柱形图上运行。
*/
function test()
{
    let myChartLine = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)
    myChartLine.Forward = 5
    myChartLine.Backward = .5
}
```

### Charts.Move

将图表移到工作簿的另一位置。

**语法**

**express.Move ( *Before* , *After* )**

\*express是一个代表 Charts 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                      |
|----------|-----------|----------|---------------------------------------------------------------------------|
| *Before* | 可选      | Variant  | 将要在其之前放置所移动图表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选      | Variant  | 将要在其之后放置所移动图表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After* ，则 ET 将新建一个工作簿并将要移动的图表移到新工作簿中。

### Charts.PrintOut

打印对象。

**语法**

**express.PrintOut ( *From* , *To* , *Copies* , *Preview* , *ActivePrinter* , *PrintToFile* , *Collate* , *PrToFileName* , *IgnorePrintAreas* )**

\*express是一个代表 Charts 对象的变量。

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
| *PrToFileName*     | 可选      | Variant  | 如果 PrintToFile 设为 True，则该参数指定要打印到的文件名。                                        |
| *IgnorePrintAreas* | 可选      | Variant  | 如果为 True，则忽略打印区域并打印整个对象。                                                       |

**说明**

*From* 和 *To* 所描述的“页”指的是要打印的页，并非指定工作表或工作簿中的全部页。

**示例**

``` JavaScript
/*此示例打印当前活动的Charts工作表。*/
Application.Charts.PrintOut()
```

### Charts.PrintPreview

按对象打印后的外观效果显示对象的预览。

**语法**

**express.PrintPreview ( *EnableChanges* )**

\*express是一个代表 Charts 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                          |
|-----------------|-----------|----------|-------------------------------------------------------------------------------|
| *EnableChanges* | 可选      | Variant  | 传递 Boolean 值，以指定用户是否可更改边距和打印预览中可用的其他页面设置选项。 |

**示例**

``` JavaScript
/*此示例在打印预览中显示Charts工作表。*/
Application.Charts.PrintPreview()
```

### Charts.Select

选择对象。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 Charts 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                                              |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

## 成员属性

### Charts.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Charts 对象的变量。

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

### Charts.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Charts 对象的变量。

### Charts.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Charts 对象的变量。

### Charts.HPageBreaks

返回一个 **HPageBreaks** 集合，它代表图表上的水平分页符。只读。

**语法**

**express.HPageBreaks**

\*express是一个代表 Charts 对象的变量。

### Charts.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Charts 对象的变量。

### Charts.VPageBreaks

返回一个 **VPageBreaks** 集合，它代表工作表上的垂直分页符。只读。

**语法**

**express.VPageBreaks**

\*express是一个代表 Charts 对象的变量。

### Charts.Visible

返回或设置一个 **Variant** 值，它确定对象是否可见。

**语法**

**express.Visible**

\*express是一个代表 Charts 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Charts](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FileExportConverters 对象

## 简介

FileExportConverter 对象的集合，这些对象代表可供保存文件时使用的所有文件转换器。

## 属性

| 序号 | 名称                                             | 说明                                                                             |
|:----:|:-------------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [Application](#FileExportConverters.Application) | 返回代表 ET 应用程序的 Application 对象。只读。                                  |
|  2   | [Count](#FileExportConverters.Count)             | 返回一个 Long 类型的值，该值代表集合中文件转换器的数量。只读。                   |
|  3   | [Creator](#FileExportConverters.Creator)         | 返回表示在其中创建特定对象的应用程序的 32 位整数。 Long 类型，只读。             |
|  4   | [Item](#FileExportConverters.Item)               | 返回集合中的单个 FileExportConverter 对象。                                      |
|  5   | [Parent](#FileExportConverters.Parent)           | 返回一个 Object 类型的值，该值代表指定 FileExportConverters 对象的父对象。只读。 |

## 成员属性

### FileExportConverters.Application

返回代表 ET 应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 FileExportConverters 对象的变量。

### FileExportConverters.Count

返回一个 **Long** 类型的值，该值代表集合中文件转换器的数量。只读。

**语法**

**express.Count**

\*express是一个代表 FileExportConverters 对象的变量。

### FileExportConverters.Creator

返回表示在其中创建特定对象的应用程序的 32 位整数。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 FileExportConverters 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回十六进制值 5843454C，该值表示字符串“XCEL”。 **Creator** 属性设计为在 ET for the Macintosh 中使用，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### FileExportConverters.Item

返回集合中的单个 **FileExportConverter** 对象。

**语法**

**express.Item**

\*express是一个代表 FileExportConverters 对象的变量。

### FileExportConverters.Parent

返回一个 **Object** 类型的值，该值代表指定 **FileExportConverters** 对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FileExportConverters 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/FileExportConverters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# HPageBreaks 对象

## 简介

打印区域内水平分页符的集合。

## 说明

每个水平分页符都由一个 HPageBreak 对象代表。

如果添加的分页符不和打印区域交叠，则新添加的 HPageBreak 对象将不会出现在打印区域的 HPageBreaks 集合中。如果重新调整打印区域或重新定义打印区域，将会改变集合的内容。

当 Application 属性、 Count 属性、 Item 属性、 Parent 属性或 Add 方法配合 HPageBreaks 属性一起使用时：

- 对于自动打印区域， HPageBreaks 属性只应用于打印区域内的分页符。
- 对于同一区域中用户定义的打印区域， HPageBreaks 属性应用于所有的分页符。

| 注释                                 |
|--------------------------------------|
| 每张工作表最多有 1026 个水平分页符。 |

使用 HPageBreaks 属性可返回 HPageBreaks 集合。使用 Add 方法可添加一个水平分页符。下例在活动单元格上方添加一个水平分页符。

``` JavaScript
Application.ActiveSheet.HPageBreaks.Add(ActiveCell)
```

## 方法

| 序号 | 名称                      | 说明                                                              |
|:----:|:--------------------------|:------------------------------------------------------------------|
|  1   | [Add](#HPageBreaks.Add)   | 添加水平分页符。返回 一个 HPageBreak 对象，它代表新的水平分页符。 |
|  2   | [Item](#HPageBreaks.Item) | 从集合中返回一个对象。                                            |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#HPageBreaks.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#HPageBreaks.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#HPageBreaks.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#HPageBreaks.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### HPageBreaks.Add

添加水平分页符。返回 一个 **HPageBreak** 对象，它代表新的水平分页符。

**语法**

**express.Add ( *Before* )**

\*express是一个代表 HPageBreaks 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                            |
|----------|-----------|----------|-------------------------------------------------|
| *Before* | 必选      | Object   | 一个 Range 对象。新的分页符将添加到该区域上方。 |

**示例**

``` JavaScript
/* 本示例在单元格 F25 上方添加水平分页符，在其左侧添加垂直分页符*/
function test() {
    let sheet2 = Application.Worksheets.Item(1)
    sheet2.HPageBreaks.Add(sheet2.Range("F25"))
    sheet2.VPageBreaks.Add(sheet2.Range("F25"))
}
```

### HPageBreaks.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 HPageBreaks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *Index* | 必选      | Long     | 对象的索引号。 |

**示例**

``` JavaScript
/* 此示例更改第一个水平分页符的位置*/
Application.Worksheets.Item(1).HPageBreaks.Item(1).Location = Worksheets.Item(1).Range("e5")
```

## 成员属性

### HPageBreaks.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 HPageBreaks 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
function test() {
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### HPageBreaks.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 HPageBreaks 对象的变量。

### HPageBreaks.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 HPageBreaks 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### HPageBreaks.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 HPageBreaks 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/HPageBreaks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# LinearGradient 对象

## 简介

LinearGradient 对象沿特定角度以线性方式在一系列颜色间转换。

## 说明

| 注释                                           |
|------------------------------------------------|
| 当使用 LinearGradient 对象时，应考虑以下几点： |

- 试图访问不具有现有渐变填充的 Interior 对象的 Gradient 属性会引起运行时错误。访问 Gradient 属性之前请注意 ` Interior.Pattern ` 属性。
- 如果将 Interior.Pattern 从渐变类型更改为非渐变类型，Gradient 对象将采用默认值。

## 属性

| 序号 | 名称                                       | 说明                                                                                                                                                               |
|:----:|:-------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#LinearGradient.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [ColorStops](#LinearGradient.ColorStops)   | 返回 LinearGradient 对象的 ColorStops 。只读。                                                                                                                     |
|  3   | [Creator](#LinearGradient.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Degree](#LinearGradient.Degree)           | 选定区域中线性渐变填充的角度。可读/写。                                                                                                                            |
|  5   | [Parent](#LinearGradient.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |

## 成员属性

### LinearGradient.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 LinearGradient 对象的变量。

### LinearGradient.ColorStops

返回 **LinearGradient** 对象的 **ColorStops** 。只读。

**语法**

**express.ColorStops**

\*express是一个代表 LinearGradient 对象的变量。

### LinearGradient.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 LinearGradient 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。Creator 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### LinearGradient.Degree

选定区域中线性渐变填充的角度。可读/写。

**语法**

**express.Degree**

\*express是一个代表 LinearGradient 对象的变量。

**说明**

使用 0 - 360 之间的值。

**示例**

``` JavaScript
function test(){
Selection.Interior.Pattern = xlPatternLinearGradient
Selection.Interior.Gradient.Degree = 45
}
```

### LinearGradient.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 LinearGradient 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/LinearGradient](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ListDataFormat 对象

## 简介

ListDataFormat 对象包含 ListColumn 对象的所有数据类型属性。这些属性是只读的。

## 说明

使用 ListColumn 对象的 ListDataFormat 属性可返回一个 ListDataFormat 对象集合。ListDataFormat 对象的默认属性为表示列表中列的数据类型的 Type 属性。这就使得用户在编写代码时无需指定 Type 属性。

如下代码示例创建一个与 SharePoint 列表链接的列表。然后检查字段 2 是否为必需（字段 1 为 ID 字段，且只读）。如果该字段是必需的文本字段，在所有现有记录中都编写了相同的数据。

| 注释                                                                                                                                                        |
|-------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 以下代码示例假设您会将 *strServerName* 和 *strListGuid* 变量替换为有效的服务器名称和列表 GUID。此外，服务器名称后面必须是“/\_vti_bin”，否则本示例将不运行。 |

``` JavaScript
function test(){
let strServerName = "http:///_vti_bin"
let strListGUID = "{}"
let objListObject = Worksheets.Item(1).ListObjects.Add(xlSrcExternal, Array(strServerName, strListGUID), true, xlYes, Range("A1"))
let objDataRange = objListObject.ListColumns.Item(2).Range.Offset(1, 0).Resize(objListObject.ListColumns.Item(2).Range.Rows.Count - 2, 1)
if(objListObject.ListColumns.Item(2).ListDataFormat.Type == xlListDataTypeText && objListObject.ListColumns.Item(2).ListDataFormat.Required){
    objDataRange.Value = "Hello World"
}
}
```

## 属性

| 序号 | 名称                                           | 说明                                                                                                                                                                                                                                                                                                                                    |
|:----:|:-----------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AllowFillIn](#ListDataFormat.AllowFillIn)     | 返回一个 Boolean 值，表示用户是否能够为提供值列表的这些列中某一列（而不是限定在值列表范围内）的单元格提供自己的数据。对于未与 SharePoint 网站链接的 列表?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。） ，返回 False 。如果未将该列指定为选择或多项选择，也返回 False 。 Boolean 类型，只读。 |
|  2   | [Application](#ListDataFormat.Application)     | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。                                                                                                         |
|  3   | [Choices](#ListDataFormat.Choices)             | 返回一个 String 值的 Array ，包含 DefaultValue 属性的 ListLookUp 、 ChoiceMulti 和 Choice 数据类型提供给用户的选项。 Variant 类型，只读。                                                                                                                                                                                               |
|  4   | [Creator](#ListDataFormat.Creator)             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                                                                                                                              |
|  5   | [DecimalPlaces](#ListDataFormat.DecimalPlaces) | 返回一个 Long 值，该值表示要为 ListColumn 对象中的数字显示的小数位数。 Long 类型，只读。                                                                                                                                                                                                                                                |
|  6   | [DefaultValue](#ListDataFormat.DefaultValue)   | 返回 Variant ，代表某列中新的一行的默认数据类型值。如果架构未指定默认值，则返回 Nothing 对象。 Variant 类型，只读。                                                                                                                                                                                                                     |
|  7   | [IsPercent](#ListDataFormat.IsPercent)         | 返回一个 Boolean 值。只有当 ListColumn 对象的数字数据以百分比的格式显示时，才返回 True 。 Boolean 类型，只读。                                                                                                                                                                                                                          |
|  8   | [lcid](#ListDataFormat.lcid)                   | 返回一个 Long 值，该值表示在架构定义中指定的 ListColumn 对象的 LCID。 Long 类型，只读。                                                                                                                                                                                                                                                 |
|  9   | [MaxCharacters](#ListDataFormat.MaxCharacters) | 如果将 Type 属性设置为 xlListDataTypeText 或 xlListDataTypeMultiLineText ，则返回 Long ，该值包含 ListColumn 对象中允许的最多字符数。 Long 类型，只读。                                                                                                                                                                                 |
|  10  | [MaxNumber](#ListDataFormat.MaxNumber)         | 返回一个 Variant ，其中包含列表列中该字段允许的最大值。 Variant 类型，只读。                                                                                                                                                                                                                                                            |
|  11  | [MinNumber](#ListDataFormat.MinNumber)         | 返回一个 Variant ，其中包含列表列中该字段允许的最小值。该值可以是负的浮点值。 Variant 类型，只读。                                                                                                                                                                                                                                      |
|  12  | [Parent](#ListDataFormat.Parent)               | 返回指定对象的父对象。只读。                                                                                                                                                                                                                                                                                                            |
|  13  | [ReadOnly](#ListDataFormat.ReadOnly)           | 如果对象以只读方式打开，则返回 True 。 Boolean 类型，只读。                                                                                                                                                                                                                                                                             |
|  14  | [Type](#ListDataFormat.Type)                   | 返回 XlListDataType 值，它代表列表列的数据类型。                                                                                                                                                                                                                                                                                        |

## 成员属性

### ListDataFormat.AllowFillIn

返回一个 **Boolean** 值，表示用户是否能够为提供值列表的这些列中某一列（而不是限定在值列表范围内）的单元格提供自己的数据。对于未与 SharePoint 网站链接的 列表?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。） ，返回 **False** 。如果未将该列指定为选择或多项选择，也返回 **False** 。 **Boolean** 类型，只读。

**语法**

**express.AllowFillIn**

\*express是一个代表 ListDataFormat 对象的变量。

### ListDataFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ListDataFormat 对象的变量。

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

### ListDataFormat.Choices

返回一个 **String** 值的 **Array** ，包含 **DefaultValue** 属性的 **ListLookUp** 、 **ChoiceMulti** 和 **Choice** 数据类型提供给用户的选项。 **Variant** 类型，只读。

**语法**

**express.Choices**

\*express是一个代表 ListDataFormat 对象的变量。

**说明**

在 ET 中，不能设置任何与 **ListDataFormat** 对象关联的属性，但可通过修改正在运行 Microsoft SharePoint Foundation 的服务器上的列表来设置这些属性。

**示例**

下例显示了与 SharePoint 列表链接的 列表 ?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。） 中第三列的 **Choice** 属性设置。在本例中，假定 **DefaultValue** 属性已设置为 **Choice** 、 **ChoiceMulti** 或 **ListLookup** 数据类型。

``` JavaScript
function test(){
   let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)

   Debug.Print (objListCol.ListDataFormat.Choices)

}
```

### ListDataFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ListDataFormat 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ListDataFormat.DecimalPlaces

返回一个 **Long** 值，该值表示要为 **ListColumn** 对象中的数字显示的小数位数。 **Long** 类型，只读。

**语法**

**express.DecimalPlaces**

\*express是一个代表 ListDataFormat 对象的变量。

**说明**

如果 **ListDataFormat.Type** 设置与小数位数不符，则返回 0。如果 Microsoft SharePoint Foundation 网站自动确定要显示在 SharePoint 列表中的小数位数，则返回 **xlAutomatic** （-4105 十进制）。

在 ET 中，不能设置任何与 **ListDataFormat** 对象关联的属性。但是可以通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

下例显示了对活动工作簿的 Sheet1 中某个 列表 ?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。） 的第三列设置的 **DecimalPlaces** 属性。

``` JavaScript
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)
   let GetDecimalPlaces = objListCol.ListDataFormat.DecimalPlaces

}
```

### ListDataFormat.DefaultValue

返回 **Variant** ，代表某列中新的一行的默认数据类型值。如果架构未指定默认值，则返回 **Nothing** 对象。 **Variant** 类型，只读。

**语法**

**express.DefaultValue**

\*express是一个代表 ListDataFormat 对象的变量。

**说明**

该属性只用于链接到 Microsoft SharePoint Foundation 网站的表格。

在 ET 中，不能设置任何与 **ListDataFormat** 对象关联的属性。但是可以通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

``` JavaScript
/*下例显示了对活动工作簿的 Sheet1 中表格的第三列设置的 DefaultValue 属性。此代码需要与 SharePoint 网站链接的列表。*/
function test(){
let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)

    if(objListCol.ListDataFormat.DefaultValue == null) {
        alert("No default value specified.")
    }
    else {
        alert(objListCol.ListDataFormat.DefaultValue)
    }

}
```

### ListDataFormat.IsPercent

返回一个 **Boolean** 值。只有当 **ListColumn** 对象的数字数据以百分比的格式显示时，才返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.IsPercent**

\*express是一个代表 ListDataFormat 对象的变量。

**说明**

此属性仅用于链接到 Microsoft SharePoint Foundation 网站的列表。

在 ET 中，不能设置任何与 **ListDataFormat** 对象关联的属性。但是可以通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

下例返回了对活动工作簿的 Sheet1 中 列表 ?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。） 第三列的 **IsPercent** 属性设置。

``` JavaScript
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)

    let GetIsPercent = objListCol.ListDataFormat.IsPercent

}
```

### ListDataFormat.lcid

返回一个 **Long** 值，该值表示在架构定义中指定的 **ListColumn** 对象的 LCID。 **Long** 类型，只读。

**语法**

**express.lcid**

\*express是一个代表 ListDataFormat 对象的变量。

**说明**

在 ET 中，LCID 表示当这是一个 **xlListDataTypeCurrency** 类型时要使用的货币符号。如果没有为列的数据类型设置区域，则返回 0（即 Language Neutral LCID）。

该属性仅用于链接到 Microsoft SharePoint Foundation 网站的表格。

在 ET 中，不能设置任何与 **ListDataFormat** 对象关联的属性。但是可以通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

``` JavaScript
/*下例显示了对活动工作簿的 Sheet1 中列表的第三列设置的 lcid 属性。 */
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)
    alert ("List LCID: " + objListCol.ListDataFormat.lcid)

}
```

### ListDataFormat.MaxCharacters

如果将 **Type** 属性设置为 **xlListDataTypeText** 或 **xlListDataTypeMultiLineText** ，则返回 **Long** ，该值包含 **ListColumn** 对象中允许的最多字符数。 **Long** 类型，只读。

**语法**

**express.MaxCharacters**

\*express是一个代表 ListDataFormat 对象的变量。

**说明**

对于 **Type** 属性被设为非文字值的列，则返回 -1。

该属性仅用于链接到 SharePoint 网站的列表。

在 ET，您不能设置任何与 **ListDataFormat** 对象关联的属性，但可通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

``` JavaScript
/*下例显示了对活动工作簿的 Sheet1 中某个列表第三列设置的 MaxCharacters 属性。*/
function test(){
  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)

   Debug.Print (objListCol.ListDataFormat.MaxCharacters)
}
```

### ListDataFormat.MaxNumber

返回一个 **Variant** ，其中包含列表列中该字段允许的最大值。 **Variant** 类型，只读。

**语法**

**express.MaxNumber**

\*express是一个代表 ListDataFormat 对象的变量。

**说明**

如果未指定最大值数字，或 **Type** 属性设置使列的最大值不适用，则返回 **Nothing** 对象。

该属性仅用于链接到 SharePoint 网站的列表。

在 ET，您不能设置任何与 **ListDataFormat** 对象关联的属性，但可通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

``` JavaScript
/*下例显示了对活动工作簿的 Sheet1 中某个列表第三列设置的 MaxNumber 属性。*/
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)

   Debug.Print (objListCol.ListDataFormat.MaxNumber)
}
```

### ListDataFormat.MinNumber

返回一个 **Variant** ，其中包含列表列中该字段允许的最小值。该值可以是负的浮点值。 **Variant** 类型，只读。

**语法**

**express.MinNumber**

\*express是一个代表 ListDataFormat 对象的变量。

**说明**

如果还没有为该字段指定值，或者 **Type** 属性的设置导致最小值不适用于该列，则该属性将返回 **Nothing** 对象。

该属性仅用于链接到 SharePoint 网站的列表。

在 ET，您不能设置任何与 **ListDataFormat** 对象关联的属性，但可通过修改 SharePoint 网站上的列表来设置这些属性。

**示例**

``` JavaScript
/*以下示例显示了对活动工作簿的 Sheet1 中某个列表第三列设置的 MinNumber 属性。*/
function test(){
 let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)

   Debug.Print (objListCol.ListDataFormat.MinNumber)
}
```

### ListDataFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ListDataFormat 对象的变量。

### ListDataFormat.ReadOnly

如果对象以只读方式打开，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.ReadOnly**

\*express是一个代表 ListDataFormat 对象的变量。

**说明**

此属性仅用于链接到 SharePoint 网站的表。

**示例**

``` JavaScript
/*下例显示了对活动工作簿的 Sheet1 中某个表第三列设置的 ReadOnly 属性。*/
function test(){
   let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
   let objListCol = wrksht.ListObjects.Item(1).ListColumns.Item(3)

   Debug.Print (objListCol.ListDataFormat.ReadOnly)
}
```

### ListDataFormat.Type

返回 **XlListDataType** 值，它代表列表列的数据类型。

**语法**

**express.Type**

\*express是一个代表 ListDataFormat 对象的变量。

**说明**

此属性仅用于链接到 SharePoint 网站的列表。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ListDataFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotCache 对象

## 简介

代表数据透视表的缓存。

## 说明

PivotCache 对象是 PivotCaches 集合的成员。

## 方法

| 序号 | 名称                                             | 说明                                                                                                      |
|:----:|:-------------------------------------------------|:----------------------------------------------------------------------------------------------------------|
|  1   | [CreatePivotTable](#PivotCache.CreatePivotTable) | 创建一个基于 PivotCache 对象的数据透视表。返回一个 PivotTable 对象。                                      |
|  2   | [MakeConnection](#PivotCache.MakeConnection)     | 为指定的数据透视表缓存建立连接。                                                                          |
|  3   | [Refresh](#PivotCache.Refresh)                   | 立即重新绘制指定的图表。                                                                                  |
|  4   | [ResetTimer](#PivotCache.ResetTimer)             | 重新设置指定的查询表或数据透视表的刷新计时器，使其时间间隔等于上次使用 RefreshPeriod 属性设置的时间间隔。 |
|  5   | [SaveAsODC](#PivotCache.SaveAsODC)               | 将数据透视表缓存的源保存为“Microsoft Office 数据连接”文件。                                               |

## 属性

| 序号 | 名称                                           | 说明                                                                                                                                                                                                                                                |
|:----:|:-----------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ADOConnection](#PivotCache.ADOConnection)     | 如果将数据透视表缓存连接到 OLE DB 数据源，则返回一个 ADO Connection 对象。 ADOConnection 属性表明 ET 与数据提供程序相连，允许用户在 ET 正在使用的同一会话的上下文内编写代码， ET 将该会话与 ADO（关系源）或 ADO MD（OLAP 源）一起使用。只读。       |
|  2   | [Application](#PivotCache.Application)         | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。                     |
|  3   | [BackgroundQuery](#PivotCache.BackgroundQuery) | 如果数据透视表的查询是异步执行（在后台执行）的，则为 True 。 Boolean 类型，可读写。                                                                                                                                                                 |
|  4   | [CommandText](#PivotCache.CommandText)         | 返回或设置指定数据源的命令串。 Variant 型，可读写。                                                                                                                                                                                                 |
|  5   | [CommandType](#PivotCache.CommandType)         | 返回或设置说明部分的下表中列出的 XlCmdType 常量之一。返回或设置的常量用于描述 CommandText 属性的值。默认值为 xlCmdSQL 。 XlCmdType 类型，可读写。                                                                                                   |
|  6   | [Connection](#PivotCache.Connection)           | 返回或设置包含下列某项的字符串：允许 ET 连接到 OLE DB 数据源的 OLE DB 设置；允许 ET 连接到 ODBC 数据源的 ODBC 设置；允许 ET 连接到 Web 数据源的 URL；或者文本文件的名称或路径，或是指定某个数据库或 Web 查询的文件名称或路径。 Variant 型，可读写。 |

## 成员方法

### PivotCache.CreatePivotTable

创建一个基于 **PivotCache** 对象的数据透视表。返回一个 **PivotTable** 对象。

**语法**

**express.CreatePivotTable ( *TableDestination* , *TableName* , *ReadData* , *DefaultVersion* )**

\*express是一个代表 PivotCache 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                                                                                                                                                                     |
|--------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *TableDestination* | 必选      | Variant  | 数据透视表目标区域（工作表中用于放置所生成的数据透视表的区域）左上角的单元格。目标区域必须位于工作簿（此工作簿包含由 expression 指定的 PivotCache 对象）的某个工作表中。 |
| *TableName*        | 可选      | Variant  | 新的数据透视表的名称。                                                                                                                                                   |
| *ReadData*         | 可选      | Variant  | 如果该值为 True，则创建一个包含外部数据库中所有记录的数据透视表高速缓存；此高速缓存可以很大。如果为 False，则允许在实际读取数据之前将某些字段设置为基于服务器的页字段。  |
| *DefaultVersion*   | 可选      | Variant  | 数据透视表的默认版本。                                                                                                                                                   |

**返回值**

PivotTable

**说明**

有关创建基于数据透视表高速缓存的数据透视表的另一种方法，请参阅 **PivotTables** 对象的 **Add** 方法。

**示例**

``` JavaScript
/*本示例在活动工作表的 A3 单元格上新建一个基于 OLAP 提供程序?（OLAP 提供程序：对特定类型的 OLAP 数据库提供访问功能的一组软件。该软件包括数据源驱动程序以及与数据库连接所必需的其他客户端软件。）的数据透视表高速缓存，然后基于该高速缓存新建一个数据透视表。*/
function test(){
　　　　With ActiveWorkbook.PivotCaches.Add(SourceType:=xlExternal)
　　　　    .Connection = _
 　　　　       "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National"
　　　　   .CommandType = xlCmdCube
　　　　   .CommandText = Array("Sales") 
 　　　　  .MaintainConnection = True
　　　　    .CreatePivotTable TableDestination:=Range("A3"), _
　　　　        TableName:= "PivotTable1"
　　　　End With
　　　　With ActiveSheet.PivotTables("PivotTable1")
　　　　    .SmallGrid = False
　　　　    .PivotCache.RefreshPeriod = 0
　　　　    With .CubeFields("[state]")
　　　　        .Orientation = xlColumnField
  　　　　      .Position = 1
　　　　    End With
　　　　    With .CubeFields("[Measures].[Count Of au_id]")
　　　　        .Orientation = xlDataField
  　　　　      .Position = 1
  　　　　  End With
　　　　End With}
```

``` JavaScript
/*本示例在活动工作表的 A3 单元格上，通过连接到 Microsoft Jet 上的 ADO 创建一个新的数据透视表高速缓存，然后再基于该高速缓存新建一个数据透视表。*/
function test(){
　　　　Dim cnnConn As ADODB.Connection
　　　　Dim rstRecordset As ADODB.Recordset
　　　　Dim cmdCommand As ADODB.Command

　　　　' Open the connection.
　　　　Set cnnConn = New ADODB.Connection
　　　　With cnnConn
 　　　　   .ConnectionString = _
　　　　        "Provider=Microsoft.Jet.OLEDB.4.0"
　　　　    .Open "C:\perfdate\record.mdb"
　　　　End With

　　　　' Set the command text.
　　　　Set cmdCommand = New ADODB.Command
　　　　Set cmdCommand.ActiveConnection = cnnConn
　　　　With cmdCommand
　　　　    .CommandText = "Select Speed, Pressure, Time From DynoRun"
　　　　    .CommandType = adCmdText
　　　　    .Execute
　　　　End With

　　　　' Open the recordset.
　　　　Set rstRecordset = New ADODB.Recordset
　　　　Set rstRecordset.ActiveConnection = cnnConn
　　　　rstRecordset.Open cmdCommand
　　　　
　　　　' Create a PivotTable cache and report.
　　　　Set objPivotCache = ActiveWorkbook.PivotCaches.Add( _
　　　　    SourceType:=xlExternal)
　　　　Set objPivotCache.Recordset = rstRecordset
　　　　With objPivotCache
　　　　    .CreatePivotTable TableDestination:=Range("A3"), _
　　　　        TableName:="Performance"
　　　　End With

　　　　With ActiveSheet.PivotTables("Performance")
　　　　    .SmallGrid = False
　　　　    With .PivotFields("Pressure")
　　　　        .Orientation = xlRowField
　　　　        .Position = 1
　　　　    End With
　　　　    With .PivotFields("Speed")
  　　　　      .Orientation = xlColumnField
　　　　        .Position = 1
   　　　　 End With
 　　　　   With .PivotFields("Time")
  　　　　      .Orientation = xlDataField
　　　　        .Position = 1
 　　　　   End With
　　　　End With

　　　　' Close the connections and clean up.
　　　　cnnConn.Close
　　　　Set cmdCommand = Nothing
　　　　Set rstRecordSet = Nothing
　　　　Set cnnConn = Nothing　　　　
}
```

### PivotCache.MakeConnection

为指定的数据透视表缓存建立连接。

**语法**

**express.MakeConnection ()**

\*express是一个代表 PivotCache 对象的变量。

**说明**

**MakeConnection** 方法可在缓存断开连接和用户想要重新建立连接之后使用。

如果没有连接缓存，各种对象和方法可能都会返回运行错误。使用该方法可确保在执行其他对象或方法之前已建立一个连接。

如果指定的数据透视表缓存的 **MaintainConnection** 属性已设置为 **False** ，指定的数据透视表缓存的 **SourceType** 属性尚未设置为 *xlExternal* ，或者连接不是 OLE DB 数据源连接，则该方法将生成一个运行时错误。

**示例**

``` JavaScript
/*以下示例确定是否将缓存连接到其源，如有必要，将与源建立连接。本示例假定数据透视表缓存位于活动工作表上。
*/
function UseMakeConnection() {
    let pvtCache = Application.ActiveWorkbook.PivotCaches().Item(1)
    // Handle run-time error if external source is not an OLE DB data source.
    try {
    // Check connection setting and make connection if necessary.
        if(pvtCache.IsConnected == true) {
            MsgBox("The MakeConnection method is not needed.")
        }
        else {
            pvtCache.MakeConnection()
            MsgBox("A connection has been made.")
        }
    }
    catch(exception) {
        MsgBox("The data source is not an OLE DB data source")
    }
}
```

### PivotCache.Refresh

立即重新绘制指定的图表。

**语法**

**express.Refresh ()**

\*express是一个代表 PivotCache 对象的变量。

**示例**

此示例刷新工作簿中第一个工作表上的第一个数据透视表的数据透视表缓存。

``` JavaScript
Worksheets.Item(1).PivotTables(1).PivotCache().Refresh()
```

### PivotCache.ResetTimer

重新设置指定的查询表或数据透视表的刷新计时器，使其时间间隔等于上次使用 **RefreshPeriod** 属性设置的时间间隔。

**语法**

**express.ResetTimer ()**

\*express是一个代表 PivotCache 对象的变量。

### PivotCache.SaveAsODC

将数据透视表缓存的源保存为“Microsoft Office 数据连接”文件。

**语法**

**express.SaveAsODC ( *ODCFileName* , *Description* , *Keywords* )**

\*express是一个代表 PivotCache 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                   |
|---------------|-----------|----------|----------------------------------------|
| *ODCFileName* | 必选      | String   | 保存文件的位置。                       |
| *Description* | 可选      | Variant  | 要保存到文件中的说明。                 |
| *Keywords*    | 可选      | Variant  | 可用于搜索该文件的以空格分隔的关键字。 |

**示例**

下例将缓存的源保存为名为“ODCFile”的 ODC 文件。此示例假定数据透视表缓存位于活动工作表上。

``` JavaScript
function UseSaveAsODC() {
　　　　Application.ActiveWorkbook.PivotCaches().Item(1).SaveAsODC ("ODCFile")                      
}
```

## 成员属性

### PivotCache.ADOConnection

如果将数据透视表缓存连接到 OLE DB 数据源，则返回一个 **ADO Connection** 对象。 **ADOConnection** 属性表明 ET 与数据提供程序相连，允许用户在 ET 正在使用的同一会话的上下文内编写代码， ET 将该会话与 ADO（关系源）或 ADO MD（OLAP 源）一起使用。只读。

**语法**

**express.ADOConnection**

\*express是一个代表 PivotCache 对象的变量。

**说明**

**ADOConnection** 属性仅供具有 OLE DB 数据源的会话使用。如果没有 ADO 会话，查询将导致运行时错误。

**ADOConnection** 属性可用于具有 ADO 的任何基于 OLE DB 的缓存。 **ADO Connection** 对象可与 ADO MD 一起使用，以查找有关缓存所基于的 OLAP 多维数据集的信息。

**示例**

``` JavaScript
/*本示例将 ADO DB Connection 对象设置为数据透视表缓存的 ADOConnection 属性。本示例假设数据透视表位于活动工作表上。 */
function test(){
　　　　Sub UseADOConnection()

　　　　    Dim ptOne As PivotTable
 　　　　   Dim cmdOne As New ADODB.Command
　　　　    Dim cfOne As CubeField

  　　　　  Set ptOne = Sheet1.PivotTables(1)
 　　　　   ptOne.PivotCache.MaintainConnection = True
  　　　　  Set cmdOne.ActiveConnection = ptOne.PivotCache.ADOConnection

  　　　　  ptOne.PivotCache.MakeConnection

    ' Create a set.
    cmdOne.CommandText = "Create Set [Warehouse].[My Set] as '{[Product].[All Products].Children}'"
    cmdOne.CommandType = adCmdUnknown
    cmdOne.Execute

   　　　　 ' Add a set to the CubeField.
　　　　   Set cfOne = ptOne.CubeFields.AddSet("My Set", "My Set")

　　　　End Sub
}
```

``` JavaScript
/*本示例将添加一个计算成员，并假设数据透视表位于活动工作表上。 */
function test(){
　　　　Sub AddMember()

  　　　　  Dim cmd As New ADODB.Command

   　　　　 If Not ActiveSheet.PivotTables(1).PivotCache.IsConnected Then
    　　　　    ActiveSheet.PivotTables(1).PivotCache.MakeConnection
  　　　　  End If

  　　　　  Set cmd.ActiveConnection = ActiveSheet.PivotTables(1).PivotCache.ADOConnection

 　　　　   ' Add a calculated member.
 　　　　   cmd.CommandText = "CREATE MEMBER [Warehouse].[Product].[All Products].[Drink and Non-Consumable] AS '[Product].[All Products].[Drink] + [Product].[All Products].[Non-Consumable]'"
   　　　　 cmd.CommandType = adCmdUnknown
  　　　　  cmd.Execute

   　　　　 ActiveSheet.PivotTables(1).PivotCache.Refresh

　　　　End Sub
}
```

### PivotCache.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotCache 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
　　　　let myObject = ActiveWorkbook
       if(myObject.Application.Value == "ET") {
 　　　　   MsgBox("This is an ET Application object.")
　　　　}
　　　　else {
　　　　    MsgBox("This is not an ET Application object.")
　　　　}
}
```

### PivotCache.BackgroundQuery

如果数据透视表的查询是异步执行（在后台执行）的，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.BackgroundQuery**

\*express是一个代表 PivotCache 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性是只读的，并且始终返回 **False** 。

**示例**

此示例使第一张工作表上的第一张数据透视表的查询在后台执行。

``` JavaScript
Worksheets.Item(1).PivotTables("Pivot1").PivotCache.BackgroundQuery = true
```

### PivotCache.CommandText

返回或设置指定数据源的命令串。 **Variant** 型，可读写。

**语法**

**express.CommandText**

\*express是一个代表 PivotCache 对象的变量。

**说明**

对于 OLE DB 数据源， **CommandType** 属性描述了 **CommandText** 属性的值。

对于 ODBC 源，设置 **CommandText** 将导致刷新数据。

### PivotCache.CommandType

返回或设置说明部分的下表中列出的 **XlCmdType** 常量之一。返回或设置的常量用于描述 **CommandText** 属性的值。默认值为 **xlCmdSQL** 。 **XlCmdType** 类型，可读写。

**语法**

**express.CommandType**

\*express是一个代表 PivotCache 对象的变量。

**说明**

| XlCmdType 可为以下这些 XlCmdType 常量之一。                                                                                                                                                                                                                                                                                                                                                  |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| xlCmdCube。包含一个 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的多维数据集?（多维数据集：一种 OLAP 数据结构。多维数据集包含维度，如“国家/地区/省（或市/自治区）/市（或县）”，还包括数据字段，如“销售额”。维度将各种类型的数据组织到带有明细数据级别的分层结构中，而数据字段度量数量。）名称。 |
| xlCmdDefault。包含 OLE DB 提供者可识别的命令文本。                                                                                                                                                                                                                                                                                                                                           |
| xlCmdSql。包含一个 SQL 语句。                                                                                                                                                                                                                                                                                                                                                                |
| xlCmdTable。包含用于访问 OLE DB 数据源的表名称。                                                                                                                                                                                                                                                                                                                                             |

只有当查询表或数据透视表缓存的 **QueryType** 属性的值为 **xlOLEDBQuery** 时，才可以设置 **CommandType** 属性。

当 **CommandType** 属性的值为 **xlCmdCube** 时，如果没有与查询表相关联的数据透视表，则不能更改该值。

**示例**

``` JavaScript
/*本示例设置第一张查询表的 ODBC 数据源的命令串。该命令串是一个 SQL 语句。*/
function test(){
　　　　let qtQtrResults = Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)
　　　　qtQtrResults.CommandType = xlCmdSql
　　　　qtQtrResults.CommandText = "Select ProductID From Products Where ProductID < 10"
　　　　qtQtrResults.Refresh()
}
```

### PivotCache.Connection

返回或设置包含下列某项的字符串：允许 ET 连接到 OLE DB 数据源的 OLE DB 设置；允许 ET 连接到 ODBC 数据源的 ODBC 设置；允许 ET 连接到 Web 数据源的 URL；或者文本文件的名称或路径，或是指定某个数据库或 Web 查询的文件名称或路径。 **Variant** 型，可读写。

**语法**

**express.Connection**

\*express是一个代表 PivotCache 对象的变量。

**说明**

在使用 脱机多维数据集文件 ? （脱机多维数据集文件：创建于硬盘或网络共享位置上的文件，用于存储数据透视表或数据透视图的 OLAP 源数据。脱机多维数据集文件允许用户在断开与 OLAP 服务器的连接后继续进行操作。） 时，请将 **UseLocalConnection** 属性设置为 **True** ，并使用 **LocalConnection** 属性，而不是用 **Connection** 属性。

另外，也可以通过选择 Microsoft ActiveX 数据对象 (ADO) 库直接访问数据源。

**示例**

``` JavaScript
/*此示例在活动工作表的 A3 单元格上新建一个基于 OLAP 提供程序?（OLAP 提供程序：对特定类型的 OLAP 数据库提供访问功能的一组软件。该软件包括数据源驱动程序以及与数据库连接所必需的其他客户端软件。）的数据透视表高速缓存，然后基于该高速缓存新建一个数据透视表。 */
function test(){
　　　　let pi = ActiveWorkbook.PivotCaches().Add(xlExternal)
   　　　　 pi.Connection = ("OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National")
   　　　　 pi.MaintainConnection = true
   　　　　 pi.CreatePivotTable(Range("A3"),"PivotTable1")
                                        
　　　　let pt = ActiveSheet.PivotTables("PivotTable1")
    　　　　pt.SmallGrid = false
   　　　　 pt.PivotCache.RefreshPeriod = 0
                                        
  　　　　  let cu = pt.CubeFields("[state]")
     　　　　   cu.Orientation = xlColumnField
    　　　　    cu.Position = 0
                                        
  　　　　  let cub = pt.CubeFields("[Measures].[Count Of au_id]")
     　　　　   cub.Orientation = xlDataField
    　　　　    cub.Position = 0
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotCache](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotFormula 对象

## 简介

代表在数据透视表中用于计算的公式。

## 说明

本对象及其相关属性和方法对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效，这是因为它不支持计算字段和计算项。

使用 PivotFormulas ( *index* )（其中 *index* 是公式左侧的公式号或字符串）可返回 PivotFormula 对象。下例将更改公式一（第一张工作表的第一个数据透视表中）的索引号，使其在公式二计算完毕后再进行计算。

``` JavaScript
Worksheets.Item(1).PivotTables(1).PivotFormulas.Item(1).Index = 2
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
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Delete">Delete</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>删除对象。</p></td>
</tr>
</tbody>
</table>

## 属性

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
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 <span>Application</span> 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Formula">Formula</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 String 值，它代表 A1 样式表示法和宏语言中的对象的公式。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Index">Index</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 Long 值，它代表 <span>PivotFormula</span> 对象在 <span>PivotFormulas</span> 集合中的索引号。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.StandardFormula">StandardFormula</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 String 值，该值指定使用标准英语（美国）格式的公式。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Value">Value</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 String 值，它代表数据透视表中指定的公式的名称。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PivotFormula.Delete

table { word-break:break-all; }

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 PivotFormula 对象的变量。

## 成员属性

### PivotFormula.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotFormula 对象的变量。

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

### PivotFormula.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotFormula 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotFormula.Formula

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

**语法**

**express.Formula**

\*express是一个代表 PivotFormula 对象的变量。

**说明**

table { word-break:break-all; }

此属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回空字符串。如果单元格包含公式， **Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

### PivotFormula.Index

table { word-break:break-all; }

返回或设置一个 **Long** 值，它代表 **PivotFormula** 对象在 **PivotFormulas** 集合中的索引号。

**语法**

**express.Index**

\*express是一个代表 PivotFormula 对象的变量。

### PivotFormula.Parent

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotFormula 对象的变量。

### PivotFormula.StandardFormula

返回或设置一个 **String** 值，该值指定使用标准英语（美国）格式的公式。可读写。

**语法**

**express.StandardFormula**

\*express是一个代表 PivotFormula 对象的变量。

**说明**

**StandardFormula** 属性主要影响具有日期或数字格式的项目名称。该属性提供了一种方法可指定或查询给定计算项的公式。

**StandardFormula** 属性是国际通用的， **Formula** 属性则不是。

**示例**

本示例向“Decimals”字段中添加 10，并将其显示为数据字段中的计算项。本示例假定数据透视表位于活动工作簿上，并且标题为“Decimals”的字段位于模拟运算表中。

``` JavaScript
function UseStandardFomula() {
    let pvtTable = ActiveSheet.PivotTables(1)
    // Change calculated field of decimals by adding '10'.
    pvtTable.CalculatedFields().Item(1).StandardFormula = "Decimals + 10"
}
```

### PivotFormula.Value

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表数据透视表中指定的公式的名称。

**语法**

**express.Value**

\*express是一个代表 PivotFormula 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotFormula](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# RectangularGradient 对象

## 简介

RectangularGradient 对象沿特定角度以线性方式在一系列颜色间转换。

## 说明

|                                                   |
|---------------------------------------------------|
| 使用 RectangularGradient 对象时，请考虑以下几点： |

- 试图访问不具有现有渐变填充的 Interior 对象的 Gradient 属性会引起运行时错误。访问 Gradient 属性之前请注意 ` Interior.Pattern ` 属性。
- 如果将 Interior.Pattern 从渐变类型更改为非渐变类型，Gradient 对象将采用默认值进行填充。

## 属性

| 序号 | 名称                                                    | 说明                                                                                                                                                               |
|:----:|:--------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#RectangularGradient.Application)         | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [ColorStops](#RectangularGradient.ColorStops)           | 返回 RectangularGradient 对象的 ColorStops 集合。只读。                                                                                                            |
|  3   | [Creator](#RectangularGradient.Creator)                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Parent](#RectangularGradient.Parent)                   | 返回指定对象的父对象。只读。                                                                                                                                       |
|  5   | [RectangleBottom](#RectangularGradient.RectangleBottom) | 代表渐变填充收敛到的点或矢量。可读/写。                                                                                                                            |
|  6   | [RectangleLeft](#RectangularGradient.RectangleLeft)     | 代表渐变填充收敛到的点或矢量。可读/写。                                                                                                                            |
|  7   | [RectangleRight](#RectangularGradient.RectangleRight)   | 代表渐变填充收敛到的点或矢量。可读/写。                                                                                                                            |
|  8   | [RectangleTop](#RectangularGradient.RectangleTop)       | 代表渐变填充收敛到的点或矢量。可读/写。                                                                                                                            |

## 成员属性

### RectangularGradient.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 RectangularGradient 对象的变量。

### RectangularGradient.ColorStops

返回 **RectangularGradient** 对象的 **ColorStops** 集合。只读。

**语法**

**express.ColorStops**

\*express是一个代表 RectangularGradient 对象的变量。

### RectangularGradient.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 RectangularGradient 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。Creator 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### RectangularGradient.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 RectangularGradient 对象的变量。

### RectangularGradient.RectangleBottom

代表渐变填充收敛到的点或矢量。可读/写。

**语法**

**express.RectangleBottom**

\*express是一个代表 RectangularGradient 对象的变量。

**说明**

与 RectangleLeft、RectangleRight 和 RectangleTop 一起使用。下表中列出了有效值。

| 属性            | 值  |
|-----------------|-----|
| RectangleLeft   | 0-1 |
| RectangleRight  | 0-1 |
| RectangleTop    | 0-1 |
| RectangleBottom | 0-1 |

### RectangularGradient.RectangleLeft

代表渐变填充收敛到的点或矢量。可读/写。

**语法**

**express.RectangleLeft**

\*express是一个代表 RectangularGradient 对象的变量。

**说明**

与 RectangleRight、RectangleTop 和 RectangleBottom 一起使用。下表中列出了有效值。

| 属性            | 值  |
|-----------------|-----|
| RectangleLeft   | 0-1 |
| RectangleRight  | 0-1 |
| RectangleTop    | 0-1 |
| RectangleBottom | 0-1 |

### RectangularGradient.RectangleRight

代表渐变填充收敛到的点或矢量。可读/写。

**语法**

**express.RectangleRight**

\*express是一个代表 RectangularGradient 对象的变量。

**说明**

与 RectangleLeft、RectangleTop 和 RectangleBottom 一起使用。下表中列出了有效值。

| 属性            | 值  |
|-----------------|-----|
| RectangleLeft   | 0-1 |
| RectangleRight  | 0-1 |
| RectangleTop    | 0-1 |
| RectangleBottom | 0-1 |

### RectangularGradient.RectangleTop

代表渐变填充收敛到的点或矢量。可读/写。

**语法**

**express.RectangleTop**

\*express是一个代表 RectangularGradient 对象的变量。

**说明**

与 RectangleLeft、RectangleRight 和 RectangleBottom 一起使用。下表中列出了有效值。

| 属性            | 值  |
|-----------------|-----|
| RectangleLeft   | 0-1 |
| RectangleRight  | 0-1 |
| RectangleTop    | 0-1 |
| RectangleBottom | 0-1 |

> WPS 开发文档： [WPS 基础接口/表格 API 参考/RectangularGradient](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SeriesCollection 对象

## 简介

指定的图表或图表组中所有 Series 对象的集合。

## 说明

使用 SeriesCollection 方法可返回 SeriesCollection 集合。

## 方法

| 序号 | 名称                                     | 说明                                               |
|:----:|:-----------------------------------------|:---------------------------------------------------|
|  1   | [Add](#SeriesCollection.Add)             | 向 SeriesCollection 集合添加一个或多个新数据系列。 |
|  2   | [Extend](#SeriesCollection.Extend)       | 向已存在的系列集合中添加新的数据点。               |
|  3   | [Item](#SeriesCollection.Item)           | 从集合中返回一个对象。                             |
|  4   | [NewSeries](#SeriesCollection.NewSeries) | 创建新系列。返回代表该新系列的 Series 对象。       |
|  5   | [Paste](#SeriesCollection.Paste)         | 将剪贴板中的数据粘贴到指定的系列集合中。           |

## 属性

| 序号 | 名称                                         | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#SeriesCollection.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#SeriesCollection.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#SeriesCollection.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#SeriesCollection.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### SeriesCollection.Add

向 **SeriesCollection** 集合添加一个或多个新数据系列。

**语法**

**express.Add ( *Source* , *Rowcol* , *SeriesLabels* , *CategoryLabels* , *Replace* )**

\*express是一个代表 SeriesCollection 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                                                                                                               |
|------------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Source*         | 必选      | Variant  | 作为 Range 对象的新数据。                                                                                                                                                          |
| *Rowcol*         | 可选      | XlRowCol | 指定新值是位于指定区域的行中还是列中。                                                                                                                                             |
| *SeriesLabels*   | 可选      | Variant  | 如果第一行或第一列包含数据系列的名称，则为 True。如果第一行或第一列包含数据系列的第一个数据点，则为 False。如果省略此参数，ET 就尝试根据第一行或第一列中的内容确定系列名称的位置。 |
| *CategoryLabels* | 可选      | Variant  | 如果第一行或第一列包含分类标签的名称，则为 True。如果第一行或第一列包含数据系列的第一个数据点，则为 False。如果省略此参数，ET 就尝试根据第一行或第一列中的内容确定分类标签的位置。 |
| *Replace*        | 可选      | Variant  | 如果 CategoryLabels 为 True 且 Replace 为 True，那么指定的分类将替换当前系列中存在的分类。如果 Replace 为 False，现有的分类将保留。默认值为 False。                                |

**返回值**

一个代表新数据系列的 Series 对象。

**说明**

实际上，此方法不会像“对象浏览器”中所述的那样返回 **Series** 对象。此方法对数据透视表无效。

**示例**

``` JavaScript
//本示例在 Chart1 中新建一个数据系列。新系列的数据源是 Sheet1 上的单元格区域 B1:B10。
Application.Charts.Item("Chart1").SeriesCollection().Add(Application.ActiveWorkbook.Worksheets.Item("Sheet1").Range("B1:B10"))

//本示例在工作表 Sheet1 上的嵌入式图表中新建一个数据系列。
function test()
{
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()
    Application.ActiveChart.SeriesCollection().Add(Application.Worksheets.Item("Sheet1").Range("B1:B10"))
}
```

### SeriesCollection.Extend

向已存在的系列集合中添加新的数据点。

**语法**

**express.Extend ( *Source* , *Rowcol* , *CategoryLabels* )**

\*express是一个代表 SeriesCollection 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                                                                                                   |
|------------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Source*         | 必选      | Variant  | 要以 Range 对象形式添加到 SeriesCollection 对象的新数据。                                                                                                              |
| *Rowcol*         | 可选      | Variant  | 指定新值是位于给定区域源的行中还是列中。可以是下列 XlRowCol 常量之一：xlRows 或 xlColumns。如果省略本参数，ET 将依据选定区域的大小和方向，或数组的维度确定数值的位置。 |
| *CategoryLabels* | 可选      | Variant  | 如果为 True，则第一行或列包含类别标签的名称。如果为 False，则第一行或列包含系列的第一个数据点。如果省略该参数，则 ET 将依据第一行或列的内容确定类别标签的位置。        |

**返回值**

Variant

**说明**

此方法对数据透视图无效。

**示例**

``` JavaScript
//本示例将工作表 Sheet1 中单元格区域 B1:B6 中的数据添加到 Chart1 中，以延伸其中的系列。
Application.Charts.Item("Chart1").SeriesCollection().Extend(Application.Worksheets.Item("Sheet1").Range("B1:B6"))
```

### SeriesCollection.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SeriesCollection 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 Series 对象。

**示例**

``` JavaScript
//本示例对 Chart1 中的趋势线向前和向后延伸的单位数进行设置。本示例应在包含单个带趋势线系列的二维柱形图上运行。
function test()
{
    let trendlines = Application.Charts.Item("Chart1").SeriesCollection().Item(1).Trendlines().Item(1)
    trendlines.Forward = 5
    trendlines.Backward = .5
}
```

### SeriesCollection.NewSeries

创建新系列。返回代表该新系列的 **Series** 对象。

**语法**

**express.NewSeries ()**

\*express是一个代表 SeriesCollection 对象的变量。

**返回值**

Series

**说明**

本方法对数据透视图无效。

**示例**

``` JavaScript
//本示例向第一个图表中添加新系列。
let ns = Application.Charts.Item(1).SeriesCollection().NewSeries()
```

### SeriesCollection.Paste

将剪贴板中的数据粘贴到指定的系列集合中。

**语法**

**express.Paste ( *Rowcol* , *SeriesLabels* , *CategoryLabels* , *Replace* , *NewSeries* )**

\*express是一个代表 SeriesCollection 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                              |
|------------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Rowcol*         | 可选      | XlRowCol | 指定与特定数据系列对应的值是在行中还是在列中。                                                                                                                                                                    |
| *SeriesLabels*   | 可选      | Variant  | 如果为 True，则用每一行的第一列（或每一列的第一行）中的单元格内容作为该行（或列）中数据系列的名称。如果为 False，则用每一行的第一列（或每一列的第一行）中的单元格内容作为数据系列的第一个数据点。默认值为 False。 |
| *CategoryLabels* | 可选      | Variant  | 如果为 True，则用选定区域的第一行（或第一列）中的内容作为图表的分类。如果为 False，则用选定区域的第一行（或第一列）中的内容作为图表的第一个数据系列。默认值为 False。                                             |
| *Replace*        | 可选      | Variant  | 如果为 True，则在用所复制区域的信息替换现有的分类时应用分类。如果为 False，则插入新分类而不替换任何原有分类。默认值为 True。                                                                                      |
| *NewSeries*      | 可选      | Variant  | 如果为 True，则将数据作为新系列粘贴。如果为 False，则将数据作为现有系列的新数据点粘贴。默认值为 True。                                                                                                            |

**返回值**

Variant

**示例**

``` JavaScript
//本示例将剪贴板中的数据粘贴到 Chart1 上的新数据系列中。
function test()
{
    Application.Worksheets.Item("Sheet1").Range("C1:C5").Copy()
    Application.Charts.Item("Chart1").SeriesCollection().Paste()
}
```

## 成员属性

### SeriesCollection.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 SeriesCollection 对象的变量。

**示例**

``` JavaScript
//本示例显示一条有关创建 myObject 的应用程序的消息。
function test()
{
    if(Application.ActiveWorkbook.Application.Value == "ET") 
    {
    alert("This is an ET Application object.")
    }
    else
    {
    alert("This is not an ET Application object.")
    }
}
```

### SeriesCollection.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 SeriesCollection 对象的变量。

### SeriesCollection.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 SeriesCollection 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### SeriesCollection.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SeriesCollection 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/SeriesCollection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Shape 对象

## 简介

代表绘图层中的对象，例如自选图形、任意多边形、OLE 对象或图片。

## 说明

Shape 对象是 Shapes 集合的成员。 Shapes 集合包含某个工作簿中的所有形状。

注释： 有三个代表形状的对象： Shapes 集合，它代表工作簿中所有的形状； ShapeRange 集合，它代表工作簿中形状的指定子集（例如， ShapeRange 对象可以代表工作簿中的形状一和形状四，或者，可以代表工作簿中所有选定的形状）； Shape 对象，它代表文档中的某一个形状。如果您需要同时处理几个形状，或处理选定区域中的多个形状，请使用 ShapeRange 集合。

使用 Shape 对象

以下各节说明了如何：

- 返回与连接符的端点相连的形状。
- 返回新建的任意多边形。
- 返回组中的单个形状。
- 返回新组成的形状组。
- 返回现有的形状。
- 返回选定区域中的形状。返回与连接符的端点相连的形状

要返回一个代表连接符所连接形状之一的 Shape 对象，请使用 BeginConnectedShape 或 EndConnectedShape 属性。

1\. 返回新建的任意多边形

使用 BuildF reeform 和 AddNodes 方法可定义一个新任意多边形的几何特性，使用 ConvertToShape 方法可创建任意多边形并返回代表它的 Shape 对象。

2\. 返回组中的单个形状

使用 GroupItems ( *index* )（其中 *index* 是形状的名称或组中的索引号）可返回一个代表一组形状中某一形状的 Shape 对象。

3\. 返回新组成的形状组

使用 Group 或 Regroup 方法可将一系列形状分成一组并返回一个 Shape 对象，该对象代表新形成的组。在形成了一个组之后，您可以按您处理其他任何形状的方法来处理该组

4\. 返回现有的形状

使用 Shapes ( *index* )（其中 *index* 是形状名称或索引号）可返回代表某个形状的 Shape 对象。

5\. 返回选定区域中的形状

使用 ` Selection.ShapeRange ` ( *index* )（其中 *index* 是形状名称或索引号）可返回一个代表选定区域中的形状的 Shape 对象。

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
<td style="text-align: left;" data-valian="middle"><a href="#Shape.Apply">Apply</a></td>
<td style="text-align: left;" data-valian="middle"><p>应用通过 PickUp 方法复制的指定形状格式。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#Shape.Copy">Copy</a></td>
<td style="text-align: left;" data-valian="middle"><p>将对象复制到剪贴板。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#Shape.Cut">Cut</a></td>
<td style="text-align: left;" data-valian="middle"><p>将对象剪切到剪贴板。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#Shape.Delete">Delete</a></td>
<td style="text-align: left;" data-valian="middle"><p>删除对象。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#Shape.Select">Select</a></td>
<td style="text-align: left;" data-valian="middle"><p>选择对象。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#Shape.Ungroup">Ungroup</a></td>
<td style="text-align: left;" data-valian="middle"><p>取消指定形状或者形状区域中组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。</p>
<p>返回值： 一个 ShapeRange 对象，它代表取消组合的形状。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#Shape.ZOrder">ZOrder</a></td>
<td style="text-align: left;" data-valian="middle"><p>将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-次序中的位置）。</p></td>
</tr>
</tbody>
</table>

## 属性

| 序号 | 名称                                      | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Shape.Application)         | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [AutoShapeType](#Shape.AutoShapeType)     | 返回或设置指定的 Shape 或 ShapeRange 对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。 MsoAutoShapeType 类型，可读写。                                                                           |
|  3   | [BottomRightCell](#Shape.BottomRightCell) | 返回一个 Range 对象，它代表位于该对象右下角下方的单元格。只读。                                                                                                                                                                 |
|  4   | [Height](#Shape.Height)                   | 返回或设置一个 Single 值，它代表对象的高度（以磅为单位）。                                                                                                                                                                      |
|  5   | [ID](#Shape.ID)                           | 返回一个 Long 值，它代表指定对象的类型。                                                                                                                                                                                        |
|  6   | [Left](#Shape.Left)                       | 返回或设置 Single 值，它代表从对象左边缘到工作表的 A 列左边缘或到图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                     |
|  7   | [Name](#Shape.Name)                       | 返回或设置一个 String 值，它代表对象的名称。                                                                                                                                                                                    |
|  8   | [Rotation](#Shape.Rotation)               | 返回或设形状的旋转角度（以度为单位）。 Single 型，可读写。                                                                                                                                                                      |
|  9   | [Top](#Shape.Top)                         | 返回或设置一个 Single 值，它代表形状范围内最上面的形状的上边缘到工作表上边缘的距离（以磅为单位）。                                                                                                                              |
|  10  | [TopLeftCell](#Shape.TopLeftCell)         | 返回一个 Range 对象，它代表位于指定对象左上角下方的单元格。只读。                                                                                                                                                               |
|  11  | [Type](#Shape.Type)                       | 返回或设置一个代表形状类型的 MsoShapeType 值。                                                                                                                                                                                  |
|  12  | [Visible](#Shape.Visible)                 | 返回或设置一个 MsoTriState 值，它确定对象是否可见。可读写。                                                                                                                                                                     |
|  13  | [Width](#Shape.Width)                     | 返回或设置一个 Single 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |

## 成员方法

### Shape.Apply

应用通过 **PickUp** 方法复制的指定形状格式。

**语法**

**express.Apply ()**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*此示例复制 myDocument 上第一个形状的格式，然后将复制的格式应用到第二个形状上。*/
function test()
{
    Application.Worksheets.Item(1).Shapes.Item(1).PickUp()
    Application.Worksheets.Item(1).Shapes.Item(2).Apply()
}
```

### Shape.Copy

将对象复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 Shape 对象的变量。

### Shape.Cut

将对象剪切到剪贴板。

**语法**

**express.Cut ()**

\*express是一个代表 Shape 对象的变量。

### Shape.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Shape 对象的变量。

### Shape.Select

选择对象。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                                              |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

### Shape.Ungroup

取消指定形状或者形状区域中组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。

返回值： 一个 **ShapeRange** 对象，它代表取消组合的形状。

**语法**

**express.Ungroup ()**

\*express是一个代表 Shape 对象的变量。

**说明**

由于一组形状被视为单个对象，因此对形状进行组合或取消组合操作时，将会更改 **Shapes** 集合中的项目数，并更改集合中受影响的项目之后的项目索引号。

**示例**

``` JavaScript
/*此示例取消 myDocument 中的所有形状组合，并分解所有的图片和 OLE 对象。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    for(let i=1;i <= myDocument.Shapes.Count;i++) {
        myDocument.Shapes.Item(i).Ungroup()
    }
}

/*此示例取消 myDocument 中的所有形状组合，但并不取消文档中图片和 OLE 对象的组合。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    for(let i=1;i <= myDocument.Shapes.Count;i++) {
        if(myDocument.Shapes.Item(i).Type == msoGroup){
            myDocument.Shapes.Item(i).Ungroup()
        }
    }
}
```

### Shape.ZOrder

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-次序中的位置）。

**语法**

**express.ZOrder ( *ZOrderCmd* )**

\*express是一个代表 Shape 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型     | 说明                                       |
|-------------|-----------|--------------|--------------------------------------------|
| *ZOrderCmd* | 必选      | MsoZOrderCmd | 指定根据其他形状将指定的形状移到什么位置。 |

**说明**

**Shape** 对象是 **Shapes** 集合的成员。 **Shapes** 集合包含某个工作簿中的所有形状。

**注释：** 有三个代表形状的对象： **Shapes** 集合，它代表工作簿中所有的形状； **ShapeRange** 集合，它代表工作簿中形状的指定子集（例如， **ShapeRange** 对象可以代表工作簿中的形状一和形状四，或者，可以代表工作簿中所有选定的形状）； **Shape** 对象，它代表文档中的某一个形状。如果您需要同时处理几个形状，或处理选定区域中的多个形状，请使用 **ShapeRange** 集合。

使用 Shape 对象

以下各节说明了如何：

- 返回与连接符的端点相连的形状。
- 返回新建的任意多边形。
- 返回组中的单个形状。
- 返回新组成的形状组。
- 返回现有的形状。
- 返回选定区域中的形状。

**1. 返回与连接符的端点相连的形状**

要返回一个代表连接符所连接形状之一的 **Shape** 对象，请使用 **BeginConnectedShape** 或 **EndConnectedShape** 属性。

2\. 返回新建的任意多边形

使用 **BuildFreeform** 和 **AddNodes** 方法可定义一个新任意多边形的几何特性，使用 **ConvertToShape** 方法可创建任意多边形并返回代表它的 **Shape** 对象。

3\. 返回组中的单个形状

使用 **GroupItems** ( *index* )（其中 *index* 是形状的名称或组中的索引号）可返回一个代表一组形状中某一形状的 **Shape** 对象。

4\. 返回新组成的形状组

使用 **Group** 或 **R egroup** 方法可将一系列形状分成一组并返回一个 **Shape** 对象，该对象代表新形成的组。在形成了一个组之后，您可以按您处理其他任何形状的方法来处理该组。

5\. 返回现有的形状

使用 **Shapes** ( *index* )（其中 *index* 是形状名称或索引号）可返回代表某个形状的 **Shape** 对象。

**6** . 返回选定区域中的形状

使用 ` Selection.ShapeRange ` ( *index* )（其中 *index* 是形状名称或索引号）可返回一个代表选定区域中的形状的 **Shape** 对象。

**示例**

## 成员属性

### Shape.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Shape 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    if(Application.ActiveWorkbook.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }else {
        alert("This is not an ET Application object.")
    }
}
```

### Shape.AutoShapeType

返回或设置指定的 **Shape** 或 **ShapeRange** 对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。 **MsoAutoShapeType** 类型，可读写。

**语法**

**express.AutoShapeType**

\*express是一个代表 Shape 对象的变量。

**说明**

改变一个形状的类型时，该形状保留其大小、颜色和其他属性。

使用 **ConnectorFormat** 对象的 **Type** 属性设置或返回连接符类型。

**示例**

``` JavaScript
/*本示例将 myDocument 中的所有 16 角星都替换为 32 角星。*/
function test()
{
    for(let i = 1; i <= Application.Worksheets.Item(1).Shapes.Count; i++) {
        if(Application.Worksheets.Item(1).Shapes.Item(i).AutoShapeType = msoShape16pointStar) {
            Application.Worksheets.Item(1).Shapes.Item(i).AutoShapeType = msoShape32pointStar
        }
    }
}
```

### Shape.BottomRightCell

返回一个 **Range** 对象，它代表位于该对象右下角下方的单元格。只读。

**语法**

**express.BottomRightCell**

\*express是一个代表 Shape 对象的变量。

### Shape.Height

返回或设置一个 **Single** 值，它代表对象的高度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 Shape 对象的变量。

### Shape.ID

返回一个 Long 值，它代表指定对象的类型。

**语法**

**express.ID**

\*express是一个代表 Shape 对象的变量。

**说明**

可将 ID 标签用作其他 HTML 文档或相同网页上的超链接引用。

### Shape.Left

返回或设置 **Single** 值，它代表从对象左边缘到工作表的 A 列左边缘或到图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 Shape 对象的变量。

### Shape.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Shape 对象的变量。

### Shape.Rotation

返回或设形状的旋转角度（以度为单位）。 **Single** 型，可读写。

**语法**

**express.Rotation**

\*express是一个代表 Shape 对象的变量。

**说明**

旋转量始终四舍五入到最接近的整数。

### Shape.Top

返回或设置一个 **Single** 值，它代表形状范围内最上面的形状的上边缘到工作表上边缘的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 Shape 对象的变量。

### Shape.TopLeftCell

返回一个 **Range** 对象，它代表位于指定对象左上角下方的单元格。只读。

**语法**

**express.TopLeftCell**

\*express是一个代表 Shape 对象的变量。

### Shape.Type

返回或设置一个代表形状类型的 **MsoShapeType** 值。

**语法**

**express.Type**

\*express是一个代表 Shape 对象的变量。

### Shape.Visible

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 Shape 对象的变量。

### Shape.Width

返回或设置一个 **Single** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 Shape 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Shape](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Style 对象

## 简介

代表区域的样式说明。

## 说明

Style 对象包含样式的所有属性（字体、数字格式、对齐方式，等等）。有几种内置样式，包括“常规”、“货币”和“百分比”。同时对多个单元格修改单元格格式属性时，使用 Style 对象是快捷高效的方法。

对于 Workbook 对象， Style 对象是 Styles 集合的成员。 Styles 集合包含该工作簿的所有已定义样式。

通过更改应用于单元格的样式的属性可更改单元格的外观。但要记住，更改样式的属性将影响所有以该样式格式化了的单元格。

样式按照名称的字母顺序排序。样式编号表明指定样式在样式名排序列表中的位置。 ` Styles(1) ` 是排序列表中的第一个样式，而 ` Styles(Styles.Count) ` 是最后一个。

有关创建和修改样式的详细信息，请参阅 Styles 对象。

使用 Style 属性可返回一个用于 Range 对象的 Style 对象。下例将“百分比”样式应用于 Sheet1 中的单元格区域 A1:A10。

``` JavaScript
Application.Worksheets.Item("Sheet1").Range("A1:A10").Style = "Percent"
```

使用 Styles ( *index* )（其中 *index* 是样式索引号或名称）可从工作簿的 Style 集合中返回一个 Styles 对象。下例通过设置活动工作簿中“常规”样式的 Bold 属性来更改该样式。

``` JavaScript
Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true
```

## 方法

| 序号 | 名称                    | 说明                    |
|:----:|:------------------------|:------------------------|
|  1   | [Delete](#Style.Delete) | 删除对象。返回Variant值 |

## 属性

| 序号 | 名称                                              | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddIndent](#Style.AddIndent)                     | 返回或设置一个 Boolean 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。                                                                                                                           |
|  2   | [Application](#Style.Application)                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  3   | [Borders](#Style.Borders)                         | 返回一个 Bor ders 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。                                                                                                                                        |
|  4   | [BuiltIn](#Style.BuiltIn)                         | 如果样式为内置样式，则为 True 。只读 Boolean 类型。                                                                                                                                                                             |
|  5   | [Creator](#Style.Creator)                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  6   | [Font](#Style.Font)                               | 返回一个 Font 对象，它代表指定对象的字体。                                                                                                                                                                                      |
|  7   | [FormulaHidden](#Style.FormulaHidden)             | 返回或设置一个 Boolean 值，它指明在工作表处于保护状态时是否隐藏公式。                                                                                                                                                           |
|  8   | [HorizontalAlignment](#Style.HorizontalAlignment) | 返回或设置一个 XlHAlign 值，它代表指定对象的水平对齐方式。                                                                                                                                                                      |
|  9   | [IncludeAlignment](#Style.IncludeAlignment)       | 如果样式包含 AddIndent 、 HorizontalAlignment 、 VerticalAlignment 、 WrapText 、 IndentLevel 和 Orientation 属性，则该属性值为 True 。 Boolean 类型，可读写。                                                                  |
|  10  | [IncludeBorder](#Style.IncludeBorder)             | 如果指定样式中包含 Color 、 ColorIndex 、 LineStyle 和 Weight 边框属性，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                             |
|  11  | [IncludeFont](#Style.IncludeFont)                 | 如果样式包含 Background 、 Bold 、 Color 、 ColorIndex 、 FontStyle 、 Italic 、 Name 、 Size 、 Strikethrough 、 Subscript 、 Superscript 和 Underline 字体属性，则此属性为 True 。 Boolean 类型，可读写。                     |
|  12  | [IncludeNumber](#Style.IncludeNumber)             | 如果样式中包含 NumberFormat 属性，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                   |
|  13  | [IncludePatterns](#Style.IncludePatterns)         | 如果指定样式中包含 Color 、 ColorIndex 、 InvertIfNegative 、 Pattern 、 PatternColor 和 PatternColorIndex 对象的内部属性，则该属性值为 True 。 Boolean 类型，可读写。                                                          |
|  14  | [IncludeProtection](#Style.IncludeProtection)     | 如果指定样式中包含 FormulaHidden 和 Locked 保护属性，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                |
|  15  | [IndentLevel](#Style.IndentLevel)                 | 返回或设置一个 Long 值，它代表样式的缩进量。                                                                                                                                                                                    |
|  16  | [Interior](#Style.Interior)                       | 返回一个 Inte rior 对象，它代表指定对象的内部。                                                                                                                                                                                 |
|  17  | [Locked](#Style.Locked)                           | 返回或设置一个 Boolean 值，它指明对象是否已被锁定。                                                                                                                                                                             |
|  18  | [MergeCells](#Style.MergeCells)                   | 如果样式包含合并的单元格，则为 True 。 Variant 型，可读写。                                                                                                                                                                     |
|  19  | [Name](#Style.Name)                               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  20  | [NameLocal](#Style.NameLocal)                     | 以用户语言返回或设置对象的名称。 String 型，只读。                                                                                                                                                                              |
|  21  | [NumberFormat](#Style.NumberFormat)               | 返回或设置一个 String 值，它代表对象的格式代码。                                                                                                                                                                                |
|  22  | [NumberFormatLocal](#Style.NumberFormatLocal)     | 以采用用户语言字符串的形式返回或设置一个 String 值，它代表对象的格式代码。                                                                                                                                                      |
|  23  | [Orientation](#Style.Orientation)                 | 返回或设置一个 XlOrientation 值，它代表文本方向。                                                                                                                                                                               |
|  24  | [Parent](#Style.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  25  | [ReadingOrder](#Style.ReadingOrder)               | 返回或设置指定对象的阅读次序。可为以下常量之一： xlRTL （从右到左）、 xlLTR （从左到右）或 xlContext 。 Long 类型，可读写。                                                                                                     |
|  26  | [ShrinkToFit](#Style.ShrinkToFit)                 | 返回或设置一个 Boolean 值，它指明文本是否可以自动收缩为适当尺寸以适应可用列宽。                                                                                                                                                 |
|  27  | [Value](#Style.Value)                             | 返回一个 String 值，它代表指定样式的名称。                                                                                                                                                                                      |
|  28  | [VerticalAlignment](#Style.VerticalAlignment)     | 返回或设置一个 XlVAlign 值，它代表指定对象的垂直对齐方式。                                                                                                                                                                      |
|  29  | [WrapText](#Style.WrapText)                       | 返回或设置一个 Boolean 值，它指明 ET 是否为对象中的文本自动换行。                                                                                                                                                               |

## 成员方法

### Style.Delete

删除对象。返回Variant值

**语法**

**express.Delete ()**

\*express是一个代表 Style 对象的变量。

## 成员属性

### Style.AddIndent

返回或设置一个 **Boolean** 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。

**语法**

**express.AddIndent**

\*express是一个代表 Style 对象的变量。

**说明**

如果将此属性的值设为 **True** ，那么在单元格中文本的对齐方式设为水平或垂直等距分布时，将自动缩进文本。

要将文本对齐方式设为等距分布，可在 Orientation 属性的值为 xlVertical 时，将 VerticalAlignment 属性设为 **xlVAlignDistributed** ；在 Orientation 属性的值为 xlHorizontal 时，将 HorizontalAlignment 属性设为 **xlHAlignDistributed** 。

### Style.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
function test(){
  /*本示例显示一条有关创建 myObject 的应用程序的消息。*/
  let myObject = Application.ActiveWorkbook
  if(myObject.Application.Value == "ET"){
      alert("This is an ET Application object.")
  }
  else{
      alert("This is not an ET Application object.")
  }
}
```

### Style.Borders

返回一个 **Bor ders** 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。

**语法**

**express.Borders**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*此示例将 Sheet1 中单元格 B2 的底部边框颜色设置为红色细边框。*/
function test(){

    let rng = Application.Worksheets.Item("Sheet1").Range("B2").Borders.Item(xlEdgeBottom)
    rng.LineStyle = xlContinuous
    rng.Weight = xlThin
    rng.ColorIndex = 3
                                
}
```

### Style.BuiltIn

如果样式为内置样式，则为 **True** 。只读 **Boolean** 类型。

**语法**

**express.BuiltIn**

\*express是一个代表 Style 对象的变量。

### Style.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Style 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Style.Font

返回一个 Font 对象，它代表指定对象的字体。

**语法**

**express.Font**

\*express是一个代表 Style 对象的变量。

### Style.FormulaHidden

返回或设置一个 **Boolean** 值，它指明在工作表处于保护状态时是否隐藏公式。

**语法**

**express.FormulaHidden**

\*express是一个代表 Style 对象的变量。

**说明**

请不要将此属性与 **Hidd en** 属性混淆。如果工作簿受保护，而工作表不受保护，将不会隐藏公式。只有在工作表受保护时，才会隐藏公式。

### Style.HorizontalAlignment

返回或设置一个 XlHAlign 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

\*express是一个代表 Style 对象的变量。

**说明**

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### Style.IncludeAlignment

如果样式包含 **AddIndent** 、 **HorizontalAlignment** 、 **VerticalAlignment** 、 **WrapText** 、 **IndentLevel** 和 **Orientation** 属性，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeAlignment**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入对齐格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeAlignment = true
```

### Style.IncludeBorder

如果指定样式中包含 **Color** 、 **ColorIndex** 、 **LineStyle** 和 **Weight** 边框属性，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeBorder**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入边框格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeBorder = true
```

### Style.IncludeFont

如果样式包含 **Background** 、 **Bold** 、 **Color** 、 **ColorIndex** 、 **FontStyle** 、 **Italic** 、 **Name** 、 **Size** 、 **Strikethrough** 、 **Subscript** 、 **Superscript** 和 **Underline** 字体属性，则此属性为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeFont**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入字体格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeFont = true
```

### Style.IncludeNumber

如果样式中包含 **NumberFormat** 属性，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeNumber**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入数字格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeNumber = true
```

### Style.IncludePatterns

如果指定样式中包含 **Color** 、 **ColorIndex** 、 **InvertIfNegative** 、 **Pattern** 、 **PatternColor** 和 **PatternColorIndex** 对象的内部属性，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludePatterns**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入图案格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludePatterns = true
```

### Style.IncludeProtection

如果指定样式中包含 **FormulaHidden** 和 **Locked** 保护属性，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeProtection**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*本示例在 Sheet1 的 A1 单元格样式中加入保护格式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeProtection = true
```

### Style.IndentLevel

返回或设置一个 **Long** 值，它代表样式的缩进量。

**语法**

**express.IndentLevel**

\*express是一个代表 Style 对象的变量。

**说明**

使用此属性将缩进量设为小于 0 或者大于 15 的数字，将导致发生错误。

### Style.Interior

返回一个 **Inte rior** 对象，它代表指定对象的内部。

**语法**

**express.Interior**

\*express是一个代表 Style 对象的变量。

### Style.Locked

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

\*express是一个代表 Style 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True** ；如果在工作表处于受保护状态时仍能修改对象，则返回 **False** 。

### Style.MergeCells

如果样式包含合并的单元格，则为 **True** 。 **Variant** 型，可读写。

**语法**

**express.MergeCells**

\*express是一个代表 Style 对象的变量。

### Style.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Style 对象的变量。

**示例**

``` JavaScript
/*首先用宏语言，然后用用户语言显示活动工作簿中的第一种样式的名称。*/
function test(){
  let sty = Application.ActiveWorkbook.Styles.Item(1)
  alert("The name of the style: " + sty.Name)
  alert("The localized name of the style: " + sty.NameLocal)
}  

/*显示活动工作簿的 Sheet1 中的默认 ListObject 对象的名称。*/
function Test(){  
    let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let oListObj = wrksht.ListObjects.Item(1)                                       
    alert(oListObj.Name)                                 
}
```

### Style.NameLocal

以用户语言返回或设置对象的名称。 **String** 型，只读。

**语法**

**express.NameLocal**

\*express是一个代表 Style 对象的变量。

**说明**

如果样式为内置样式，则此属性以当前系统所使用的地区语言返回样式的名称。

**示例**

``` JavaScript
function test(){
  /*此示例显示活动工作簿上样式一的原名称和本地化以后的名称。*/
  let sty = Application.ActiveWorkbook.Styles.Item(1)
  alert("The name of the style is " + sty.Name)
  alert("The localized name of the style is " + sty.NameLocal)
}
```

### Style.NumberFormat

返回或设置一个 **String** 值，它代表对象的格式代码。

**语法**

**express.NumberFormat**

\*express是一个代表 Style 对象的变量。

**说明**

格式代码与 **“设置单元格格式”** 对话框中的 **“格式代码”** 选项是同一个字符串。 **Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

### Style.NumberFormatLocal

以采用用户语言字符串的形式返回或设置一个 **String** 值，它代表对象的格式代码。

**语法**

**express.NumberFormatLocal**

\*express是一个代表 Style 对象的变量。

**说明**

**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

### Style.Orientation

返回或设置一个 **XlOrientation** 值，它代表文本方向。

**语法**

**express.Orientation**

\*express是一个代表 Style 对象的变量。

### Style.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Style 对象的变量。

### Style.ReadingOrder

返回或设置指定对象的阅读次序。可为以下常量之一： **xlRTL** （从右到左）、 **xlLTR** （从左到右）或 **xlContext** 。 **Long** 类型，可读写。

**语法**

**express.ReadingOrder**

\*express是一个代表 Style 对象的变量。

### Style.ShrinkToFit

返回或设置一个 **Boolean** 值，它指明文本是否可以自动收缩为适当尺寸以适应可用列宽。

**语法**

**express.ShrinkToFit**

\*express是一个代表 Style 对象的变量。

### Style.Value

返回一个 **String** 值，它代表指定样式的名称。

**语法**

**express.Value**

\*express是一个代表 Style 对象的变量。

### Style.VerticalAlignment

返回或设置一个 **XlVAlign** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

\*express是一个代表 Style 对象的变量。

**说明**

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### Style.WrapText

返回或设置一个 **Boolean** 值，它指明 ET 是否为对象中的文本自动换行。

**语法**

**express.WrapText**

\*express是一个代表 Style 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Style](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TableStyles 对象

## 简介

代表可应用于表格的样式。

## 说明

表格样式提供了一种为整个表格或数据透视图设置格式的方式。表格样式取代了用于为整个表格设置格式的现有的自动套用格式功能。

表格样式与自动套用格式在以下几个方面不同：

- 可以创建和重用自定义表格样式。
- 表格样式可处理主题。
- 如果更改文档主题的配色方案和/或字体方案，则将更改内置表格样式的外观。
- 当对象发生变化时，表格样式可以将样式重新应用于像数据透视表和表格之类的对象。该表格将记住应用于对象的样式，在添加、删除、隐藏和显示单元格时，表格将相应地进行重新显示。
- 表格样式在功能区中具有可见的用户界面。

## 方法

| 序号 | 名称                      | 说明                                           |
|:----:|:--------------------------|:-----------------------------------------------|
|  1   | [Add](#TableStyles.Add)   | 创建新的 TableStyle 对象，并将其添加到集合中。 |
|  2   | [Item](#TableStyles.Item) | 从集合中返回一个对象。                         |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                               |
|:----:|:----------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#TableStyles.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Count](#TableStyles.Count)             | 返回集合中对象的数目。只读 Long 类型。                                                                                                                             |
|  3   | [Creator](#TableStyles.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型                                                                                           |
|  4   | [Parent](#TableStyles.Parent)           | 返回指定对象的父对象。只读                                                                                                                                         |

## 成员方法

### TableStyles.Add

创建新的 **TableStyle** 对象，并将其添加到集合中。

**语法**

**express.Add ( *TableStyleName* )**

\*express是一个代表 TableStyles 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明           |
|------------------|-----------|----------|----------------|
| *TableStyleName* | 必选      | String   | 表样式的名称。 |

**返回值**

TableStyle

### TableStyles.Item

从集合中返回一个对象。

**语法**

**express.Item ( *index* )**

\*express是一个代表 TableStyles 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *index* | 必选      | Variant  | 对象的名称或索引号。 |

**返回值**

TableStyle

## 成员属性

### TableStyles.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 TableStyles 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### TableStyles.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 TableStyles 对象的变量。

### TableStyles.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型

**语法**

**express.Creator**

\*express是一个代表 TableStyles 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### TableStyles.Parent

返回指定对象的父对象。只读

**语法**

**express.Parent**

\*express是一个代表 TableStyles 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/TableStyles](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# WorkbookConnection 对象

## 简介

连接是从 WPS Office ET 2007 工作簿之外的外部数据源获取数据所需要的一组信息。

## 说明

连接可存储在 ET 工作簿中。工作簿打开时，ET 创建连接的内存副本，称为连接对象。连接对象包含服务器名称和服务器上打开的对象的名称等信息。连接对象也可以选择包含身份验证凭据和/或一个命令，该命令可传递给服务器执行（例如：由 SQL Server 执行的 SELECT 语句）。

| 注释                               |
|------------------------------------|
| 连接的确切形式取决于数据检索的机制 |

## 方法

| 序号 | 名称                                   | 说明             |
|:----:|:---------------------------------------|:-----------------|
|  1   | [Delete](#WorkbookConnection.Delete)   | 删除工作簿连接。 |
|  2   | [Refresh](#WorkbookConnection.Refresh) | 刷新工作簿连接。 |

## 属性

| 序号 | 名称                                                   | 说明                                                                                                                                                               |
|:----:|:-------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#WorkbookConnection.Application)         | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#WorkbookConnection.Creator)                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Description](#WorkbookConnection.Description)         | 返回或设置 WorkbookConnection 对象的简要说明。可读/写 String 类型。                                                                                                |
|  4   | [Name](#WorkbookConnection.Name)                       | 返回或设置工作簿连接对象的名称。可读/写 String 类型。                                                                                                              |
|  5   | [ODBCConnection](#WorkbookConnection.ODBCConnection)   | 返回指定的 WorkbookConnection 对象的 ODBC 连接详细信息。只读 ODBCConnection 类型。                                                                                 |
|  6   | [OLEDBConnection](#WorkbookConnection.OLEDBConnection) | 返回指定的 WorkbookConnection 对象的 OLEDB 连接详细信息。只读 OLEDBConnection 类型。                                                                               |
|  7   | [Parent](#WorkbookConnection.Parent)                   | 返回指定对象的父对象。只读。                                                                                                                                       |
|  8   | [Ranges](#WorkbookConnection.Ranges)                   | 返回指定的 WorkbookConnection 对象的对象区域。只读 Ranges 类型。                                                                                                   |
|  9   | [Type](#WorkbookConnection.Type)                       | 返回工作簿连接类型。只读 XlConnectionType 。                                                                                                                       |

## 成员方法

### WorkbookConnection.Delete

删除工作簿连接。

**语法**

**express.Delete ()**

\*express是一个代表 WorkbookConnection 对象的变量。

**说明**

使用此方法可以删除一个外部数据连接。此方法不适用于与其他工作簿之间的链接。

删除连接将不会删除正在使用该连接的任何对象。删除连接将不会导致从文件系统中删除任何连接文件。如果编辑任何这些对象以使用另一个连接，则一切都将重新开始工作。

那些使用已删除链接的对象所表现出的行为就好像无法建立连接一样。

### WorkbookConnection.Refresh

刷新工作簿连接。

**语法**

**express.Refresh ()**

\*express是一个代表 WorkbookConnection 对象的变量。

**说明**

如果 **DisplayAlerts** 属性为 **False** ，则不显示对话框， **Refresh** 方法将失败，并引发“连接信息无效”异常。

刷新连接失败对刷新其他操作没有任何影响。

## 成员属性

### WorkbookConnection.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 WorkbookConnection 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### WorkbookConnection.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 WorkbookConnection 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### WorkbookConnection.Description

返回或设置 **WorkbookConnection** 对象的简要说明。可读/写 **String** 类型。

**语法**

**express.Description**

\*express是一个代表 WorkbookConnection 对象的变量。

**说明**

在 **“连接属性”** 对话框中，用户可以编辑连接的名称和/或说明。在该对话框中更改名称和说明只会更改 ET 连接对象中的相应字段。

说明的最大长度为 255 个字符。如果用户指定的连接文件的说明长于 255 个字符，则说明会被截断，以满足 255 个字符的限制。

### WorkbookConnection.Name

返回或设置工作簿连接对象的名称。可读/写 **String** 类型。

**语法**

**express.Name**

\*express是一个代表 WorkbookConnection 对象的变量。

### WorkbookConnection.ODBCConnection

返回指定的 **WorkbookConnection** 对象的 ODBC 连接详细信息。只读 **ODBCConnection** 类型。

**语法**

**express.ODBCConnection**

\*express是一个代表 WorkbookConnection 对象的变量。

### WorkbookConnection.OLEDBConnection

返回指定的 **WorkbookConnection** 对象的 OLEDB 连接详细信息。只读 **OLEDBConnection** 类型。

**语法**

**express.OLEDBConnection**

\*express是一个代表 WorkbookConnection 对象的变量。

### WorkbookConnection.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 WorkbookConnection 对象的变量。

### WorkbookConnection.Ranges

返回指定的 **WorkbookConnection** 对象的对象区域。只读 **Ranges** 类型。

**语法**

**express.Ranges**

\*express是一个代表 WorkbookConnection 对象的变量。

### WorkbookConnection.Type

返回工作簿连接类型。只读 **XlConnectionType** 。

**语法**

**express.Type**

\*express是一个代表 WorkbookConnection 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/WorkbookConnection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Bookmarks 对象

## 简介

Bookmark 对象的集合，该集合代表指定选定内容、区域或文档中的书签。

## 方法

| 序号 | 名称                        | 说明                                                     |
|:----:|:----------------------------|:---------------------------------------------------------|
|  1   | [Add](#Bookmarks.Add)       | 返回一个 Bookmark 对象，该对象代表添加到区域中的书签。   |
|  2   | [Exists](#Bookmarks.Exists) | 判断指定书签是否存在。如果存在指定的书签，则返回 True 。 |
|  3   | [Item](#Bookmarks.Item)     | 返回集合中的单个 Bookmark 对象。                         |

## 属性

| 序号 | 名称                                        | 说明                                                                                                 |
|:----:|:--------------------------------------------|:-----------------------------------------------------------------------------------------------------|
|  1   | [Application](#Bookmarks.Application)       | 返回一个 Application 对象，该对象代表 WPS 应用程序。                                                 |
|  2   | [Count](#Bookmarks.Count)                   | 返回 Bookmarks 集合中的项目数。 Number 类型，只读。                                                  |
|  3   | [Creator](#Bookmarks.Creator)               | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。                         |
|  4   | [DefaultSorting](#Bookmarks.DefaultSorting) | 返回或设置 “插入” 菜单上 “书签” 对话框中显示的书签名的排序依据选项。 WdBookmarkSortBy 类型，可读写。 |
|  5   | [Parent](#Bookmarks.Parent)                 | 返回一个 Object 类型的值，该值代表指定 Bookmarks 集合中的父对象。                                    |
|  6   | [ShowHidden](#Bookmarks.ShowHidden)         | 如果 Bookmarks 集合中包括隐藏书签，则该属性值为 True 。 Boolean 类型，可读写。                       |

## 成员方法

### Bookmarks.Add

返回一个 **Bookmark** 对象，该对象代表添加到区域中的书签。

**语法**

**express.Add ( *Name* , *Range* )**

\*express是一个代表 Bookmarks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Name*  | 必选      | String   | 书签名。书签名不能多于一个单词。                             |
| *Range* | 可选      | Variant  | 书签标记的文本区域。可将书签设置到一个折叠的区域（插入点）。 |

**示例**

``` JavaScript
/* 在当前插入点插入名为 "point" 的书签 */
Application.ActiveDocument.Bookmarks.Add("point",null)
```

### Bookmarks.Exists

判断指定书签是否存在。如果存在指定的书签，则返回 **True** 。

**语法**

**express.Exists ( *Name* )**

\*express是一个代表 Bookmarks 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明   |
|--------|-----------|----------|--------|
| *Name* | 必选      | String   | 书签名 |

**示例**

``` JavaScript
/* 如果存在名为 "temp" 的书签，则选中 */
function test() {
    let bks = Application.ActiveDocument.Bookmarks;
    if(bks.Exists("temp") === true) {
        bks.Item("temp").Select();
    }
}
```

### Bookmarks.Item

返回集合中的单个 **Bookmark** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Bookmarks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                       |
|---------|-----------|----------|--------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Number 类型值，或代表单个对象名称的 String 类型值。 |

**示例**

``` JavaScript
/* 选中名为"temp"的书签 */
Application.ActiveDocument.Bookmarks.Item("temp").Select()
```

## 成员属性

### Bookmarks.Application

返回一个 Application 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Bookmarks 对象的变量。

### Bookmarks.Count

返回 **Bookmarks** 集合中的项目数。 **Number** 类型，只读。

**语法**

**express.Count**

\*express是一个代表 Bookmarks 对象的变量。

### Bookmarks.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Bookmarks 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

### Bookmarks.DefaultSorting

返回或设置 **“插入”** 菜单上 **“书签”** 对话框中显示的书签名的排序依据选项。 **WdBookmarkSortBy** 类型，可读写。

**语法**

**express.DefaultSorting**

\*express是一个代表 Bookmarks 对象的变量。

**说明**

该属性不影响 **Bookmark** 对象在 **Bookmarks** 集合中的顺序。

**示例**

``` JavaScript
/*本示例按位置对书签进行排序，然后显示“书签”对话框。*/
function test()
{
    Application.ActiveDocument.Bookmarks.DefaultSorting = wdSortByLocation
    Dialogs.Item(wdDialogInsertBookmark).Show()
}
```

### Bookmarks.Parent

返回一个 **Object** 类型的值，该值代表指定 **Bookmarks** 集合中的父对象。

**语法**

**express.Parent**

\*express是一个代表 Bookmarks 对象的变量。

### Bookmarks.ShowHidden

如果 **Bookmarks** 集合中包括隐藏书签，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowHidden**

\*express是一个代表 Bookmarks 对象的变量。

**说明**

**ShowHidden** 属性还可控制是否在 **“插入”** 菜单的 **“书签”** 对话框中列出隐藏书签。在文档中插入交叉引用时，会自动插入隐藏书签。

**示例**

``` JavaScript
/* 显示“书签”对话框，并设置显示隐藏书签 */
function test() {
    Application.ActiveDocument.Bookmarks.ShowHidden = false;
    Application.Dialogs.Item(wdDialogInsertBookmark).Show();
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Bookmarks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Breaks 对象

## 简介

页面中分页符、分栏符和分节符的集合。使用 Breaks 集合及相关对象和属性可通过编程方式定义文档的页面版式。

## 方法

| 序号 | 名称                 | 说明                          |
|:----:|:---------------------|:------------------------------|
|  1   | [Item](#Breaks.Item) | 返回集合中的单个 Break 对象。 |

## 属性

| 序号 | 名称                               | 说明                                                                         |
|:----:|:-----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Breaks.Application) | 返回一个 Application 对象，该对象代表 WPS 应用程序。                         |
|  2   | [Count](#Breaks.Count)             | 返回 Breaks 集合中的项目数。 Long 类型，只读。                               |
|  3   | [Creator](#Breaks.Creator)         | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。 |
|  4   | [Parent](#Breaks.Parent)           | 返回一个 Object 类型的值，该值代表指定 Breaks 集合中的父对象。               |

## 成员方法

### Breaks.Item

返回集合中的单个 **Break** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Breaks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

## 成员属性

### Breaks.Application

返回一个 Application 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Breaks 对象的变量。

### Breaks.Count

返回 **Breaks** 集合中的项目数。 **Long** 类型，只读。

**语法**

**express.Count**

\*express是一个代表 Breaks 对象的变量。

### Breaks.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Breaks 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

### Breaks.Parent

返回一个 **Object** 类型的值，该值代表指定 **Breaks** 集合中的父对象。

**语法**

**express.Parent**

\*express是一个代表 Breaks 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Breaks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ColorFormat 对象

## 简介

代表单色对象的颜色，或以渐变/图案填充的对象的前景色或背景色。

## 说明

使用下表中列出的属性之一返回 ColorFormat 对象。

| 使用此属性     | 结合此对象       | 返回一个代表以下颜色的 ColorFormat 对象    |
|----------------|------------------|--------------------------------------------|
| BackColor      | FillFormat       | 背景填充颜色（在底纹或带图案的填充中使用） |
| ForeColor      | FillFormat       | 前景填充颜色（或单色填充的填充颜色）       |
| BackColor      | LineFormat       | 背景线条颜色（在带图案线条中使用）         |
| ForeColor      | LineFormat       | 前景线条颜色（或实线的线条颜色）           |
| [ForeColor]()  | [ShadowFormat]() | 阴影颜色                                   |
| ExtrusionColor | ThreeDFormat     | 突出对象的边的颜色                         |

## 属性

| 序号 | 名称                                              | 说明                                                                          |
|:----:|:--------------------------------------------------|:------------------------------------------------------------------------------|
|  1   | [Application](#ColorFormat.Application)           | 返回一个 Application 对象，该对象代表 WPS 应用程序。                          |
|  2   | [Brightness](#ColorFormat.Brightness)             | 返回一个 Number 类型的值，代表指定形状颜色的亮度。可读/写。                   |
|  3   | [ObjectThemeColor](#ColorFormat.ObjectThemeColor) | 返回或设置一个 WdThemeColorIndex 常量，该常量表示颜色格式的主题颜色。可读写。 |
|  4   | [RGB](#ColorFormat.RGB)                           | 返回或设置指定颜色的 RGB 值（红-绿-蓝值）， Number 类型，可读写。             |
|  5   | [Type](#ColorFormat.Type)                         | 返回或设置形状的颜色类型。 MsoColorType 类型，只读。                          |

## 成员属性

### ColorFormat.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 ColorFormat 对象的变量。

### ColorFormat.Brightness

返回一个 **Number** 类型的值，代表指定形状颜色的亮度。可读/写。

**语法**

**express.Brightness**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

可以为 **Brightness** 属性输入 -1（最暗）到 1（最亮）之间的数字，0 为中间值。

**示例**

``` JavaScript
/* 设置当前文档第一个对象的前景色以及背景色和亮度 */
function test() {
    let shps = Application.ActiveDocument.Shapes;
    if (shps.Count >= 1) {
        let shp = shps.Item(1);
        let fillFmt = shp.Fill;
        fillFmt.ForeColor.RGB = 0x00ff00;
        fillFmt.BackColor.RGB = 0x0000ff;
        fillFmt.BackColor.Brightness = 0;
        fillFmt.ForeColor.Brightness = 0.4;
    }
}
```

### ColorFormat.ObjectThemeColor

返回或设置一个 **WdThemeColorIndex** 常量，该常量表示颜色格式的主题颜色。可读写。

**语法**

**express.ObjectThemeColor**

\*express是一个代表 ColorFormat 对象的变量。

### ColorFormat.RGB

返回或设置指定颜色的 RGB 值（红-绿-蓝值）， **Number** 类型，可读写。

**语法**

**express.RGB**

\*express是一个代表 ColorFormat 对象的变量。

**示例**

``` JavaScript
/* 本示例将活动文档的第一个图形的颜色设置为灰色 */
Application.ActiveDocument.Shapes.Item(1).Fill.ForeColor.RGB = 0x7f7f7f
```

### ColorFormat.Type

返回或设置形状的颜色类型。 **MsoColorType** 类型，只读。

**语法**

**express.Type**

\*express是一个代表 ColorFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/ColorFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Column 对象

## 简介

代表单个表格列。 Column 对象是 Columns 集合的成员。 Columns 集合包括一个表格、选定内容或区域中的所有列。

## 说明

使用 Columns ( *Index* ) 返回单个 Column 对象（其中 *Index* 是索引编号）。索引编号代表该列在 Columns 集合中的位置（从左至右计数）。

以下示例选定活动文档中的表 1 的第一列。

``` JavaScript
Application.ActiveDocument.Tables.Item(1).Columns.Item(1).Select()
```

可将 Column 属性与 Cell 对象配合使用，返回 Column 对象。以下示例删除单元格 1 中的文本，然后插入新文本并排序整个列。

``` JavaScript
function test() {
    let rng = Application.ActiveDocument.Tables.Item(1).Cell(1, 1)
    rng.Range.Delete()
    rng.Range.InsertBefore("Sales")
    rng.Column.Sort()
}
```

使用 Add 方法在表格中添加一列。以下示例在活动文档的第一个表格中添加一列，然后使各列宽度相等。

``` JavaScript
function test() {
    if (Application.ActiveDocument.Tables.Count >= 1) {
        let myTable = Application.ActiveDocument.Tables.Item(1)
        myTable.Columns.Add(myTable.Columns.Item(1))
        myTable.Columns.DistributeWidth
    }
}
```

说明可将 Information 属性与 Selection 对象配合使用，返回当前列编号。以下示例选定当前列并在消息框中显示其列编号。

``` JavaScript
function test() {
    if(Application.Selection.Information(wdWithInTable) == true){
        Application.Selection.Columns.Item(1).Select()
        alert("Column " + Application.Selection.Information(wdStartOfRangeColumnNumber))
    }
}
```

## 方法

| 序号 | 名称                         | 说明                                                               |
|:----:|:-----------------------------|:-------------------------------------------------------------------|
|  1   | [AutoFit](#Column.AutoFit)   | 改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。 |
|  2   | [Delete](#Column.Delete)     | 删除指定列。                                                       |
|  3   | [Select](#Column.Select)     | 选择指定表格列。                                                   |
|  4   | [SetWidth](#Column.SetWidth) | 设置表格列的宽度。                                                 |
|  5   | [Sort](#Column.Sort)         | 对指定表格列进行排序。                                             |

## 属性

| 序号 | 名称                                             | 说明                                                                                                     |
|:----:|:-------------------------------------------------|:---------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Column.Application)               | 返回一个代表 WPS 应用程序的 Application 对象。                                                           |
|  2   | [Borders](#Column.Borders)                       | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                                    |
|  3   | [Cells](#Column.Cells)                           | 返回一个 Cells 集合，该集合代表表格列中的表格单元格。只读。                                              |
|  4   | [Creator](#Column.Creator)                       | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                             |
|  5   | [Index](#Column.Index)                           | 返回一个 Number 类型的值，该值代表项目在集合中的位置。只读。                                             |
|  6   | [IsFirst](#Column.IsFirst)                       | 如果指定的列是表格的第一列，则返回 True 。只读 Boolean 类型。                                            |
|  7   | [IsLast](#Column.IsLast)                         | 如果指定的列是表格的最后一列，则返回 True 。只读 Boolean 类型。                                          |
|  8   | [NestingLevel](#Column.NestingLevel)             | 返回指定列的嵌套层。只读 Long 类型。                                                                     |
|  9   | [Next](#Column.Next)                             | 返回表格列集合中的下一列。只读。                                                                         |
|  10  | [Parent](#Column.Parent)                         | 返回一个 Object 类型值，该值代表指定 Column 对象的父对象。                                               |
|  11  | [PreferredWidth](#Column.PreferredWidth)         | 返回或设置指定的单元格、列或表格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。 Single 类型，可读写。 |
|  12  | [PreferredWidthType](#Column.PreferredWidthType) | 返回或设置用于指定表格列的宽度的首选度量单位。可读/写 WdPreferredWidthType 类型。                        |
|  13  | [Previous](#Column.Previous)                     | 返回表格列集合中的前一列。只读。                                                                         |
|  14  | [Shading](#Column.Shading)                       | 返回一个引用指定列的底纹格式的 Shading 对象。                                                            |
|  15  | [Width](#Column.Width)                           | 返回或设置指定列的宽度，以磅为单位。可读/写 Long 类型。                                                  |

## 成员方法

### Column.AutoFit

改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。

**语法**

**express.AutoFit ()**

\*express是一个代表 Column 对象的变量。

**说明**

如果表格的宽度已等于从左边界到右边界的距离，则此方法无效。

**示例**

``` JavaScript
/*本示例在新文档中创建一个 3x3 表格，然后调整第一列的宽度，使之与文本的宽度相称。*/
function test() {
    let docNew = Application.Documents.Add()
    let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)
    tableNew.Cell(1, 1).Range.InsertAfter("First cell")
    tableNew.Columns.Item(1).AutoFit()
}
```

``` JavaScript
/*本示例在新文档中创建一个 3x3 表格，然后调整所有列的宽度，使之与文本的宽度相称。*/
function test() {
    let docNew = Application.Documents.Add()
    let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)
    tableNew.Cell(1, 1).Range.InsertAfter("First cell")
    tableNew.Cell(1, 2).Range.InsertAfter("This is cell (1,2)")
    tableNew.Cell(1, 3).Range.InsertAfter("(1,3)")
    tableNew.Columns.AutoFit()
}
```

### Column.Delete

删除指定列。

**语法**

**express.Delete ()**

\*express是一个代表 Column 对象的变量。

### Column.Select

选择指定表格列。

**语法**

**express.Select ()**

\*express是一个代表 Column 对象的变量。

### Column.SetWidth

设置表格列的宽度。

**语法**

**express.SetWidth ( *ColumnWidth* , *RulerStyle* )**

\*express是一个代表 Column 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型     | 说明                            |
|---------------|-----------|--------------|---------------------------------|
| *ColumnWidth* | 必选      | Single       | 指定列的宽度，以磅为单位。      |
| *RulerStyle*  | 必选      | WdRulerStyle | 控制 WPS 调整单元格宽度的方式。 |

**说明**

上述 **WdRulerStyle** 行为应用于左对齐的表格。 **WdRulerStyle** 行为应用于中对齐和右对齐的表格时可能会出现未知效果，因此应谨慎使用 **SetWidth** 方法。

### Column.Sort

对指定表格列进行排序。

**语法**

**express.Sort ( *ExcludeHeader* , *SortFieldType* , *SortOrder* , *CaseSensitive* , *BidiSort* , *IgnoreThe* , *IgnoreKashida* , *IgnoreDiacritics* , *IgnoreHe* , *LanguageID* )**

\*express是一个代表 Column 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                                                                                                                                            |
|--------------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------|
| *ExcludeHeader*    | 可选      | Variant  | 如果该属性值为 True，则不对首行或首段标题进行排序。默认值为 False。                                                                             |
| *SortFieldType*    | 可选      | Variant  | 列的排序类型。可以是 WdSortFieldType 常量之一。                                                                                                 |
| *SortOrder*        | 可选      | Variant  | 列的排序顺序。可以是 WdSortOrder 常量之一。                                                                                                     |
| *CaseSensitive*    | 可选      | Variant  | 如果该属性值为 True，则排序时区分大小写。默认值为 False。                                                                                       |
| *BidiSort*         | 可选      | Variant  | 如果该属性值为 True，则基于从右向左排列的语言规则进行排序。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。                       |
| *IgnoreThe*        | 可选      | Variant  | 如果该属性值为 True，则在从右向左排列的语言的文本排序中忽略阿拉伯字符“alef lam”。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *IgnoreKashida*    | 可选      | Variant  | 如果该属性值为 True，则在从右向左排列的语言的文本排序中忽略“kashidas”。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。           |
| *IgnoreDiacritics* | 可选      | Variant  | 如果该属性值为 True，则在从右向左排列的语言的文本排序中忽略双向控制字符。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。         |
| *IgnoreHe*         | 可选      | Variant  | 如果该属性值为 True，则在从右向左排列的语言的文本排序中忽略希伯来字符“he”。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。       |
| *LanguageID*       | 可选      | Variant  | 指定排序语言。可以是 WdLanguageID 常量之一。                                                                                                    |

**说明**

如果要对表格单元格中的段落进行排序，则只能包括段落标记，不能包括单元格结束标记；如果在选定内容或区域中包括了结束单元格标记，然后试图对段落进行排序， WPS 将显示一条消息，说明未找到进行排序的有效记录。

## 成员属性

### Column.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Column 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Column.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Column 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Column.Cells

返回一个 **Cells** 集合，该集合代表表格列中的表格单元格。只读。

**语法**

**express.Cells**

\*express是一个代表 Column 对象的变量。

### Column.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Column 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Column.Index

返回一个 **Number** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Column 对象的变量。

### Column.IsFirst

如果指定的列是表格的第一列，则返回 **True** 。只读 **Boolean** 类型。

**语法**

**express.IsFirst**

\*express是一个代表 Column 对象的变量。

**示例**

``` JavaScript
/* 本示例判断选定部分的第一列是否为表格中的第一列 */
alert(Application.Selection.Columns.Item(1).IsFirst);
```

### Column.IsLast

如果指定的列是表格的最后一列，则返回 **True** 。只读 **Boolean** 类型。

**语法**

**express.IsLast**

\*express是一个代表 Column 对象的变量。

**示例**

``` JavaScript
/*  本示例判断选定部分的第一列是否为表格的最后一列 */
alert(Application.Selection.Columns.Item(1).IsLast);
```

### Column.NestingLevel

返回指定列的嵌套层。只读 **Long** 类型。

**语法**

**express.NestingLevel**

\*express是一个代表 Column 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

**示例**

``` JavaScript
/*本示例新建一个文档，创建一个三层嵌套表格，并在每个表格的第一个单元格中填入该表格所在的嵌套层数。*/
function test() {
    Application.Documents.Add()
    Application.ActiveDocument.Tables.Add(Selection.Range, 3, 3, wdWord9TableBehavior, wdAutoFitContent)
    let rng = Application.ActiveDocument.Tables.Item(1).Range
    rng.Copy()
    rng.Cells.Item(1).Range.Text = rng.Cells.Item(1).NestingLevel
    rng.Cells.Item(5).Range.PasteAsNestedTable()

    let rng_rng = rng.Cells.Item(5).Tables.Item(1).Range
    rng_rng.Cells.Item(1).Range.Text = rng_rng.Cells.Item(1).NestingLevel
    rng_rng.Cells.Item(5).Range.PasteAsNestedTable()

    let rng_rng_rng = rng_rng.Cells.Item(5).Tables.Item(1).Range
    rng_rng_rng.Cells.Item(1).Range.Text = rng_rng_rng.Cells.Item(1).NestingLevel
}
```

### Column.Next

返回表格列集合中的下一列。只读。

**语法**

**express.Next**

\*express是一个代表 Column 对象的变量。

**示例**

``` JavaScript
/*如果选定内容位于表格之中，则本示例选定下一表格列的内容。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        Application.Selection.Columns.Item(1).Next.Select()
    }
}
```

### Column.Parent

返回一个 **Object** 类型值，该值代表指定 **Column** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Column 对象的变量。

### Column.PreferredWidth

返回或设置指定的单元格、列或表格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。 **Single** 类型，可读写。

**语法**

**express.PreferredWidth**

\*express是一个代表 Column 对象的变量。

**说明**

如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPoints** ，则 **PreferredWidth** 属性将返回或设置以磅为单位的宽度值；如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPercent** ，则 **PreferredWidth** 属性将返回或设置用窗口宽度百分比表示的宽度值。

### Column.PreferredWidthType

返回或设置用于指定表格列的宽度的首选度量单位。可读/写 **WdPreferredWidthType** 类型。

**语法**

**express.PreferredWidthType**

\*express是一个代表 Column 对象的变量。

### Column.Previous

返回表格列集合中的前一列。只读。

**语法**

**express.Previous**

\*express是一个代表 Column 对象的变量。

**示例**

``` JavaScript
/*如果选定内容位于表格中，则本示例选定前一表格列的内容。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        Application.Selection.Columns.Item(1).Previous.Select()
    }
}
```

### Column.Shading

返回一个引用指定列的底纹格式的 **Shading** 对象。

**语法**

**express.Shading**

\*express是一个代表 Column 对象的变量。

**示例**

``` JavaScript
/*本示例对活动文档中第一张表格的第一列应用水平线纹理。*/
function test() {
    if (Application.ActiveDocument.Tables.Count >= 1) {
        let fpx = Application.ActiveDocument.Tables.Item(1).Columns.Item(1).Shading
        fpx.Texture = wdTextureHorizontal
    }
}
```

### Column.Width

返回或设置指定列的宽度，以磅为单位。可读/写 **Long** 类型。

**语法**

**express.Width**

\*express是一个代表 Column 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Column](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Hyperlinks 对象

## 简介

代表文档、范围或所选内容中 Hyperlink 对象的集合。

## 方法

| 序号 | 名称                     | 说明                                                                        |
|:----:|:-------------------------|:----------------------------------------------------------------------------|
|  1   | [Add](#Hyperlinks.Add)   | 返回一个 Hyperlink 对象，该对象代表添加到区域、选定内容或文档的新的超链接。 |
|  2   | [Item](#Hyperlinks.Item) | 返回集合中的单个 Hyperlink 对象。                                           |

## 属性

| 序号 | 名称                                   | 说明                                                                         |
|:----:|:---------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Hyperlinks.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Hyperlinks.Count)             | 返回一个 Long 类型的值，该值代表集合中超链接的数目。只读。                   |
|  3   | [Creator](#Hyperlinks.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#Hyperlinks.Parent)           | 返回一个 Object 类型值，该值代表指定 Hyperlinks 对象的父对象。               |

## 成员方法

### Hyperlinks.Add

返回一个 **Hyperlink** 对象，该对象代表添加到区域、选定内容或文档的新的超链接。

**语法**

**express.Add ( *Anchor* , *Address* , *SubAddress* , *ScreenTip* , *TextToDisplay* , *Target* )**

\*express是一个代表 Hyperlinks 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                |
|-----------------|-----------|----------|-----------------------------------------------------------------------------------------------------|
| *Anchor*        | 必选      | Object   | 要转换为超链接的文本或图形。                                                                        |
| *Address*       | 可选      | Variant  | 指定链接的地址。此地址可以是电子邮件地址、Internet 地址或文件名。请注意，WPS 不检查该地址的正确性。 |
| *SubAddress*    | 可选      | Variant  | 目标文件内的位置名，如书签、已命名的区域或幻灯片编号。                                              |
| *ScreenTip*     | 可选      | Variant  | 当鼠标指针放在指定超链接上时显示为“屏幕提示”的文本。默认值为“地址”。                                |
| *TextToDisplay* | 可选      | Variant  | 指定超链接的显示文本。该参数的值将取代由 Anchor 指定的文本或图形。                                  |
| *Target*        | 可选      | Variant  | 要在其中加载指定超链接的框架或窗口的名字。                                                          |

**示例**

``` JavaScript
/*本示例将选定内容转换为指向万维网上 Microsoft 公司地址的超链接。*/
Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "https://www.wps.cn")

/*本示例将选定内容转换为链接到 MyFile.doc 中名为 MyBookMark 的书签的超链接。*/
Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "C:\\My Documents\\MyFile.doc", "MyBookMark")

/*本示例将选定内容中的第一个图形定义为超链接。*/
Application.ActiveDocument.Hyperlinks.Add(Selection.ShapeRange.Item(1), "https://www.wps.cn")
```

### Hyperlinks.Item

返回集合中的单个 **Hyperlink** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Hyperlinks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

## 成员属性

### Hyperlinks.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 Hyperlinks 对象的变量。

### Hyperlinks.Count

返回一个 **Long** 类型的值，该值代表集合中超链接的数目。只读。

**语法**

**express.Count**

\*express是一个代表 Hyperlinks 对象的变量。

### Hyperlinks.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Hyperlinks 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Hyperlinks.Parent

返回一个 **Object** 类型值，该值代表指定 **Hyperlinks** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Hyperlinks 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Hyperlinks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# LineFormat 对象

## 简介

代表线条和箭头的格式设置。对于线条， LineFormat 对象包含该线条自身的格式信息；对于具有边框的形状，该对象包含形状的边框的格式信息。

## 说明

可以使用 Line 属性返回一个 LineFormat 对象。以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形。

``` JavaScript
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

## 方法

| 序号 | 名称                               | 说明                                                |
|:----:|:-----------------------------------|:----------------------------------------------------|
|  1   | [BreakLink](#LineFormat.BreakLink) | 断开源文件与指定 OLE 对象、图片或链接域之间的链接。 |
|  2   | [Update](#LineFormat.Update)       | 更新指定的链接格式。                                |

## 属性

| 序号 | 名称                                                     | 说明                                                                      |
|:----:|:---------------------------------------------------------|:--------------------------------------------------------------------------|
|  1   | [Application](#LineFormat.Application)                   | 返回一个代表 WPS 应用程序的 Application 对象。                            |
|  2   | [BackColor](#LineFormat.BackColor)                       | 返回或设置一个 [ColorFormat]() 对象，该对象代表图案线条的背景色。可读写。 |
|  3   | [BeginArrowheadLength](#LineFormat.BeginArrowheadLength) | 返回或设置指定线条开头的箭头长度。 MsoArrowheadLength 类型，可读写。      |
|  4   | [BeginArrowheadStyle](#LineFormat.BeginArrowheadStyle)   | 返回或设置指定线条开头的箭头样式。 MsoArrowheadStyle 类型，可读写。       |
|  5   | [BeginArrowheadWidth](#LineFormat.BeginArrowheadWidth)   | 返回或设置指定线条开头的箭头宽度。 MsoArrowheadWidth 类型，可读写。       |
|  6   | [DashStyle](#LineFormat.DashStyle)                       | 返回或设置指定线条的虚线样式。 MsoLineDashStyle 类型，可读写。            |
|  7   | [EndArrowheadLength](#LineFormat.EndArrowheadLength)     | 返回或设置指定线条末尾箭头的长度。 MsoArrowheadLength 类型，可读写。      |
|  8   | [EndArrowheadStyle](#LineFormat.EndArrowheadStyle)       | 返回或设置指定线条末尾箭头的样式。 MsoArrowheadStyle 类型，可读写。       |
|  9   | [EndArrowheadWidth](#LineFormat.EndArrowheadWidth)       | 返回或设置指定线条末尾箭头的宽度。 MsoArrowheadWidth 类型，可读写。       |
|  10  | [ForeColor](#LineFormat.ForeColor)                       | 返回或设置一个 [ColorFormat]() 对象，该对象代表线条的前景色。可读写。     |
|  11  | [Pattern](#LineFormat.Pattern)                           | 设置或返回代表应用于指定线条的图案的值。可读写 MsoPatternType 类型。      |
|  12  | [Style](#LineFormat.Style)                               | 返回或设置线条格式样式。可读写 MsoLineStyle 类型。                        |
|  13  | [Transparency](#LineFormat.Transparency)                 | 返回或设置线条的透明度。可读写 Number 类型。                              |
|  14  | [Weight](#LineFormat.Weight)                             | 返回或设置指定线条的粗细（以磅为单位）。 Number 类型，可读写。            |

## 成员方法

### LineFormat.BreakLink

断开源文件与指定 OLE 对象、图片或链接域之间的链接。

**语法**

**express.BreakLink ()**

\*express是一个代表 LineFormat 对象的变量。

**说明**

使用该方法后，在源文件发生变化的情况下，链接结果不会自动更新。

**示例**

``` JavaScript
/*本示例更新活动文档中所有的 OLE 对象，然后断开链接。*/
function test()
{
let myShapes = ActiveDocument.Shapes

for(let i = 1; i <= myShapes.Count; i++){
    if(myShapes.Item(i).Type == msoLinkedOLEObject){
        myShapes.Item(i).Update()
        myShapes.Item(i).LinkFormat.BreakLink()
    }
}
}
```

### LineFormat.Update

更新指定的链接格式。

**语法**

**express.Update ()**

\*express是一个代表 LineFormat 对象的变量。

## 成员属性

### LineFormat.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.BackColor

返回或设置一个 **[ColorFormat]()** 对象，该对象代表图案线条的背景色。可读写。

**语法**

**express.BackColor**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 本示例为活动文档添加一条带图案线条 */
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.BackColor.RGB = 0x00ff00;
    lnFormat.Pattern = msoPatternDarkDownwardDiagonal;
}
```

### LineFormat.BeginArrowheadLength

返回或设置指定线条开头的箭头长度。 **MsoArrowheadLength** 类型，可读写。

**语法**

**express.BeginArrowheadLength**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.BeginArrowheadStyle

返回或设置指定线条开头的箭头样式。 **MsoArrowheadStyle** 类型，可读写。

**语法**

**express.BeginArrowheadStyle**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.BeginArrowheadWidth

返回或设置指定线条开头的箭头宽度。 **MsoArrowheadWidth** 类型，可读写。

**语法**

**express.BeginArrowheadWidth**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.DashStyle

返回或设置指定线条的虚线样式。 **MsoLineDashStyle** 类型，可读写。

**语法**

**express.DashStyle**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.EndArrowheadLength

返回或设置指定线条末尾箭头的长度。 **MsoArrowheadLength** 类型，可读写。

**语法**

**express.EndArrowheadLength**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.EndArrowheadStyle

返回或设置指定线条末尾箭头的样式。 **MsoArrowheadStyle** 类型，可读写。

**语法**

**express.EndArrowheadStyle**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.EndArrowheadWidth

返回或设置指定线条末尾箭头的宽度。 **MsoArrowheadWidth** 类型，可读写。

**语法**

**express.EndArrowheadWidth**

\*express是一个代表 LineFormat 对象的变量。

**示例**

``` JavaScript
/* 以下示例向活动文档添加一条蓝色虚线。在该线条的起点有一个短而窄的椭圆，在该线条的终点有一个长而宽的三角形*/
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.ForeColor.RGB = 0xff00ff;
    lnFormat.BeginArrowheadLength = msoArrowheadShort;
    lnFormat.BeginArrowheadStyle = msoArrowheadOval;
    lnFormat.BeginArrowheadWidth = msoArrowheadNarrow
    lnFormat.EndArrowheadLength = msoArrowheadLong
    lnFormat.EndArrowheadStyle = msoArrowheadTriangle
    lnFormat.EndArrowheadWidth = msoArrowheadWide
}
```

### LineFormat.ForeColor

返回或设置一个 **[ColorFormat]()** 对象，该对象代表线条的前景色。可读写。

**语法**

**express.ForeColor**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.Pattern

设置或返回代表应用于指定线条的图案的值。可读写 **MsoPatternType** 类型。

**语法**

**express.Pattern**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.Style

返回或设置线条格式样式。可读写 **MsoLineStyle** 类型。

**语法**

**express.Style**

\*express是一个代表 LineFormat 对象的变量。

### LineFormat.Transparency

返回或设置线条的透明度。可读写 **Number** 类型。

**语法**

**express.Transparency**

\*express是一个代表 LineFormat 对象的变量。

**说明**

此属性的值可以取 0.0（不透明）至 1.0（透明）之间的值。该属性只影响单色线条的外观；而不影响带图案的线条的外观。

**示例**

``` JavaScript
/* 添加一条直线，并且设置透明度为0.5 */
function test() {
    let lnFormat = Application.ActiveDocument.Shapes.AddLine(100,100,200,300).Line;
    lnFormat.DashStyle = msoLineDashDotDot;
    lnFormat.BackColor.RGB = 0x00ff00;
    lnFormat.Transparency = 0.5
}
```

### LineFormat.Weight

返回或设置指定线条的粗细（以磅为单位）。 **Number** 类型，可读写。

**语法**

**express.Weight**

\*express是一个代表 LineFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/LineFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# List 对象

## 简介

代表已应用于文档中指定段落的单一列表格式。 List 对象是 Lists 集合的一个成员。

## 说明

使用 Lists ( *Index* ) 可返回一个 List 对象，其中 *Index* 是索引号。以下示例返回活动文档中第二个列表中的项目数。

``` JavaScript
Application.ActiveDocument.Lists.Item(2).CountNumberedItems()
```

要返回具有列表格式的所有段落，请使用 ListParagraphs 属性。要将这些段落作为范围返回，请使用 Range 属性。

要对现有列表应用不同的列表格式，请使用 ApplyListTemplate 方法和 List 对象。要向文档添加新列表，请对指定范围使用 ApplyListTemplate 方法和 [ListFormat]() 对象。

使用 CanContinuePreviousList 方法可确定是否能继续使用以前应用于文档的列表中的列表格式。

使用 CountNumberedItems 方法可返回编号列表或项目符号列表中的项目数，包括 LISTNUM 域。

要确定一个列表是否包含多个列表模板，请使用 SingleListTemplate 属性。

您可以处理文档中的单个 List 对象，但是要更精确地控制，则应使用 ListFormat 对象。

注释：图片项目符号列表不包含在 Lists 集合中，并且不能使用 List 对象进行处理。

## 方法

| 序号 | 名称                                                     | 说明                                                                                                                |
|:----:|:---------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------|
|  1   | [ApplyListTemplate](#List.ApplyListTemplate)             | 为指定的 ListFormat 对象设置列表格式。                                                                              |
|  2   | [CanContinuePreviousList](#List.CanContinuePreviousList) | 返回一个 WdContinue 常量（ wdContinueDisabled 、 wdResetList 或 wdContinueList ），该常量表明是否接续上一列表格式。 |
|  3   | [ConvertNumbersToText](#List.ConvertNumbersToText)       | 更改指定的 List 对象中的列表编号和 LISTNUM 域。                                                                     |
|  4   | [CountNumberedItems](#List.CountNumberedItems)           | 返回指定 List 对象中的项目符号或编号项目和 LISTNUM 域的数量。                                                       |
|  5   | [RemoveNumbers](#List.RemoveNumbers)                     | 删除指定列表中的编号或项目符号。                                                                                    |

## 属性

| 序号 | 名称                                           | 说明                                                                             |
|:----:|:-----------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [ListParagraphs](#List.ListParagraphs)         | 返回一个 ListParagraphs 集合，该集合代表列表、文档或范围中的所有编号段落。只读。 |
|  2   | [Range](#List.Range)                           | 返回一个 Range 对象，该对象代表指定对象所包含的文档部分。                        |
|  3   | [SingleListTemplate](#List.SingleListTemplate) | 如果整个列表都使用同一列表模板，则该属性值为 True 。只读 Boolean 类型。          |

## 成员方法

### List.ApplyListTemplate

为指定的 **ListFormat** 对象设置列表格式。

**语法**

**express.ApplyListTemplate ( *ListTemplate* , *ContinuePreviousList* , *ApplyTo* , *DefaultListBehavior* )**

\*express是一个代表 List 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型     | 说明                                                                                                                                                                                                                                                                                                                                                                  |
|------------------------|-----------|--------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ListTemplate*         | 必选      | ListTemplate | 要应用的列表模板。                                                                                                                                                                                                                                                                                                                                                    |
| *ContinuePreviousList* | 可选      | Variant      | 如果为 True，则接续前一列表的编号；如果为 False，则新建一个列表。                                                                                                                                                                                                                                                                                                     |
| *ApplyTo*              | 可选      | Variant      | 要应用列表模板的列表部分。可以是下列 WdListApplyTo 常量之一：wdListSelection、wdListWholeList 或 wdListThisPointForward。                                                                                                                                                                                                                                             |
| *DefaultListBehavior*  | 可选      | Variant      | 设置可指定 WPS 是否使用新的网页格式，以便列表具有更好的显示效果的值。其值可取下列 WdDefaultListBehavior 常量之一：wdWord8ListBehavior（使用与 WPS 97 兼容的格式）或 wdWord9ListBehavior（使用适于网页的格式）。考虑到兼容问题，默认常量为 wdWord8ListBehavior。但是在新建过程中若要建立缩进式多级列表，应当使用 wdWord9ListBehavior，以利用改进过的适用于网页的格式。 |

**示例**

``` JavaScript
/* 本示例将变量 myRange 设置为活动文档的一个区域，然后检查该区域，判断其中是否存在列表格式。如果没有列表格式，则为该区域设置第四种符号列表模板的格式 */
function test() {
    let actDoc = Application.ActiveDocument;
    let rng = actDoc.Range(actDoc.Paragraphs.Item(1).Range.Start, actDoc.Paragraphs.Item(2).Range.End);
    if (rng.ListFormat.ListType === wdListNoNumbering) {
        rng.ListFormat.ApplyListTemplate(Application.ListGalleries.Item(2).ListTemplates.Item(4))
    }
}
```

### List.CanContinuePreviousList

返回一个 **WdContinue** 常量（ **wdContinueDisabled** 、 **wdResetList** 或 **wdContinueList** ），该常量表明是否接续上一列表格式。

**语法**

**express.CanContinuePreviousList ( *ListTemplate* )**

\*express是一个代表 List 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型     | 说明                                 |
|----------------|-----------|--------------|--------------------------------------|
| *ListTemplate* | 必选      | ListTemplate | 一个应用于文档中先前段落的列表模板。 |

**说明**

本方法返回指定列表格式的 **“继续前一列表”** 和 **“重新开始编号”** 选项状态，这些选项位于 **“项目符号和编号”** 对话框中。设置 **ApplyListTemplate** 方法的 ContinuePreviousList 参数可更改这些选项的设置。

**示例**

``` JavaScript
/* 本示例检查是否可以接续前面得编号，则应用当前的列表模板，其编号接续先前列表的编号。所选内容必须位于第二个列表中，否则导致出错 */
function test() {
    let selListFmt = Application.Selection.Range.ListFormat
    let canConti = selListFmt.CanContinuePreviousList(selListFmt.ListTemplate)
    if (canConti !== wdContinueDisabled ) {
        selListFmt.ApplyListTemplate(selListFmt.ListTemplate, true)
    }
}
```

### List.ConvertNumbersToText

更改指定的 **List** 对象中的列表编号和 LISTNUM 域。

**语法**

**express.ConvertNumbersToText ()**

\*express是一个代表 List 对象的变量。

**示例**

``` JavaScript
/* 本示例将第一个列表的编号转换为文本 */
Application.ActiveDocument.Lists.Item(1).ConvertNumbersToText()
```

### List.CountNumberedItems

返回指定 **List** 对象中的项目符号或编号项目和 LISTNUM 域的数量。

**语法**

**express.CountNumberedItems ()**

\*express是一个代表 List 对象的变量。

**示例**

``` JavaScript
/* 本示例返回第一个List对象中数量 */
Application.ActiveDocument.Lists.Item(1).CountNumberedItems()
```

### List.RemoveNumbers

删除指定列表中的编号或项目符号。

**语法**

**express.RemoveNumbers ( *NumberType* )**

\*express是一个代表 List 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型     | 说明               |
|--------------|-----------|--------------|--------------------|
| *NumberType* | 可选      | WdNumberType | 要删除的编号类型。 |

**示例**

``` JavaScript
/* 移除第一个列表List得项目编号或符号 */
Application.ActiveDocument.Lists.Item(1).RemoveNumbers()
```

## 成员属性

### List.ListParagraphs

返回一个 **ListParagraphs** 集合，该集合代表列表、文档或范围中的所有编号段落。只读。

**语法**

**express.ListParagraphs**

\*express是一个代表 List 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/* 将第一个列表中得第一个段落设置为双下划线 */
Application.ActiveDocument.Lists.Item(1).ListParagraphs.Item(1).Range.Underline = wdUnderlineDouble
```

### List.Range

返回一个 **Range** 对象，该对象代表指定对象所包含的文档部分。

**语法**

**express.Range**

\*express是一个代表 List 对象的变量。

### List.SingleListTemplate

如果整个列表都使用同一列表模板，则该属性值为 **True** 。只读 **Boolean** 类型。

**语法**

**express.SingleListTemplate**

\*express是一个代表 List 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/List](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Options 对象

## 简介

代表 WPS 中的应用程序和文档选项。 Options 对象的许多属性与 “选项” 对话框中的项对应。

## 说明

使用 Options 属性可返回 Options 对象。以下示例设置了 WPS 的三个应用程序选项。

``` JavaScript
function test() {
    Application.Options.AllowDragAndDrop = true
    Application.Options.ConfirmConversions = false
    Application.Options.MeasurementUnit = wdPoints
}
```

## 方法

| 序号 | 名称                                        | 说明                                                                       |
|:----:|:--------------------------------------------|:---------------------------------------------------------------------------|
|  1   | [DefaultFilePath](#Options.DefaultFilePath) | 返回或设置各项（例如文档、模板和图形）的默认文件夹。 String 类型，可读写。 |

## 属性

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
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AddBiDirectionalMarksWhenSavingTextFile">AddBiDirectionalMarksWhenSavingTextFile</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在将文档保存为文本文件时添加双向控制字符，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AddControlCharacters">AddControlCharacters</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在剪切和复制文本时添加双向控制字符，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AddHebDoubleQuote">AddHebDoubleQuote</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 会将数字格式括在双引号 (") 中，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AllowAccentedUppercase">AllowAccentedUppercase</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当一个法语字符改成大写字母时重音仍保留，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AllowClickAndTypeMouse">AllowClickAndTypeMouse</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果启用“即点即输”功能，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AllowCombinedAuxiliaryForms">AllowCombinedAuxiliaryForms</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在朝鲜语文档中执行拼写检查时将忽略辅助动词形式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AllowCompoundNounProcessing">AllowCompoundNounProcessing</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在朝鲜语文档中执行拼写检查时将忽略复合名词，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AllowDragAndDrop">AllowDragAndDrop</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果可以使用拖动来移动或复制所选内容，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AllowOpenInDraftView">AllowOpenInDraftView</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 Boolean 值，该值表示是否允许用户在草稿视图中打开文档。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AllowPixelUnits">AllowPixelUnits</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 用像素作为默认度量单位（这是 HTML 功能所支持的度量单位），则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AllowReadingMode">AllowReadingMode</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则表明 WPS 将文档在“阅读版式”视图中打开。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AlwaysUseClearType">AlwaysUseClearType</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 Boolean 值，该值表示是否使用 ClearType 来显示菜单、功能区和对话框文本中的字体。可读/写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AnimateScreenMovements">AnimateScreenMovements</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 激活动态鼠标功能、使用动态光标，并且激活后台保存与查找和替换操作之类的动态操作效果，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个代表 WPS 应用程序的 <span>Application</span> 对象。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ApplyFarEastFontsToAscii">ApplyFarEastFontsToAscii</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 将中文字体应用于西文文字，则该属性值为 True 。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">16</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ArabicMode">ArabicMode</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置阿拉伯语的拼写检查模式。 WdAraSpeller 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">17</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ArabicNumeral">ArabicNumeral</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一篇阿拉伯语文档的数字样式。 WdArabicNumeral 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">18</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoCreateNewDrawings">AutoCreateNewDrawings</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 在 <span>绘图画布</span> 上画出新创建的图形。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">19</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatApplyBulletedLists">AutoFormatApplyBulletedLists</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当 WPS 自动设置文档或区域格式时，会用 “格式” 菜单的 “项目符号和编号” 对话框中的项目符号替换列表段落的开始字符（例如星号、连字符和大于号符号），则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">20</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatApplyFirstIndents">AutoFormatApplyFirstIndents</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 在自动设置文档或区域的格式时，用首行缩进替换在段落开始处输入的空格。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">21</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatApplyHeadings">AutoFormatApplyHeadings</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当 WPS 自动设置文档或区域格式时，将自动为标题设置样式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">22</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatApplyLists">AutoFormatApplyLists</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当 WPS 自动设置文档或区域格式时，自动为列表应用样式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">23</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatApplyOtherParas">AutoFormatApplyOtherParas</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当 WPS 自动设置文档或区域格式时，自动为非标题项或列表项的段落设置样式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">24</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeApplyBorders">AutoFormatAsYouTypeApplyBorders</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果按下 Enter 后，三个或更多的连字符 (-)、等号 (=) 或下划线 (_) 将自动替换成指定的边框线，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">25</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeApplyBulletedLists">AutoFormatAsYouTypeApplyBulletedLists</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在键入时用 “格式” 菜单的 “项目符号和编号” 对话框中的项目符号替换当前的项目符号（例如星号、连字符和大于号），则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">26</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeApplyClosings">AutoFormatAsYouTypeApplyClosings</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 会在键入时自动为信函的结束语应用“结束语”样式。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">27</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeApplyDates">AutoFormatAsYouTypeApplyDates</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 自动为键入的日期应用“日期”样式。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">28</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeApplyFirstIndents">AutoFormatAsYouTypeApplyFirstIndents</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 在键入时自动用首行缩进替代段落开始处输入的空格。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">29</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeApplyHeadings">AutoFormatAsYouTypeApplyHeadings</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在键入标题时，自动为标题设置样式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">30</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeApplyNumberedLists">AutoFormatAsYouTypeApplyNumberedLists</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果根据键入的内容，以 “格式” 菜单的 “项目符号和编号”对话框中的编号方案将段落自动设置为编号列表，则该属性值为 True 。例如，如果一个段落以“1.1”和一个制表符开始，则当按下 Enter 后 WPS 将自动插入“1.2”和一个制表符。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">31</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeApplyTables">AutoFormatAsYouTypeApplyTables</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则当您键入一个加号、一串连字符、另一个加号等，然后按下 Enter 后， WPS 将自动创建一张表格。加号变为列边框，连字符变为列宽度。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">32</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeAutoLetterWizard">AutoFormatAsYouTypeAutoLetterWizard</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则在用户输入信函称呼或结束语时，WPS 会自动启动“英文信函向导”。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">33</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeDefineStyles">AutoFormatAsYouTypeDefineStyles</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 自动根据手动设置的格式新建一种样式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">34</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeDeleteAutoSpaces">AutoFormatAsYouTypeDeleteAutoSpaces</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 会在用户键入时自动删除日文和拉丁文文字之间插入的空格。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">35</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeFormatListItemBeginning">AutoFormatAsYouTypeFormatListItemBeginning</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 将列表项开始处的字符格式重复应用于下一个列表项，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">36</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeInsertClosings">AutoFormatAsYouTypeInsertClosings</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则当用户输入备忘录标题时，WPS 会自动插入相应的备忘录结束语。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">37</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeInsertOvers">AutoFormatAsYouTypeInsertOvers</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则当用户输入“ ”或“ ”时，WPS 将自动插入“ ”。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">38</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeMatchParentheses">AutoFormatAsYouTypeMatchParentheses</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，WPS 会自动更正不正确成对的括号，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">39</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeReplaceFarEastDashes">AutoFormatAsYouTypeReplaceFarEastDashes</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，WPS 会自动更正长元音和划线，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">40</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeReplaceFractions">AutoFormatAsYouTypeReplaceFractions</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在您键入时，将使用当前字符集的分数替换键入的分数，则该属性值为 True 。例如，将“1/2”替换为“?”。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">41</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeReplaceHyperlinks">AutoFormatAsYouTypeReplaceHyperlinks</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果本属性值为 True ，则 WPS 会在键入时自动将电子邮件地址、服务器和共享名称（即 UNC 路径）和 Internet 地址（即 URL）改为超链接。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">42</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeReplaceOrdinals">AutoFormatAsYouTypeReplaceOrdinals</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则键入时， WPS 会用相同字母的上标格式替换序数后缀“st”、“nd”、“rd”和“th”。例如，“1st”被替换为“1”后接上标“st”。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">43</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeReplacePlainTextEmphasis">AutoFormatAsYouTypeReplacePlainTextEmphasis</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果键入时，手动强调的字符将自动替换为相应字符格式，则该属性值为 True 。例如，将“*bold*”改为“ bold ”、将“_underline_”改为“underline”。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">44</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeReplaceQuotes">AutoFormatAsYouTypeReplaceQuotes</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果键入时自动将所键入的直引号替换为弯引号，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">45</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatAsYouTypeReplaceSymbols">AutoFormatAsYouTypeReplaceSymbols</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果键入的两个连续的连字符 (--) 将替换为短划线 (–) 或长划线 (—)，则该属性值为 True 。 Boolean 类型，可读写。如果键入连字符时带有前导和尾部空格，则 WPS 将连字符替换为短划线。如果没有尾部空格，则连字符替换为长划线。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">46</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatDeleteAutoSpaces">AutoFormatDeleteAutoSpaces</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在自动设置文档或区域的格式时，删除日文和西文文字之间插入的空格，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">47</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatMatchParentheses">AutoFormatMatchParentheses</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在自动设置文档或区域的格式时，会更正不恰当匹配的括号，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">48</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatPlainTextWordMail">AutoFormatPlainTextWordMail</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在 WPS 中打开纯文本格式的电子邮件时， WPS 将自动进行格式设置，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">49</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatPreserveStyles">AutoFormatPreserveStyles</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当 WPS 自动设置文档或区域格式时，保留先前所用的样式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">50</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatReplaceFarEastDashes">AutoFormatReplaceFarEastDashes</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在自动设置文档或区域的格式时，会更正长元音和划线的使用，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">51</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatReplaceFractions">AutoFormatReplaceFractions</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当 WPS 自动设置文档或区域格式时，会用当前字符集中的分数替换所键入的分数，则该属性值为 True 。例如，将“1/2”替换为“?”。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">52</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatReplaceHyperlinks">AutoFormatReplaceHyperlinks</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 在对某一文档或范围自动套用格式时，将对其中的电子邮件地址、服务器和共享名称（即 UNC 路径）和 Internet 地址（即 URL）自动进行格式设置。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">53</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatReplaceOrdinals">AutoFormatReplaceOrdinals</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则当 WPS 自动设置文档或区域格式时，用相同字母的上标格式替换序数后缀“st”、“nd”、“rd”和“th”。例如，将“1st”替换为“1”后接上标“st”。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">54</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatReplacePlainTextEmphasis">AutoFormatReplacePlainTextEmphasis</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当 WPS 自动设置文档或区域的格式时，将手动强调的字符替换为相应的字符格式（例如，“*bold*”更改为“bold”，“_underline_”更改为“underline”），则该属性值为 True 。 Boolean 类型，可读写。</p>
<p>如果当 WPS 自动设置文档或区域的格式时，将手动强调的字符替换为相应的字符格式（例如，“*bold*”更改为“bold”，“_underline_”更改为“underline”），则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">55</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatReplaceQuotes">AutoFormatReplaceQuotes</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当 WPS 自动设置文档或区域格式时，直引号会自动更改为弯引号，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">56</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoFormatReplaceSymbols">AutoFormatReplaceSymbols</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当 WPS 自动设置文档或区域的格式时，将两个连续的连字号 (--) 替换为短破折号 (</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">57</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoKeyboardSwitching">AutoKeyboardSwitching</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 随时自动切换键盘语言并与键入的内容匹配，则该属性值为 True 。可读/写 Boolean 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">58</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.AutoWordSelection">AutoWordSelection</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在拖动时，一次选定一个单词，而不是选定一个字符，则该属性为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">59</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.BackgroundOpen">BackgroundOpen</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，WPS 在后台打开 Web 文档。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">60</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.BackgroundSave">BackgroundSave</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 在后台保存文档。当 WPS 在后台保存文档时，用户可以继续键入内容或执行命令。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">61</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.BibliographySort">BibliographySort</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 String 类型的值，该值代表在 “源管理器” 对话框中显示源的次序。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">62</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.BibliographyStyle">BibliographyStyle</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 String 类型的值，该值代表用于设置书目格式的样式的名称。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">63</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.BrazilReform">BrazilReform</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置巴西葡萄牙语拼写器的模式。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">64</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ButtonFieldClicks">ButtonFieldClicks</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个数字，该数字是运行 GOTOBUTTON 或 MACROBUTTON 域所需的点击次数（单击或双击）。 Long 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">65</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.CheckGrammarAsYouType">CheckGrammarAsYouType</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在键入时自动对文档进行语法检查并自动标记错误，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">66</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.CheckGrammarWithSpelling">CheckGrammarWithSpelling</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在检查拼写的同时也检查语法，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">67</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.CheckHangulEndings">CheckHangulEndings</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 自动检测朝鲜文字结尾，并在将朝鲜文字转换为朝鲜文汉字过程中忽略朝鲜文字结尾，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">68</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.CheckSpellingAsYouType">CheckSpellingAsYouType</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在您键入的同时自动进行拼写检查并标记错误，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">69</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.CommentsColor">CommentsColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>该属性返回或设置一个 WdColorIndex 常量，代表文档中批注的颜色。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">70</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ConfirmConversions">ConfirmConversions</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在打开或者插入一个非 WPS 文档或模板的文件之前，会显示一个 “转换文件” 对话框，则该属性值为 True 。在 “转换文件” 对话框中，用户可以选择需要转换的文件格式。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">71</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ContextualSpeller">ContextualSpeller</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 Boolean 值，该值表示是否使用上下文拼写检查器，基于单词上下文和它前后的单词来检查拼写。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">72</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ConvertHighAnsiToFarEast">ConvertHighAnsiToFarEast</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在打开文档时，WPS 将与东亚字体相关的文字转换为相应的字体，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">73</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.CreateBackup">CreateBackup</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在每次保存文档时都建立一个备份，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">74</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">75</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.CtrlClickHyperlinkToOpen">CtrlClickHyperlinkToOpen</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 需要单击同时按住 Ctrl 键来打开超链接，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">76</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.CursorMovement">CursorMovement</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置在双向文字中插入点移动的方式。 WdCursorMovement 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">77</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DefaultBorderColor">DefaultBorderColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置用于新的 <span>Border</span> 对象的默认 24 位颜色。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">78</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DefaultBorderColorIndex">DefaultBorderColorIndex</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置边框的默认线条颜色。 WdColorIndex 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">79</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DefaultBorderLineStyle">DefaultBorderLineStyle</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置默认边框的样式。 WdLineStyle 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">80</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DefaultBorderLineWidth">DefaultBorderLineWidth</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置边框的默认线条宽度。 WdLineWidth 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">81</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DefaultEPostageApp">DefaultEPostageApp</a></td>
<td style="text-align: left;" data-valian="middle"><p>设置或返回一个 String 类型的值，该值代表默认的电子邮政应用程序的路径和文件名。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">82</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DefaultFilePath">DefaultFilePath</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置各项（例如文档、模板和图形）的默认文件夹。 String 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">83</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DefaultHighlightColorIndex">DefaultHighlightColorIndex</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置突出显示文本所用的颜色，用 “格式” 工具栏的 “突出显示” 按钮可将文本格式设为突出显示。 WdColorIndex 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">84</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DefaultOpenFormat">DefaultOpenFormat</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置打开文档的默认转换器。可为由 OpenFormat 属性返回的一个数字，也可为下列 WdOpenFormat 常量之一。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">85</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DefaultTextEncoding">DefaultTextEncoding</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 MsoEncoding 常量，该常量代表 WPS 用于所有保存为编码文本文件的文档的代码页或字符集。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">86</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DefaultTray">DefaultTray</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置默认纸盒，以便打印机使用该纸盒打印文档。 String 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">87</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DefaultTrayID">DefaultTrayID</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置用于打印文档的默认纸盒。 WdPaperTray 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">88</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DeletedCellColor">DeletedCellColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 WdCellColor 常量，该常量代表删除的单元格的颜色。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">89</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DeletedTextColor">DeletedTextColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>当激活修订选项时，返回或设置要删除的文字的颜色。 WdColorIndex 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">90</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DeletedTextMark">DeletedTextMark</a></td>
<td style="text-align: left;" data-valian="middle"><p>当激活修订选项时，返回或设置要删除的文字的格式。 WdDeletedTextMark 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">91</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DiacriticColorVal">DiacriticColorVal</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置用于从右向左语言文档中的音调符号的 24 位颜色。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">92</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DisableFeaturesbyDefault">DisableFeaturesbyDefault</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 将在所有文档中禁用 <span>DisableFeaturesIntroducedAfterbyDefault</span> 中指定的 WPS 版本后引入的所有功能。默认值为 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">93</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DisableFeaturesIntroducedAfterbyDefault">DisableFeaturesIntroducedAfterbyDefault</a></td>
<td style="text-align: left;" data-valian="middle"><p>对所有文档禁用指定版本后引入的所有功能。 WdDisableFeaturesIntroducedAfter 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">94</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DisplayGridLines">DisplayGridLines</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 显示文档网格，则该属性值为 True 。该属性相当于 “视图” 菜单中的 “网格线” 命令。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">95</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DisplayPasteOptions">DisplayPasteOptions</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 直接在新粘贴的文本下显示 “粘贴选项” 按钮。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">96</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DisplaySmartTagButtons">DisplaySmartTagButtons</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 在鼠标指针置于智能标记上方时显示一个按钮。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">97</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DocumentViewDirection">DocumentViewDirection</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置整篇文档的对齐方式和阅读顺序。 WdDocumentViewDirection 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">98</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.DoNotPromptForConvert">DoNotPromptForConvert</a></td>
<td style="text-align: left;" data-valian="middle"><p>设置或返回一个 Boolean ，它代表在为处于兼容模式下的文档调用“转换”命令时是否出现一个警告对话框。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">99</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.EnableHangulHanjaRecentOrdering">EnableHangulHanjaRecentOrdering</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在进行朝鲜文字和朝鲜文汉字间的转换时，在建议列表的顶端显示最近用过的单词，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">100</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.EnableLegacyIMEMode">EnableLegacyIMEMode</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Boolean 类型的值，该值代表是否启用旧式 IME 模式。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">101</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.EnableLivePreview">EnableLivePreview</a></td>
<td style="text-align: left;" data-valian="middle"><p>设置或返回 Boolean 值，该值表示显示还是隐藏在使用支持预览的库时所出现的库预览。如果为 True ，将在应用命令之前在文档中显示预览。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">102</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.EnableMisusedWordsDictionary">EnableMisusedWordsDictionary</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在检查文档的拼写和语法时，也检查误用的单词，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">103</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.EnableSound">EnableSound</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在导致出错时使计算机发出声音进行响应，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">104</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.EnvelopeFeederInstalled">EnvelopeFeederInstalled</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果当前的打印机为信封指定了送纸器，则该属性值为 True 。 Boolean 类型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">105</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.FormatScanning">FormatScanning</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 对文档中的所有格式进行跟踪。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">106</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.FrenchReform">FrenchReform</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 <span>WdFrenchSpeller</span> 常量，该常量代表要对语言格式设置为法语的文本区域使用哪个拼写词典。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">107</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.GridDistanceHorizontal">GridDistanceHorizontal</a></td>
<td style="text-align: left;" data-valian="middle"><p>当在新文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置 WPS 使用的不可见网格线的水平间距。 Single 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">108</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.GridDistanceVertical">GridDistanceVertical</a></td>
<td style="text-align: left;" data-valian="middle"><p>在新文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置 WPS 所用的不可见网格线间的垂直间距。 Single 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">109</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.GridOriginHorizontal">GridOriginHorizontal</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置不可见网格相对于页面左边的起点位置，当在新文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。 Single 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">110</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.GridOriginVertical">GridOriginVertical</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置不可见网格相对于页面顶边的起点位置，当在新文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。 Single 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">111</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.HangulHanjaFastConversion">HangulHanjaFastConversion</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在朝鲜文字和朝鲜文汉字间进行转换时，自动转换具有单独建议的单词，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">112</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.HebrewMode">HebrewMode</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置希伯来语拼写检查工具的模式。 WdHebSpellStart 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">113</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.HideChartDraftModeNotification">HideChartDraftModeNotification</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在草图模式下插入的图表上隐藏 “草图模式” 通知标签，则该属性值为 “True” 。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">114</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.IgnoreInternetAndFileAddresses">IgnoreInternetAndFileAddresses</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则进行拼写检查时忽略文件扩展名、MS-DOS 路径、电子邮件地址、服务器名称、共享名（也称为 UNC 路径）和 Internet 地址（也称为 URL）。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">115</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.IgnoreMixedDigits">IgnoreMixedDigits</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果进行拼写检查时忽略包含数字的单词，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">116</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.IgnoreUppercase">IgnoreUppercase</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果进行拼写检查时忽略全部大写的单词，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">117</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.IMEAutomaticControl">IMEAutomaticControl</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果设置 WPS 自动打开和关闭日文输入法编辑器 (IME)，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">118</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.InlineConversion">InlineConversion</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 将中文输入法 (IME) 中的未确认字符串显示为已有（已确认的）字符串之间的插入项，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">119</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.InsertChartInDraftMode">InsertChartInDraftMode</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果使用草图模式在文档中插入图表，则该属性值为 True 。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">120</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.InsertedCellColor">InsertedCellColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 WdCellColor 常量，该常量代表插入的表格单元格的颜色。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">121</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.InsertedTextColor">InsertedTextColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>当启用修订选项时，返回或设置插入文字的颜色。 WdColorIndex 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">122</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.InsertedTextMark">InsertedTextMark</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置启用修订时 WPS 设置插入文本格式的方式（ TrackRevisions 属性为 True ）。 WdInsertedTextMark 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">123</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.INSKeyForOvertype">INSKeyForOvertype</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果可以用 Ins 键来打开和关闭“改写”，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">124</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.INSKeyForPaste">INSKeyForPaste</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果可用 INS 键来粘贴“剪贴板”的内容，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">125</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.InterpretHighAnsi">InterpretHighAnsi</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置高端字符译码方式。 WdHighAnsiText 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">126</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.LabelSmartTags">LabelSmartTags</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，WPS 使用智能标记信息标记文档中的文本。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">127</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.LocalNetworkFile">LocalNetworkFile</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在编辑保存在网络服务器上的文件时，WPS 在用户的计算机上创建文件的一个本地副本，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">128</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MapPaperSize">MapPaperSize</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则自动调整采用其他国家/地区标准纸张大小（例如 A4）的文档格式，以便按照用户所在国家/地区的标准纸张大小（例如，Letter）正确打印该文档。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">129</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyAY%20">MatchFuzzyAY</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在搜索过程中 WPS 将忽略 行和 行字符后“ ”和“ ”之间的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">130</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyBV">MatchFuzzyBV</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 将在搜索过程中忽略“ ”与“ ”之间以及“ ”与“ ”之间的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">131</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyByte">MatchFuzzyByte</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在搜索过程中 WPS 将忽略全角和半角字符（西文或日文）的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">132</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyCase">MatchFuzzyCase</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在搜索过程中不区分字母大小写，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">133</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyDash">MatchFuzzyDash</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在搜索过程中忽略减号、划线以及长元音符号之间的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">134</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyDZ">MatchFuzzyDZ</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 将在搜索过程中忽略“ ”与“ ”之间以及“ ”与“ ”之间的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">135</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyHF">MatchFuzzyHF</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 将在搜索过程中忽略“ ”与“ ”之间以及“ ”与“ ”之间的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">136</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyHiragana">MatchFuzzyHiragana</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在搜索过程中将忽略平假名与片假名之间的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">137</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyIterationMark">MatchFuzzyIterationMark</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在搜索过程中将忽略重复的标记类型间的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">138</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyKanji">MatchFuzzyKanji</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在搜索过程中忽略标准和非标准日文汉字的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">139</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyKiKu">MatchFuzzyKiKu</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在搜索过程中 WPS 将忽略 行字符前“ ”和“ ”之间的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">140</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyOldKana">MatchFuzzyOldKana</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在搜索过程中将忽略新旧假名字符之间的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">141</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyProlongedSoundMark">MatchFuzzyProlongedSoundMark</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在搜索过程中 WPS 将忽略短元音和长元音的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">142</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyPunctuation">MatchFuzzyPunctuation</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在搜索过程中 WPS 忽略标点类型的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">143</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzySmallKana">MatchFuzzySmallKana</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在搜索过程中忽略双元音和双辅音的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">144</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzySpace">MatchFuzzySpace</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在搜索过程中忽略空格标记的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">145</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyTC">MatchFuzzyTC</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在搜索过程中 WPS 将忽略“ ”、“ ”与“ ”之间以及“ ”与“ ”之间的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">146</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MatchFuzzyZJ">MatchFuzzyZJ</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 将在搜索过程中忽略“ ”与“ ”之间以及“ ”与“ ”之间的差异，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">147</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MeasurementUnit">MeasurementUnit</a></td>
<td style="text-align: left;" data-valian="middle"><p>设置或者返回 WPS 的标准度量单位。 WdMeasurementUnits 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">148</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MergedCellColor">MergedCellColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 WdCellColor 常量，该常量代表合并的表格单元格的颜色。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">149</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MonthNames">MonthNames</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置朝鲜文字和朝鲜文汉字之间的转换方向。 WdMonthNames 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">150</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MoveFromTextColor">MoveFromTextColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 <span>WdColorIndex</span> 常量，该常量代表所移动文本的颜色。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">151</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MoveFromTextMark">MoveFromTextMark</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 <span>WdMoveFromTextMark</span> 常量，该常量代表要用于所移动文本的修订标记的类型。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">152</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MoveToTextColor">MoveToTextColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 <span>WdColorIndex</span> 常量，该常量代表所移动文本的颜色。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">153</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MoveToTextMark">MoveToTextMark</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 <span>WdMoveToTextMark</span> 常量，该常量代表要用于所移动文本的修订标记的类型。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">154</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.MultipleWordConversionsMode">MultipleWordConversionsMode</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置朝鲜文字和朝鲜文汉字之间的转换方向。 WdMultipleWordConversionsMode 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">155</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.OMathAutoBuildUp">OMathAutoBuildUp</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Boolean 类型的值，该值代表 WPS 是否自动将公式转换为专业格式。如果为 True ，则表示 WPS 自动将公式转换为专业格式。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">156</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.OMathCopyLF">OMathCopyLF</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Boolean 类型的值，该值代表如何以纯文本的形式表示公式。如果为 True ，则表示以线性格式表示公式。如果为 False ，则表示以 MathML 格式表示公式。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">157</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.OptimizeForWord97byDefault">OptimizeForWord97byDefault</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 将禁用所有不兼容的格式，以便在 WPS 97 中查看新文档而对其进行优化，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">158</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.Overtype">Overtype</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果激活改写模式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">159</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.Pagination">Pagination</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在后台重新为文档分页，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">160</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Object 类型值，该值代表指定 Options 对象的父对象。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">161</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteAdjustParagraphSpacing">PasteAdjustParagraphSpacing</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在剪切和粘贴选定内容时，WPS 自动调整段落间距，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">162</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteAdjustTableFormatting">PasteAdjustTableFormatting</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在剪切和粘贴选定内容时，WPS 自动调整表格的格式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">163</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteAdjustWordSpacing">PasteAdjustWordSpacing</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在剪切和粘贴选定内容时，WPS 自动调整字符间距，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">164</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteFormatBetweenDocuments">PasteFormatBetweenDocuments</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 WdPasteOptions 常量，该常量代表在从另一个 WPS Office WPS 文档复制文本时粘贴文本的方式。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">165</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteFormatBetweenStyledDocuments">PasteFormatBetweenStyledDocuments</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 <span>WdPasteOptions</span> 常量，该常量代表在从使用样式的文档中复制文本时粘贴文本的方式。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">166</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteFormatFromExternalSource">PasteFormatFromExternalSource</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 WdPasteOptions 常量，该常量代表在从外部源（如网页）复制文本时粘贴文本的方式。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">167</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteFormatWithinDocument">PasteFormatWithinDocument</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 WdPasteOptions 常量，该常量代表在同一文档内复制或剪切文本然后进行粘贴时的文本粘贴方式。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">168</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteMergeFromPPT">PasteMergeFromPPT</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则从 Microsoft PowerPoint 粘贴文本时合并文本格式。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">169</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteMergeFromXL">PasteMergeFromXL</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则从 ET 粘贴表格时合并表格的格式。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">170</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteMergeLists">PasteMergeLists</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则合并粘贴列表和周围列表的格式。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">171</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteOptionKeepBulletsAndNumbers">PasteOptionKeepBulletsAndNumbers</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 Boolean 值，该值表示在从 “粘贴选项” 上下文菜单中选择 “仅保留文本” 时是否保留项目符号和编号。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">172</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteSmartCutPaste">PasteSmartCutPaste</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在文档中智能粘贴选定内容，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">173</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PasteSmartStyleBehavior">PasteSmartStyleBehavior</a></td>
<td style="text-align: left;" data-valian="middle"><p>从其他文档粘贴选定文本时，如果 WPS 智能合并样式，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">174</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PictureEditor">PictureEditor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置用于编辑图片的应用程序的名称。 String 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">175</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PictureWrapType">PictureWrapType</a></td>
<td style="text-align: left;" data-valian="middle"><p>设置或返回一个 WdWrapTypeMerged 类型的值，该值表示 WPS 在图片周围环绕文字的方式。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">176</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PortugalReform">PortugalReform</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置欧洲葡萄牙语拼写器的模式。可读/写 <span>WdPortugueseReform</span> 。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">177</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrecisePositioning">PrecisePositioning</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Boolean 类型的值，该值表示 WPS 是否针对打印版式而不是针对屏幕可读性优化字符位置。如果该属性的值为 True ，则禁用压缩字符间距以便增强屏幕可读性的默认设置，并启用适用于打印介质的字符间距。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">178</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrintBackground">PrintBackground</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在后台打印，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">179</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrintBackgrounds">PrintBackgrounds</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Boolean 类型的值，该值代表打印文档时，是否打印背景颜色和图像。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">180</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrintComments">PrintComments</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在文档后另起一页打印批注，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">181</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrintDraft">PrintDraft</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 以最简单的格式打印文档，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">182</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrintDrawingObjects">PrintDrawingObjects</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 打印图形对象，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">183</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrintEvenPagesInAscendingOrder">PrintEvenPagesInAscendingOrder</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在手动双面打印模式下按升序打印偶数页，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">184</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrintFieldCodes">PrintFieldCodes</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 打印域代码而不打印域的结果，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">185</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrintHiddenText">PrintHiddenText</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果打印隐藏文字，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">186</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrintProperties">PrintProperties</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在文档结尾处另起一页打印文档摘要信息，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">187</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrintReverse">PrintReverse</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 逆页序打印页面，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">188</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PrintXMLTag">PrintXMLTag</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Boolean 类型的值，该值代表打印文档时是否打印 XML 标记。对应于 “选项” 对话框中 “打印” 选项卡上的 “XML 标记” 复选框。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">189</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.PromptUpdateStyle">PromptUpdateStyle</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则在改变样式的格式时显示一条消息，请用户确认是否要重新设置样式的格式或重新应用原样式格式。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">190</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.RepeatWPS">RepeatWPS</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 Boolean 值，该值表示在检查拼写时是否标记重复的单词。如果为 True ，将标记重复的单词。可读/写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">191</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ReplaceSelection">ReplaceSelection</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果键入或粘贴的内容替换选定内容，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">192</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.RevisedLinesColor">RevisedLinesColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>当激活修订选项时，返回或设置文档中修订线的颜色。 WdColorIndex 类型，可读/写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">193</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.RevisedLinesMark">RevisedLinesMark</a></td>
<td style="text-align: left;" data-valian="middle"><p>当激活修订选项时，返回或设置修订线在文档中的位置。 WdRevisedLinesMark 类型，可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">194</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.RevisedPropertiesColor">RevisedPropertiesColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>在启用修订功能时，返回或设置用于标记格式更改情况的颜色。 WdColorIndex 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">195</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.RevisedPropertiesMark">RevisedPropertiesMark</a></td>
<td style="text-align: left;" data-valian="middle"><p>在启用修订功能时，返回或设置用于显示格式修改情况的标记。 WdRevisedPropertiesMark 类型，可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">196</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.RevisionsBalloonPrintOrientation">RevisionsBalloonPrintOrientation</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 WdRevisionsBalloonPrintOrientation 常量，该常量表示打印时修订和批注气球的方向。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">197</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.RTFInClipboard">RTFInClipboard</a></td>
<td style="text-align: left;" data-valian="middle"><p>您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 Options 对象的 RTFInClipboard 属性的信息，请参阅 WPS OfficeMacintosh Edition 中附带的语言参考帮助。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">198</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SaveInterval">SaveInterval</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置保存“自动恢复”信息的时间间隔（以分钟为单位）。 Long 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">199</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SaveNormalPrompt">SaveNormalPrompt</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在关闭之前提示用户确认是否保存对 Normal 模板所做的更改，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">200</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SavePropertiesPrompt">SavePropertiesPrompt</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在保存新的文档时，提示该文档的属性信息，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">201</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SendMailAttach">SendMailAttach</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果使用 “文件” 菜单上的 “发送” 命令将活动文档作为附件插入邮件中，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">202</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SequenceCheck">SequenceCheck</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则检查南亚文本独立字符的顺序。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">203</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ShortMenuNames">ShortMenuNames</a></td>
<td style="text-align: left;" data-valian="middle"><p>您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 Options 对象的 ShortMenuNames 属性的信息，请参阅 WPS OfficeMacintosh Edition 中附带的语言参考帮助。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">204</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ShowControlCharacters">ShowControlCharacters</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果显示当前文档中的双向控制符，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">205</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ShowDevTools">ShowDevTools</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 Boolean 类型的值，该值代表是否在功能区中显示 “开发人员” 选项卡。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">206</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ShowDiacritics">ShowDiacritics</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在使用从右向左语言的文档中显示音调符号，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">207</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ShowFormatError">ShowFormatError</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，WPS 通过在文本（该文本的格式与文档中使用频繁的其他格式相似）下加弯曲下划线来标记不一致的格式。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">208</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ShowMarkupOpenSave">ShowMarkupOpenSave</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Boolean 类型的值，该值代表打开或保存文件时 WPS 是否显示隐藏标记。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">209</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ShowMenuFloaties">ShowMenuFloaties</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 Boolean 值，该值表示当用户在文档窗口中右键单击时是否显示浮动工具栏。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">210</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ShowReadabilityStatistics">ShowReadabilityStatistics</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在结束语法检查时，显示统计信息的摘要列表（包括可读性程度），则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">211</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.ShowSelectionFloaties">ShowSelectionFloaties</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 Boolean 值，该值表示当用户选择文本时是否显示浮动工具栏。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">212</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SmartCursoring">SmartCursoring</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Boolean 类型的值，该值表示是否启用智能指针。如果值为 True ，则启用智能指针。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">213</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SmartCutPaste">SmartCutPaste</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在剪切和粘贴时，WPS 自动调整字词和标点符号之间的间距，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">214</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SmartParaSelection">SmartParaSelection</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则在选择段落中大多数或全部内容时，WPS 会在选定内容中包含段落标记。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">215</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SnapToGrid">SnapToGrid</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果绘制、移动自选图形或中文字符或者调整其大小时，它们自动与不可见的网格对齐，则该属性值为 True 。可读写 Boolean 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">216</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SnapToShapes">SnapToShapes</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 自动将自选图形或中文字符与不可见的网格线对齐，这些网格线穿过其他自选图形或中文字符的垂直和水平边缘，则该属性值为 True 。可读写 Boolean 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">217</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SpanishMode">SpanishMode</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置西班牙语拼写器的模式。可读/写 <span>WdSpanishSpeller</span> 。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">218</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SplitCellColor">SplitCellColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 WdCellColor 常量，该常量代表拆分的表格单元格的颜色。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">219</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.StoreRSIDOnSave">StoreRSIDOnSave</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 在每次保存文档时为文档中的变化指定一个随机数，以便比较与合并文档。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">220</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.StrictFinalYaa">StrictFinalYaa</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果拼写检查使用针对以字母“yaa”结尾的阿拉伯语单词的拼写规则，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">221</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.StrictInitialAlefHamza">StrictInitialAlefHamza</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果拼写检查使用针对以“alef hamza”开头的阿拉伯语单词的拼写规则，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">222</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.StrictRussianE">StrictRussianE</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果拼写检查器使用有关使用严格 ? 字符的俄语单词的拼写规则，则该属性值为 True 。可读/写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">223</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.StrictTaaMarboota">StrictTaaMarboota</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果拼写检查器使用拼写规则来标记以 haa（而不是 taa marboota）结尾的阿拉伯语单词，则该属性值为 True 。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">224</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SuggestFromMainDictionaryOnly">SuggestFromMainDictionaryOnly</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 仅根据主词典提供拼写建议，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">225</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.SuggestSpellingCorrections">SuggestSpellingCorrections</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在检查拼写时，总是为每一个拼写错误的单词提供可选的拼写建议，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">226</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.TabIndentKey">TabIndentKey</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果可分别使用 Tab 和 Backspace 键增加和减少段落的左缩进，并且可使用 Backspace 键将右对齐的段落更改为居中的段落而将居中的段落更改为左对齐的段落，则该属性值为 True 。可读写 Boolean 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">227</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.TypeNReplace">TypeNReplace</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 将替换非法的南亚字符。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">228</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.UpdateFieldsAtPrint">UpdateFieldsAtPrint</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在打印文档前自动更新域，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">229</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.UpdateFieldsWithTrackedChangesAtPrint">UpdateFieldsWithTrackedChangesAtPrint</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 2015 允许在打印前更新包含修订的字段，则该属性值为 True 。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">230</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.UpdateLinksAtOpen">UpdateLinksAtOpen</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在打开文档时自动更新嵌入的所有 OLE 链接，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">231</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.UpdateLinksAtPrint">UpdateLinksAtPrint</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在打印文档前更新指向其他文件的嵌入链接，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">232</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.UpdateStyleListBehavior">UpdateStyleListBehavior</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 <span>WdUpdateStyleListBehavior</span> 常量，该常量指定在更新某种样式以与包含编号或项目符号的所选内容相匹配时， WPS 2015 应采取的行为。可读/写。rs"&gt;版本信息</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">233</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.UseCharacterUnit">UseCharacterUnit</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在当前文档中使用字符作为默认度量单位，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">234</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.UseDiffDiacColor">UseDiffDiacColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果可以设置当前文档中音调符号的颜色，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">235</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.UseGermanSpellingReform">UseGermanSpellingReform</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果 WPS 在检查拼写时，使用德语后期修订拼写规则，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">236</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.UseNormalStyleForList">UseNormalStyleForList</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 Boolean 值，该值表示 WPS 是否对项目符号和编号使用“正文”样式。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">237</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.VisualSelection">VisualSelection</a></td>
<td style="text-align: left;" data-valian="middle"><p>根据从右向左式语言文档中的可视光标移动，返回或设置选择行为。 WdVisualSelection 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">238</td>
<td style="text-align: left;" data-valian="middle"><a href="#Options.WarnBeforeSavingPrintingSendingMarkup">WarnBeforeSavingPrintingSendingMarkup</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则 WPS 在保存、打印或以电子邮件形式发送包含批注或修订的文档时，显示警告。 Boolean 类型，可读写。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### Options.DefaultFilePath

返回或设置各项（例如文档、模板和图形）的默认文件夹。 **String** 类型，可读写。

**语法**

**express.DefaultFilePath ( *Path* )**

\*express是一个代表 Options 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型          | 说明               |
|--------|-----------|-------------------|--------------------|
| *Path* | 必选      | WdDefaultFilePath | 设置默认的文件夹。 |

**说明**

您可以使用空字符串 ("") 从 Windows 注册表中删除设置。新设置会立即生效。

**示例**

``` JavaScript
/*本示例为 WPS 文档设置默认文件夹。*/
Options.DefaultFilePath(wdDocumentsPath) = "C:\\Documents"
```

``` JavaScript
/*本示例返回用户模板当前的默认路径（相当于“选项”对话框中“文件位置”选项卡上的默认路径设置）。*/
let strPath = Options.DefaultFilePath(wdUserTemplatesPath)
```

## 成员属性

### Options.AddBiDirectionalMarksWhenSavingTextFile

如果 WPS 在将文档保存为文本文件时添加双向控制字符，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AddBiDirectionalMarksWhenSavingTextFile**

\*express是一个代表 Options 对象的变量。

**说明**

如果保存文本文件时带有双向控制符，则可保留从右向左和从左向右的属性，以及中性字符的顺序。

**示例**

``` JavaScript
/*本示例将设定 WPS 在将文档保存为文本文件时添加双向控制字符。*/
Options.AddBiDirectionalMarksWhenSavingTextFile = true
```

### Options.AddControlCharacters

如果 WPS 在剪切和复制文本时添加双向控制字符，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AddControlCharacters**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在剪切和复制文本时添加双向控制字符。*/
Options.AddControlCharacters = true
```

### Options.AddHebDoubleQuote

如果 WPS 会将数字格式括在双引号 (") 中，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AddHebDoubleQuote**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 将数字格式括在双引号 (") 中。*/
Options.AddHebDoubleQuote = true
```

### Options.AllowAccentedUppercase

如果当一个法语字符改成大写字母时重音仍保留，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AllowAccentedUppercase**

\*express是一个代表 Options 对象的变量。

**说明**

本属性只影响标准法语文本。对于其他语言的文本，尽管将 **AllowAccentedUppercase** 属性设为 **False** ，仍然保留重音。

当一个重音符号被删除之后，如果把一个字符改回小写字母，则重音不会再出现。

**示例**

``` JavaScript
/*当法语字符被改为大写字母时，本示例将设置 WPS 删除重音符号。*/
Options.AllowAccentedUppercase = false
```

``` JavaScript
/*本示例返回在“选项”对话框中的“编辑”选项卡上的“允许法语中大写字母带重音符号”选项的状态。*/
let blnUppercaseAccents = Options.AllowAccentedUppercase
```

### Options.AllowClickAndTypeMouse

如果启用“即点即输”功能，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AllowClickAndTypeMouse**

\*express是一个代表 Options 对象的变量。

**说明**

有关“即点即输”功能的详细信息，请参阅“关于即点即输”。

**示例**

``` JavaScript
/*本示例检查并确认是否已启用“即点即输”功能。如果未启用该功能，则根据用户的选择来设置该功能。*/
function test()
{
if(Options.AllowClickAndTypeMouse == false) { 
    x = MsgBox("Do you want to use Click and Type?", jsYesNo)
    if(x == jsResultYes) {
        Options.AllowClickAndTypeMouse = true
        MsgBox("Click and Type enabled!")
    }
}
}
```

### Options.AllowCombinedAuxiliaryForms

如果 WPS 在朝鲜语文档中执行拼写检查时将忽略辅助动词形式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AllowCombinedAuxiliaryForms**

\*express是一个代表 Options 对象的变量。

**说明**

有关在 WPS 中使用亚洲语言的详细信息，请参阅 WPS 中关于亚洲语言的功能。

**示例**

``` JavaScript
/*本示例询问用户：WPS 在朝鲜语文档中执行拼写检查时，是否忽略辅助动词形式。*/
function test()
{
if(Options.AllowCombinedAuxiliaryForms == false){
    let x = MsgBox("Do you want to ignore auxiliary verb forms when checking spelling?", jsYesNo)
    if(x == jsResultYes){
        Options.AllowCombinedAuxiliaryForms = true
        MsgBox("Auxiliary verb forms will be ignored!")
    }
}
}
```

### Options.AllowCompoundNounProcessing

如果 WPS 在朝鲜语文档中执行拼写检查时将忽略复合名词，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AllowCompoundNounProcessing**

\*express是一个代表 Options 对象的变量。

**说明**

有关在 WPS 中使用亚洲语言的详细信息，请参阅 WPS 中关于亚洲语言的功能。

**示例**

``` JavaScript
/*本示例询问用户：WPS 在朝鲜语文档中执行拼写检查时，是否忽略复合名词。*/
function test()
{
if(Options.AllowCompoundNounProcessing == false){
    let x = MsgBox("Do you want to ignore compound "+"nouns when checking spelling?",jsYesNo)
    if(x == jsResultYes){
        Options.AllowCompoundNounProcessing = true
        MsgBox("Compound nouns will be ignored!")
    }
}
}
```

### Options.AllowDragAndDrop

如果可以使用拖动来移动或复制所选内容，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AllowDragAndDrop**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例打开拖放编辑功能。*/
Options.AllowDragAndDrop = true
```

``` JavaScript
/*本示例返回“选项”对话框中“编辑”选项卡上的“拖放式文字编辑”选项的状态。*/
let blnDragAndDrop = Options.AllowDragAndDrop
```

### Options.AllowOpenInDraftView

返回或设置 **Boolean** 值，该值表示是否允许用户在草稿视图中打开文档。可读/写。

**语法**

**express.AllowOpenInDraftView**

\*express是一个代表 Options 对象的变量。

**说明**

此属性对应于 **“ WPS 选项”** 对话框中的 **“允许在草稿视图中打开文档”** 复选框。

### Options.AllowPixelUnits

如果 WPS 用像素作为默认度量单位（这是 HTML 功能所支持的度量单位），则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AllowPixelUnits**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：使 WPS 可以将像素作为 HTML 功能的默认度量单位。*/
Options.AllowPixelUnits = true
```

### Options.AllowReadingMode

如果为 **True** ，则表明 WPS 将文档在“阅读版式”视图中打开。 **Boolean** 类型，可读写。

**语法**

**express.AllowReadingMode**

\*express是一个代表 Options 对象的变量。

**说明**

对应于 **“选项”** 对话框中 **“常规”** 选项卡上的 **“允许从‘阅读版式’启动”** 复选框。

**示例**

``` JavaScript
/*下列示例切换“允许从‘阅读版式’启动”复选框。*/
function ToggleReadingMode(){
    if(Options.AllowReadingMode == true){
        Options.AllowReadingMode = false
    }
    else{
        Options.AllowReadingMode = true
    }
}
```

### Options.AlwaysUseClearType

返回或设置 **Boolean** 值，该值表示是否使用 ClearType 来显示菜单、功能区和对话框文本中的字体。可读/写。

**语法**

**express.AlwaysUseClearType**

\*express是一个代表 Options 对象的变量。

**说明**

即使关闭 Microsoft Windows 的 ClearType 设置，将此属性设置为 **True** 也会替代 Windows 设置，并在所有 WPS Office应用程序中使用 ClearType。此属性对应于 **“ WPS 选项”** 对话框的 **“高级”** 选项卡中的 **“总是使用 ClearType”** 复选框。

### Options.AnimateScreenMovements

如果 WPS 激活动态鼠标功能、使用动态光标，并且激活后台保存与查找和替换操作之类的动态操作效果，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AnimateScreenMovements**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 以激活屏幕上的动态移动功能。*/
Options.AnimateScreenMovements = true
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“常规”选项卡的“提供动画反馈”选项的当前状态。*/
let blnAnimation = Options.AnimateScreenMovements
```

### Options.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Options 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Options.ApplyFarEastFontsToAscii

如果 WPS 将中文字体应用于西文文字，则该属性值为 **True** 。

**语法**

**express.ApplyFarEastFontsToAscii**

\*express是一个代表 Options 对象的变量。

**说明**

仅在选择一种东亚语言进行编辑时才应用该属性。如果为 **False** ，并且对某个特定区域应用中文字体，则 WPS 将不对该区域中的任何西文文字应用该字体。

**示例**

``` JavaScript
/*本示例设置 WPS 对西文文字应用中文字体。*/
Options.ApplyFarEastFontsToAscii = true
```

### Options.ArabicMode

返回或设置阿拉伯语的拼写检查模式。 **WdAraSpeller** 类型，可读写。

**语法**

**express.ArabicMode**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置拼写检查忽略阿拉伯语中关于单词以 alef hamza 开头的拼写规则。*/
Options.ArabicMode = wdInitialAlef
```

### Options.ArabicNumeral

返回或设置一篇阿拉伯语文档的数字样式。 **WdArabicNumeral** 类型，可读写。

**语法**

**express.ArabicNumeral**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例把数字样式设为印地语样式。*/
Options.ArabicNumeral = wdNumeralHindi
```

### Options.AutoCreateNewDrawings

如果为 **True** ，则 WPS 在 绘图画布 上画出新创建的图形。 **Boolean** 类型，可读写。

**语法**

**express.AutoCreateNewDrawings**

\*express是一个代表 Options 对象的变量。

**说明**

**AutoCreateNewDrawings** 属性只能影响从 WPS 内部添加的图形。如果图形是通过 示例代码 代码添加的，则此选项设置为 **True** 或 **False** 都没有关系，因为它们是按照代码的指定添加的。

**示例**

``` JavaScript
/*本示例设置 WPS 直接在文档中添加新创建的图形，而不通过绘图画布。*/
function NewDrawings(){
    Application.Options.AutoCreateNewDrawings = false
}
```

### Options.AutoFormatApplyBulletedLists

如果当 WPS 自动设置文档或区域格式时，会用 **“格式”** 菜单的 **“项目符号和编号”** 对话框中的项目符号替换列表段落的开始字符（例如星号、连字符和大于号符号），则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatApplyBulletedLists**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例用项目符号替换当前选定内容中列表段落的段首字符。*/
function test()
{
Options.AutoFormatApplyBulletedLists = true
Selection.Range.AutoFormat()
}
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“自动项目符号列表”选项的状态。*/
let blnAutoFormat = Options.AutoFormatApplyBulletedLists
```

### Options.AutoFormatApplyFirstIndents

如果为 **True** ，则 WPS 在自动设置文档或区域的格式时，用首行缩进替换在段落开始处输入的空格。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatApplyFirstIndents**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 用首行缩进替换在段落开始处输入的空格，并自动设置选定区域的格式。*/
function test()
{
Options.AutoFormatApplyFirstIndents = true
Selection.Range.AutoFormat()
}
```

### Options.AutoFormatApplyHeadings

如果当 WPS 自动设置文档或区域格式时，将自动为标题设置样式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatApplyHeadings**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例对选定内容中的标题应用标题 1 到标题 9 的样式。*/
function test()
{
Options.AutoFormatApplyHeadings = true
Selection.Range.AutoFormat()
}
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“标题”选项的状态。*/
let blnAutoFormat = Options.AutoFormatApplyHeadings
```

### Options.AutoFormatApplyLists

如果当 WPS 自动设置文档或区域格式时，自动为列表应用样式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatApplyLists**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例为当前选定内容中的所有列表应用样式。*/
function test()
{
Options.AutoFormatApplyLists = true
Selection.Range.AutoFormat()
}
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“列表”选项的状态。*/
let blnAutoFormat = Options.AutoFormatApplyLists
```

### Options.AutoFormatApplyOtherParas

如果当 WPS 自动设置文档或区域格式时，自动为非标题项或列表项的段落设置样式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatApplyOtherParas**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例自动为当前选定段落应用样式。*/
function test()
{
Options.AutoFormatApplyOtherParas = true
Selection.Range.AutoFormat()
}
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“其他段落样式”选项的状态。*/
let blnAutoFormat = Options.AutoFormatApplyOtherParas
```

### Options.AutoFormatAsYouTypeApplyBorders

如果按下 Enter 后，三个或更多的连字符 (-)、等号 (=) 或下划线 (\_) 将自动替换成指定的边框线，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyBorders**

\*express是一个代表 Options 对象的变量。

**说明**

以一条 0.75 磅的直线替换连字符 (-)，以一条 0.75 磅的双线替换等号符号 (=)，以一条 1.5 磅的直线替换下划线 (\_)。

**示例**

``` JavaScript
/*本示例将三个或更多的连字符 (-)、等号符号 (=) 或下划线字符 (_) 序列替换成边框。*/
Options.AutoFormatAsYouTypeApplyBorders = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“边框”选项的当前设置。*/
MsgBox(Options.AutoFormatAsYouTypeApplyBorders)
```

### Options.AutoFormatAsYouTypeApplyBulletedLists

如果在键入时用 **“格式”** 菜单的 **“项目符号和编号”** 对话框中的项目符号替换当前的项目符号（例如星号、连字符和大于号），则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyBulletedLists**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例用项目符号替换在列表中键入的字符。*/
Options.AutoFormatAsYouTypeApplyBulletedLists = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“自动项目符号列表”选项的状态。*/
let blnAutoFormat = Options.AutoFormatAsYouTypeApplyBulletedLists
```

### Options.AutoFormatAsYouTypeApplyClosings

如果为 **True** ，则 WPS 会在键入时自动为信函的结束语应用“结束语”样式。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyClosings**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在键入时自动为信函的结束语应用“结束语”样式。*/
function AutoClosings(){
    Options.AutoFormatAsYouTypeApplyClosings = true
}
```

### Options.AutoFormatAsYouTypeApplyDates

如果为 **True** ，则 WPS 自动为键入的日期应用“日期”样式。可读写。

**语法**

**express.AutoFormatAsYouTypeApplyDates**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 自动为键入时输入的日期应用“日期”样式。*/
function AutoApplyDates(){
    Options.AutoFormatAsYouTypeApplyDates = true
}
```

### Options.AutoFormatAsYouTypeApplyFirstIndents

如果为 **True** ，则 WPS 在键入时自动用首行缩进替代段落开始处输入的空格。可读写。

**语法**

**express.AutoFormatAsYouTypeApplyFirstIndents**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 自动将键入时在段落开始处输入的空格替换为首行缩进。*/
function ApplyFirstIndents(){
    Options.AutoFormatAsYouTypeApplyFirstIndents = true
}
```

### Options.AutoFormatAsYouTypeApplyHeadings

如果在键入标题时，自动为标题设置样式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyHeadings**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*当键入文字时，本示例对选定内容中的标题应用标题 1 到标题 9 的样式。*/
Options.AutoFormatAsYouTypeApplyHeadings = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“标题”选项的状态。*/
let blnAutoFormat = Options.AutoFormatAsYouTypeApplyHeadings
```

### Options.AutoFormatAsYouTypeApplyNumberedLists

如果根据键入的内容，以 **“格式”** 菜单的 “项目符号和编号”对话框中的编号方案将段落自动设置为编号列表，则该属性值为 **True** 。例如，如果一个段落以“1.1”和一个制表符开始，则当按下 Enter 后 WPS 将自动插入“1.2”和一个制表符。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyNumberedLists**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*键入文字时，本示例自动对列表进行编号。*/
Options.AutoFormatAsYouTypeApplyNumberedLists = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“自动编号列表”选项的状态。*/
let blnAutoFormat = Options.AutoFormatAsYouTypeApplyNumberedLists
```

### Options.AutoFormatAsYouTypeApplyTables

如果为 **True** ，则当您键入一个加号、一串连字符、另一个加号等，然后按下 Enter 后， WPS 将自动创建一张表格。加号变为列边框，连字符变为列宽度。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyTables**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在您键入时自动创建表格。*/
Options.AutoFormatAsYouTypeApplyTables = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“表格”选项的状态。*/
let blnAutoFormat = Options.AutoFormatAsYouTypeApplyTables
```

### Options.AutoFormatAsYouTypeAutoLetterWizard

如果为 **True** ，则在用户输入信函称呼或结束语时，WPS 会自动启动“英文信函向导”。可读写。

**语法**

**express.AutoFormatAsYouTypeAutoLetterWizard**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在用户输入信函称呼或结束语时，自动启动“英文信函向导”。*/
function AutoLeterWizard(){
    Options.AutoFormatAsYouTypeAutoLetterWizard = true
}
```

### Options.AutoFormatAsYouTypeDefineStyles

如果 WPS 自动根据手动设置的格式新建一种样式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeDefineStyles**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*键入文本时，本示例使 WPS 自动创建样式。*/
Options.AutoFormatAsYouTypeDefineStyles = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“基于所用格式定义样式”选项的状态。*/
let blnAutoFormat = Options.AutoFormatAsYouTypeDefineStyles
```

### Options.AutoFormatAsYouTypeDeleteAutoSpaces

如果为 **True** ，则 WPS 会在用户键入时自动删除日文和拉丁文文字之间插入的空格。可读写。

**语法**

**express.AutoFormatAsYouTypeDeleteAutoSpaces**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在用户键入时自动删除日文和拉丁文文字之间插入的空格。*/
function AutoDeleteSpaces() {
    Options.AutoFormatAsYouTypeDeleteAutoSpaces = true
}
```

### Options.AutoFormatAsYouTypeFormatListItemBeginning

如果 WPS 将列表项开始处的字符格式重复应用于下一个列表项，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeFormatListItemBeginning**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 自动重复列表项开始处的字符格式。*/
Options.AutoFormatAsYouTypeFormatListItemBeginning = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“将列表项开始的格式设为与其前一项相似”选项的状态。*/
let blnAutoFormat = Options.AutoFormatAsYouTypeFormatListItemBeginning
```

### Options.AutoFormatAsYouTypeInsertClosings

如果为 **True** ，则当用户输入备忘录标题时，WPS 会自动插入相应的备忘录结束语。可读写。

**语法**

**express.AutoFormatAsYouTypeInsertClosings**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在用户输入备忘录标题时自动插入相应的备忘录结束语。*/
function AutoInsertClosings() {
    Options.AutoFormatAsYouTypeInsertClosings = true
}
```

### Options.AutoFormatAsYouTypeInsertOvers

如果为 **True** ，则当用户输入“ ”或“ ”时，WPS 将自动插入“ ”。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeInsertOvers**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在用户输入“ ki”或“ an”时自动插入“ ijou ijou”。*/
Options.AutoFormatAsYouTypeInsertOvers = true
```

### Options.AutoFormatAsYouTypeMatchParentheses

如果为 **True** ，WPS 会自动更正不正确成对的括号，可读写。

**语法**

**express.AutoFormatAsYouTypeMatchParentheses**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在用户键入时自动更正不正确成对的括号。*/
function AutoMatchParentheses() {
    Options.AutoFormatAsYouTypeMatchParentheses = true
}
```

### Options.AutoFormatAsYouTypeReplaceFarEastDashes

如果为 **True** ，WPS 会自动更正长元音和划线，可读写。

**语法**

**express.AutoFormatAsYouTypeReplaceFarEastDashes**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在用户键入时自动更正长元音和划线。*/
function AutoFarEastDashes() {
    Options.AutoFormatAsYouTypeReplaceFarEastDashes = true
}
```

### Options.AutoFormatAsYouTypeReplaceFractions

如果在您键入时，将使用当前字符集的分数替换键入的分数，则该属性值为 **True** 。例如，将“1/2”替换为“?”。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeReplaceFractions**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例关闭键入分数的自动替换功能。*/
Options.AutoFormatAsYouTypeReplaceFractions = false
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“分数 (1/2) 替换为分数字符 (?)”选项的状态。*/
let blnAutoFormat = Options.AutoFormatAsYouTypeReplaceFractions
```

### Options.AutoFormatAsYouTypeReplaceHyperlinks

如果本属性值为 **True** ，则 WPS 会在键入时自动将电子邮件地址、服务器和共享名称（即 UNC 路径）和 Internet 地址（即 URL）改为超链接。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeReplaceHyperlinks**

\*express是一个代表 Options 对象的变量。

**说明**

WPS 将任何看起来像电子邮件地址、UNC 或 URL 的文字更改为超链接。 WPS 并不检查超链接是否有效。

**示例**

``` JavaScript
/*本示例使 WPS 自动将键入的任何 Internet 或者网络路径替换为超链接。*/
Options.AutoFormatAsYouTypeReplaceHyperlinks = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“Internet 及网络路径替换为超链接”选项的状态。*/
let blnAutoFormat = Options.AutoFormatAsYouTypeReplaceHyperlinks
```

### Options.AutoFormatAsYouTypeReplaceOrdinals

如果为 **True** ，则键入时， WPS 会用相同字母的上标格式替换序数后缀“st”、“nd”、“rd”和“th”。例如，“1st”被替换为“1”后接上标“st”。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeReplaceOrdinals**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例启用将序数字母自动替换为上标字母的功能。*/
Options.AutoFormatAsYouTypeReplaceOrdinals = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“序号 (1st) 替换为上标”选项的状态。*/
let blnAutoFormat = Options.AutoFormatAsYouTypeReplaceOrdinals
```

### Options.AutoFormatAsYouTypeReplacePlainTextEmphasis

如果键入时，手动强调的字符将自动替换为相应字符格式，则该属性值为 **True** 。例如，将“\*bold\*”改为“ **bold** ”、将“\_underline\_”改为“underline”。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeReplacePlainTextEmphasis**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例启用将手动强调的字符替换为相应字符格式的功能。*/
Options.AutoFormatAsYouTypeReplacePlainTextEmphasis = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“*加粗* 和 _下划线_ 替换为真正格式”选项的状态。*/
let blnAutoFormat = Options.AutoFormatAsYouTypeReplacePlainTextEmphasis
 
```

### Options.AutoFormatAsYouTypeReplaceQuotes

如果键入时自动将所键入的直引号替换为弯引号，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeReplaceQuotes**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例启用将所键入的直引号自动替换为弯引号的功能。*/
Options.AutoFormatAsYouTypeReplaceQuotes = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“直引号替换为弯引号”选项的状态。*/
let blnAutoFormat = Options.AutoFormatReplaceQuotes
```

### Options.AutoFormatAsYouTypeReplaceSymbols

如果键入的两个连续的连字符 (--) 将替换为短划线 (–) 或长划线 (—)，则该属性值为 **True** 。 **Boolean** 类型，可读写。如果键入连字符时带有前导和尾部空格，则 WPS 将连字符替换为短划线。如果没有尾部空格，则连字符替换为长划线。

**语法**

**express.AutoFormatAsYouTypeReplaceSymbols**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例启用将键入的将连字符替换为符号的功能。*/
Options.AutoFormatAsYouTypeReplaceSymbols = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“连字符 (--) 替换为长划线 (—)”选项的状态。*/
let blnAutoFormat = Options.AutoFormatAsYouTypeReplaceSymbols
```

### Options.AutoFormatDeleteAutoSpaces

如果 WPS 在自动设置文档或区域的格式时，删除日文和西文文字之间插入的空格，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatDeleteAutoSpaces**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 自动删除日文和西文文字之间的空格，然后设置当前选定内容的格式。*/
function test()
{
Options.AutoFormatDeleteAutoSpaces = true
Selection.Range.AutoFormat()
}
```

### Options.AutoFormatMatchParentheses

如果 WPS 在自动设置文档或区域的格式时，会更正不恰当匹配的括号，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatMatchParentheses**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 自动更正不恰当匹配的括号，然后设置当前选定内容的格式。*/
function test()
{
Options.AutoFormatMatchParentheses = true
Selection.Range.AutoFormat()
}
```

### Options.AutoFormatPlainTextWordMail

如果在 WPS 中打开纯文本格式的电子邮件时， WPS 将自动进行格式设置，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatPlainTextWordMail**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为可对打开的纯文本格式的电子邮件自动进行格式设置。*/
Options.AutoFormatPlainTextWordMail = true
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“纯文本型 WordMail 文档”选项的状态。*/
let blnAutoFormat = Options.AutoFormatPlainTextWordMail
```

### Options.AutoFormatPreserveStyles

如果当 WPS 自动设置文档或区域格式时，保留先前所用的样式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatPreserveStyles**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例中，先设置 Word，使其可保留原有样式，并自动设置标题、列表和其他段落的样式，然后 WPS 自动设置当前所选内容的格式。*/
function test()
{
Options.AutoFormatPreserveStyles = true
Options.AutoFormatApplyHeadings = true
Options.AutoFormatApplyLists = true
Options.AutoFormatApplyOtherParas = true

Selection.Range.AutoFormat()
}
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“样式”选项的状态。*/
let blnAutoFormat = Options.AutoFormatPreserveStyles
```

### Options.AutoFormatReplaceFarEastDashes

如果 WPS 在自动设置文档或区域的格式时，会更正长元音和划线的使用，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplaceFarEastDashes**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 自动更正长元音和划线的使用，然后设置当前选定内容的格式。*/
function test()
{
Options.AutoFormatReplaceFarEastDashes = true
Selection.Range.AutoFormat()
}
```

### Options.AutoFormatReplaceFractions

如果当 WPS 自动设置文档或区域格式时，会用当前字符集中的分数替换所键入的分数，则该属性值为 **True** 。例如，将“1/2”替换为“?”。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplaceFractions**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例激活键入分数的替换功能，然后自动设置当前选定内容的格式。*/
function test()
{
Options.AutoFormatReplaceFractions = true
Selection.Range.AutoFormat()
}
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“分数 (1/2) 替换为分数字符 (?)”选项的状态。*/
let blnAutoFormat = Options.AutoFormatReplaceFractions
```

### Options.AutoFormatReplaceHyperlinks

如果为 **True** ，则 WPS 在对某一文档或范围自动套用格式时，将对其中的电子邮件地址、服务器和共享名称（即 UNC 路径）和 Internet 地址（即 URL）自动进行格式设置。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplaceHyperlinks**

\*express是一个代表 Options 对象的变量。

**说明**

WPS 将任何看起来像电子邮件地址、UNC 或 URL 的文字更改为超链接。 WPS 并不检查超链接是否有效。

**示例**

``` JavaScript
/*本示例使 WPS 能自动将任何 Internet 或网络路径替换为超链接，然后对选定内容自动套用格式。*/
function test()
{
Options.AutoFormatReplaceHyperlinks = true
Selection.Range.AutoFormat()
}
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“Internet 及网络路径替换为超链接”选项的状态。*/
let blnAutoFormat = Options.AutoFormatReplaceHyperlinks
```

### Options.AutoFormatReplaceOrdinals

如果为 **True** ，则当 WPS 自动设置文档或区域格式时，用相同字母的上标格式替换序数后缀“st”、“nd”、“rd”和“th”。例如，将“1st”替换为“1”后接上标“st”。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplaceOrdinals**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将序数字母自动替换为上标，然后自动设置当前选定内容的格式。*/
function test()
{
Options.AutoFormatReplaceOrdinals = true
Selection.Range.AutoFormat()
}
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“序号(1st)替换为上标”选项的状态。*/
let blnAutoFormat = Options.AutoFormatReplaceOrdinals
```

### Options.AutoFormatReplacePlainTextEmphasis

如果当 WPS 自动设置文档或区域的格式时，将手动强调的字符替换为相应的字符格式（例如，“\*bold\*”更改为“bold”，“\_underline\_”更改为“underline”），则该属性值为 **True** 。 **Boolean** 类型，可读写。

如果当 WPS 自动设置文档或区域的格式时，将手动强调的字符替换为相应的字符格式（例如，“\*bold\*”更改为“bold”，“\_underline\_”更改为“underline”），则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplacePlainTextEmphasis**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例启用将手动强调字符替换为相应字符格式的功能。*/
function test()
{
Options.AutoFormatReplacePlainTextEmphasis = true
Selection.Range.AutoFormat()
}
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“*加粗*和 _下划线_ 替换为真正格式”选项的状态。*/
let blnAutoFormat = Options.AutoFormatReplacePlainTextEmphasis
```

### Options.AutoFormatReplaceQuotes

如果当 WPS 自动设置文档或区域格式时，直引号会自动更改为弯引号，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplaceQuotes**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例启用将直引号自动替换为弯引号的功能，然后自动设置当前所选内容的格式。*/
function test()
{
Options.AutoFormatReplaceQuotes = true
Selection.Range.AutoFormat()
}
```

``` JavaScript
/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“直引号替换为弯引号”选项的状态。*/
let blnAutoFormat = Options.AutoFormatReplaceQuotes
```

### Options.AutoFormatReplaceSymbols

如果当 WPS 自动设置文档或区域的格式时，将两个连续的连字号 (--) 替换为短破折号 (

**语法**

**express.AutoFormatReplaceSymbols**

\*express是一个代表 Options 对象的变量。

### Options.AutoKeyboardSwitching

如果 WPS 随时自动切换键盘语言并与键入的内容匹配，则该属性值为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.AutoKeyboardSwitching**

\*express是一个代表 Options 对象的变量。

**说明**

若要使用该属性，必须将 **CheckLanguage** 属性设为 **True** 。

**示例**

``` JavaScript
/*本示例请用户选择是否启用多语言文档的键盘自动切换功能。*/
function test()
{
let x = MsgBox("Enable automatic keyboard switching?", jsYesNo)
if (x == jsResultYes) {
    Application.CheckLanguage = true
    Options.AutoKeyboardSwitching = true
    MsgBox("Automatic keyboard switching enabled!")
}
}
```

### Options.AutoWordSelection

如果在拖动时，一次选定一个单词，而不是选定一个字符，则该属性为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoWordSelection**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在拖动时不选定整个单词，而是选定单个字符。*/
Options.AutoWordSelection = false
```

``` JavaScript
/*本示例返回“选项”对话框中“编辑”选项卡上“选定时自动选定整个单词”选项的状态。*/
let blnAutoSelect = Options.AutoWordSelection
```

### Options.BackgroundOpen

如果为 **True** ，WPS 在后台打开 Web 文档。 **Boolean** 类型，可读写。

**语法**

**express.BackgroundOpen**

\*express是一个代表 Options 对象的变量。

**说明**

当 WPS 在后台打开较大的 Web 文档时，用户可以在另一个文档中继续键入和选择命令。然而，Web 文档完全打开后，则不能在打开的文档中使用 WPS 的 示例代码 函数。

**示例**

``` JavaScript
/*本示例在不在后台打开较大的 Web 文档和在后台打开较大文档之间切换。*/
function BackOpen() {
    if ( Options.BackgroundOpen == false ) {
        Options.BackgroundOpen = true
    }
    else {
        Options.BackgroundOpen  = false
    }
}
```

### Options.BackgroundSave

如果为 **True** ，则 WPS 在后台保存文档。当 WPS 在后台保存文档时，用户可以继续键入内容或执行命令。 **Boolean** 类型，可读写。

**语法**

**express.BackgroundSave**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例让用户在 WPS 保存文档时可以继续处理文档。*/
Options.BackgroundSave = true
```

``` JavaScript
/*本示例返回“选项”对话框中“保存”选项卡上“允许后台保存”选项的当前状态。*/
let blnAutoSave = Options.BackgroundSave
```

### Options.BibliographySort

返回或设置一个 **String** 类型的值，该值代表在 **“源管理器”** 对话框中显示源的次序。可读写。

**语法**

**express.BibliographySort**

\*express是一个代表 Options 对象的变量。

**说明**

可能的 **BibliographySort** 属性值有 ` Author ` 、 ` Title ` 、 ` Tag ` 或 ` Year ` 。

### Options.BibliographyStyle

返回或设置一个 **String** 类型的值，该值代表用于设置书目格式的样式的名称。可读写。

**语法**

**express.BibliographyStyle**

\*express是一个代表 Options 对象的变量。

### Options.BrazilReform

返回或设置巴西葡萄牙语拼写器的模式。可读/写。

**语法**

**express.BrazilReform**

\*express是一个代表 Options 对象的变量。

**说明**

设置该属性与在 **“ WPS 选项”** 对话框中 **“巴西葡萄牙语模式:”** 旁边的下拉框中选择项目具有相同作用（ **“校对”** 、 **“在 WPS Office程序中更正拼写时”** ）。

| 注释                                                                                                       |
|------------------------------------------------------------------------------------------------------------|
| 该属性不设置欧洲葡萄牙语拼写器的模式。若要设置欧洲葡萄牙语拼写器模式，请使用 Options.PortugalReform 属性。 |

### Options.ButtonFieldClicks

返回或设置一个数字，该数字是运行 GOTOBUTTON 或 MACROBUTTON 域所需的点击次数（单击或双击）。 **Long** 类型，可读写。

**语法**

**express.ButtonFieldClicks**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将运行 GOTOBUTTON 或 MACROBUTTON 域所需的点击次数设置为 1。*/
Options.ButtonFieldClicks = 1
```

### Options.CheckGrammarAsYouType

如果 WPS 在键入时自动对文档进行语法检查并自动标记错误，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CheckGrammarAsYouType**

\*express是一个代表 Options 对象的变量。

**说明**

本属性标记语法错误，但必须在将 **ShowGrammaticalErrors** 的属性设置为 **True** 后才能在屏幕上显示错误标记。

**示例**

``` JavaScript
/*本示例对 WPS 进行设置，使其在您键入时检查活动文档中的语法错误并显示发现的所有错误。*/
function test()
{
Options.CheckGrammarAsYouType = true
ActiveDocument.ShowGrammaticalErrors = true
}
```

``` JavaScript
/*本示例返回“工具”菜单的“选项”对话框中“拼写和语法”选项卡上“键入时检查语法”选项的状态。*/
let blnCheck = Options.CheckGrammarAsYouType
```

### Options.CheckGrammarWithSpelling

如果 WPS 在检查拼写的同时也检查语法，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CheckGrammarWithSpelling**

\*express是一个代表 Options 对象的变量。

**说明**

当使用 **“工具”** 菜单中的 **“拼写”** 命令检查拼写时，可用本属性控制 WPS 是否同时检查语法。

若要通过 Visual Basic 过程检查拼写或语法，可用 **CheckSpelling** 方法只检查拼写，而用 **CheckGrammar** 方法同时检查拼写和语法。

**示例**

``` JavaScript
/*本示例返回“选项”对话框“拼写和语法”选项卡上“随拼写检查语法”选项的状态。如果选中该选项，则检查过程会检查活动文档的拼写和语法；否则只检查拼写。*/
function test()
{
if (Options.CheckGrammarWithSpelling == true) {
    ActiveDocument.CheckGrammar()
}
else {
    ActiveDocument.CheckSpelling()
}
}
```

### Options.CheckHangulEndings

如果 WPS 自动检测朝鲜文字结尾，并在将朝鲜文字转换为朝鲜文汉字过程中忽略朝鲜文字结尾，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CheckHangulEndings**

\*express是一个代表 Options 对象的变量。

**说明**

如果将朝鲜文汉字转换为朝鲜文字，则忽略该属性。

**示例**

``` JavaScript
/*本示例询问用户是否设置 WPS 自动检测朝鲜文字结尾，并在将朝鲜文字转换为朝鲜文汉字过程中忽略朝鲜文字结尾。*/
function test()
{
let x = MsgBox("Check Hangul endings during conversion from Hangul to Hanja?", jsYesNo)
if ( x == jsResultYes ) {
    Options.CheckHangulEndings = true
}
else {
    Options.CheckHangulEndings = false
}
}
```

### Options.CheckSpellingAsYouType

如果 WPS 在您键入的同时自动进行拼写检查并标记错误，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CheckSpellingAsYouType**

\*express是一个代表 Options 对象的变量。

**说明**

本属性将标记拼写错误。如要在屏幕上查看标记，则必须将 **ShowSpellingErrors** 属性设置为 **True** 。

**示例**

``` JavaScript
/*本示例关闭 WPS 中的自动拼写检查功能。*/
Options.CheckSpellingAsYouType = false
```

``` JavaScript
/*本示例设置 WPS 在键入的同时检查拼写错误，并在活动文档中显示所找到的错误。*/
function test()
{
Options.CheckSpellingAsYouType = true
ActiveDocument.ShowSpellingErrors = true
}
```

``` JavaScript
/*本示例返回“工具”菜单的“选项”对话框中“拼写和语法”选项卡上“键入时检查拼写”选项的状态。*/
let blnCheck = Options.CheckSpellingAsYouType
```

### Options.CommentsColor

该属性返回或设置一个 **WdColorIndex** 常量，代表文档中批注的颜色。可读写。

**语法**

**express.CommentsColor**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 的全局选项，根据批注的作者设置批注的颜色。*/
function ColorCodeComments() {
    Options.CommentsColor = wdByAuthor
}
```

### Options.ConfirmConversions

如果 WPS 在打开或者插入一个非 WPS 文档或模板的文件之前，会显示一个 **“转换文件”** 对话框，则该属性值为 **True** 。在 **“转换文件”** 对话框中，用户可以选择需要转换的文件格式。 **Boolean** 类型，可读写。

**语法**

**express.ConfirmConversions**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在打开一个非 WPS 文档或模板的文件之前，显示“转换文件”对话框。*/
Options.ConfirmConversions = true
```

``` JavaScript
/*本示例返回“选项”对话框中“常规”选项卡上“打开时确认转换”选项的当前状态。*/
let blnConfirm = Options.ConfirmConversions
```

### Options.ContextualSpeller

返回或设置 **Boolean** 值，该值表示是否使用上下文拼写检查器，基于单词上下文和它前后的单词来检查拼写。可读/写。

**语法**

**express.ContextualSpeller**

\*express是一个代表 Options 对象的变量。

**说明**

上下文拼写检查器可以识别出单词的结构和它们在句子中的位置，以确定拼写是否正确。它可以找到标准拼写检查器找不到的错误。例如，用户可能键入“superb owl”，而不是“super bowl”。由于“superb”和“owl”都是拼写正确的单词，因此标准拼写检查器找不到错误。基于句子上下文中的单词以及它们前后的单词，上下文拼写检查器就能确定正确单词是“super”和“bowl”，并自动进行更改。

此属性对应于 **“ WPS 选项”** 对话框的 **“校对”** 选项卡中的 **“使用上下文拼写检查”** 复选框。

### Options.ConvertHighAnsiToFarEast

如果在打开文档时，WPS 将与东亚字体相关的文字转换为相应的字体，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ConvertHighAnsiToFarEast**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在打开文档时，将与东亚字体相关的文字转换为相应的字体。*/
Options.ConvertHighAnsiToFarEast = true
```

### Options.CreateBackup

如果 WPS 在每次保存文档时都建立一个备份，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CreateBackup**

\*express是一个代表 Options 对象的变量。

**说明**

**CreateBackup** 和 **AllowFastSave** 属性不能同时设置为 **True** 。

**示例**

``` JavaScript
/*本示例设置 WPS 自动创建一个备份，然后开始保存活动文档。*/
function test()
{
Options.CreateBackup = true
ActiveDocument.Save()
}
```

### Options.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Options 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Options.CtrlClickHyperlinkToOpen

如果 WPS 需要单击同时按住 Ctrl 键来打开超链接，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CtrlClickHyperlinkToOpen**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例禁用需要在单击的同时按住 Ctrl 键来打开超链接的选项。*/
function ToggleHyperlinkOption() {
    Options.CtrlClickHyperlinkToOpen = false
}
```

### Options.CursorMovement

返回或设置在双向文字中插入点移动的方式。 **WdCursorMovement** 类型，可读写。

**语法**

**express.CursorMovement**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*当插入点在双向文字中移动时，本示例使其前进至下一个可见的邻近字符。*/
Options.CursorMovement = wdCursorMovementVisual
```

### Options.DefaultBorderColor

返回或设置用于新的 **Border** 对象的默认 24 位颜色。可读写。

**语法**

**express.DefaultBorderColor**

\*express是一个代表 Options 对象的变量。

**说明**

可以是任何有效的 **WdColor** 常量或 Visual Basic 的 **RGB** 函数返回的值。

**示例**

``` JavaScript
/*本示例将新边框的默认颜色设为蓝绿色。*/
Options.DefaultBorderColor = wdColorTeal
```

### Options.DefaultBorderColorIndex

返回或设置边框的默认线条颜色。 **WdColorIndex** 类型，可读写。

**语法**

**express.DefaultBorderColorIndex**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例更改边框的默认颜色和样式，并为活动文档的第一段设置一个边框。*/
function test()
{
ActiveDocument.Paragraphs.Item(1).Borders.Enable = true
Options.DefaultBorderColorIndex = wdRed
Options.DefaultBorderLineStyle = wdLineStyleDouble

}
```

### Options.DefaultBorderLineStyle

返回或设置默认边框的样式。 **WdLineStyle** 类型，可读写。

**语法**

**express.DefaultBorderLineStyle**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将默认样式设置为双线。*/
Options.DefaultBorderLineStyle = wdLineStyleDouble
```

``` JavaScript
/*本示例返回当前默认的样式。*/
let lngTemp = Options.DefaultBorderLineStyle
```

### Options.DefaultBorderLineWidth

返回或设置边框的默认线条宽度。 **WdLineWidth** 类型，可读写。

**语法**

**express.DefaultBorderLineWidth**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例更改边框的默认线条宽度，然后为选定内容的每一段周围添加边框。*/
function test()
{
Options.DefaultBorderLineWidth = wdLineWidth050pt
Selection.Borders.Enable = true
}
```

### Options.DefaultEPostageApp

设置或返回一个 **String** 类型的值，该值代表默认的电子邮政应用程序的路径和文件名。可读写。

**语法**

**express.DefaultEPostageApp**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例指定默认的电子邮政应用程序的路径和文件名。*/
function DefaultEPostage() {
    Application.Options.DefaultEPostageApp = "C:\\MyApp\\EPostage.exe"
}
```

### Options.DefaultFilePath

返回或设置各项（例如文档、模板和图形）的默认文件夹。 **String** 类型，可读写。

**语法**

**express.DefaultFilePath**

\*express是一个代表 Options 对象的变量。

**说明**

| **名称** | **必选/可选** | **数据类型**          | **说明**           |
|----------|---------------|-----------------------|--------------------|
| *Path*   | 必选          | **WdDefaultFilePath** | 设置默认的文件夹。 |

您可以使用空字符串 ("") 从 Windows 注册表中删除设置。新设置会立即生效。

**示例**

``` JavaScript
/*本示例为 WPS 文档设置默认文件夹。*/
Application.Options.DefaultFilePath(wdDocumentsPath,"C:\\Documents") 

/*本示例返回用户模板当前的默认路径（相当于“选项”对话框中“文件位置”选项卡上的默认路径设置）。*/
let strPath = Application.Options.DefaultFilePath(wdUserTemplatesPath)
```

### Options.DefaultHighlightColorIndex

返回或设置突出显示文本所用的颜色，用 **“格式”** 工具栏的 **“突出显示”** 按钮可将文本格式设为突出显示。 **WdColorIndex** 类型，可读写。

**语法**

**express.DefaultHighlightColorIndex**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将默认的突出显示颜色设置为浅绿色。新的颜色不应用于以前突出显示的任何文字。*/
Options.DefaultHighlightColorIndex = wdBrightGreen
```

``` JavaScript
/*本示例返回当前默认的突出显示颜色索引。*/
let lngTemp = Options.DefaultHighlightColorIndex
```

### Options.DefaultOpenFormat

返回或设置打开文档的默认转换器。可为由 **OpenFormat** 属性返回的一个数字，也可为下列 **WdOpenFormat** 常量之一。

**语法**

**express.DefaultOpenFormat**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例首先将 WPS 文档格式设置为打开文档的默认转换器，然后打开文档 Forecast.doc。*/
function test()
{
Options.DefaultOpenFormat = wdOpenFormatDocument
Documents.Open("C:\\Sales\\Forecast.doc")
}
```

``` JavaScript
/*本示例将打开文档所需的默认转换器设为在打开文档时自动选用相应的文件转换器。*/
Options.DefaultOpenFormat = wdOpenFormatAuto
```

``` JavaScript
/*本示例将 WordPerfect 6.x 格式设为打开文档的默认转换器。*/
function test()
{
Options.DefaultOpenFormat =
    FileConverters.Item("WordPerfect6x").OpenFormat
}
```

### Options.DefaultTextEncoding

返回或设置一个 **MsoEncoding** 常量，该常量代表 WPS 用于所有保存为编码文本文件的文档的代码页或字符集。可读写。

**语法**

**express.DefaultTextEncoding**

\*express是一个代表 Options 对象的变量。

**说明**

使用 **TextEncoding** 属性设置个别文档的编码。若要设置 HTML 文档的编码，请使用 **Encoding** 属性。

**示例**

``` JavaScript
/*本示例将全局文本编码设为 Western 代码页，这表示 WPS 将使用 Western 代码页来保存所有编码文本文件。*/
function test()
{
function DefaultEncode() {
    Application.Options.DefaultTextEncoding = msoEncodingWestern
}
}
```

### Options.DefaultTray

返回或设置默认纸盒，以便打印机使用该纸盒打印文档。 **String** 类型，可读写。

**语法**

**express.DefaultTray**

\*express是一个代表 Options 对象的变量。

**说明**

要设置该属性，必须在 **“选项”** 对话框 **“打印”** 选项卡上 **“默认纸盒”** 框中指定一个字符串。使用 **DefaultTrayID** 属性并指定一个 **WdPaperTray** 常量也可设置该选项。

**示例**

``` JavaScript
/*本示例设置 WPS 使用下层纸盒。*/
Options.DefaultTray = "Lower tray"
```

``` JavaScript
/*本示例返回在“选项”对话框中“打印”选项卡上“默认纸盒”框中找到的字符串。*/
Msgbox(Options.DefaultTray)
```

### Options.DefaultTrayID

返回或设置用于打印文档的默认纸盒。 **WdPaperTray** 类型，可读写。

**语法**

**express.DefaultTrayID**

\*express是一个代表 Options 对象的变量。

**说明**

将 **DefaultTray** 属性与 **“选项”** 对话框 **“打印”** 选项卡上 **“默认纸盒”** 框中的字符串结合使用，也可设置该选项。

**示例**

``` JavaScript
/**/
function test()
{
Options.DefaultTrayID = wdPrinterUpperBin
ActiveDocument.PrintOut()
}
```

``` JavaScript
/*本示例返回 “选项”对话框中“打印”选项卡上“默认纸盒”选项的当前设置。*/
let lngTray = Options.DefaultTrayID
```

### Options.DeletedCellColor

返回或设置一个 **WdCellColor** 常量，该常量代表删除的单元格的颜色。可读写。

**语法**

**express.DeletedCellColor**

\*express是一个代表 Options 对象的变量。

### Options.DeletedTextColor

当激活修订选项时，返回或设置要删除的文字的颜色。 **WdColorIndex** 类型，可读写。

**语法**

**express.DeletedTextColor**

\*express是一个代表 Options 对象的变量。

**说明**

如果将 **DeletedTextColor** 属性设置为 **wdByAuthor** ， WPS 会自动为文档的前八名审阅者分配不同的颜色。

**示例**

``` JavaScript
/*本示例将要删除文本的颜色设置为鲜绿色。*/
Options.DeletedTextColor = wdBrightGreen
```

``` JavaScript
/*本示例返回“选项”对话框中“修订”选项卡上“删除的文字”下“颜色”选项的当前状态。*/
let lngTemp = Options.DeletedTextColor
```

### Options.DeletedTextMark

当激活修订选项时，返回或设置要删除的文字的格式。 **WdDeletedTextMark** 类型，可读写。

**语法**

**express.DeletedTextMark**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将要删除的文本的格式设置为删除线。*/
Options.DeletedTextMark = wdDeletedTextMarkStrikeThrough
```

``` JavaScript
/*本示例返回“选项”对话框中“修订”选项卡上“删除的文字”下“标记”选项的当前状态。*/
let lngTemp = Options.DeletedTextMark
```

### Options.DiacriticColorVal

返回或设置用于从右向左语言文档中的音调符号的 24 位颜色。

**语法**

**express.DiacriticColorVal**

\*express是一个代表 Options 对象的变量。

**说明**

此属性可以是任何有效的 **WdColor** 常量或 Visual Basic 的 **RGB** 函数返回的值。必须将 **UseDiffDiacColor** 属性的值设为 **True** 才能使用此属性。

**示例**

``` JavaScript
/*本示例将音调字符的颜色设为鲜绿色。*/
function test()
{
if ( Options.UseDiffDiacColor == true ) {
    Options.DiacriticColorVal = wdColorBrightGreen
}
}
```

### Options.DisableFeaturesbyDefault

如果为 **True** ，则 WPS 将在所有文档中禁用 **DisableFeaturesIntroducedAfterbyDefault** 中指定的 WPS 版本后引入的所有功能。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.DisableFeaturesbyDefault**

\*express是一个代表 Options 对象的变量。

**说明**

**DisableFeaturesByDefault** 属性设置应用程序的全局选项。如果仅对该文档禁用 WPS 97 for Windows 版本后引入的所有功能，可使用 **DisableFeatures** 属性。

**示例**

``` JavaScript
/*本示例对所有文档禁用 WPS for Windows 的 7.0 和 7.0a 版之后引入的所有功能。*/
function FeaturesDisableByDefault() {
    let  options = Application.Options

    //Checks whether features are disabled
    if ( options.DisableFeaturesbyDefault == true ) {

        //If they are, disables all features after  WPS for Windows
        options.DisableFeaturesIntroducedAfterbyDefault = wd70
    }
    else {

        //If not, turns on the disable features option and disables
        //all features introduced after  WPS for Windows
        options.DisableFeaturesbyDefault = true
        options.DisableFeaturesIntroducedAfterbyDefault = wd70
    }
}
```

### Options.DisableFeaturesIntroducedAfterbyDefault

对所有文档禁用指定版本后引入的所有功能。 **WdDisableFeaturesIntroducedAfter** 类型，可读写。

**语法**

**express.DisableFeaturesIntroducedAfterbyDefault**

\*express是一个代表 Options 对象的变量。

**说明**

设置 **DisableFeaturesIntroducedAfterByDefault** 属性之前，必须先将 **DisableFeaturesByDefault** 属性设为 **True** 。否则该设置无效，并保持默认设置 WPS 97 for Windows。

**DisableFeaturesIntroducedAfterByDefault** 属性设置应用程序的全局选项并影响所有文档。如果仅对某文档禁用指定版本后引入的所有功能，可使用 **DisableFeaturesIntroducedAfter** 属性。

**示例**

``` JavaScript
/*本示例对所有文档禁用 WPS for Windows 的 7.0 和 7.0a 版之后引入的所有功能。*/
function FeaturesDisableByDefault() {
    let options = Application.Options

    //Checks whether features are disabled
    if ( options.DisableFeaturesbyDefault == true ) {

        //If they are, disables all features after  WPS for Windows
        options.DisableFeaturesIntroducedAfterbyDefault = wd70
    }
    else {

        //If not, turns on the disable features option and disables
        //all features introduced after  WPS for Windows
        options.DisableFeaturesbyDefault = true
        options.DisableFeaturesIntroducedAfterbyDefault = wd70
    }
}
```

### Options.DisplayGridLines

如果 WPS 显示文档网格，则该属性值为 **True** 。该属性相当于 **“视图”** 菜单中的 **“网格线”** 命令。 **Boolean** 类型，可读写。

**语法**

**express.DisplayGridLines**

\*express是一个代表 Options 对象的变量。

**说明**

该属性只影响文档网格。对于表格虚框，请使用 **TableGridlines** 属性。

**示例**

``` JavaScript
/*本示例实现的功能是：在显示或隐藏活动窗口的文档网格之间切换。*/
Options.DisplayGridLines = ( !Options.DisplayGridLines )
```

### Options.DisplayPasteOptions

如果为 **True** ，则 WPS 直接在新粘贴的文本下显示 **“粘贴选项”** 按钮。 **Boolean** 类型，可读写。

**语法**

**express.DisplayPasteOptions**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*如果“粘贴选项”按钮无效，则本示例将启用该选项。*/
function ShowPasteOptionsButton() {
    if ( Options.DisplayPasteOptions == false ) {
        Options.DisplayPasteOptions = true
    }
}
```

### Options.DisplaySmartTagButtons

如果为 **True** ，则 WPS 在鼠标指针置于智能标记上方时显示一个按钮。 **Boolean** 类型，可读写。

**语法**

**express.DisplaySmartTagButtons**

\*express是一个代表 Options 对象的变量。

**说明**

智能标记按钮提供一个下拉菜单，以便用户访问智能标记选项和操作。

**示例**

``` JavaScript
/*本示例隐藏在鼠标指针置于智能标记上方时按钮的显示。*/
function DontShowSmartTagButton() {
    Options.DisplaySmartTagButtons = false
}
```

### Options.DocumentViewDirection

返回或设置整篇文档的对齐方式和阅读顺序。 **WdDocumentViewDirection** 类型，可读写。

**语法**

**express.DocumentViewDirection**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将整篇文档设为右对齐方式和从右向左的阅读顺序。*/
Options.DocumentViewDirection = wdDocumentViewRtl
```

### Options.DoNotPromptForConvert

设置或返回一个 **Boolean** ，它代表在为处于兼容模式下的文档调用“转换”命令时是否出现一个警告对话框。可读写。

**语法**

**express.DoNotPromptForConvert**

\*express是一个代表 Options 对象的变量。

### Options.EnableHangulHanjaRecentOrdering

如果 WPS 在进行朝鲜文字和朝鲜文汉字间的转换时，在建议列表的顶端显示最近用过的单词，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableHangulHanjaRecentOrdering**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例询问用户是否设置 WPS 在进行朝鲜文字和朝鲜文汉字间的转换时，在建议列表的顶端显示最近用过的单词。*/
function test()
{
let x = MsgBox("Display most recently used words at the top of the suggestions list?", jsYesNo)
if ( x == jsResultYes ) {
    Options.EnableHangulHanjaRecentOrdering = true
}
else {
    Options.EnableHangulHanjaRecentOrdering = false
}
}
```

### Options.EnableLegacyIMEMode

返回或设置一个 **Boolean** 类型的值，该值代表是否启用旧式 IME 模式。可读写。

**语法**

**express.EnableLegacyIMEMode**

\*express是一个代表 Options 对象的变量。

### Options.EnableLivePreview

设置或返回 **Boolean** 值，该值表示显示还是隐藏在使用支持预览的库时所出现的库预览。如果为 **True** ，将在应用命令之前在文档中显示预览。可读/写。

**语法**

**express.EnableLivePreview**

\*express是一个代表 Options 对象的变量。

**说明**

此属性对应于 **“ WPS 选项”** 对话框中的 **“启用实时预览”** 复选框。

### Options.EnableMisusedWordsDictionary

如果 WPS 在检查文档的拼写和语法时，也检查误用的单词，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableMisusedWordsDictionary**

\*express是一个代表 Options 对象的变量。

**说明**

在检查误用单词时， WPS 将检查以下内容：形容词和副词、比较级和最高级、“like”作为连词、“nor”与“or”、“what”与“which”、“who”与“whom”、度量单位、连词、介词和代名词的不正确用法。

**示例**

``` JavaScript
/*本示例设定 WPS 在拼写和语法检查中忽略误用的单词。*/
Options.EnableMisusedWordsDictionary = false
```

### Options.EnableSound

如果 WPS 在导致出错时使计算机发出声音进行响应，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableSound**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例根据用户的输入，设置“选项”对话框“常规”选项卡上的“提供声音反馈”选项。*/
function test()
{
if ( MsgBox("Do you want  WPS to beep on errors?", 36) == jsResultYes ) {
    Options.EnableSound = true
}
else {
    Options.EnableSound = false
}
}
```

### Options.EnvelopeFeederInstalled

如果当前的打印机为信封指定了送纸器，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.EnvelopeFeederInstalled**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*如果已安装了信封送纸器，本示例以信封方式打印活动文档。*/
function test()
{
if ( Options.EnvelopeFeederInstalled == true ) {
    ActiveDocument.Envelope.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,InchesToPoints(3),InchesToPoints(1.5))
}
else {
    Msgbox("No envelope feeder available.")
}
}
```

### Options.FormatScanning

如果为 **True** ，则 WPS 对文档中的所有格式进行跟踪。 **Boolean** 类型，可读写。

**语法**

**express.FormatScanning**

\*express是一个代表 Options 对象的变量。

**说明**

启用 **FormatScanning** 属性允许您识别文档中所有单独的格式，这样您可以方便地将相同的格式应用于新文本，并快速替换或修改文档中出现的所有给定的格式。

**示例**

``` JavaScript
/*本示例启用 WPS 对文档中格式的跟踪，但禁止在文字下方显示波浪下划线，这些文字的格式类似于文档中使用频率更高的其他格式。*/
function ShowFormatErrors() {
    Options.FormatScanning = true
    Options.ShowFormatError = false  //Disables displaying squiggly underline
}
```

### Options.FrenchReform

返回或设置一个 **WdFrenchSpeller** 常量，该常量代表要对语言格式设置为法语的文本区域使用哪个拼写词典。可读写。

**语法**

**express.FrenchReform**

\*express是一个代表 Options 对象的变量。

### Options.GridDistanceHorizontal

当在新文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置 WPS 使用的不可见网格线的水平间距。 **Single** 类型，可读写。

**语法**

**express.GridDistanceHorizontal**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设置网格线的水平和垂直间距，并为新文档启动“对象与网格对齐”功能。*/
function test()
{
Options.GridDistanceHorizontal = InchesToPoints(0.2)
Options.GridDistanceVertical = InchesToPoints(0.2)
Options.SnapToGrid = true
Documents.Add()
}
```

### Options.GridDistanceVertical

在新文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置 WPS 所用的不可见网格线间的垂直间距。 **Single** 类型，可读写。

**语法**

**express.GridDistanceVertical**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设置网格线的水平和垂直间距，并为新文档启动“对象与网格对齐”功能。*/
function test()
{
Options.GridDistanceHorizontal = InchesToPoints(0.2)
Options.GridDistanceVertical = InchesToPoints(0.2)
Options.SnapToGrid = true
Documents.Add()
}
```

### Options.GridOriginHorizontal

返回或设置不可见网格相对于页面左边的起点位置，当在新文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。 **Single** 类型，可读写。

**语法**

**express.GridOriginHorizontal**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设置网格的水平和垂直起点，以及网格线之间的水平和垂直距离，并为新文档启用“对象与网格对齐”功能。*/
function test()
{
Options.GridOriginHorizontal = InchesToPoints(1)
Options.GridOriginVertical = InchesToPoints(2)
Options.GridDistanceHorizontal = InchesToPoints(0.1)
Options.GridDistanceVertical = InchesToPoints(0.1)
Options.SnapToGrid = true
Documents.Add()
}
```

### Options.GridOriginVertical

返回或设置不可见网格相对于页面顶边的起点位置，当在新文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。 **Single** 类型，可读写。

**语法**

**express.GridOriginVertical**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设置网格的水平和垂直起点，以及网格线之间的水平和垂直距离，并为新文档启用“对象与网格对齐”功能。*/
function test()
{
Options.GridOriginHorizontal = InchesToPoints(1)
Options.GridOriginVertical = InchesToPoints(2)
Options.GridDistanceHorizontal = InchesToPoints(0.2)
Options.GridDistanceVertical = InchesToPoints(0.2)
Options.SnapToGrid = true
Documents.Add()
}
```

### Options.HangulHanjaFastConversion

如果 WPS 在朝鲜文字和朝鲜文汉字间进行转换时，自动转换具有单独建议的单词，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.HangulHanjaFastConversion**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例询问用户，在朝鲜文字和朝鲜文汉字间进行转换时，是否设置 WPS 使用快速转换功能。*/
function test()
{
let x = MsgBox("Use fast conversion?", jsYesNo)
if ( x == jsResultYes ) {
    Options.HangulHanjaFastConversion = true
}
else {
    Options.HangulHanjaFastConversion = false
}
}
```

### Options.HebrewMode

返回或设置希伯来语拼写检查工具的模式。 **WdHebSpellStart** 类型，可读写。

**语法**

**express.HebrewMode**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置拼写检查工具在进行拼写检查时，遵循希伯来语言学院规定的包含音调符号的文字的常规书写规则。*/
Options.HebrewMode = wdFullScript
```

### Options.HideChartDraftModeNotification

如果在草图模式下插入的图表上隐藏 **“草图模式”** 通知标签，则该属性值为 **“True”** 。可读/写。

**语法**

**express.HideChartDraftModeNotification**

\*express是一个代表 Options 对象的变量。

**说明**

设置该属性不会影响先前在草图模式下插入的图表。

### Options.IgnoreInternetAndFileAddresses

如果为 **True** ，则进行拼写检查时忽略文件扩展名、MS-DOS 路径、电子邮件地址、服务器名称、共享名（也称为 UNC 路径）和 Internet 地址（也称为 URL）。 **Boolean** 类型，可读写。

**语法**

**express.IgnoreInternetAndFileAddresses**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 忽略文件名和 Internet 地址，然后对活动文档进行拼写检查。*/
function test()
{
Options.IgnoreInternetAndFileAddresses = true
ActiveDocument.CheckSpelling()
}
```

``` JavaScript
/*本示例返回“选项”对话框中“拼写和语法”选项卡上“忽略 Internet 和文件地址”选项的当前状态。*/
let blnTemp = Options.IgnoreInternetAndFileAddresses
```

### Options.IgnoreMixedDigits

如果进行拼写检查时忽略包含数字的单词，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IgnoreMixedDigits**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 忽略包含数字的单词，然后对活动文档进行拼写检查。*/
function test()
{
Options.IgnoreMixedDigits = true
ActiveDocument.CheckSpelling()
}
```

``` JavaScript
/*本示例返回“选项”对话框中“拼写和语法”选项卡上“忽略带数字的单词”选项当前的状态。*/
let blnTemp = Options.IgnoreMixedDigits
```

### Options.IgnoreUppercase

如果进行拼写检查时忽略全部大写的单词，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IgnoreUppercase**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 忽略全部大写的单词，然后对活动文档进行拼写检查。*/
function test()
{
Options.IgnoreUppercase = true
ActiveDocument.CheckSpelling()
}
```

``` JavaScript
/*本示例返回“选项”对话框中“拼写和语法”选项卡上“忽略所有字母都大写的单词”选项当前的状态。*/
let blnTemp = Options.IgnoreUppercase
```

### Options.IMEAutomaticControl

如果设置 WPS 自动打开和关闭日文输入法编辑器 (IME)，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IMEAutomaticControl**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 自动打开和关闭日文输入法编辑器 (IME)。*/
Options.IMEAutomaticControl = true
```

### Options.InlineConversion

如果 WPS 将中文输入法 (IME) 中的未确认字符串显示为已有（已确认的）字符串之间的插入项，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.InlineConversion**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 将中文输入法 (IME) 中的未确认字符串显示为已有（已确认的）字符串之间的插入项。*/
Options.InlineConversion = true
```

### Options.InsertChartInDraftMode

如果使用草图模式在文档中插入图表，则该属性值为 **True** 。可读/写。

**语法**

**express.InsertChartInDraftMode**

\*express是一个代表 Options 对象的变量。

**说明**

设置 **InsertChartInDraftMode** 属性与单击 WPS 2015 中 **“ WPS 选项”** 对话框中的 **“使用草图模式插入图表”** 具有相同作用（单击“高级”，滚动到“图表”）。

### Options.InsertedCellColor

返回或设置一个 **WdCellColor** 常量，该常量代表插入的表格单元格的颜色。可读写。

**语法**

**express.InsertedCellColor**

\*express是一个代表 Options 对象的变量。

### Options.InsertedTextColor

当启用修订选项时，返回或设置插入文字的颜色。 **WdColorIndex** 类型，可读写。

**语法**

**express.InsertedTextColor**

\*express是一个代表 Options 对象的变量。

**说明**

如果将 **InsertedTextColor** 属性设置为 **wdByAuthor** ，WPS 会自动为文档的前八个审阅者各赋予唯一的颜色。

**示例**

``` JavaScript
/*本示例将插入文本的颜色设置为深红色。*/
Options.InsertedTextColor = wdDarkRed
```

``` JavaScript
/*本示例返回“选项”对话框中“修订”选项卡上“修订选项”下“颜色”选项的当前状态。*/
let lngColor = Options.InsertedTextColor
```

### Options.InsertedTextMark

返回或设置启用修订时 WPS 设置插入文本格式的方式（ **TrackRevisions** 属性为 **True** ）。 **WdInsertedTextMark** 类型，可读写。

**语法**

**express.InsertedTextMark**

\*express是一个代表 Options 对象的变量。

**说明**

如果未启用修订，则会忽略该属性。将该属性与 **InsertedTextColor** 属性结合使用可控制文档中插入文本的外观。

要在编辑过程中查看插入文本的格式， **ShowRevisions** 属性必须为 **True** 。如果要 WPS 在打印文档时使用插入文本的格式， **PrintRevisions** 属性必须为 **True** 。

**示例**

``` JavaScript
/*本示例设置 WPS 为插入的文本应用斜体格式。*/
Options.InsertedTextMark = wdInsertedTextMarkItalic
```

``` JavaScript
/*本示例设置 WPS 为插入的文本应用加粗格式（如果尚未将其设置为加粗格式）。*/
function test()
{
if ( Options.InsertedTextMark != wdInsertedTextMarkBold ) {
    Options.InsertedTextMark = wdInsertedTextMarkBold
}
else {
    MsgBox("Inserted text is already bold!")
}
}
```

### Options.INSKeyForOvertype

如果可以用 Ins 键来打开和关闭“改写”，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.INSKeyForOvertype**

\*express是一个代表 Options 对象的变量。

### Options.INSKeyForPaste

如果可用 INS 键来粘贴“剪贴板”的内容，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.INSKeyForPaste**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例启用 INS 来粘贴“剪贴板”的内容。*/
Options.INSKeyForPaste = true
```

``` JavaScript
/*本示例返回“选项”对话框中“编辑”选项卡上“用 INS 粘贴”选项的状态。*/
let blnTemp = Options.INSKeyForPaste
```

### Options.InterpretHighAnsi

返回或设置高端字符译码方式。 **WdHighAnsiText** 类型，可读写。

**语法**

**express.InterpretHighAnsi**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 不将所有高端字符译码为东亚字符。*/
Options.InterpretHighAnsi = wdHighAnsiIsFarEast
```

### Options.LabelSmartTags

如果为 **True** ，WPS 使用智能标记信息标记文档中的文本。 **Boolean** 类型，可读写。

**语法**

**express.LabelSmartTags**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例在文档中关闭智能标记功能。*/
function MarkSmartTags() {
    Application.Options.LabelSmartTags = false
}
```

### Options.LocalNetworkFile

如果在编辑保存在网络服务器上的文件时，WPS 在用户的计算机上创建文件的一个本地副本，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.LocalNetworkFile**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例指示 WPS 不对保存在服务器上的文件创建本地副本文件。*/
function LocalFile() {
    Application.Options.LocalNetworkFile = false
}
```

### Options.MapPaperSize

如果为 **True** ，则自动调整采用其他国家/地区标准纸张大小（例如 A4）的文档格式，以便按照用户所在国家/地区的标准纸张大小（例如，Letter）正确打印该文档。 **Boolean** 类型，可读写。

**语法**

**express.MapPaperSize**

\*express是一个代表 Options 对象的变量。

**说明**

本属性只影响您的文档的打印输出，文档的格式不变。

**示例**

``` JavaScript
/*本示例允许 WPS 根据国家/地区设置调整纸张大小。*/
Options.MapPaperSize = true
```

``` JavaScript
/*本示例返回“工具”菜单的“选项”对话框中“打印”选项卡上“允许重调 A4/Letter 纸型”选项的状态。*/
let temp = Options.MapPaperSize
```

### Options.MatchFuzzyAY

如果在搜索过程中 WPS 将忽略 行和 行字符后“ ”和“ ”之间的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyAY**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 在搜索过程中忽略位于 i 行和 e 行字符后面的字符“ a”与“ ya”的差异。*/
Options.MatchFuzzyAY = true
```

### Options.MatchFuzzyBV

如果 WPS 将在搜索过程中忽略“ ”与“ ”之间以及“ ”与“ ”之间的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyBV**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 在搜索过程中忽略“ ba”与“ vua vua”之间以及“ ha”与“ fua fua”之间的差异。*/
Options.MatchFuzzyBV = true
```

### Options.MatchFuzzyByte

如果在搜索过程中 WPS 将忽略全角和半角字符（西文或日文）的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyByte**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在搜索过程中忽略全角和半角字符（西文或日文）的差异。*/
Options.MatchFuzzyByte = true
```

### Options.MatchFuzzyCase

如果 WPS 在搜索过程中不区分字母大小写，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyCase**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 在搜索过程中不区分字母的大小写。*/
Options.MatchFuzzyCase = true
```

### Options.MatchFuzzyDash

如果 WPS 在搜索过程中忽略减号、划线以及长元音符号之间的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyDash**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在搜索过程中忽略减号、划线以及长元音符号之间的差异。*/
Options.MatchFuzzyDash = true
```

### Options.MatchFuzzyDZ

如果 WPS 将在搜索过程中忽略“ ”与“ ”之间以及“ ”与“ ”之间的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyDZ**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 在搜索过程中忽略“ di”与“ zi”之间以及“ du”与“ zu”之间的差异。*/
Options.MatchFuzzyDZ = true
```

### Options.MatchFuzzyHF

如果 WPS 将在搜索过程中忽略“ ”与“ ”之间以及“ ”与“ ”之间的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyHF**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 在搜索过程中忽略“ heyu heyu”与“ fuyu fuyu”之间以及“ beyu beyu”与“ vuyu vuyu”之间的差异。*/
Options.MatchFuzzyHF = true
```

### Options.MatchFuzzyHiragana

如果 WPS 在搜索过程中将忽略平假名与片假名之间的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyHiragana**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在搜索过程中忽略平假名与片假名之间的差异。*/
Options.MatchFuzzyHiragana = true
```

### Options.MatchFuzzyIterationMark

如果 WPS 在搜索过程中将忽略重复的标记类型间的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyIterationMark**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在搜索过程中忽略重复的标记类型间的差异。*/
Options.MatchFuzzyIterationMark = true
```

### Options.MatchFuzzyKanji

如果 WPS 在搜索过程中忽略标准和非标准日文汉字的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyKanji**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在搜索过程中忽略标准和非标准日文汉字的差异。*/
Options.MatchFuzzyKanji = true
```

### Options.MatchFuzzyKiKu

如果在搜索过程中 WPS 将忽略 行字符前“ ”和“ ”之间的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyKiKu**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 在搜索过程中忽略位于 sa 行字符之前的字符“ ki”和“ ku”之间的差异。*/
Options.MatchFuzzyKiKu = true
```

### Options.MatchFuzzyOldKana

如果 WPS 在搜索过程中将忽略新旧假名字符之间的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyOldKana**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在搜索过程中忽略新旧假名字符的差异。*/
Options.MatchFuzzyOldKana = true
```

### Options.MatchFuzzyProlongedSoundMark

如果在搜索过程中 WPS 将忽略短元音和长元音的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyProlongedSoundMark**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在搜索过程中忽略短元音和长元音的差异。*/
Options.MatchFuzzyProlongedSoundMark = true
```

### Options.MatchFuzzyPunctuation

如果在搜索过程中 WPS 忽略标点类型的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyPunctuation**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在搜索过程中忽略标点的类型差异。*/
Options.MatchFuzzyPunctuation = true
```

### Options.MatchFuzzySmallKana

如果 WPS 在搜索过程中忽略双元音和双辅音的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzySmallKana**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在搜索过程中忽略双元音和双辅音的差异。*/
Options.MatchFuzzySmallKana = true
```

### Options.MatchFuzzySpace

如果 WPS 在搜索过程中忽略空格标记的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzySpace**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在搜索过程中忽略使用的空格标记之间的差异。*/
Options.MatchFuzzySpace = true
```

### Options.MatchFuzzyTC

如果在搜索过程中 WPS 将忽略“ ”、“ ”与“ ”之间以及“ ”与“ ”之间的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyTC**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 在搜索过程中忽略“ tsui tsui”、“ tei tei”与“ chi”之间以及“ dei dei”与“ ji”之间的差异。*/
Options.MatchFuzzyTC = true
```

### Options.MatchFuzzyZJ

如果 WPS 将在搜索过程中忽略“ ”与“ ”之间以及“ ”与“ ”之间的差异，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyZJ**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 在搜索过程中忽略“ se”与“ shie shie”之间以及“ ze”与“ jie jie”之间的差异。*/
Options.MatchFuzzyZJ = true
```

### Options.MeasurementUnit

设置或者返回 WPS 的标准度量单位。 **WdMeasurementUnits** 类型，可读写。

**语法**

**express.MeasurementUnit**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 的标准度量单位设置为磅。*/
Options.MeasurementUnit = wdPoints
```

``` JavaScript
/*本示例返回“选项”对话框（“工具”菜单）中“常用”选项卡上当前所选定的度量单位。*/
let CurrUnit = Options.MeasurementUnit
```

### Options.MergedCellColor

返回或设置一个 **WdCellColor** 常量，该常量代表合并的表格单元格的颜色。可读写。

**语法**

**express.MergedCellColor**

\*express是一个代表 Options 对象的变量。

**说明**

此属性仅适用于对其运行了 **CompareDocuments** 、 **Compare** 或 **Merge** 方法的文档。

### Options.MonthNames

返回或设置朝鲜文字和朝鲜文汉字之间的转换方向。 **WdMonthNames** 类型，可读写。

**语法**

**express.MonthNames**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 在默认情况下将朝鲜文字转换成朝鲜文汉字。*/
Options.MultipleWordConversionsMode = wdHangulToHanja
```

### Options.MoveFromTextColor

返回或设置一个 **WdColorIndex** 常量，该常量代表所移动文本的颜色。可读写。

**语法**

**express.MoveFromTextColor**

\*express是一个代表 Options 对象的变量。

### Options.MoveFromTextMark

返回或设置一个 **WdMoveFromTextMark** 常量，该常量代表要用于所移动文本的修订标记的类型。可读写。

**语法**

**express.MoveFromTextMark**

\*express是一个代表 Options 对象的变量。

### Options.MoveToTextColor

返回或设置一个 **WdColorIndex** 常量，该常量代表所移动文本的颜色。可读写。

**语法**

**express.MoveToTextColor**

\*express是一个代表 Options 对象的变量。

### Options.MoveToTextMark

返回或设置一个 **WdMoveToTextMark** 常量，该常量代表要用于所移动文本的修订标记的类型。可读写。

**语法**

**express.MoveToTextMark**

\*express是一个代表 Options 对象的变量。

### Options.MultipleWordConversionsMode

返回或设置朝鲜文字和朝鲜文汉字之间的转换方向。 **WdMultipleWordConversionsMode** 类型，可读写。

**语法**

**express.MultipleWordConversionsMode**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 在默认情况下将朝鲜文字转换成朝鲜文汉字。*/
Options.MultipleWordConversionsMode = wdHangulToHanja
```

### Options.OMathAutoBuildUp

返回或设置一个 **Boolean** 类型的值，该值代表 WPS 是否自动将公式转换为专业格式。如果为 **True** ，则表示 WPS 自动将公式转换为专业格式。可读写。

**语法**

**express.OMathAutoBuildUp**

\*express是一个代表 Options 对象的变量。

### Options.OMathCopyLF

返回或设置一个 **Boolean** 类型的值，该值代表如何以纯文本的形式表示公式。如果为 **True** ，则表示以线性格式表示公式。如果为 **False** ，则表示以 MathML 格式表示公式。可读写。

**语法**

**express.OMathCopyLF**

\*express是一个代表 Options 对象的变量。

### Options.OptimizeForWord97byDefault

如果 WPS 将禁用所有不兼容的格式，以便在 WPS 97 中查看新文档而对其进行优化，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.OptimizeForWord97byDefault**

\*express是一个代表 Options 对象的变量。

**说明**

若要为 WPS 97 优化单个文档，可用 **OptimizeForWord97** 属性。

**示例**

``` JavaScript
/*本示例实现的功能是：先设定 WPS 禁用新文档中所有与 WPS 97 不兼容的格式，然后新建一个文档，其 OptimizeForWord97 属性自动设置为 True。*/
function test()
{
Options.OptimizeForWord97byDefault = true
MsgBox(Documents.Add(undefined, undefined, wdNewBlankDocument).OptimizeForWord97)
}
```

### Options.Overtype

如果激活改写模式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Overtype**

\*express是一个代表 Options 对象的变量。

**说明**

在改写模式中，键入的字符逐一替代已有的字符。如果改写模式没有被激活，则键入的字符将已有文本右移。

**示例**

``` JavaScript
/*如果“改写”模式已激活，则本示例显示一个信息框，询问是否将“改写”模式改为非激活状态。如果用户单击“是”按钮，则“改写”模式变为非激活状态。*/
function test()
{
if ( Options.Overtype == true ) {
    let aButton = MsgBox("Overtype is on. Turn off?", 4)
    if ( aButton == jsResultYes ) {
        Options.Overtype = false
    }
}
}
```

### Options.Pagination

如果 WPS 在后台重新为文档分页，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Pagination**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在后台重新分页。*/
Options.Pagination = true
```

``` JavaScript
/*本示例返回“工具”菜单的“选项”对话框中“常规”选项卡上“后台重新分页”选项的当前状态。*/
let temp = Options.Pagination
```

### Options.Parent

返回一个 **Object** 类型值，该值代表指定 **Options** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Options 对象的变量。

### Options.PasteAdjustParagraphSpacing

如果在剪切和粘贴选定内容时，WPS 自动调整段落间距，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PasteAdjustParagraphSpacing**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*如果禁用该选项，本示例设置 WPS 在剪切和粘贴选定内容时自动调整段落间距。*/
function AdjustParaSpace() {
    if ( Options.PasteAdjustParagraphSpacing == false ) {
        Options.PasteAdjustParagraphSpacing = true
    }
}
```

### Options.PasteAdjustTableFormatting

如果在剪切和粘贴选定内容时，WPS 自动调整表格的格式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PasteAdjustTableFormatting**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*如果禁用该选项，本示例设置 WPS 在剪切和粘贴时自动设置表格的格式。*/
function AdjustTableFormatting() {
    if ( Options.PasteAdjustTableFormatting == false ) {
        Options.PasteAdjustTableFormatting = true
    }
}
```

### Options.PasteAdjustWordSpacing

如果在剪切和粘贴选定内容时，WPS 自动调整字符间距，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PasteAdjustWordSpacing**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*如果禁用该选项，本示例设置 WPS 在剪切和粘贴选定文本时自动调整字符间距。*/
function AdjustWordSpace() {
    if ( Options.PasteAdjustWordSpacing == false ) {
        Options.PasteAdjustWordSpacing = true
    }
}
```

### Options.PasteFormatBetweenDocuments

返回或设置一个 WdPasteOptions 常量，该常量代表在从另一个 WPS Office WPS 文档复制文本时粘贴文本的方式。可读写。

**语法**

**express.PasteFormatBetweenDocuments**

\*express是一个代表 Options 对象的变量。

**说明**

相当于 **“ WPS 选项”** 对话框的 **“高级”** 选项卡中的 **“在两个文档之间粘贴(不带样式)”** 选项。

### Options.PasteFormatBetweenStyledDocuments

返回或设置一个 **WdPasteOptions** 常量，该常量代表在从使用样式的文档中复制文本时粘贴文本的方式。可读写。

**语法**

**express.PasteFormatBetweenStyledDocuments**

\*express是一个代表 Options 对象的变量。

**说明**

相当于 **“ WPS 选项”** 对话框的 **“高级”** 选项卡中的 **“在两个文档之间粘贴(带样式)”** 选项。

### Options.PasteFormatFromExternalSource

返回或设置一个 WdPasteOptions 常量，该常量代表在从外部源（如网页）复制文本时粘贴文本的方式。可读写。

**语法**

**express.PasteFormatFromExternalSource**

\*express是一个代表 Options 对象的变量。

**说明**

相当于 **“ WPS 选项”** 对话框的 **“高级”** 选项卡中的 **“从其他程序粘贴”** 选项。

### Options.PasteFormatWithinDocument

返回或设置一个 WdPasteOptions 常量，该常量代表在同一文档内复制或剪切文本然后进行粘贴时的文本粘贴方式。可读写。

**语法**

**express.PasteFormatWithinDocument**

\*express是一个代表 Options 对象的变量。

**说明**

相当于 **“ WPS 选项”** 对话框的 **“高级”** 选项卡中的 **“在同一文档内粘贴”** 选项。

### Options.PasteMergeFromPPT

如果为 **True** ，则从 Microsoft PowerPoint 粘贴文本时合并文本格式。 **Boolean** 类型，可读写。

**语法**

**express.PasteMergeFromPPT**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*如果禁用该选项，本示例设置 WPS 从 PowerPoint 粘贴内容时自动合并文本格式。*/
function AdjustPPTFormatting() {
    if ( Options.PasteMergeFromPPT == false ) {
        Options.PasteMergeFromPPT = true
    }
}
```

### Options.PasteMergeFromXL

如果为 **True** ，则从 ET 粘贴表格时合并表格的格式。 **Boolean** 类型，可读写。

**语法**

**express.PasteMergeFromXL**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*如果禁用该选项，本示例设置 WPS 在粘贴 ET 表格时自动合并表格的格式。*/
function AdjustExcelFormatting() {
    if ( Options.PasteMergeFromXL == false ) {
        Options.PasteMergeFromXL = true
    }
}
```

### Options.PasteMergeLists

如果为 **True** ，则合并粘贴列表和周围列表的格式。 **Boolean** 类型，可读写。

**语法**

**express.PasteMergeLists**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*如果禁用该选项，则本示例设置 WPS 自动合并粘贴列表和周围列表的格式。*/
function UseSmartStyle() {
    if ( Options.PasteMergeLists == false ) {
        Options.PasteMergeLists = true
    }
}
```

### Options.PasteOptionKeepBulletsAndNumbers

返回或设置 **Boolean** 值，该值表示在从 **“粘贴选项”** 上下文菜单中选择 **“仅保留文本”** 时是否保留项目符号和编号。可读/写。

**语法**

**express.PasteOptionKeepBulletsAndNumbers**

\*express是一个代表 Options 对象的变量。

**说明**

此属性对应于 **“ WPS 选项”** 对话框的 **“高级”** 选项卡中的 **“使用‘仅保留文本’选项粘贴文本时保留项目符号和编号”** 复选框。

### Options.PasteSmartCutPaste

如果 WPS 在文档中智能粘贴选定内容，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PasteSmartCutPaste**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*如果禁用该选项，本示例设置 WPS 智能粘贴选定文本。*/
function EnableSmartCutPaste() {
    if ( Options.PasteSmartCutPaste == false ) {
        Options.PasteSmartCutPaste = true
    }
}
```

### Options.PasteSmartStyleBehavior

从其他文档粘贴选定文本时，如果 WPS 智能合并样式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PasteSmartStyleBehavior**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*如果禁用该选项，本示例会将 WPS 设置为从其他文档的选定文本中智能粘贴样式。*/
function UseSmartStyle() {
    if ( Options.PasteSmartStyleBehavior == false ) {
        Options.PasteSmartStyleBehavior = true
    }
}
```

### Options.PictureEditor

返回或设置用于编辑图片的应用程序的名称。 **String** 类型，可读写。

**语法**

**express.PictureEditor**

\*express是一个代表 Options 对象的变量。

**说明**

必须使用 **“工具”** 菜单 **“选项”** 对话框中 **“编辑”** 选项卡上“图片编辑器”框中显示的正确单词。否则，将使用默认设置“WPS ”。

如果所需的图形应用程序的名称没有显示在列表中，请和图形应用程序的制造商联系咨询。

**示例**

``` JavaScript
/*本示例设置用于编辑图片的应用程序。*/
Options.PictureEditor = "WPS "
```

``` JavaScript
/*本示例返回用于编辑图片的应用程序的名称。*/
MsgBox(Options.PictureEditor)
```

### Options.PictureWrapType

设置或返回一个 **WdWrapTypeMerged** 类型的值，该值表示 WPS 在图片周围环绕文字的方式。可读写。

**语法**

**express.PictureWrapType**

\*express是一个代表 Options 对象的变量。

**说明**

这是默认的选项设置并对所有插入的图片有效，除非为图片单独定义图片环绕。

**示例**

``` JavaScript
/*如果尚未指定嵌入格式，本示例将 WPS 设置为以嵌入方式在文本中插入并粘贴所有图片。*/
function PicWrap() {
    let  options = Application.Options
    if ( options.PictureWrapType != wdWrapMergeInline ) {
        options.PictureWrapType = wdWrapMergeInline
    }
}
```

### Options.PortugalReform

返回或设置欧洲葡萄牙语拼写器的模式。可读/写 WdPortugueseReform 。

**语法**

**express.PortugalReform**

\*express是一个代表 Options 对象的变量。

**说明**

设置该属性与在 **“ WPS 选项”** 对话框中 **“欧洲葡萄牙语模式:”** 旁边的下拉框中选择项目具有相同作用（ **“校对”** 、 **“在 WPS Office程序中更正拼写时”** ）。

| 注释                                                                                                     |
|----------------------------------------------------------------------------------------------------------|
| 该属性不设置巴西葡萄牙语拼写器的模式。若要设置巴西葡萄牙语拼写器模式，请使用 Options.BrazilReform 属性。 |

### Options.PrecisePositioning

返回或设置一个 **Boolean** 类型的值，该值表示 WPS 是否针对打印版式而不是针对屏幕可读性优化字符位置。如果该属性的值为 **True** ，则禁用压缩字符间距以便增强屏幕可读性的默认设置，并启用适用于打印介质的字符间距。可读/写。

**语法**

**express.PrecisePositioning**

\*express是一个代表 Options 对象的变量。

### Options.PrintBackground

如果 WPS 在后台打印，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintBackground**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在后台打印文档，然后打印活动文档。*/
function test()
{
Options.PrintBackground = true
ActiveDocument.PrintOut()
}
```

``` JavaScript
/*本示例返回“工具”菜单的“选项”对话框中“打印”选项卡上“后台打印”选项的当前状态。*/
let temp = Options.PrintBackground
```

### Options.PrintBackgrounds

返回一个 **Boolean** 类型的值，该值代表打印文档时，是否打印背景颜色和图像。

**语法**

**express.PrintBackgrounds**

\*express是一个代表 Options 对象的变量。

**说明**

如果为 **True** ，代表打印背景颜色和图像。如果为 **False** ，则代表不打印背景颜色和图像。

**示例**

``` JavaScript
/*下列示例指定打印文档的同时要打印背景颜色和图像。*/
Options.PrintBackgrounds = true
```

### Options.PrintComments

如果 WPS 在文档后另起一页打印批注，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintComments**

\*express是一个代表 Options 对象的变量。

**说明**

如果将 **PrintComments** 属性设置为 **True** ，则 **PrintHiddenText** 属性将自动设置为 **True** 。但是，如果将 **PrintComments** 属性设置为 **False** ，不会影响到 **PrintHiddenText** 属性的设置。

**示例**

``` JavaScript
/*本示例将 WPS 设置为打印批注，然后打印活动文档。*/
function test()
{
Options.PrintComments = true
ActiveDocument.PrintOut()
}
```

### Options.PrintDraft

如果 WPS 以最简单的格式打印文档，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintDraft**

\*express是一个代表 Options 对象的变量。

**说明**

不是所有的打印机都支持草稿输出。

**示例**

``` JavaScript
/*本示例将 WPS 设置为使用草稿输出，然后打印活动文档。*/
function test()
{
Options.PrintDraft = true
ActiveDocument.PrintOut()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“草稿输出”选项的当前状态。*/
let temp = Options.PrintDraft
```

### Options.PrintDrawingObjects

如果 WPS 打印图形对象，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintDrawingObjects**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为打印图形对象，然后打印活动文档。*/
function test()
{
Options.PrintDrawingObjects = true
ActiveDocument.PrintOut()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“图形对象”选项的当前状态。*/
let temp = Options.PrintDrawingObjects
```

### Options.PrintEvenPagesInAscendingOrder

如果 WPS 在手动双面打印模式下按升序打印偶数页，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintEvenPagesInAscendingOrder**

\*express是一个代表 Options 对象的变量。

**说明**

如果将 **PrintOut** 方法的 *ManualDuplexPrint* 参数设为 **False** ，该属性将被忽略。

**示例**

``` JavaScript
/*本示例先设定 WPS 在手动双面打印模式下按升序打印奇数页，按降序打印偶数页，然后以此方式打印活动文档。*/
function test()
{
Options.PrintOddPagesInAscendingOrder = true
Options.PrintEvenPagesInAscendingOrder = false
ActiveDocument.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,null,null,true)
}
```

### Options.PrintFieldCodes

如果 WPS 打印域代码而不打印域的结果，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintFieldCodes**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为打印域代码，然后打印活动文档。*/
function test()
{
Options.PrintFieldCodes = true
ActiveDocument.PrintOut()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“域代码”选项的当前状态。*/
let temp = Options.PrintFieldCodes
```

### Options.PrintHiddenText

如果打印隐藏文字，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintHiddenText**

\*express是一个代表 Options 对象的变量。

**说明**

如果将 **PrintHiddenText** 属性设置为 **False** ，则 **PrintComments** 属性将自动设置为 **False** 。但是，如果将 **PrintHiddenText** 属性设置为 **True** ，不会影响到 **PrintComments** 属性的设置。

**示例**

``` JavaScript
/*本示例将 WPS 设置为打印隐藏文字，然后打印活动文档。*/
Application.Options.PrintHiddenText = true
Application.ActiveDocument.PrintOut()

/*本示例返回“选项”对话框中“打印”选项卡上“隐藏文字”选项的当前状态。*/
let temp = Application.Options.PrintHiddenText
```

### Options.PrintProperties

如果 WPS 在文档结尾处另起一页打印文档摘要信息，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintProperties**

\*express是一个代表 Options 对象的变量。

**说明**

如果不打印文档摘要信息，则为 **False** 。可在 **“文件”** 菜单的 **“属性”** 对话框中找到摘要信息。

**示例**

``` JavaScript
/*本示例设置 WPS 在文档结尾另起一页打印文档的摘要信息，然后打印活动文档。*/
function test()
{
Options.PrintProperties = true
ActiveDocument.PrintOut()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“文档属性”选项的当前状态。*/
let temp = Options.PrintProperties
```

### Options.PrintReverse

如果 WPS 逆页序打印页面，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintReverse**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为逆页序打印页面，然后打印活动文档。*/
function test()
{
Options.PrintReverse = true
ActiveDocument.PrintOut()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“逆页序打印”选项的当前状态。*/
let temp = Options.PrintReverse
```

### Options.PrintXMLTag

返回一个 **Boolean** 类型的值，该值代表打印文档时是否打印 XML 标记。对应于 **“选项”** 对话框中 **“打印”** 选项卡上的 **“XML 标记”** 复选框。

**语法**

**express.PrintXMLTag**

\*express是一个代表 Options 对象的变量。

**说明**

如果为 **True** ，则代表打印标记。如果为 **False** ，则代表不打印标记。

**示例**

``` JavaScript
/*下列示例指定打印文档的同时要打印标记。*/
Options.PrintXMLTag = true
```

### Options.PromptUpdateStyle

如果为 **True** ，则在改变样式的格式时显示一条消息，请用户确认是否要重新设置样式的格式或重新应用原样式格式。 **Boolean** 类型，可读写。

**语法**

**express.PromptUpdateStyle**

\*express是一个代表 Options 对象的变量。

**说明**

如果为 **False** ，则对选定内容重新应用样式的格式，而无需确认用户是否要改变样式。

**示例**

``` JavaScript
/*本示例检查在更新样式时用户是否会收到消息，如果没有，则启用该功能。*/
function UpdateStylePrompt() {
    let opt = Application.Options
    if(opt.PromptUpdateStyle == false) {
       opt.PromptUpdateStyle = true
    }
}
```

### Options.RepeatWPS

返回或设置 **Boolean** 值，该值表示在检查拼写时是否标记重复的单词。如果为 **True** ，将标记重复的单词。可读/写。

**语法**

**express.RepeatWPS**

\*express是一个代表 Options 对象的变量。

**说明**

此属性对应于 **“ WPS 选项”** 对话框的 **“校对”** 选项卡上的 **“标记重复单词”** 复选框。

### Options.ReplaceSelection

如果键入或粘贴的内容替换选定内容，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ReplaceSelection**

\*express是一个代表 Options 对象的变量。

**说明**

如果在选定内容之前添加键入或粘贴的内容，并保留选定内容，则为 **False** 。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在选定内容之前添加键入或粘贴的内容，并保留选定内容。*/
Application.Options.ReplaceSelection = false

/*本示例返回“工具”菜单“选项”对话框中“编辑”选项卡上“键入内容替换选定内容”选项的状态。*/
let temp = Application.Options.ReplaceSelection
```

### Options.RevisedLinesColor

当激活修订选项时，返回或设置文档中修订线的颜色。 **WdColorIndex** 类型，可读/写。

**语法**

**express.RevisedLinesColor**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将修订线的颜色设置为粉红色。*/
Options.RevisedLinesColor = wdPink
```

``` JavaScript
/*本示例返回“选项”对话框（“工具”菜单）中“修订”选项卡上的“修订行”下的“颜色”选项的当前状态。*/
let temp = Options.RevisedLinesColor
```

### Options.RevisedLinesMark

当激活修订选项时，返回或设置修订线在文档中的位置。 **WdRevisedLinesMark** 类型，可读/写。

**语法**

**express.RevisedLinesMark**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置修订线显示在每页的左边距上。*/
Options.RevisedLinesMark = wdRevisedLinesMarkLeftBorder
```

``` JavaScript
/*本示例返回“选项”对话框（“工具”菜单）中“修订”选项卡上的“修订行”下的“标记”选项的当前状态。*/
let temp = Options.RevisedLinesMark
```

### Options.RevisedPropertiesColor

在启用修订功能时，返回或设置用于标记格式更改情况的颜色。 **WdColorIndex** 类型，可读写。

**语法**

**express.RevisedPropertiesColor**

\*express是一个代表 Options 对象的变量。

**说明**

如果删除或插入的文本在格式上有所更改，则使用 **DeletedTextColor** 或 **InsertedTextColor** 属性覆盖 **RevisedPropertiesColor** 属性。

**示例**

``` JavaScript
/*本示例跟踪活动文档的修订，将已改变格式的文本的颜色设置为兰绿色，并将加粗格式应用于所选内容。*/
function test()
{
ActiveDocument.TrackRevisions = true
Options.RevisedPropertiesColor = wdTeal
Selection.Font.Bold = true
}
```

``` JavaScript
/*本示例返回“选项”对话框（“工具”菜单）中“修订”选项卡上“修订选项”下“颜色”框中所选中的选项。*/
let temp = Options.RevisedPropertiesColor
```

### Options.RevisedPropertiesMark

在启用修订功能时，返回或设置用于显示格式修改情况的标记。 **WdRevisedPropertiesMark** 类型，可读/写。

**语法**

**express.RevisedPropertiesMark**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例在启用修订功能后，用双下划线标记修改过格式的文本。*/
Options.RevisedPropertiesMark = wdRevisedPropertiesMarkDoubleUnderline
 
```

``` JavaScript
/*本示例返回“选项”对话框（“工具”菜单）中“修订”选项卡上“修订”下“格式”框中所选中的选项。*/
let temp = Options.RevisedPropertiesMark
```

### Options.RevisionsBalloonPrintOrientation

返回或设置一个 **WdRevisionsBalloonPrintOrientation** 常量，该常量表示打印时修订和批注气球的方向。可读写。

**语法**

**express.RevisionsBalloonPrintOrientation**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例以横向格式打印带有批注的文档，修订和批注气球在页面的一侧，文本在另一侧。*/
function PrintLandscapeCommentBalloons() {
    Options.RevisionsBalloonPrintOrientation = wdBalloonPrintOrientationForceLandscape
}
```

### Options.RTFInClipboard

您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 **Options** 对象的 **RTFInClipboard** 属性的信息，请参阅 WPS OfficeMacintosh Edition 中附带的语言参考帮助。

**语法**

**express.RTFInClipboard**

\*express是一个代表 Options 对象的变量。

### Options.SaveInterval

返回或设置保存“自动恢复”信息的时间间隔（以分钟为单位）。 **Long** 类型，可读写。

**语法**

**express.SaveInterval**

\*express是一个代表 Options 对象的变量。

**说明**

将 **SaveInterval** 属性设置为 0（零）时，将关闭保存“自动恢复”信息的功能。

**示例**

``` JavaScript
/*本示例将 WPS 设置为每隔五分钟保存所有打开文档的“自动恢复”信息。*/
Options.SaveInterval = 5
```

``` JavaScript
/*本示例将 WPS 设置为不保存“自动恢复”信息。*/
Options.SaveInterval = 0
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“保存”选项卡上“自动保存时间间隔:”选项的当前状态。*/
let temp = Options.SaveInterval
```

### Options.SaveNormalPrompt

如果 WPS 在关闭之前提示用户确认是否保存对 Normal 模板所做的更改，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SaveNormalPrompt**

\*express是一个代表 Options 对象的变量。

**说明**

如果 WPS 在关闭之前自动保存对 Normal 模板所做的更改，则为 **False** 。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在关闭之前自动保存对 Normal 模板所做的更改，然后退出。*/
function test()
{
Options.SaveNormalPrompt = false
Application.Quit()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“保存”选项卡上“提示保存 Normal 模板”选项的当前状态。*/
let temp = Options.SaveNormalPrompt
```

### Options.SavePropertiesPrompt

如果 WPS 在保存新的文档时，提示该文档的属性信息，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SavePropertiesPrompt**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在保存新文档时，提示该文档的属性信息。*/
Options.SavePropertiesPrompt = true
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“保存”选项卡上“提示保存文档属性”选项的当前状态。*/
let temp = Options.SavePropertiesPrompt
```

### Options.SendMailAttach

如果使用 **“文件”** 菜单上的 **“发送”** 命令将活动文档作为附件插入邮件中，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SendMailAttach**

\*express是一个代表 Options 对象的变量。

**说明**

如果使用 **“发送”** 命令将活动文档的内容作为文本插入邮件中，则为 **False** 。

**示例**

``` JavaScript
/*本示例打开一个新电子邮件，并将活动文档作为邮件附件。*/
function test()
{
Options.SendMailAttach = true
ActiveDocument.SendMail()
}
```

``` JavaScript
/*本示例返回“选项”对话框中“常规”选项卡上“作为附件发送”选项的状态。*/
MsgBox(Options.SendMailAttach)
```

### Options.SequenceCheck

如果为 **True** ，则检查南亚文本独立字符的顺序。 **Boolean** 类型，可读写。

**语法**

**express.SequenceCheck**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例启用顺序检查，使用户可以键入有效的独立字符顺序以组成南亚文本的有效字符单元格。*/
function CheckSequence() {
    Options.SequenceCheck = true
}
```

### Options.ShortMenuNames

您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 **Options** 对象的 **ShortMenuNames** 属性的信息，请参阅 WPS OfficeMacintosh Edition 中附带的语言参考帮助。

**语法**

**express.ShortMenuNames**

\*express是一个代表 Options 对象的变量。

### Options.ShowControlCharacters

如果显示当前文档中的双向控制符，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowControlCharacters**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例在当前文档中隐藏双向控制符。*/
Options.ShowControlCharacters = false
```

### Options.ShowDevTools

返回或设置 **Boolean** 类型的值，该值代表是否在功能区中显示 **“开发人员”** 选项卡。可读/写。

**语法**

**express.ShowDevTools**

\*express是一个代表 Options 对象的变量。

**说明**

此属性对应于 **“ WPS 选项”** 对话框中的 **“在功能区中显示‘开发人员’选项卡”** 复选框。

### Options.ShowDiacritics

如果在使用从右向左语言的文档中显示音调符号，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowDiacritics**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例隐藏当前文档中的音调符号。*/
Options.ShowDiacritics = false
```

### Options.ShowFormatError

如果为 **True** ，WPS 通过在文本（该文本的格式与文档中使用频繁的其他格式相似）下加弯曲下划线来标记不一致的格式。 **Boolean** 类型，可读写。

**语法**

**express.ShowFormatError**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例使 WPS 可以保留文档中格式的修订，但不在文本下显示弯曲的下划线。*/
function ShowFormatErrors() {
    let opt =  Options
    opt.FormatScanning = true  //Enables keeping track of formatting
    opt.ShowFormatError = false
}
```

### Options.ShowMarkupOpenSave

返回或设置一个 **Boolean** 类型的值，该值代表打开或保存文件时 WPS 是否显示隐藏标记。

**语法**

**express.ShowMarkupOpenSave**

\*express是一个代表 Options 对象的变量。

**说明**

**ShowMarkupOpenSave** 属性对应于 **“选项”** 对话框的 **“安全性”** 选项卡上的 **“打开或保存时标记可见”** 选项。

**示例**

``` JavaScript
/*以下示例启用“打开或保存时标记可见”选项。*/
Options.ShowMarkupOpenSave = true
```

### Options.ShowMenuFloaties

返回或设置 **Boolean** 值，该值表示当用户在文档窗口中右键单击时是否显示浮动工具栏。可读/写。

**语法**

**express.ShowMenuFloaties**

\*express是一个代表 Options 对象的变量。

### Options.ShowReadabilityStatistics

如果 WPS 在结束语法检查时，显示统计信息的摘要列表（包括可读性程度），则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowReadabilityStatistics**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为显示可读性统计信息，然后检查活动文档的拼写和语法。*/
function test()
{
Options.ShowReadabilityStatistics = true
ActiveDocument.CheckGrammar()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“拼写和语法”选项卡上“显示可读性统计信息”选项的当前状态。*/
let temp = Options.ShowReadabilityStatistics
```

### Options.ShowSelectionFloaties

返回或设置 **Boolean** 值，该值表示当用户选择文本时是否显示浮动工具栏。可读/写。

**语法**

**express.ShowSelectionFloaties**

\*express是一个代表 Options 对象的变量。

**说明**

对应于 **“ WPS 选项”** 对话框中的 **“选择时显示浮动工具栏”** 复选框。

### Options.SmartCursoring

返回或设置一个 **Boolean** 类型的值，该值表示是否启用智能指针。如果值为 **True** ，则启用智能指针。

**语法**

**express.SmartCursoring**

\*express是一个代表 Options 对象的变量。

**说明**

**SmartCursoring** 属性对应于 **“选项”** 对话框的 **“编辑”** 选项卡中的 **“使用智能指针”** 选项，该选项在默认情况下处于选中状态。

**SmartCursoring** 属性为 **True** 时，使用 Page Down 键在文档中滚动会使指针移到当前页面。如果 **SmartCursoring** 属性为 **False** ，则指针停留在上次编辑的位置。

**示例**

``` JavaScript
/*以下示例禁用智能指针。*/
Options.SmartCursoring = false
```

### Options.SmartCutPaste

如果在剪切和粘贴时，WPS 自动调整字词和标点符号之间的间距，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SmartCutPaste**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在剪切和粘贴时自动调整字词和标点符号之间的间距，然后删除和粘贴新建文档中的部分文本。如果将 SmartCutPaste 属性设置为 False，那么第二个和第三个词将一起处理。*/
function test()
{
Options.SmartCutPaste = true
let myDoc = Documents.Add()
myDoc.Content.InsertAfter("The brown quick fox")
myDoc.Words.Item(2).Cut()
myDoc.Characters.Item(10).Paste()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“编辑”选项卡上“智能剪贴”选项的状态。*/
let temp = Options.SmartCutPaste
```

### Options.SmartParaSelection

如果为 **True** ，则在选择段落中大多数或全部内容时，WPS 会在选定内容中包含段落标记。 **Boolean** 类型，可读写。

**语法**

**express.SmartParaSelection**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例禁用智能段落选择功能。*/
function SetSmartParagraphSelection() {
    Options.SmartParaSelection = false
}
```

### Options.SnapToGrid

如果绘制、移动自选图形或中文字符或者调整其大小时，它们自动与不可见的网格对齐，则该属性值为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.SnapToGrid**

\*express是一个代表 Options 对象的变量。

**说明**

绘制或移动自选图形或者调整其大小时，按下 Alt 可以暂时使该设置无效。

**示例**

``` JavaScript
/*以下示例设置 Word，使新文档的自选图形自动与不可见的网格对齐。*/
function test()
{
Options.SnapToGrid = true
Documents.Add()
}
```

``` JavaScript
/*以下示例返回“对齐网格”对话框中的“对齐网格”选项的状态。*/
let Temp = Options.SnapToGrid
```

### Options.SnapToShapes

如果 WPS 自动将自选图形或中文字符与不可见的网格线对齐，这些网格线穿过其他自选图形或中文字符的垂直和水平边缘，则该属性值为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.SnapToShapes**

\*express是一个代表 Options 对象的变量。

**说明**

此属性为每个自选图形创建附加的不可见网格线。 **SnapToShapes** 属性与 **SnapToGrid** 属性无关。

**示例**

``` JavaScript
/*以下示例设置 Word，使新文档中的自选图形自动与不可见的网格线对齐，这些网格线穿过其他自选图形的垂直和水平边缘。*/
function test()
{
Options.SnapToShapes = true
Documents.Add()
}
```

### Options.SpanishMode

返回或设置西班牙语拼写器的模式。可读/写 WdSpanishSpeller 。

**语法**

**express.SpanishMode**

\*express是一个代表 Options 对象的变量。

**说明**

设置该属性与选择 **“ WPS 选项”** 对话框中 **“西班牙语模式:”** 下的选项具有相同作用（ **“校对”** 、 **“在 WPS Office程序中更正拼写时”** ）。

### Options.SplitCellColor

返回或设置一个 **WdCellColor** 常量，该常量代表拆分的表格单元格的颜色。可读写。

**语法**

**express.SplitCellColor**

\*express是一个代表 Options 对象的变量。

**说明**

此属性仅适用于对其运行了 **CompareDocuments** 、 **Compare** 或 **Merge** 方法的文档。

### Options.StoreRSIDOnSave

如果为 **True** ，则 WPS 在每次保存文档时为文档中的变化指定一个随机数，以便比较与合并文档。 **Boolean** 类型，可读写。

**语法**

**express.StoreRSIDOnSave**

\*express是一个代表 Options 对象的变量。

**说明**

WPS 将这些随机数保存在表格中并在每次保存后更新该表格。 **StoreRSIDOnSave** 属性的默认值为 **True** 。但是不保存有关 HTML 文档的 RSID 信息。

如果要删除文档中有关作者和审阅者的信息，可使用 **RemovePersonalInformation** 属性。

**示例**

``` JavaScript
/*本示例在保存文档时关闭保存随机数功能。*/
function SaveRandomNumber() {
    Application.Options.StoreRSIDOnSave = false
}
```

### Options.StrictFinalYaa

如果拼写检查使用针对以字母“yaa”结尾的阿拉伯语单词的拼写规则，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.StrictFinalYaa**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置拼写检查使用针对以字符 yaa 结尾的阿拉伯语单词的拼写规则。*/
Options.StrictFinalYaa = true
```

### Options.StrictInitialAlefHamza

如果拼写检查使用针对以“alef hamza”开头的阿拉伯语单词的拼写规则，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.StrictInitialAlefHamza**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例设置拼写检查使用针对以“alef hamza”开头的阿拉伯语单词的拼写规则。*/
Options.StrictInitialAlefHamza = true
```

### Options.StrictRussianE

如果拼写检查器使用有关使用严格 ? 字符的俄语单词的拼写规则，则该属性值为 **True** 。可读/写。

**语法**

**express.StrictRussianE**

\*express是一个代表 Options 对象的变量。

**说明**

设置该属性与选中或取消选中 **“ WPS 选项”** 对话框中的 **“俄语: 强制严格 ?”** 复选框具有相同作用（ **“校对”** 、 **“在 WPS Office程序中更正拼写时”** ）。

### Options.StrictTaaMarboota

如果拼写检查器使用拼写规则来标记以 haa（而不是 taa marboota）结尾的阿拉伯语单词，则该属性值为 **True** 。可读/写。

**语法**

**express.StrictTaaMarboota**

\*express是一个代表 Options 对象的变量。

**说明**

设置该属性与选中或清除 **“ WPS 选项”** 对话框中的 **“阿拉伯语: 严格 taa marboota”** 复选框具有相同作用（ **“校对”** 、 **“在 WPS Office程序中更正拼写时”** ）。

### Options.SuggestFromMainDictionaryOnly

如果 WPS 仅根据主词典提供拼写建议，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SuggestFromMainDictionaryOnly**

\*express是一个代表 Options 对象的变量。

**说明**

如果根据主词典和添加的任何自定义词典提供拼写建议，则该属性值返回 **False** 。

**示例**

``` JavaScript
/*本示例将 WPS 设置为仅根据主词典选择建议单词，然后在活动文档中进行拼写检查。*/
function test()
{
Options.SuggestFromMainDictionaryOnly = true
ActiveDocument.CheckSpelling()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“拼写和语法”选项卡上“仅根据主词典提供建议”选项的当前状态。*/
let temp = Options.SuggestFromMainDictionaryOnly
```

### Options.SuggestSpellingCorrections

如果 WPS 在检查拼写时，总是为每一个拼写错误的单词提供可选的拼写建议，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.SuggestSpellingCorrections**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为总是为拼写错误的单词提出可选的拼写建议，然后在活动文档中进行拼写检查。*/
function test()
{
Options.SuggestSpellingCorrections = true
ActiveDocument.CheckSpelling()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“拼写和语法”选项卡上“总提出更正建议”选项的当前状态。 */
let temp = Options.SuggestSpellingCorrections
```

### Options.TabIndentKey

如果可分别使用 Tab 和 Backspace 键增加和减少段落的左缩进，并且可使用 Backspace 键将右对齐的段落更改为居中的段落而将居中的段落更改为左对齐的段落，则该属性值为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.TabIndentKey**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*以下示例设置 WPS 使用 Tab 和 Backspace 键来设置段落的左缩进而不是插入和删除制表位。*/
Options.TabIndentKey = true
```

### Options.TypeNReplace

如果为 **True** ，则 WPS 将替换非法的南亚字符。 **Boolean** 类型，可读写。

**语法**

**express.TypeNReplace**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为替换非法的南亚字符。*/
function TypeReplace() {
    Application.Options.TypeNReplace = true
}
```

### Options.UpdateFieldsAtPrint

如果 WPS 在打印文档前自动更新域，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.UpdateFieldsAtPrint**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在打印文档前自动更新域，然后打印活动文档。*/
function test()
{
Options.UpdateFieldsAtPrint = true
ActiveDocument.PrintOut()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“更新域”选项的当前状态。*/
let temp = Options.UpdateFieldsAtPrint
```

### Options.UpdateFieldsWithTrackedChangesAtPrint

如果 WPS 2015 允许在打印前更新包含修订的字段，则该属性值为 **True** 。可读/写。

**语法**

**express.UpdateFieldsWithTrackedChangesAtPrint**

\*express是一个代表 Options 对象的变量。

**说明**

设置该属性与在 WPS 2015 中选中或取消选中 **“ WPS 选项”** 对话框中 **“允许在打印之前更新包含修订的字段”** （ **“高级”** 选项卡， **“打印”** ）旁边的复选框具有相同作用。

### Options.UpdateLinksAtOpen

如果 WPS 在打开文档时自动更新嵌入的所有 OLE 链接，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.UpdateLinksAtOpen**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在打开文档时自动更新嵌入的 OLE 链接。*/
Options.UpdateLinksAtOpen = true
```

``` JavaScript
/*本示例返回“选项”对话框中“常规”选项卡上“打开时更新自动方式的链接”选项的当前状态。*/
let temp = Options.UpdateLinksAtOpen
```

### Options.UpdateLinksAtPrint

如果 WPS 在打印文档前更新指向其他文件的嵌入链接，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.UpdateLinksAtPrint**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为在打印文档前自动更新嵌入链接，然后打印活动文档。*/
function test()
{
Options.UpdateLinksAtPrint = true
ActiveDocument.PrintOut()
}
```

``` JavaScript
/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“更新链接”选项的当前状态。*/
let temp = Options.UpdateLinksAtPrint
```

### Options.UpdateStyleListBehavior

返回或设置 WdUpdateStyleListBehavior 常量，该常量指定在更新某种样式以与包含编号或项目符号的所选内容相匹配时， WPS 2015 应采取的行为。可读/写。rs"\>版本信息

**语法**

**express.UpdateStyleListBehavior**

\*express是一个代表 Options 对象的变量。

**说明**

设置该属性与选择 WPS 2015 **“选项”** 对话框中下拉列表中的项目具有相同作用（ **“高级”** 选项卡， **“编辑选项”** ， **“更新样式以匹配所选内容:”** ）。

### Options.UseCharacterUnit

如果 WPS 在当前文档中使用字符作为默认度量单位，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.UseCharacterUnit**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为使用字符作为默认度量单位。*/
Options.UseCharacterUnit = true
```

### Options.UseDiffDiacColor

如果可以设置当前文档中音调符号的颜色，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.UseDiffDiacColor**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例在设置当前选定内容中的音调符号的颜色前，检查 UseDiffDiacColor 属性的值。*/
function test()
{
if(Options.UseDiffDiacColor == true) {
    Selection.Font.DiacriticColor = wdColorBlue
}
}
```

### Options.UseGermanSpellingReform

如果 WPS 在检查拼写时，使用德语后期修订拼写规则，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.UseGermanSpellingReform**

\*express是一个代表 Options 对象的变量。

**说明**

由于您选择或安装的语言支持不同（例如美国英语），该属性可能无法使用。

**示例**

``` JavaScript
/*本示例将 WPS 设置为使用德语后期修订规则检查拼写。*/
Options.UseGermanSpellingReform = true
```

### Options.UseNormalStyleForList

返回或设置 **Boolean** 值，该值表示 WPS 是否对项目符号和编号使用“正文”样式。可读写。

**语法**

**express.UseNormalStyleForList**

\*express是一个代表 Options 对象的变量。

**说明**

对应于 **“ WPS 选项”** 对话框的 **“高级”** 选项卡中的 **“对项目符号或编号列表使用正文样式”** 复选框。

### Options.VisualSelection

根据从右向左式语言文档中的可视光标移动，返回或设置选择行为。 **WdVisualSelection** 类型，可读写。

**语法**

**express.VisualSelection**

\*express是一个代表 Options 对象的变量。

**说明**

必须将 **CursorMovement** 属性设置为 **wdCursorMovementVisual** 才能使用此属性。

**示例**

``` JavaScript
/*本示例设置选定行为，以使选定内容换行。*/
function test()
{
if(Options.CursorMovement == wdCursorMovementVisual) {
   Options.VisualSelection = wdVisualSelectionContinuous
}
}
```

### Options.WarnBeforeSavingPrintingSendingMarkup

如果为 **True** ，则 WPS 在保存、打印或以电子邮件形式发送包含批注或修订的文档时，显示警告。 **Boolean** 类型，可读写。

**语法**

**express.WarnBeforeSavingPrintingSendingMarkup**

\*express是一个代表 Options 对象的变量。

**示例**

``` JavaScript
/*本示例打印活动文档，但如果文档包含修订或批注，则允许用户停止打印。*/
function SaferPrint() {
    //Save old state in variable
    let blnOldState = Application.Options.WarnBeforeSavingPrintingSendingMarkup
                                
    //Turn on warning
    Application.Options.WarnBeforeSavingPrintingSendingMarkup = true
    
    //Print document                            
    ActiveDocument.PrintOut()
                                
    //Restore original warning state
    Application.Options.WarnBeforeSavingPrintingSendingMarkup = blnOldState
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Options](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Revisions 对象

## 简介

Revision 对象 [的集合](https://docs.microsoft.com/zh-cn/office/vba/api/word.revision) ，这些对象代表区域或文档中标有修订标记的更改。

## 方法

| 序号 | 名称                              | 说明                                                               |
|:----:|:----------------------------------|:-------------------------------------------------------------------|
|  1   | [AcceptAll](#Revisions.AcceptAll) | 接受文档或区域中的所有修订、删除所有修订标记并将更改合并到文档中。 |
|  2   | [Item](#Revisions.Item)           | 返回集合 中的单个 Revisions 对象。                                 |
|  3   | [RejectAll](#Revisions.RejectAll) | 拒绝接受文档的修订集合                                             |

## 属性

| 序号 | 名称                      | 说明                                                 |
|:----:|:--------------------------|:-----------------------------------------------------|
|  1   | [Count](#Revisions.Count) | 返回一 个 Long ，代表集合中的修订数。 此为只读属性。 |

## 成员方法

### Revisions.AcceptAll

接受文档或区域中的所有修订、删除所有修订标记并将更改合并到文档中。

**语法**

**express.AcceptAll ()**

\*express是一个代表 Revisions 对象的变量。

**示例**

``` JavaScript
//接受选中区域的所有修订的
Application.Selection.Range.Revisions.AcceptAll()
```

### Revisions.Item

返回集合 中的单个 Revisions 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Revisions 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明       |
|---------|-----------|----------|------------|
| *Index* | 可选      | Long     | 对象的索引 |

**示例**

``` JavaScript
//获取选区的第一个修订对象
Application.Selection.Reivesion.Item(1)
```

### Revisions.RejectAll

拒绝接受文档的修订集合

**语法**

**express.RejectAll ()**

\*express是一个代表 Revisions 对象的变量。

**示例**

``` JavaScript
//拒绝接受选区的修订集合
Application.Selection.Range.Revisions.RejectAll()
```

## 成员属性

### Revisions.Count

返回一 个 Long ，代表集合中的修订数。 此为只读属性。

**语法**

**express.Count**

\*express是一个代表 Revisions 对象的变量。

**示例**

``` JavaScript
//获取激活文档的修订个数。
Application.ActiveDocument.Revisions.Count
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Revisions](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Section 对象

## 简介

代表所选内容、范围或文档中的一节

## 说明

代表所选内容、范围或文档中的一节。 Section 对象是 Sections 集合的一个成员。 Sections 集合包括所选内容、范围或文档中的所有节。

## 属性

| 序号 | 名称                                            | 说明                                                                         |
|:----:|:------------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Section.Application)             | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Borders](#Section.Borders)                     | 返回 Borders 集合，该集合代表节中的所有边框。                                |
|  3   | [Creator](#Section.Creator)                     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Footers](#Section.Footers)                     | 返回一个 HeadersFooters 集合，该集合代表指定节的页脚。只读。                 |
|  5   | [Headers](#Section.Headers)                     | 返回一个 HeadersFooters 集合，该集合代表指定节的页眉。只读。                 |
|  6   | [Index](#Section.Index)                         | 返回一个 Long 类型的值，该值代表项目在集合中的位置。只读。                   |
|  7   | [Pagesetup](#Section.Pagesetup)                 | 代表页面设置说明                                                             |
|  8   | [Parent](#Section.Parent)                       | 返回一个 Object 类型值，该值代表指定 Section 对象的父对象。                  |
|  9   | [ProtectedForForms](#Section.ProtectedForForms) | 如果保护指定节中的窗体，则该属性值为 True 。 Boolean 类型，可读写。          |
|  10  | [Range](#Section.Range)                         | 返回一个 Range 对象，该对象代表指定对象中包含的文档部分。                    |

## 成员属性

### Section.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Section 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Section.Borders

返回 **Borders** 集合，该集合代表节中的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Section 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Section.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Section 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Section.Footers

返回一个 **HeadersFooters** 集合，该集合代表指定节的页脚。只读。

**语法**

**express.Footers**

\*express是一个代表 Section 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。要返回代表指定节的页眉的 **HeadersFooters** 集合，请使用 **Headers** 属性。

**示例**

``` JavaScript
/*本示例向活动文档第一节中的主页脚添加一个右对齐的页码。*/
function test()
{
let myFooters = ActiveDocument.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary)
myFooters.PageNumbers.Add(wdAlignPageNumberRight)
}
```

### Section.Headers

返回一个 **HeadersFooters** 集合，该集合代表指定节的页眉。只读。

**语法**

**express.Headers**

\*express是一个代表 Section 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。要返回代表指定节的页脚的 **HeadersFooters** 集合，请使用 **Footers** 属性。

**示例**

``` JavaScript
/*本示例为活动文档中除首页之外的每页添加居中的页码（为首页创建一个独立的页眉）。*/
function test()
{
let myHeaders = ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary)
myHeaders.PageNumbers.Add(wdAlignPageNumberCenter, false)
}
```

``` JavaScript
/*本示例将向活动文档首页页眉添加文本。*/
function test()
{
ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = true
let myHeaders = ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterFirstPage)
myHeaders.Range.InsertAfter("First Page Text")
myHeaders.Range.Paragraphs.Alignment = wdAlignParagraphRight
}
```

### Section.Index

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Section 对象的变量。

**示例**

``` JavaScript
Application.ActiveDocument.Sections.Item(1).Index
```

### Section.Pagesetup

代表页面设置说明

**语法**

**express.Pagesetup**

\*express是一个代表 Section 对象的变量。

**说明**

**代表页面设置说明。PageSetup** 对象包含文档的所有页面设置属性（如左边距、下边距和纸张大小）。

**示例**

``` JavaScript
//以下示例将活动文档的第一节设为横向
Application.ActiveDocument.Sections.Item(1).PageSetup.Orientation =  wdOrientLandscape
```

### Section.Parent

返回一个 **Object** 类型值，该值代表指定 **Section** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Section 对象的变量。

### Section.ProtectedForForms

如果保护指定节中的窗体，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ProtectedForForms**

\*express是一个代表 Section 对象的变量。

**说明**

如果保护节中的窗体，则只能选择和修改窗体域中的文本。要保护整个文档，请使用 **Document** 对象的 **Protect** 方法。

**示例**

``` JavaScript
/*本示例打开活动文档中第二节的窗体保护。*/
function test()
{
if(ActiveDocument.Sections.Count >= 2) {
    ActiveDocument.Sections.Item(2).ProtectedForForms = true
}
}
```

``` JavaScript
/*本示例取消对选定内容中第一节的保护。*/
Selection.Sections.Item(1).ProtectedForForms = false
```

``` JavaScript
/*本示例切换对选定内容中第一节的保护状态。*/
Selection.Sections.Item(1).ProtectedForForms = !(Selection.Sections.Item(1).ProtectedForForms)
```

### Section.Range

返回一个 **Range** 对象，该对象代表指定对象中包含的文档部分。

**语法**

**express.Range**

\*express是一个代表 Section 对象的变量。

**示例**

``` JavaScript
//下示例在活动文档的开头将文本作为一个新段落插入
function test()
{
    var myRange = Application.ActiveDocument.Sections.Item(1).Range
    myRange.MoveEnd(wdCharacter,-1)
    myRange.Collapse(wdCollapseEnd)
    myRange.InsertParagraphAfter()
    myRange.InsertAfter("End of section")
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Section](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# XMLChildNodeSuggestion 对象

## 简介

表示根据架构可能作为当前元素子元素的节点，但不保证其有效性。

## 说明

在用户将元素插入文档之前，WPS 无法验证元素的显示顺序（或出现次数的上限或下限或其他约束）。如果选中 “仅列出当前元素的子元素” 复选框，则每个 XMLChildNodeSuggestion 对象都是 “XML 结构” 任务窗格底部 XML 元素列表中的项目。

使用 XMLSchemaReference 属性可访问节点所引用的单个架构。使用 Item 方法可从 XMLChildNodeSuggestion 对象的集合中返回一个 XMLChildNodeSuggestion 对象。以下代码将所有允许的元素插入活动文档中。

``` JavaScript
function test()
{
let objNode = Selection.XMLParentNode 

for(let i = 1; i <= objNode.ChildNodeSuggestions.Count; i++){
    objNode.ChildNodeSuggestions.Item(i).Insert()
    Selection.MoveRight()
}
}
```

使用 NamespaceURI 属性可返回 XMLChildNodeSuggestion 对象中引用的架构的命名空间。使用 BaseName 属性可返回代表 XMLChildNodeSuggestion 对象的元素的名称。

``` JavaScript
function test()
{
let objSuggestion = ActiveDocument.ChildNodeSuggestions
    
for(let i = 1; i <= objSuggestion.Count; i++){
        
    if(objSuggestion.Item(i).BaseName == "example"){
        objSuggestion.Item(i).Insert()
    }             
}
}
```

## 方法

| 序号 | 名称                                     | 说明                                   |
|:----:|:-----------------------------------------|:---------------------------------------|
|  1   | [Insert](#XMLChildNodeSuggestion.Insert) | 插入代表 XML 元素节点的 XMLNode 对象。 |

## 属性

| 序号 | 名称                                                             | 说明                                                                                      |
|:----:|:-----------------------------------------------------------------|:------------------------------------------------------------------------------------------|
|  1   | [Application](#XMLChildNodeSuggestion.Application)               | 返回一个代表 WPS 应用程序的 Application 对象。                                            |
|  2   | [BaseName](#XMLChildNodeSuggestion.BaseName)                     | 返回一个 String 类型的值，该值代表不带任何前缀的元素名称。                                |
|  3   | [Creator](#XMLChildNodeSuggestion.Creator)                       | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。              |
|  4   | [NamespaceURI](#XMLChildNodeSuggestion.NamespaceURI)             | 返回一个 String 类型的值，该值代表指定对象的架构命名空间的统一资源标识符 (URI)。只读。    |
|  5   | [Parent](#XMLChildNodeSuggestion.Parent)                         | 返回一个 Object 类型值，该值代表指定 XMLChildNodeSuggestion 对象的父对象。                |
|  6   | [XMLSchemaReference](#XMLChildNodeSuggestion.XMLSchemaReference) | 返回一个 XMLSchemaReference 对象，代表指定的 XMLChildNodeSuggestion 对象所属的 XML 架构。 |

## 成员方法

### XMLChildNodeSuggestion.Insert

插入代表 XML 元素节点的 **XMLNode** 对象。

**语法**

**express.Insert ( *Range* )**

\*express是一个代表 XMLChildNodeSuggestion 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                      |
|---------|-----------|----------|-------------------------------------------|
| *Range* | 可选      | Variant  | 环绕其放置打开和关闭 XML 元素的文字范围。 |

## 成员属性

### XMLChildNodeSuggestion.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 XMLChildNodeSuggestion 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### XMLChildNodeSuggestion.BaseName

返回一个 **String** 类型的值，该值代表不带任何前缀的元素名称。

**语法**

**express.BaseName**

\*express是一个代表 XMLChildNodeSuggestion 对象的变量。

**示例**

``` JavaScript
/*下例将 author 属性添加到活动文档中的 book 元素，然后设置属性的值。*/
function AddIDAttribute(){
    let objElement = ActiveDocument.XMLNodes

    for(let i = 1; i <= objElement.Count; i++){
        if(objElement.Item(i).NodeType == wdXMLNodeElement){
        if(objElement.Item(i).BaseName == "book"){
                
                let objAttribute = objElement.Item(i).Attributes.Add("author", objElement.Item(i).NamespaceURI)
                objAttribute.NodeValue = "David Barber"
                
            }
        }
    }
}
```

### XMLChildNodeSuggestion.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 XMLChildNodeSuggestion 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### XMLChildNodeSuggestion.NamespaceURI

返回一个 **String** 类型的值，该值代表指定对象的架构命名空间的统一资源标识符 (URI)。只读。

**语法**

**express.NamespaceURI**

\*express是一个代表 XMLChildNodeSuggestion 对象的变量。

**说明**

| 注释                                                                            |
|---------------------------------------------------------------------------------|
| 如果创建的是用于 WPS 的 XML 架构，则强烈建议在架构中指定 targetNamespace 设置。 |

### XMLChildNodeSuggestion.Parent

返回一个 **Object** 类型值，该值代表指定 **XMLChildNodeSuggestion** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 XMLChildNodeSuggestion 对象的变量。

**说明**

------------------------------------------------------------------------

### XMLChildNodeSuggestion.XMLSchemaReference

返回一个 **XMLSchemaReference** 对象，代表指定的 **XMLChildNodeSuggestion** 对象所属的 XML 架构。

**语法**

**express.XMLSchemaReference**

\*express是一个代表 XMLChildNodeSuggestion 对象的变量。

**示例**

``` JavaScript
/*下列示例为活动文档中的第一个 XMLChildNodeSuggestion 对象重新加载架构。*/
function test()
{
let objSuggestion = ActiveDocument.XMLChildNodeSuggestions.Item(1)
objSuggestion.XMLSchemaReference.Reload()
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/XMLChildNodeSuggestion](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# XMLMapping 对象

## 简介

表示 ContentControl 对象上自定义 XML 与内容控件之间的 XML 映射。XML 映射是内容控件中的文本与此文档的自定义 XML 数据存储中的 XML 元素之间的链接。

## 说明

使用 SetMapping 方法可以添加或更改使用 XPath 字符串的内容控件的 XML 映射。以下示例设置文档作者这一内置文档属性，在活动文档中插入新的内容控件，然后将此控件的 XML 映射设置为该内置文档属性。

``` JavaScript
function test() {
    Application.ActiveDocument.BuiltInDocumentProperties.Item("Author").Value = "David";
    let objcc = Application.ActiveDocument.ContentControls.Add(wdContentControlDate, Application.ActiveDocument.Paragraphs.Item(1).Range);
    let objMap = objcc.XMLMapping;
    let bInMap = objMap.SetMapping("/ns1:coreProperties[1]/ns0:createdate[1]");
} 
```

使用 SetMappingByNode 方法可以添加或更改使用 CustomXMLNode 对象的内容控件的 XML 映射。以下示例具有与上一个示例相同的功能，但使用 SetMappingByNode 方法。

``` JavaScript
function test() {
    Application.ActiveDocument.BuiltInDocumentProperties.Item("Author").Value = "David Jaffe"
    let rng = Application.ActiveDocument.Paragraphs.Item(1).Range
    let objcc = Application.ActiveDocument.ContentControls.Add(wdContentControlDate, rng)

    let npURL = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    let objNP =  Application.ActiveDocument.CustomXMLParts.SelectByNamespace(npURL).Item(1)
    let objNode = objNP.DocumentElement.ChildNodes.Item(1)

    let objMap = objcc.XMLMapping
    let blnMap = objMap.SetMappingByNode(objNode)
} 
```

以下示例创建新的 CustomXMLPart 对象，将自定义 XML 加载到该对象中，然后创建两个新的内容控件，并将它们分别映射到自定义 XML 内的不同 XML 元素。

``` JavaScript
function test() {
    objCustomPart = Application.ActiveDocument.CustomXMLParts.Add();
    objCustomPart.LoadXML ("<books><book><author>Matt Hink</author>" +
            "<title>Migration Paths of the Red Breasted Robin</title>" +
            "<genre>non-fiction</genre><price>29.95</price>" +
            "<pub_date>2/1/2007</pub_date><abstract>You see them in " +
            "the spring outside your windows.  You hear their lovely " +
            "songs wafting in the warm spring air.  Now follow the path " +
            "of the red breasted robin as it migrates to warmer climes " +
            "in the fall, and then back to your back yard in the spring." +
            "</abstract></book></books>")

    Application.ActiveDocument.Range().InsertParagraphBefore()
    objRange = Application.ActiveDocument.Paragraphs.Item(1).Range
    objCustomNode = objCustomPart.SelectSingleNode("/books/book/title")
    objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange)
    objCustomControl.XMLMapping.SetMappingByNode(objCustomNode)
    
    objRange.InsertParagraphAfter()
    objRange = Application.ActiveDocument.Paragraphs.Item(2).Range
    objCustomNode = objCustomPart.SelectSingleNode("/books/book/abstract")
    objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange)
    objCustomControl.XMLMapping.SetMappingByNode(objCustomNode)
    
    alert(objCustomControl.XMLMapping.IsMapped)
}
```

使用 Delete 方法可以删除内容控件的 XML 映射。删除内容控件的 XML 映射时，只会删除内容控件与 XML 数据之间的连接。内容控件和 XML 数据都会留在文档中。以下示例删除活动文档内当前映射的所有内容控件的 XML 映射。

``` JavaScript
Application.ActiveDocument.ContentControls.Item(1).XMLMapping.Delete()
```

使用 IsMapped 属性可以确定内容控件是否映射到文档的数据存储中的 XML 节点。以下示例删除活动文档中所有映射的内容控件的 XML 映射。

``` JavaScript
Application.ActiveDocument.ContentControls.Item(1).XMLMapping.IsMapped
```

使用 Custom XMLNode 属性可以访问内容控件映射到的 XML 节点。使用 CustomXMLPart 属性可以访问内容控件映射到的 XML 部件。有关使用 CustomXMLNode 和 CustomXMLPart 对象的详细信息，请参阅各自的对象主题。

## 方法

| 序号 | 名称                                             | 说明                                                                                                                           |
|:----:|:-------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Delete](#XMLMapping.Delete)                     | 从父内容控件中删除 XML 映射。                                                                                                  |
|  2   | [SetMapping](#XMLMapping.SetMapping)             | 允许创建或更改内容控件上的 XML 映射。如果 WPS 将内容控件映射到文档的自定义 XML 数据存储中的自定义 XML 节点，将返回 True 。     |
|  3   | [SetMappingByNode](#XMLMapping.SetMappingByNode) | 允许创建或更改内容控件上的 XML 数据映射。如果 WPS 将内容控件映射到文档的自定义 XML 数据存储中的自定义 XML 节点，将返回 True 。 |

## 属性

| 序号 | 名称                                         | 说明                                                                                           |
|:----:|:---------------------------------------------|:-----------------------------------------------------------------------------------------------|
|  1   | [Application](#XMLMapping.Application)       | 返回一个 Application 对象，该对象代表 WPS 应用程序。                                           |
|  2   | [Creator](#XMLMapping.Creator)               | 返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 Long 类型。                   |
|  3   | [CustomXMLNode](#XMLMapping.CustomXMLNode)   | 返回 CustomXMLNode 对象，该对象表示文档中的内容控件映射到的自定义 XML 节点（位于数据存储中）。 |
|  4   | [CustomXMLPart](#XMLMapping.CustomXMLPart)   | 返回 CustomXMLPart 对象，该对象表示文档中的内容控件映射到的自定义 XML 部件。                   |
|  5   | [IsMapped](#XMLMapping.IsMapped)             | 返回 Boolean 值，该值表示文档中的内容控件是否映射到文档的 XML 数据存储内的 XML 节点。只读。    |
|  6   | [Parent](#XMLMapping.Parent)                 | 返回一个 Object 类型的值，该值代表指定 XMLMapping 对象的父对象。                               |
|  7   | [PrefixMappings](#XMLMapping.PrefixMappings) | 返回 String 值，该值表示用于为当前 XML 映射计算 XPath 的前缀映射。只读。                       |
|  8   | [XPath](#XMLMapping.XPath)                   | 返回 String 值，该值表示 XML 映射的 XPath ，其计算结果是当前映射的 XML 节点。只读。            |

## 成员方法

### XMLMapping.Delete

从父内容控件中删除 XML 映射。

**语法**

**express.Delete ()**

\*express是一个代表 XMLMapping 对象的变量。

**说明**

此操作删除 XML 映射。XML 数据和内容控件仍保留在文档中。

### XMLMapping.SetMapping

允许创建或更改内容控件上的 XML 映射。如果 WPS 将内容控件映射到文档的自定义 XML 数据存储中的自定义 XML 节点，将返回 **True** 。

**语法**

**express.SetMapping ( *XPath* , *PrefixMapping* , *Source* )**

\*express是一个代表 XMLMapping 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型      | 说明                                                                                                                                                                        |
|-----------------|-----------|---------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *XPath*         | 必选      | String        | 指定 XPath 字符串，该字符串表示内容控件所映射到的 XML 节点。无效的 XPath 字符串将导致运行时错误。                                                                           |
| *PrefixMapping* | 可选      | String        | 指定在查询 XPath 参数所提供的表达式时要使用的前缀映射。如果省略， WPS 将使用当前文档中指定的自定义 XML 部件的前缀映射集。                                                   |
| *Source*        | 可选      | CustomXMLPart | 指定内容控件要映射到的自定义 XML 数据。如果省略此参数，将依据当前文档中的所有自定义 XML 来计算 XPath，并且用 XPath 在其中解析为 XML 节点的第一个 CustomXMLPart 来建立映射。 |

**说明**

如果 XML 映射已经存在， WPS 将替换现有 XML 映射，并以新映射的 XML 节点的内容替换内容控件的文本。如果指定的 XPath 的计算结果不是指定的自定义 XML 部件中的 XML 节点，那么仍可以指定映射，并且将创建一个映射。当指定的 XPath 的计算结果为指定的自定义 XML 部件中的 XML 节点时，此映射将自动链接。

请参阅 SetMappingByNode 方法。

| 注释                                         |
|----------------------------------------------|
| 为格式文本内容控件创建映射将导致运行时错误。 |

**示例**

``` JavaScript
/* 以下示例插入自定义 XML 部件并为该自定义部件设置 XML，然后在文档开头插入两个内容控件，并将控件的内容映射为自定义部件中 XML 元素的内容*/
function test() {
    let objRange
    let objCustomPart
    let objCustomControl

    objCustomPart = Application.ActiveDocument.CustomXMLParts.Add()
    objCustomPart.LoadXML ("<books><book><author>Matt Hink</author>" +
        "<title>Migration Paths of the Red Breasted Robin</title>" +
        "<genre>non-fiction</genre><price>29.95</price>" +
        "<pub_date>2/1/2007</pub_date><abstract>You see them in " +
        "the spring outside your windows.  You hear their lovely " +
        "songs wafting in the warm spring air.  Now follow the path "+
        "of the red breasted robin as it migrates to warmer climes " +
        "in the fall, and then back to your back yard in the spring." +
        "</abstract></book></books>")

    Application.ActiveDocument.Range().InsertParagraphBefore()
    objRange = Application.ActiveDocument.Paragraphs.Item(1).Range
    objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange)
    objCustomControl.XMLMapping.SetMapping("/books/book/title", null, objCustomPart)

    objRange.InsertParagraphAfter()
    objRange = Application.ActiveDocument.Paragraphs.Item(2).Range
    objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange)
    objCustomControl.XMLMapping.SetMapping("/books/book/abstract", null, objCustomPart)
}
```

### XMLMapping.SetMappingByNode

允许创建或更改内容控件上的 XML 数据映射。如果 WPS 将内容控件映射到文档的自定义 XML 数据存储中的自定义 XML 节点，将返回 **True** 。

**语法**

**express.SetMappingByNode ( *Node* )**

\*express是一个代表 XMLMapping 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型      | 说明                                  |
|--------|-----------|---------------|---------------------------------------|
| *Node* | 必选      | CustomXMLNode | 指定当前内容控件所映射到的 XML 节点。 |

**说明**

如果 XML 映射已经存在，则 WPS 将替换现有 XML 映射，并以新映射的 XML 节点的内容替换内容控件的文本。请参阅 **SetMapping** 方法。

| 注释                                         |
|----------------------------------------------|
| 为格式文本内容控件创建映射将导致运行时错误。 |

**示例**

``` JavaScript
/* 以下示例设置文档作者这一内置文档属性，在活动文档中插入新的内容控件，然后将此控件的 XML 映射设置为该内置文档属性*/
function test() {
    let objcc
    let objNode
    let objMap
    let blnMap
    let objNP

    Application.ActiveDocument.BuiltInDocumentProperties.Item("Author").Value = "David Jaffe"

    objcc = Application.ActiveDocument.ContentControls.Add(wdContentControlDate, ActiveDocument.Paragraphs.Item(1).Range)
    objNP = Application.ActiveDocument.CustomXMLParts.SelectByNamespace("http://schemas.openxmlformats.org/package/2006/metadata/core-properties").Item(1)
    objNode = objNP.DocumentElement.ChildNodes.Item(1)

    objMap = objcc.XMLMapping
    blnMap = objMap.SetMappingByNode(objNode)

    /* 以下示例创建自定义 XML 部件，然后创建两个内容控件，并将每个内容控件映射到自定义 XML 中的特定节点*/
    let objRange
    let objCustomPart
    let objCustomControl
    let objCustomNode

    objCustomPart = Application.ActiveDocument.CustomXMLParts.Add()
    objCustomPart.LoadXML ("<books><book><author>Matt Hink</author>" +
        "<title>Migration Paths of the Red Breasted Robin</title><genre>non-fiction</genre>" +
        "<price>29.95</price><pub_date>2007-02-01</pub_date><abstract>" +
        "You see them in the spring outside your windows.  You hear their lovely " +
        "songs wafting in the warm spring air.  Now follow the path of the red breasted robin " +
        "as it migrates to warmer climes in the fall, and then back to your back yard " +
        "in the spring.</abstract></book></books>")

    Application.ActiveDocument.Range().InsertParagraphBefore()
    objRange = Application.ActiveDocument.Paragraphs.Item(1).Range
    objCustomNode = objCustomPart.SelectSingleNode("/books/book/title")
    objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange)
    objCustomControl.XMLMapping.SetMappingByNode(objCustomNode)

    objRange.InsertParagraphAfter()
    objRange = Application.ActiveDocument.Paragraphs.Item(2).Range
    objCustomNode = objCustomPart.SelectSingleNode("/books/book/abstract")
    objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange)
    objCustomControl.XMLMapping.SetMappingByNode(objCustomNode)
}
```

## 成员属性

### XMLMapping.Application

返回一个 Application 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 XMLMapping 对象的变量。

### XMLMapping.Creator

返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 XMLMapping 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

| 注释                                    |
|-----------------------------------------|
| 该值也可用常量 **wdCreatorCode** 表示。 |

### XMLMapping.CustomXMLNode

返回 **CustomXMLNode** 对象，该对象表示文档中的内容控件映射到的自定义 XML 节点（位于数据存储中）。

**语法**

**express.CustomXMLNode**

\*express是一个代表 XMLMapping 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动文档中插入新的内容控件和自定义 XML 部件，将内容控件映射到自定义 XML 部件中的某个节点，然后设置映射的 XML 节点的值。*/
function test() {
    let objCC
    let objPart
    let objNode
    let objMap

    objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlText)
    objPart = Application.ActiveDocument.CustomXMLParts.Add("<books><book>" +
        "<author></author><title></title><genre></genre><price></price>" +
        "<pub_date></pub_date><abstract></abstract></book></books>")

    objMap = objCC.XMLMapping
    objMap.SetMapping("/books/book/author", null, objPart)

    objNode = objMap.CustomXMLNode
    objNode.Text = "Matt Hink"

    objCC.Range.Text = objNode.Text
}
```

### XMLMapping.CustomXMLPart

返回 **CustomXMLPart** 对象，该对象表示文档中的内容控件映射到的自定义 XML 部件。

**语法**

**express.CustomXMLPart**

\*express是一个代表 XMLMapping 对象的变量。

**示例**

``` JavaScript
/* 以下示例访问活动文档中的第一个内容控件以及它映射到的自定义 XML 部件，然后设置自定义 XML 部件中包含的某个 XML 节点的文本值*/
/* 此示例假定活动文档中至少有一个内容控件映射到包含 XPath 字符串指定的 XML 节点的自定义 XML 部件*/
function test() {
    let objCC
    let objPart
    let objNode

    objCC = Application.ActiveDocument.ContentControls.Item(1)
    objPart = objCC.XMLMapping.CustomXMLPart
    objNode = objPart.SelectSingleNode("/books/book/title")
    objNode.Text = "Mystery of the Empty Chair"
}
```

### XMLMapping.IsMapped

返回 **Boolean** 值，该值表示文档中的内容控件是否映射到文档的 XML 数据存储内的 XML 节点。只读。

**语法**

**express.IsMapped**

\*express是一个代表 XMLMapping 对象的变量。

**示例**

``` JavaScript
/* 以下示例删除活动文档内当前映射的所有内容控件的 XML 映射*/
function test() {
    let objCC = Application.ActiveDocument.ContentControls

    for(let i = 1; i <= objCC.Count; i++){
        if(objCC.Item(i).XMLMapping.IsMapped == true){
            objCC.XMLMapping.Delete()
        }
    }
}
```

### XMLMapping.Parent

返回一个 **Object** 类型的值，该值代表指定 **XMLMapping** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 XMLMapping 对象的变量。

### XMLMapping.PrefixMappings

返回 **String** 值，该值表示用于为当前 XML 映射计算 XPath 的前缀映射。只读。

**语法**

**express.PrefixMappings**

\*express是一个代表 XMLMapping 对象的变量。

**说明**

若要为内容控件设置映射，请使用 **SetMapping** 方法或 **SetMappingByNode** 方法。

**示例**

``` JavaScript
请使用 SetMapping 方法或 SetMappingByNode 方法。
```

### XMLMapping.XPath

返回 **String** 值，该值表示 **XML** 映射的 **XPath** ，其计算结果是当前映射的 **XML** 节点。只读。

**语法**

**express.XPath**

\*express是一个代表 XMLMapping 对象的变量。

**说明**

若要为内容控件设置映射，请使用 SetMapping 方法或 SetMappingByNode 方法。如果映射不是活动的，则使用此属性将返回错误。

**示例**

``` JavaScript
/* 以下示例检查活动文档中第一个内容控件是否是日期控件，以及 XPath 字符串是否设置为特定的内置文档属性。如果 XPath 不匹配并且控件是日期控件，它将设置到该控件的映射*/
function test() {
    let blnMap
    let objCC = Application.ActiveDocument.ContentControls.Item(1)
    let objMap = objCC.XMLMapping
    if((objCC.Type == wdContentControlDate) && (objMap.XPath != "/ns1:coreProperties[1]/ns0:createdate[1]")) {
        blnMap = objMap.SetMapping("/ns1:coreProperties[1]/ns0:createdate[1]")
        if(blnMap == false) {
            alert("Unable to map the content control.")
        }
    }
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/XMLMapping](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# XMLNodes 对象

## 简介

XMLNode 对象的集合，该集合代表 “XML 结构” 任务窗格的树视图中的所有节点，这些节点表示用户已应用于文档的元素。树视图中的每个节点都是 XMLNode 对象的一个实例。树视图的层次结构表明节点是否包含子节点。

## 方法

| 序号 | 名称                   | 说明                                            |
|:----:|:-----------------------|:------------------------------------------------|
|  1   | [Add](#XMLNodes.Add)   | 返回一个 XMLNode 对象，该对象代表新添加的元素。 |
|  2   | [Item](#XMLNodes.Item) | 返回集合中的单个 XMLNode 对象。                 |

## 属性

| 序号 | 名称                                 | 说明                                                                         |
|:----:|:-------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#XMLNodes.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#XMLNodes.Count)             | 返回一个 Number 类型的值，该值代表集合中 XML 元素的数量。只读。              |
|  3   | [Creator](#XMLNodes.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#XMLNodes.Parent)           | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |

## 成员方法

### XMLNodes.Add

返回一个 **XMLNode** 对象，该对象代表新添加的元素。

**语法**

**express.Add ( *Name* , *Namespace* , *Range* )**

\*express是一个代表 XMLNodes 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                            |
|-------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Name*      | 必选      | String   | 在 Namespace 参数指定的 XML 架构中的元素的名称。由于 XML 区分大小写，因此，Name 参数所指定的元素的拼写必须和它在架构中所显示的完全一样。如果与 Namespace 参数中指定的架构中的任何元素名称都不匹配，则显示错误。 |
| *Namespace* | 必选      | String   | 架构中定义的架构名称。Namespace 参数区分大小写，因此必须严格按照它在架构中所显示的那样进行拼写。如果在附加到文档的任何架构中都找不到指定的命名空间，则显示错误。                                                |
| *Range*     | 可选      | Range    | 要应用元素的区域。默认为将元素标记置于插入点或选定内容的周围（如果已选择文本）。                                                                                                                                |

**示例**

``` JavaScript
/* 以下示例将选定文本添加到示例元素，该示例元素来自活动文档中所引用的第一个架构*/
function test(){
    Application.ActiveDocument.XMLNodes.Add ("example",ActiveDocument.XMLSchemaReferences.Item(1).NamespaceURI,Selection.Range)
}
```

### XMLNodes.Item

返回集合中的单个 **XMLNode** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 XMLNodes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                           |
|---------|-----------|----------|----------------------------------------------------------------|
| *Index* | 必选      | Number   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

## 成员属性

### XMLNodes.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 XMLNodes 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### XMLNodes.Count

返回一个 **Number** 类型的值，该值代表集合中 XML 元素的数量。只读。

**语法**

**express.Count**

\*express是一个代表 XMLNodes 对象的变量。

### XMLNodes.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 XMLNodes 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### XMLNodes.Parent

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Parent**

\*express是一个代表 XMLNodes 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/XMLNodes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Zoom 对象

## 简介

包含窗口或窗格的显示比例选项（例如，显示比例）。 Zoom 对象是 Zooms 集合的一个成员。

## 说明

使用 View 对象的 Zoom 属性可返回一个 Zoom 对象。以下示例将活动窗口的显示比例设置为 110%。

``` JavaScript
ActiveDocument.ActiveWindow.View.Zoom.Percentage = 110
```

使用 Zooms ( *Index* ) 可返回一个 Zoom 对象，其中 *Index* 标识视图类型。由 index 指定的视图类型可以是下列 WdViewType 常量之一： wdMasterView 、 wdNormalView 、 wdOutlineView 、 wdPrintPreview 、 wdPrintView 或 wdWebView 。以下示例设置活动窗口的缩放比例，以便看到整个页面。

``` JavaScript
ActiveDocument.ActiveWindow.ActivePane.Zooms.Item(wdPrintView).PageFit = wdPageFitFullPage
```

Add 方法不可用于 Zooms 集合。对于每个不同的视图类型（如大纲视图、普通视图或页面视图）， Zooms 集合都只包含一个 Zoom 对象。

## 属性

| 序号 | 名称                             | 说明                                                                                |
|:----:|:---------------------------------|:------------------------------------------------------------------------------------|
|  1   | [Application](#Zoom.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                      |
|  2   | [Creator](#Zoom.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。        |
|  3   | [PageColumns](#Zoom.PageColumns) | 返回或设置需要在页面视图或打印预览中在屏幕上并排显示的页数。 Long 类型，可读写。    |
|  4   | [PageFit](#Zoom.PageFit)         | 返回或设置窗口的显示比例，以便能查看整页或整个页宽的范围。 WdPageFit 类型，可读写。 |
|  5   | [PageRows](#Zoom.PageRows)       | 返回或设置在页面视图或打印预览中在屏幕上垂直显示的页数。 Long 类型，可读写。        |
|  6   | [Parent](#Zoom.Parent)           | 返回一个 Object 类型值，该值代表指定 Zoom 对象的父对象。                            |
|  7   | [Percentage](#Zoom.Percentage)   | 用百分比的形式返回或设置窗口的显示比例。 Long 类型，可读写。                        |

## 成员属性

### Zoom.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 Zoom 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Zoom.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Zoom 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Zoom.PageColumns

返回或设置需要在页面视图或打印预览中在屏幕上并排显示的页数。 Long 类型，可读写。

**语法**

**express.PageColumns**

\*express是一个代表 Zoom 对象的变量。

**示例**

``` JavaScript
/*本示例将活动窗口切换到页面视图，并显示并列竖排的两页。*/
function test() {
    let view = Application.ActiveDocument.ActiveWindow.View
    view.Type = wdPrintView
    view.Zoom.PageColumns = 2
    view.Zoom.PageRows = 1
}

/*本示例将 Hello.doc 的文档窗口切换到页面视图，并显示整页。*/
function test() {
    let view = Application.Windows.Item("Hello.doc").View
    view.Type = wdPrintView
    let z = view.Zoom
    z.PageColumns = 1
    z.PageRows = 1
    z.PageFit = wdPageFitFullPage
} 
```

### Zoom.PageFit

返回或设置窗口的显示比例，以便能查看整页或整个页宽的范围。 **WdPageFit** 类型，可读写。

**语法**

**express.PageFit**

\*express是一个代表 Zoom 对象的变量。

**说明**

如果文档不在页面视图内，则 **wdPageFitFullPage** 常量无效。

如果将 **PageFit** 属性设置为 **wdPageFitBestFit** ，则在每次改变文档窗口的大小时，都将自动重新计算显示比例。如果该属性设置为 **wdPageFitNone** ，则在改变文档窗口的大小时，不会重新计算显示比例。

**示例**

``` JavaScript
/*本示例更改 Letter.doc 的窗口的显示比例，以便可查看整个文本宽度的范围。*/
let view = Application.Windows.Item("Letter.doc").View
view.Type = wdNormalView
view.Zoom.PageFit = wdPageFitBestFit

/*本示例将活动窗口切换到页面视图并更改显示比例，以便可查看整页。*/
let view = Application.ActiveDocument.ActiveWindow.View
view.Type = wdPrintView
view.Zoom.PageFit = wdPageFitFullPage
```

### Zoom.PageRows

返回或设置在页面视图或打印预览中在屏幕上垂直显示的页数。 **Long** 类型，可读写。

**语法**

**express.PageRows**

\*express是一个代表 Zoom 对象的变量。

**示例**

``` JavaScript
/*本示例将活动窗口切换到打印预览并垂直显示两页。*/
PrintPreview = true
let zoom = Application.ActiveDocument.ActiveWindow.View.Zoom
zoom.PageColumns = 1
zoom.PageRows = 2
```

### Zoom.Parent

返回一个 **Object** 类型值，该值代表指定 **Zoom** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Zoom 对象的变量。

### Zoom.Percentage

用百分比的形式返回或设置窗口的显示比例。 **Long** 类型，可读写。

**语法**

**express.Percentage**

\*express是一个代表 Zoom 对象的变量。

**示例**

``` JavaScript
/*本示例将活动窗口切换到普通视图并将显示比例设置为 80%。*/
let view = Application.ActiveDocument.ActiveWindow.View
view.Type = wdNormalView
view.Zoom.Percentage = 80

/*本示例将活动窗口的显示比例增加 10%。*/
let myZoom = Application.ActiveDocument.ActiveWindow.View.Zoom
myZoom.Percentage = myZoom.Percentage + 10
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Zoom](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PrintRanges 对象

## 简介

指定的演示文稿中所有 PrintRange 对象的集合。每个 PrintRange 对象代表一个要打印的连续幻灯片或页的范围。

## 方法

| 序号 | 名称                              | 说明                                                                                              |
|:----:|:----------------------------------|:--------------------------------------------------------------------------------------------------|
|  1   | [Add](#PrintRanges.Add)           | 返回一个代表新动画行为的 PrintRange 对象。                                                        |
|  2   | [ClearAll](#PrintRanges.ClearAll) | 从 PrintRanges 集合中清除所有打印范围。使用 PrintRanges 集合的 Add 方法可向该集合中添加打印范围。 |
|  3   | [Item](#PrintRanges.Item)         | 从指定集合中返回单个对象。                                                                        |

## 属性

| 序号 | 名称                                    | 说明                                                  |
|:----:|:----------------------------------------|:------------------------------------------------------|
|  1   | [Application](#PrintRanges.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者 |
|  2   | [Count](#PrintRanges.Count)             | 返回指定集合中的对象数目。                            |
|  3   | [Parent](#PrintRanges.Parent)           | 返回指定对象的父对象。                                |

## 成员方法

### PrintRanges.Add

返回一个代表新动画行为的 **PrintRange** 对象。

**语法**

**express.Add ( *Start* , *End* )**

\*express是一个代表 PrintRanges 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                           |
|---------|-----------|----------|--------------------------------|
| *Start* | 必选      | Long     | 范围中起始幻灯片的幻灯片编号。 |
| *End*   | 必选      | Long     | 范围中终止幻灯片的幻灯片编号。 |

**返回值**

PrintRange

### PrintRanges.ClearAll

从 **PrintRanges** 集合中清除所有打印范围。使用 **PrintRanges** 集合的 **Add** 方法可向该集合中添加打印范围。

**语法**

**express.ClearAll ()**

\*express是一个代表 PrintRanges 对象的变量。

**返回值**

Nothing

**示例**

``` JavaScript
/*本示例清除当前演示文稿中所有以前定义的打印范围；创建新的打印范围，其中包含第一张，第三张到第五张，第八张和第九张幻灯片；打印新定义的幻灯片范围；然后清除新定义的打印范围。*/
function test(){
　　　　let print = ActivePresentation.PrintOptions
    　　　　print.RangeType = ppPrintSlideRange
   　　　　 let ranges = print.Ranges
    　　　　    ranges.ClearAll()
   　　　　     ranges.Add(1, 1)
     　　　　   ranges.Add(3, 5)
      　　　　  ranges.Add(8, 9)
      　　　　  ranges.Parent.Parent.PrintOut()
      　　　　  ranges.ClearAll()
}
```

### PrintRanges.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 PrintRanges 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                             |
|---------|-----------|----------|----------------------------------|
| *Index* | 必选      | Long     | 集合中要返回的单个对象的索引号。 |

**返回值**

PrintRange

## 成员属性

### PrintRanges.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者

**语法**

**express.Application**

\*express是一个代表 PrintRanges 对象的变量。

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
function test(){
　　　　let slides = ActivePresentation.Slides.Item(1).Shapes
       for(let i = 1; i <= slides.Count; i++){
    　　　　if(slides.Item(i).Type == msoLinkedOLEObject){
     　　　　   MsgBox(slides.Item(i).OLEFormat.Application.Name)
　　　　    }
　　　　}
}
```

### PrintRanges.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 PrintRanges 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口。*/
function test(){
　　　　let windows = Application.Windows
　　　　for(let i = 2; i <= windows.Count; i++){
　　　　    window.Item(i).Close()
　　　　}
}
```

### PrintRanges.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 PrintRanges 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes
　　　　let addshape = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    　　　　addshape.TextRange.Text = "Test text"
    　　　　addshape.Parent.Rotation = 45
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/PrintRanges](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomXMLNode 对象

## 简介

代表文档中树的一个 XML 节点。 CustomXMLNode 对象是 CustomXMLNodes 集合的一个成员。

## 说明

CustomXMLNode 对象设计为具有与 IXMLDOMNode 接口等效的功能。此外，它包含 XPath 属性，这是对 MSXML 提供的对象的重大改进。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

以下示例通过使用 XPath 表达式从 CustomXMLPart 对象中选择单个节点，并将它分配给 CustomXMLNode 对象。

``` JavaScript
function test() {
let ment = Application.ActiveDocument
// Returns the first custom xml part with the given root namespace.
let cxp1 = ment.CustomXMLParts.Item("urn:invoice:namespace")     
// Get the first node matching the XPath expression.                             
let cxn = cxp1.SelectSingleNode("//*[@quantity < 4]")
}
```

## 方法

| 序号 | 名称                                                      | 说明                                                                                                                                                               |
|:----:|:----------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AppendChildNode](#CustomXMLNode.AppendChildNode)         | 将单个节点作为树中上下文元素节点下的最后一个子节点追加。                                                                                                           |
|  2   | [AppendChildSubtree](#CustomXMLNode.AppendChildSubtree)   | 将子树作为树中上下文元素节点下的最后一个子节点添加。                                                                                                               |
|  3   | [Delete](#CustomXMLNode.Delete)                           | 从树中删除当前节点（包括其所有子项，如果有）。                                                                                                                     |
|  4   | [HasChildNodes](#CustomXMLNode.HasChildNodes)             | 如果当前元素节点包含子元素节点，则获取 True 。                                                                                                                     |
|  5   | [InsertNodeBefore](#CustomXMLNode.InsertNodeBefore)       | 紧邻树中上下文节点之前插入一个新节点。                                                                                                                             |
|  6   | [InsertSubtreeBefore](#CustomXMLNode.InsertSubtreeBefore) | 将指定的子树插入到紧邻上下文节点之前的位置。                                                                                                                       |
|  7   | [RemoveChild](#CustomXMLNode.RemoveChild)                 | 从树中删除指定的子节点。                                                                                                                                           |
|  8   | [ReplaceChildNode](#CustomXMLNode.ReplaceChildNode)       | 从主树中删除指定的子节点（及其子树），并将它替换为相同位置中的其他节点。                                                                                           |
|  9   | [ReplaceChildSubtree](#CustomXMLNode.ReplaceChildSubtree) | 从主树中删除指定节点（及其子树），并将它替换为相同位置中的其他子树。                                                                                               |
|  10  | [SelectNodes](#CustomXMLNode.SelectNodes)                 | 选择与 XPath 表达式匹配的节点的集合。此方法与 CustomXMLPart . SelectNodes 方法的不同之处在于，它将从作为上下文节点的“表达式”节点开始计算 XPath 表达式。            |
|  11  | [SelectSingleNode](#CustomXMLNode.SelectSingleNode)       | 从与 XPath 表达式匹配的集合中选择一个节点。此方法与 CustomXMLPart . SelectSingleNode 方法的不同之处在于，它将从作为上下文节点的“表达式”节点开始计算 XPath 表达式。 |

## 属性

| 序号 | 名称                                              | 说明                                                                                                                                                        |
|:----:|:--------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLNode.Application)         | 获取一个代表 CustomXMLNode 的容器应用程序的 Application 对象。只读。                                                                                        |
|  2   | [Attributes](#CustomXMLNode.Attributes)           | 获取一个代表当前节点中的当前元素的属性的 CustomXMLNodes 集合。只读。                                                                                        |
|  3   | [BaseName](#CustomXMLNode.BaseName)               | 获取文档对象模型 (DOM) 中节点的不带命名空间前缀的基本名称（如果存在）。只读。                                                                               |
|  4   | [ChildNodes](#CustomXMLNode.ChildNodes)           | 获取一个包含当前节点中所有子元素的 CustomXMLNodes 集合。只读。                                                                                              |
|  5   | [Creator](#CustomXMLNode.Creator)                 | 获取一个 32 位整数，指示创建 CustomXMLNode 对象时所使用的应用程序。只读。                                                                                   |
|  6   | [FirstChild](#CustomXMLNode.FirstChild)           | 获取与当前节点中第一个子元素对应的 CustomXMLNode 对象。如果该节点没有子元素（或者其类型不为 msoCustomXMLNodeElement ），则返回 Nothing 。只读。             |
|  7   | [LastChild](#CustomXMLNode.LastChild)             | 获取与当前节点中的最后一个子元素对应的 CustomXMLNode 对象。如果该节点没有子元素（或者其类型不为 msoCustomXMLNodeElement ），则该属性将返回 Nothing 。只读。 |
|  8   | [NamespaceURI](#CustomXMLNode.NamespaceURI)       | 获取 CustomXMLNode 对象的命名空间的唯一地址标识符。只读。                                                                                                   |
|  9   | [NextSibling](#CustomXMLNode.NextSibling)         | 获取当前节点的下一个同级节点（元素、注释或处理指令）。如果该节点是其所在级别的最后一个同级节点，则该属性将返回 Nothing 。只读。                             |
|  10  | [NodeType](#CustomXMLNode.NodeType)               | 获取当前节点的类型。只读。                                                                                                                                  |
|  11  | [NodeValue](#CustomXMLNode.NodeValue)             | 获取或设置当前节点的值。可读/写。                                                                                                                           |
|  12  | [OwnerDocument](#CustomXMLNode.OwnerDocument)     | 获取代表与此节点关联的 ET 工作簿、Microsoft PowerPoint 演示文稿或 WPS 文档的对象。只读。                                                                    |
|  13  | [OwnerPart](#CustomXMLNode.OwnerPart)             | 获取代表与此节点关联的部件的对象。只读。                                                                                                                    |
|  14  | [Parent](#CustomXMLNode.Parent)                   | 获取 CustomXMLNode 对象的 Parent 对象。只读。                                                                                                               |
|  15  | [ParentNode](#CustomXMLNode.ParentNode)           | 获取当前节点的父元素节点。如果当前节点位于根级，则该属性将返回 Nothing 。只读。                                                                             |
|  16  | [PreviousSibling](#CustomXMLNode.PreviousSibling) | 获取当前节点的前一个同级节点（元素、注释或处理指令）。如果当前节点为其所在级别的第一个节点，则该属性将返回 Nothing 。只读。                                 |
|  17  | [Text](#CustomXMLNode.Text)                       | 获取或设置当前节点的文本。可读/写。                                                                                                                         |
|  18  | [XML](#CustomXMLNode.XML)                         | 获取当前节点及其子项（如果存在）的 XML 表示形式。只读。                                                                                                     |
|  19  | [XPath](#CustomXMLNode.XPath)                     | 获取一个 String 类型的值，其中包含当前节点的规范化 XPath。如果该节点已不在文档对象模型 (DOM) 中，则该属性将返回一条错误消息。只读。                         |

## 成员方法

### CustomXMLNode.AppendChildNode

将单个节点作为树中上下文元素节点下的最后一个子节点追加。

**语法**

**express.AppendChildNode ( *Name* , *NamespaceURI* , *NodeType* , *NodeValue* )**

\*express是一个代表 CustomXMLNode 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型             | 说明                                                                                                                                       |
|----------------|-----------|----------------------|--------------------------------------------------------------------------------------------------------------------------------------------|
| *Name*         | 可选      | String               | 代表要追加的元素的基本名称。                                                                                                               |
| *NamespaceURI* | 可选      | String               | 代表要追加的元素的命名空间。如果要追加 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 类型的节点，则此参数是必需的，否则它将被忽略。 |
| *NodeType*     | 可选      | MsoCustomXMLNodeType | 指定要追加的节点的类型。如果未指定该参数，则假定节点类型为 msoCustomXMLNodeElement。                                                       |
| *NodeValue*    | 可选      | String               | 用于为允许文本的那些节点设置所追加节点的值。如果节点不允许文本，则忽略该参数。                                                             |

**说明**

如果上下文节点是除 **msoXMLNodeElement** 之外的任何类型，或者如果操作将导致树结构无效，则不执行追加操作，并显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例说明了如何将 CustomXMLNode 对象追加到其他节点。 
function test()
{                  
    let cxn = Application.ActivePresentation.CustomXMLParts.Item(1).SelectSingleNode("/ns1:coreProperties[1]") 
    // Append a child node to the single node selected previously.
    cxn.AppendChildNode("discount", "urn:invoice:namespace", Application.Enum.msoCustomXMLNodeElement, "0.10") 
}
```

### CustomXMLNode.AppendChildSubtree

将子树作为树中上下文元素节点下的最后一个子节点添加。

**语法**

**express.AppendChildSubtree ( *XML* )**

\*express是一个代表 CustomXMLNode 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明               |
|-------|-----------|----------|--------------------|
| *XML* | 必选      | String   | 代表要添加的子树。 |

**说明**

如果上下文节点是除 **msoXMLNodeElement** 之外的任何类型，则不执行追加操作，并显示一条错误消息。如果根据架构验证 CustomXMLNode，但该操作将导致树结构无效，则不执行追加操作，并显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例说明了如何将节点追加到现有节点。
function test()
{
    let tion = Application.ActiveDocument
    // Add and populate a custom xml part
    let cxp1 = tion.CustomXMLParts.Add("<invoice />")        
    // Get nodes using XPath.                             
    let cxn = cxp1.SelectSingleNode("/ns1:coreProperties[1]")      // Append a child subtree to the single node selected previously.
    cxn.AppendChildSubtree("0.10")  
}
```

### CustomXMLNode.Delete

从树中删除当前节点（包括其所有子项，如果有）。

**语法**

**express.Delete ()**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

如果操作将导致树结构无效，则不会执行删除操作并显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例演示了使用各种方法添加自定义 XML 部件、使用不同的标准选择部件和节点、附加子树，以及删除部件和节点的过程。
function test()
{
    try 
    {
        let tion = Application.ActiveDocument
        // Example written for Word.
        // Adding a custom XML part.
        tion.CustomXMLParts.Add("")        
        // Add and then load from a file.
        let cxp1 = tion.CustomXMLParts.Add()
        // Returns the first custom XML part with the given root namespace.
        let cxp2 = tion.CustomXMLParts.Item(1)            
        // Access all with a given root namespace; returns the entire collection.
        let cxps = tion.CustomXMLParts.SelectByNamespace("urn:invoice:namespace")      
        // DOM-type operations.
        // Get the XML.
        strXml = cxp2.XML
        // Get the root namespace
        strUri = cxp2.NamespaceURI 
        // Get nodes using XPath.                             
        let cxn = cxp2.SelectSingleNode("//*[@quantity < 4]") 
        let cxns = cxp2.SelectNodes("//*[@unitPrice > 20]")
        // Append a child subtree to the single node selected previously.
        cxn.AppendChildSubtree("<rebates><rebate>0.10</rebate></rebates>")                
        // Delete custom XML part and node and its children.
        cxp2.Delete()
        cxn.Delete()
    }
    // Exception handling. Show the message and resume.
   catch(exception) 
   {
       alert(exception.Description)
   }
 
}
```

### CustomXMLNode.HasChildNodes

如果当前元素节点包含子元素节点，则获取 **True** 。

**语法**

**express.HasChildNodes ()**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

当 **CustomXMLNode** 的节点类型不是 **msoCustomXMLNodeElement** 时，此方法将始终返回 **False** 。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例演示了使用各种方法添加自定义 XML 部件、使用不同的标准选择部件和节点、附加子树，以及删除部件和节点的过程。
function test()
{
    try 
    {
        let tion = Application.ActiveDocument
        // Example written for Word.
        // Adding a custom XML part.
        tion.CustomXMLParts.Add("")        
        // Add and then load from a file.
        let cxp1 = tion.CustomXMLParts.Add()   
        // Returns the first custom XML part with the given root namespace.
        let cxp2 = tion.CustomXMLParts.Item(1)            
        // Access all with a given root namespace; returns the entire collection.
        let cxps = tion.CustomXMLParts.SelectByNamespace("urn:invoice:namespace")      
        // DOM-type operations.
        // Get the XML.
        strXml = cxp2.XML
        // Get the root namespace
        strUri = cxp2.NamespaceURI 
        // Get nodes using XPath.                             
        let cxn = cxp2.SelectSingleNode("//*[@quantity < 4]") 
        let cxns = cxp2.SelectNodes("//*[@unitPrice > 20]")
        // Append a child subtree to the single node selected previously.
        cxn.AppendChildSubtree("<rebates><rebate>0.10</rebate></rebates>")                
        // Delete custom XML part and node and its children.
        cxp2.Delete()
        cxn.Delete()
    }
    // Exception handling. Show the message and resume.
   catch(exception) 
   {
       alert(exception.Description)
   }
 
}
```

### CustomXMLNode.InsertNodeBefore

紧邻树中上下文节点之前插入一个新节点。

**语法**

**express.InsertNodeBefore ( *Name* , *NamespaceURI* , *NodeType* , *NodeValue* , *NextSibling* )**

\*express是一个代表 CustomXMLNode 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型             | 说明                                                                                                                                       |
|----------------|-----------|----------------------|--------------------------------------------------------------------------------------------------------------------------------------------|
| *Name*         | 可选      | String               | 代表要添加的节点的基本名称。                                                                                                               |
| *NamespaceURI* | 可选      | String               | 代表要添加的元素的命名空间。如果要添加 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 类型的节点，则此参数是必需的，否则它将被忽略。 |
| *NodeType*     | 可选      | MsoCustomXMLNodeType | 指定要添加的节点的类型。如果未指定参数，则假定它是 msoCustomXMLNodeElement 类型的节点。                                                    |
| *NodeValue*    | 可选      | String               | 用于为允许文本的那些节点设置要添加节点的值。如果节点不允许文本，则忽略该参数。                                                             |
| *NextSibling*  | 可选      | CustomXMLNode        | 代表上下文节点。                                                                                                                           |

**说明**

如果在添加 **msoCustomXMLNodeElement** 、 **msoCustomXMLNodeComment** 或 **msoCustomXMLNodeProcessingInstruction** 类型的节点时上下文节点不存在，则将新节点添加到上下文节点的最后一个子节点。如果操作将导致树结构无效，则不执行插入操作，并显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例添加一个自定义部件，然后通过使用 XPath 表达式查找该部件中的一个节点。代码然后在找到的节点之前插入一个节点。
function test()
{
    let tion = Application.ActiveDocument
    // Add a custom xml part.
    tion.CustomXMLParts.Add("<invoice>")        
    // Returns the first custom xml part with the given root namespace.
    let cxp1 = tion.CustomXMLParts.Item("urn:invoice:namespace")             
    // Get node using XPath.                             
    let cxn = cxp1.SelectSingleNode("//*[@supplier = "Contoso"]") 
    // Insert a node before the single node selected previously.
    cxn.InsertNodeBefore("discount", "urn:invoice:namespace")
```

### CustomXMLNode.InsertSubtreeBefore

将指定的子树插入到紧邻上下文节点之前的位置。

**语法**

**express.InsertSubtreeBefore ( *XML* , *NextSibling* )**

\*express是一个代表 CustomXMLNode 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型      | 说明               |
|---------------|-----------|---------------|--------------------|
| *XML*         | 必选      | String        | 代表要添加的子树。 |
| *NextSibling* | 可选      | CustomXMLNode | 指定上下文节点。   |

**说明**

如果 *NextSibling* 参数不是上下文节点的子节点，或者操作将导致树结构无效，则不执行插入操作，并显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例添加一个自定义部件，然后通过使用 XPath 表达式查找该部件中的一个节点。代码然后在找到的节点之后插入一个节点。
function test()
{
    let tion = Application.ActiveDocument
    // Add a custom xml part.
    tion.CustomXMLParts.Add("<invoice>"）       
    // Returns the first custom xml part with the given root namespace.
    let cxp1 = tion.CustomXMLParts.Item("urn:invoice:namespace")             
    // Get nodes using XPath.                             
    Set cxn = cxp1.SelectSingleNode("//*[@supplier = "Contoso"]") 
    // Insert a node before the single node selected previously.
    cxn.InsertNodeAfter("discount", "urn:invoice:namespace")
}
```

### CustomXMLNode.RemoveChild

从树中删除指定的子节点。

**语法**

**express.RemoveChild ( *Child* )**

\*express是一个代表 CustomXMLNode 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型      | 说明                     |
|---------|-----------|---------------|--------------------------|
| *Child* | 必选      | CustomXMLNode | 代表上下文节点的子节点。 |

**说明**

如果在 *Child* 参数中指定的节点不是上下文节点的子节点，或者操作将导致树无效，则不执行删除操作，并显示一条错误消息

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例选择一个自定义部件，然后选择该部件中的一个节点。代码然后删除该节点的子节点。
function test()
{
    let tion = Application.ActiveDocument
    // Return the first part with the given root namespace.
    let cxp1 = tion.CustomXMLParts.Item("urn:invoice:namespace")         
    // Get node using XPath expression.                             
    let cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]") 
    // Remove a child node.
    cxn.RemoveChild(cxn.SelectSingleNode("//discount")) 
}
```

### CustomXMLNode.ReplaceChildNode

从主树中删除指定的子节点（及其子树），并将它替换为相同位置中的其他节点。

**语法**

**express.ReplaceChildNode ( *OldNode* , *Name* , *NamespaceURI* , *NodeType* , *NodeValue* )**

\*express是一个代表 CustomXMLNode 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型             | 说明                                                                                                                                       |
|----------------|-----------|----------------------|--------------------------------------------------------------------------------------------------------------------------------------------|
| *OldNode*      | 必选      | CustomXMLNode        | 代表要替换的子节点。                                                                                                                       |
| *Name*         | 可选      | String               | 代表要添加的元素的基本名称。                                                                                                               |
| *NamespaceURI* | 可选      | String               | 代表要添加的元素的命名空间。如果要添加 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 类型的节点，则此参数是必需的，否则它将被忽略。 |
| *NodeType*     | 可选      | MsoCustomXMLNodeType | 指定要添加的节点的类型。如果未指定参数，则假定节点类型为 msoCustomXMLNodeElement。                                                         |
| *NodeValue*    | 可选      | String               | 用于为允许文本的那些节点设置要添加的节点值。如果节点不允许文本，则忽略该参数。                                                             |

**说明**

如果 *OldNode* 参数不是上下文节点的子节点，或者操作将导致树结构无效，则不执行替换操作，并显示一条错误消息。此外，如果要添加的节点已存在，则不执行替换操作，并显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例选择一个自定义部件，然后选择该部件中的一个节点。代码然后将该节点的子节点替换为其他节点。
function test()
{
    let tion = Application.ActiveDocument
    // Return the first custom xml part with the given root namespace.
    let cxp1 = tion.CustomXMLParts.Item("urn:invoice:namespace")                                 
    let cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]") 
    // Replace a child node.
    cxn.ReplaceChildNode(cxn.SelectSingleNode("//discount", "rebate")
}
```

### CustomXMLNode.ReplaceChildSubtree

从主树中删除指定节点（及其子树），并将它替换为相同位置中的其他子树。

**语法**

**express.ReplaceChildSubtree ( *XML* , *OldNode* )**

\*express是一个代表 CustomXMLNode 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型      | 说明                 |
|-----------|-----------|---------------|----------------------|
| *XML*     | 必选      | String        | 代表要添加的子树。   |
| *OldNode* | 必选      | CustomXMLNode | 代表要替换的子节点。 |

**说明**

如果操作将导致树结构无效，则不执行替换操作，并显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例选择一个自定义部件，然后选择该部件中的一个节点。代码然后将该节点的子子树替换为其他子树。
function test()
{
    let tion = Application.ActiveDocument
    // Return the first custom xml part with the given root namespace.
    let cxp1 = tion.CustomXMLParts.Item("urn:invoice:namespace")           
    // Get node using XPath expression.                             
    let cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]") 
    // Replace one subtree and its children with another.
    cxn.ReplaceChildSubtree("<rebates><rebate>0.10</rebate></rebates>", "//discounts")
}
```

### CustomXMLNode.SelectNodes

选择与 XPath 表达式匹配的节点的集合。此方法与 **CustomXMLPart** . **SelectNodes** 方法的不同之处在于，它将从作为上下文节点的“表达式”节点开始计算 XPath 表达式。

**语法**

**express.SelectNodes ( *XPath* )**

\*express是一个代表 CustomXMLNode 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                    |
|---------|-----------|----------|-------------------------|
| *XPath* | 必选      | String   | 包含一个 XPath 表达式。 |

**返回值**

CustomXMLNodes

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例说明了添加自定义 XML 部件，选择与命名空间 URI 匹配的部件，然后在该部件内选择与 XPath 表达式匹配的节点。
function test()
{
    // Add a custom xml part.
    Application.ActiveDocument.CustomXMLParts.Add("")
    // Return the first custom xml part with the given namespace.
    let cxp1 = Application.ActiveDocument.CustomXMLParts.Item("urn:invoice:namespace") 
    // Get all of the nodes matching an XPath expression.
    let cxns = cxp1.SelectNodes("//*[@unitPrice > 20]")
}
```

### CustomXMLNode.SelectSingleNode

从与 XPath 表达式匹配的集合中选择一个节点。此方法与 **CustomXMLPart** . **SelectSingleNode** 方法的不同之处在于，它将从作为上下文节点的“表达式”节点开始计算 XPath 表达式。

**语法**

**express.SelectSingleNode ( *XPath* )**

\*express是一个代表 CustomXMLNode 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                    |
|---------|-----------|----------|-------------------------|
| *XPath* | 必选      | String   | 包含一个 XPath 表达式。 |

**说明**

XPath 表达式的前缀映射是从 **NamespaceManager** 属性检索的。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例说明了添加自定义 XML 部件，选择与命名空间 URI 匹配的部件，然后在该部件内选择与 XPath 表达式匹配的节点。
function test()
{
    // Add a custom xml part.
    Application.ActiveDocument.CustomXMLParts.Add("")
    // Return the first custom xml part with the given namespace.
    let cxp1 = Application.ActiveDocument.CustomXMLParts.Item("urn:invoice:namespace")        
    // Get a node using XPath.                             
    let cxn = cxp1.Item(1).SelectSingleNode("//*[@supplierID = 1]")
}
```

## 成员属性

### CustomXMLNode.Application

获取一个代表 **CustomXMLNode** 的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.Attributes

获取一个代表当前节点中的当前元素的属性的 **CustomXMLNodes** 集合。只读。

**语法**

**express.Attributes**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.BaseName

获取文档对象模型 (DOM) 中节点的不带命名空间前缀的基本名称（如果存在）。只读。

**语法**

**express.BaseName**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

这是 **CustomXMLNode** 对象的默认成员。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.ChildNodes

获取一个包含当前节点中所有子元素的 **CustomXMLNodes** 集合。只读。

**语法**

**express.ChildNodes**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.Creator

获取一个 32 位整数，指示创建 **CustomXMLNode** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.FirstChild

获取与当前节点中第一个子元素对应的 **CustomXMLNode** 对象。如果该节点没有子元素（或者其类型不为 **msoCustomXMLNodeElement** ），则返回 **Nothing** 。只读。

**语法**

**express.FirstChild**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.LastChild

获取与当前节点中的最后一个子元素对应的 **CustomXMLNode** 对象。如果该节点没有子元素（或者其类型不为 **msoCustomXMLNodeElement** ），则该属性将返回 **Nothing** 。只读。

**语法**

**express.LastChild**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.NamespaceURI

获取 **CustomXMLNode** 对象的命名空间的唯一地址标识符。只读。

**语法**

**express.NamespaceURI**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.NextSibling

获取当前节点的下一个同级节点（元素、注释或处理指令）。如果该节点是其所在级别的最后一个同级节点，则该属性将返回 **Nothing** 。只读。

**语法**

**express.NextSibling**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.NodeType

获取当前节点的类型。只读。

**语法**

**express.NodeType**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.NodeValue

获取或设置当前节点的值。可读/写。

**语法**

**express.NodeValue**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                               |
|--------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常 |

### CustomXMLNode.OwnerDocument

获取代表与此节点关联的 ET 工作簿、Microsoft PowerPoint 演示文稿或 WPS 文档的对象。只读。

**语法**

**express.OwnerDocument**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.OwnerPart

获取代表与此节点关联的部件的对象。只读。

**语法**

**express.OwnerPart**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.Parent

获取 **CustomXMLNode** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.ParentNode

获取当前节点的父元素节点。如果当前节点位于根级，则该属性将返回 **Nothing** 。只读。

**语法**

**express.ParentNode**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.PreviousSibling

获取当前节点的前一个同级节点（元素、注释或处理指令）。如果当前节点为其所在级别的第一个节点，则该属性将返回 **Nothing** 。只读。

**语法**

**express.PreviousSibling**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.Text

获取或设置当前节点的文本。可读/写。

**语法**

**express.Text**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.XML

获取当前节点及其子项（如果存在）的 XML 表示形式。只读。

**语法**

**express.XML**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLNode.XPath

获取一个 **String** 类型的值，其中包含当前节点的规范化 XPath。如果该节点已不在文档对象模型 (DOM) 中，则该属性将返回一条错误消息。只读。

**语法**

**express.XPath**

\*express是一个代表 CustomXMLNode 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLNode](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# EncryptionProvider 对象

## 简介

提供用于设置权限、应用基础加密和解密的加密技术以及进行用户身份验证的方法。

## 说明

加密提供程序通过自定义的 COM 加载项实现。Office 文档内提供了存储区来存储加载项特定的信息，可以在其中存储加密、解密、应用权限和显示权限设置或验证用户界面所需的任何信息。

## 方法

| 序号 | 名称                                                       | 说明                                                                                                                                       |
|:----:|:-----------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Authenticate](#EncryptionProvider.Authenticate)           | 用于确定用户是否具有打开加密文档的适当权限。                                                                                               |
|  2   | [CloneSession](#EncryptionProvider.CloneSession)           | 为即将保存的文件创建 EncryptionProvider 对象加密会话的另一个工作副本。                                                                     |
|  3   | [DecryptStream](#EncryptionProvider.DecryptStream)         | 对文档的加密数据流进行解密并将其返回。                                                                                                     |
|  4   | [EncryptStream](#EncryptionProvider.EncryptStream)         | 加密并返回文档的数据流。                                                                                                                   |
|  5   | [EndSession](#EncryptionProvider.EndSession)               | 结束当前加密会话。                                                                                                                         |
|  6   | [GetProviderDetail](#EncryptionProvider.GetProviderDetail) | 显示有关当前文档的加密的信息。                                                                                                             |
|  7   | [NewSession](#EncryptionProvider.NewSession)               | 由 EncryptionProvider 对象在创建新的加密会话时使用。当文档被调入内存中后，提供程序将使用此会话缓存与加密、用户和权限相关的文档特定的信息。 |
|  8   | [Save](#EncryptionProvider.Save)                           | 保存加密的文档。                                                                                                                           |
|  9   | [ShowSettings](#EncryptionProvider.ShowSettings)           | 用于显示当前文档的加密设置的对话框。                                                                                                       |

## 成员方法

### EncryptionProvider.Authenticate

用于确定用户是否具有打开加密文档的适当权限。

**语法**

**express.Authenticate ( *ParentWindow* , *EncryptionData* , *PermissionsMask* )**

\*express是一个代表 EncryptionProvider 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型   | 说明                               |
|-------------------|-----------|------------|------------------------------------|
| *ParentWindow*    | 必选      | IUnknown   | 指定为了显示加密设置而调用的窗口。 |
| *EncryptionData*  | 必选      | IUnknown   | 包含当前文档的加密数据。           |
| *PermissionsMask* | 必选      | 无符号整数 | 加密提供程序加载项显示的用户界面。 |

**返回值**

Long

**说明**

它是 COM 加载项加密提供程序的显示位置，与适于应用加密的用户界面无关。例如，密码加密提供程序将提示用户输入密码。

### EncryptionProvider.CloneSession

为即将保存的文件创建 **EncryptionProvider** 对象加密会话的另一个工作副本。

**语法**

**express.CloneSession ( *SessionHandle* )**

\*express是一个代表 EncryptionProvider 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明              |
|-----------------|-----------|----------|-------------------|
| *SessionHandle* | 必选      | Long     | 克隆的会话的 ID。 |

**说明**

此方法的输出是另一个具有相同加密设置的会话句柄。当执行自动保存操作时，将向您提供此会话句柄。

### EncryptionProvider.DecryptStream

对文档的加密数据流进行解密并将其返回。

**语法**

**express.DecryptStream ( *SessionHandle* , *StreamName* , *EncryptedStream* , *UnencryptedStream* )**

\*express是一个代表 EncryptionProvider 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明               |
|---------------------|-----------|----------|--------------------|
| *SessionHandle*     | 必选      | Long     | 当前会话的 ID。    |
| *StreamName*        | 必选      | String   | 数据流的 ID。      |
| *EncryptedStream*   | 必选      | IUnknown | 已加密数据流。     |
| *UnencryptedStream* | 必选      | IUnknown | 解密之前的数据流。 |

**说明**

此方法由 COM 加载项调用，调用时间为文档已打开并且在加载项已验证打开文档的用户已经过身份验证后。此方法是 EncryptStream 方法的逆向操作，它将加密数据转换回纯（未加密）数据。

### EncryptionProvider.EncryptStream

加密并返回文档的数据流。

**语法**

**express.EncryptStream ( *SessionHandle* , *StreamName* , *UnencryptedStream* , *EncryptedStream* )**

\*express是一个代表 EncryptionProvider 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明                     |
|---------------------|-----------|----------|--------------------------|
| *SessionHandle*     | 必选      | Long     | 当前会话的 ID。          |
| *StreamName*        | 必选      | String   | 加密的文档数据流的名称。 |
| *UnencryptedStream* | 必选      | IUnknown | 加密之前的数据流。       |
| *EncryptedStream*   | 必选      | IUnknown | 加密后的数据流信息。     |

**说明**

此方法通常是在执行保存操作的过程中由 COM 加载项调用。

### EncryptionProvider.EndSession

结束当前加密会话。

**语法**

**express.EndSession ( *SessionHandle* )**

\*express是一个代表 EncryptionProvider 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明            |
|-----------------|-----------|----------|-----------------|
| *SessionHandle* | 必选      | Long     | 当前会话的 ID。 |

**说明**

在执行保存操作期间，COM 加载项将调用 **CloneSession** 方法，以便为即将保存的文件创建 **EncryptionProvider** 对象加密会话的另一个工作副本。接下来调用 **Save** 方法，以获取要保留的有关加密设置的所有自定义信息。在以后再次打开此文档时，便可以使用此信息。然后调用 **EncryptStream** 方法，向提供程序提供文档的全部内容。最后，为克隆的会话句柄调用 **EndSession** 方法来完成该过程。

### EncryptionProvider.GetProviderDetail

显示有关当前文档的加密的信息。

**语法**

**express.GetProviderDetail ( *encprovdet* )**

\*express是一个代表 EncryptionProvider 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型                 | 说明                 |
|--------------|-----------|--------------------------|----------------------|
| *encprovdet* | 必选      | EncryptionProviderDetail | 指定所需的加密信息。 |

**返回值**

Variant

**说明**

通过此方法，可以查询 **EncryptionProvider** 对象以获取类似下面的信息：未安装您的自定义 COM 加载项的用户的下载 URL、实现的算法以及使用的加密模式。

### EncryptionProvider.NewSession

由 **EncryptionProvider** 对象在创建新的加密会话时使用。当文档被调入内存中后，提供程序将使用此会话缓存与加密、用户和权限相关的文档特定的信息。

**语法**

**express.NewSession ( *ParentWindow* )**

\*express是一个代表 EncryptionProvider 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                               |
|----------------|-----------|----------|------------------------------------|
| *ParentWindow* | 必选      | IUnknown | 指定为了显示加密设置而调用的窗口。 |

**返回值**

Long

**说明**

此方法由 COM 加载项调用。

### EncryptionProvider.Save

保存加密的文档。

**语法**

**express.Save ( *SessionHandle* , *EncryptionData* )**

\*express是一个代表 EncryptionProvider 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明            |
|------------------|-----------|----------|-----------------|
| *SessionHandle*  | 必选      | Long     | 当前会话的 ID。 |
| *EncryptionData* | 必选      | IUnknown | 包含加密信息。  |

**返回值**

Long

**说明**

将文件保存为 Office Open XML 文件格式（它是支持自定义文件加密的唯一格式）时，COM 加载项将调用提供程序来加密文档。如果您尝试保存的目标文件格式不支持自定义文件加密，而您具有执行此操作的相应权限，则 WPS Office将在不加密的情况下保存该文档。这样便可以将文档导出为不支持加密或权限管理的格式。

### EncryptionProvider.ShowSettings

用于显示当前文档的加密设置的对话框。

**语法**

**express.ShowSettings ( *SessionHandle* , *ParentWindow* , *ReadOnly* , *Remove* )**

\*express是一个代表 EncryptionProvider 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                          |
|-----------------|-----------|----------|---------------------------------------------------------------|
| *SessionHandle* | 必选      | Long     | 当前会话的 ID。                                               |
| *ParentWindow*  | 必选      | IUnknown | 指定为了显示加密设置而调用的窗口。                            |
| *ReadOnly*      | 必选      | Boolean  | 指定是否希望用户能更改加密设置。                              |
| *Remove*        | 必选      | Boolean  | 如果为 True，则在执行下一个保存操作的过程中将删除文档的加密。 |

**说明**

只能对已加密的文档调用此方法。您可以在 COM 加载项中使用此方法，以基于用户的任务显示您所希望的用户体验。例如，在纯加密情况中，可以显示一个对话框来更改文档的密码。在权限管理情况中，可以决定是显示用于更改权限的对话框，还是显示用户的权限。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/EncryptionProvider](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FileDialogFilter 对象

## 简介

返回通过 FileDialog 对象显示的文件对话框中的一个文件筛选。每个文件筛选确定文件对话框中显示的文件。

## 说明

使用 FileDialogFilters 集合的 Item 方法可返回一个 FileDialogFilter 对象。使用 Add 方法可向 FileDialogFilters 集合中添加一个 FileDialogFilter 对象。使用 Extensions 属性可返回 FileDialogFilter 对象用于筛选文件的扩展名，使用 Description 属性可以返回筛选的说明，但是这两个属性均为只读。如果要设置扩展名和说明，必须使用 Add 方法。

示例

下面的示例循环访问 “SaveAs” 对话框的默认筛选器，并显示每个包括 ET 文件的筛选器的说明。

``` JavaScript
function test() {

    //Declare a variable as a FileDialogFilters collection.

    //Declare a variable as a FileDialogFilter object.

    //Set the FileDialogFilters collection variable to
    //the FileDialogFilters collection of the SaveAs dialog box.
    let fdfs = Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters

    //Iterate through the description and extensions of each
    //default filter in the SaveAs dialog box.
    for (let i = 1; i <= fdfs.Count; i++) {
        let fdf = fdfs.Item(i)
        //Display the description of filters that include
        //ET files.
        if (fdf.Extensions.indexOf("xls") > -1) {
            alert("Description of filter: " + fdf.Description)
        }
    }
}
```

## 属性

| 序号 | 名称                                         | 说明                                                                                                                                    |
|:----:|:---------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FileDialogFilter.Application) | 获取一个 Application 对象，代表 FileDialogFilter 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Creator](#FileDialogFilter.Creator)         | 获取一个 32 位整数，指示创建 FileDialogFilter 对象时所使用的应用程序。只读。                                                            |
|  3   | [Description](#FileDialogFilter.Description) | 获取一个 String 类型的值，代表每个 Filter 对象的说明。该说明为文件对话框中显示的文本。只读。                                            |
|  4   | [Extensions](#FileDialogFilter.Extensions)   | 获取一个包含扩展名的值，这些扩展名决定对于每个 Filter 对象在文件对话框中显示的文件。只读。                                              |
|  5   | [Parent](#FileDialogFilter.Parent)           | 获取 FileDialogFilter 对象的 Parent 对象。只读。                                                                                        |

## 成员属性

### FileDialogFilter.Application

获取一个 Application 对象，代表 **FileDialogFilter** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 FileDialogFilter 对象的变量。

**说明**

### FileDialogFilter.Creator

获取一个 32 位整数，指示创建 **FileDialogFilter** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 FileDialogFilter 对象的变量。

### FileDialogFilter.Description

获取一个 **String** 类型的值，代表每个 **Filter** 对象的说明。该说明为文件对话框中显示的文本。只读。

**语法**

**express.Description**

\*express是一个代表 FileDialogFilter 对象的变量。

**示例**

**示例**

下面的示例循环访问 **“SaveAs”** 对话框的默认筛选器，并显示每个包括 ET 文件的筛选器的说明。Extensions 属性用于查找相应的筛选对象。

``` JavaScript
function test() {

    //Declare a variable as a FileDialogFilters collection.

    //Declare a variable as a FileDialogFilter object.

    //Set the FileDialogFilters collection variable to
    //the FileDialogFilters collection of the SaveAs dialog box.
    let fdfs = Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters

    //Iterate through the description and extensions of each
    //default filter in the SaveAs dialog box.
    for(let i = 1;i <= fdfs.Count;i++) {
        let fdf = fdfs.Item(i)
        //Display the description of filters that include
        //ET files
        if(fdf.Extensions.indexOf("xls") > -1) {
            alert("Description of filter: " + fdf.Description)
        }
    }

}
```

### FileDialogFilter.Extensions

获取一个包含扩展名的值，这些扩展名决定对于每个 **Filter** 对象在文件对话框中显示的文件。只读。

**语法**

**express.Extensions**

\*express是一个代表 FileDialogFilter 对象的变量。

**示例**

下面的示例通过循环访问 **“SaveAs”** 对话框中的筛选器来显示 ET 文件的扩展名和说明。

``` JavaScript
function test() {

    //Declare a variable as a FileDialogFilters collection.

    //Declare a variable as a FileDialogFilter object.

    //Set the FileDialogFilters collection variable to
    //the FileDialogFilters collection of the SaveAs dialog box.
    let fdfs = Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters

    //Iterate through the description and extensions of each
    //default filter in the SaveAs dialog box.
    for(let i = 1;i <= fdfs.Count;i++) {
        let fdf = fdfs.Item(i)
        //Display the description of filters that include
        //ET files
        if(fdf.Extensions.indexOf("xls") > -1) {
            alert("Description of filter: " + fdf.Description)
        }
    }

}
```

### FileDialogFilter.Parent

获取 **FileDialogFilter** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FileDialogFilter 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/FileDialogFilter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FileTypes 对象

## 简介

msoFileType 类型的值的集合，这些值决定在搜索期间返回的文件类型。

## 说明

所有搜索都只有一个 FileTypes 集合，因此在执行搜索之前清除 FileTypes 集合很重要，除非希望搜索以前搜索的文件类型。清除该集合的最简便方法是将 FileType 属性设置为要搜索的第一个文件类型。还可以使用 Remove 方法删除单个类型。要确定集合中每一项的文件类型，请使用 Item 方法返回 msoFileType 值。

## 方法

| 序号 | 名称                        | 说明                             |
|:----:|:----------------------------|:---------------------------------|
|  1   | [Add](#FileTypes.Add)       | 向文件搜索中新添一个文件类型。   |
|  2   | [Remove](#FileTypes.Remove) | 从集合中删除一个 FileType 对象。 |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                             |
|:----:|:--------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FileTypes.Application) | 获取一个 Application 对象，代表 FileTypes 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Creator](#FileTypes.Creator)         | 获取一个 32 位整数，指示创建 FileTypes 对象时所使用的应用程序。只读。                                                            |
|  3   | [Item](#FileTypes.Item)               | 获取一个指示要搜索的文件类型的值。只读。                                                                                         |

## 成员方法

### FileTypes.Add

向文件搜索中新添一个文件类型。

**语法**

**express.Add ( *FileType* )**

\*express是一个代表 FileTypes 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型    | 说明                   |
|------------|-----------|-------------|------------------------|
| *FileType* | 必选      | MsoFileType | 指定要搜索的文件类型。 |

### FileTypes.Remove

从集合中删除一个 **FileType** 对象。

**语法**

**express.Remove ( *Index* )**

\*express是一个代表 FileTypes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                       |
|---------|-----------|----------|----------------------------|
| *Index* | 必选      | Long     | 要删除的文件类型的索引号。 |

## 成员属性

### FileTypes.Application

获取一个 **Application** 对象，代表 **FileTypes** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 FileTypes 对象的变量。

### FileTypes.Creator

获取一个 32 位整数，指示创建 **FileTypes** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 FileTypes 对象的变量。

**类型**

Long

### FileTypes.Item

获取一个指示要搜索的文件类型的值。只读。

**语法**

**express.Item**

\*express是一个代表 FileTypes 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/FileTypes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
