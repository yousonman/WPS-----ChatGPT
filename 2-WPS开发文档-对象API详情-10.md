# Env 对象

## 简介

Env 对象主要用于取系统环境基本信息，这个对象目前提供了取用户目录、临时目录等相关信息，后续可能会根据实际需求不断扩充。

## 说明

以下代码示例，取到系统临时目录并打印到控制台。

``` JavaScript
let tempPath = Application.Env.GetTempPath()
console.log(tempPath)  //如果是jde环境，则是Debug.Print(tempPath)
```

## 方法

| 序号 | 名称                                            | 说明                                                  |
|:----:|:------------------------------------------------|:------------------------------------------------------|
|  1   | [GetAppDataPath](#Env.GetAppDataPath)           | 取系统%AppData%/Roming目录的路径，仅Windows平台适用。 |
|  2   | [GetDesktopDpi](#Env.GetDesktopDpi)             | 取系统DPI。                                           |
|  3   | [GetHomePath](#Env.GetHomePath)                 | 取系统HOME目录的路径，代表当前用户主目录。            |
|  4   | [GetProgramDataPath](#Env.GetProgramDataPath)   | 取系统ProgramData目录的路径，仅windows平台适用。      |
|  5   | [GetProgramFilesPath](#Env.GetProgramFilesPath) | 取系统ProgramFiles目录的路径，仅Windows平台适用。     |
|  6   | [GetRootPath](#Env.GetRootPath)                 | 取系统根目录的路径。                                  |
|  7   | [GetTempPath](#Env.GetTempPath)                 | 取系统临时目录的路径。                                |

## 成员方法

### Env.GetAppDataPath

取系统%AppData%/Roming目录的路径，仅Windows平台适用。

**语法**

**express.GetAppDataPath ()**

\*express是一个代表 Env 对象的变量。

### Env.GetDesktopDpi

取系统DPI。

**语法**

**express.GetDesktopDpi ()**

\*express是一个代表 Env 对象的变量。

### Env.GetHomePath

取系统HOME目录的路径，代表当前用户主目录。

**语法**

**express.GetHomePath ()**

\*express是一个代表 Env 对象的变量。

### Env.GetProgramDataPath

取系统ProgramData目录的路径，仅windows平台适用。

**语法**

**express.GetProgramDataPath ()**

\*express是一个代表 Env 对象的变量。

### Env.GetProgramFilesPath

取系统ProgramFiles目录的路径，仅Windows平台适用。

**语法**

**express.GetProgramFilesPath ()**

\*express是一个代表 Env 对象的变量。

### Env.GetRootPath

取系统根目录的路径。

**语法**

**express.GetRootPath ()**

\*express是一个代表 Env 对象的变量。

### Env.GetTempPath

取系统临时目录的路径。

**语法**

**express.GetTempPath ()**

\*express是一个代表 Env 对象的变量。

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/系统环境/Env](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# HLayout 对象

## 简介

将内部的控件按照水平方向排布，一列一个。

## 说明

将内部的控件按照水平方向排布，一列一个。

## 方法

| 序号 | 名称                  | 说明                       |
|:----:|:----------------------|:---------------------------|
|  1   | [Move](#HLayout.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                              | 说明                                                 |
|:----:|:--------------------------------------------------|:-----------------------------------------------------|
|  1   | [Enabled](#HLayout.Enabled)                       | 您可以通过该属性设置或返回是否启用该控件             |
|  2   | [Height](#HLayout.Height)                         | 获取或设置指定对象的高度                             |
|  3   | [LayoutBottomMargin](#HLayout.LayoutBottomMargin) | 获取或设置容器内部的控件与容器底部的距离             |
|  4   | [LayoutLeftMargin](#HLayout.LayoutLeftMargin)     | 获取或设置容器内部的控件与容器左边缘的距离           |
|  5   | [LayoutRightMargin](#HLayout.LayoutRightMargin)   | 获取或设置容器内部的控件与容器右边缘的距离           |
|  6   | [LayoutSpacing](#HLayout.LayoutSpacing)           | 获取或设置容器内部多个控件之间的水平间距             |
|  7   | [LayoutTopMargin](#HLayout.LayoutTopMargin)       | 获取或设置容器内部的控件与容器顶部的距离             |
|  8   | [Left](#HLayout.Left)                             | 您可以通过该属性指定控件在对象或报表中的位置         |
|  9   | [Name](#HLayout.Name)                             | 您可以通过该属性指定或确定表示对象名称的字符串，只读 |
|  10  | [Top](#HLayout.Top)                               | 您可以通过该属性指定控件在对象或报表中的位置         |
|  11  | [Visible](#HLayout.Visible)                       | 返回或设置该对象是否可见                             |
|  12  | [Width](#HLayout.Width)                           | 获取或设置指定对象的宽度                             |

## 成员方法

### HLayout.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 HLayout 对象的变量。

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
/*将HLayout1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.HLayout1.Move(100,100,100,100)
}
```

## 成员属性

### HLayout.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 HLayout 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.HLayout1.Enabled = true
}
```

### HLayout.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将HLayout1控件的高度设置为100*/
function func1()
{
    UserForm1.HLayout1.Height = 100
}
```

### HLayout.LayoutBottomMargin

获取或设置容器内部的控件与容器底部的距离

**语法**

**express.LayoutBottomMargin**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器底部的距离

**示例**

``` JavaScript
/*将HLayout1中控件到容器底部的距离设为10*/
function func1()
{
    UserForm1.HLayout1.LayoutBottomMargin = 10
}
```

### HLayout.LayoutLeftMargin

获取或设置容器内部的控件与容器左边缘的距离

**语法**

**express.LayoutLeftMargin**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器左边缘的距离

**示例**

``` JavaScript
/*将HLayout1中控件到容器左边缘的距离设为10*/
function func1()
{
    UserForm1.HLayout1.LayoutLeftMargin = 10
}
```

### HLayout.LayoutRightMargin

获取或设置容器内部的控件与容器右边缘的距离

**语法**

**express.LayoutRightMargin**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器右边缘的距离

**示例**

``` JavaScript
/*将HLayout1中控件到容器右边缘的距离设为10*/
function func1()
{
    UserForm1.HLayout1.LayoutRightMargin = 10
}
```

### HLayout.LayoutSpacing

获取或设置容器内部多个控件之间的水平间距

**语法**

**express.LayoutSpacing**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置容器内部多个控件之间的水平间距

**示例**

``` JavaScript
/*将HLayout1中多个控件之间的距离设为10*/
function func1()
{
    UserForm1.HLayout1.LayoutSpacing = 10
}
```

### HLayout.LayoutTopMargin

获取或设置容器内部的控件与容器顶部的距离

**语法**

**express.LayoutTopMargin**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置容器内部的控件与容器顶部的距离

**示例**

``` JavaScript
/*将HLayout1中控件到容器顶部的距离设为10*/
function func1()
{
    UserForm1.HLayout1.LayoutTopMargin = 10
}
```

### HLayout.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 HLayout 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将HLayout1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.HLayout1.Left = 100
}
```

### HLayout.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 HLayout 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### HLayout.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 HLayout 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将HLayout1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.HLayout1.Top = 100
}
```

### HLayout.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 HLayout 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将HLayout1设置为可见*/
function func1()
{
    UserForm1.HLayout1.Visible = true
}
```

### HLayout.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 HLayout 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将HLayout1的宽度设置为100*/
function func1()
{
    UserForm1.HLayout1.Width = 100
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/HLayout](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# UserForm 对象

## 简介

UserForm 对象是构成应用程序用户界面的一部分的窗口或对话框。

## 说明

用户窗体对象，是一个容器。代表用户创建的窗体，用户可以向其中添加控件，通过该可以调用添加的控件对象，修改对象的属性，调用对象的方法。

## 方法

| 序号 | 名称                           | 说明                                           |
|:----:|:-------------------------------|:-----------------------------------------------|
|  1   | [Controls](#UserForm.Controls) | 返回控件的窗体、子窗体、报告或者区块的控件合集 |
|  2   | [Hide](#UserForm.Hide)         | 隐藏UserForm对象但不卸载该对象                 |
|  3   | [Move](#UserForm.Move)         | 将对象移动到参数指定的位置                     |
|  4   | [Show](#UserForm.Show)         | 显示一个用户窗体对象                           |

## 属性

| 序号 | 名称                                       | 说明                                                                |
|:----:|:-------------------------------------------|:--------------------------------------------------------------------|
|  1   | [BackColor](#UserForm.BackColor)           | 返回或指定当前对象的背景颜色                                        |
|  2   | [BorderColor](#UserForm.BorderColor)       | 返回或 当前对象的边框颜色，只读                                     |
|  3   | [Caption](#UserForm.Caption)               | 返回或指定用来识别或描述对象的显示在对象上的描述性文字，只读        |
|  4   | [Enabled](#UserForm.Enabled)               | 指定该控件是否能接收焦点或者对用户生成的事件进行响应，只读          |
|  5   | [Font](#UserForm.Font)                     | 返回或 指定对象的字体，只读                                         |
|  6   | [ForeColor](#UserForm.ForeColor)           | 返回或 指定对象的前景颜色                                           |
|  7   | [Height](#UserForm.Height)                 | 返回或 指定以像素点表示的对象的高                                   |
|  8   | [Left](#UserForm.Left)                     | 返回或 指定控件到包含它的窗体的左边缘的距离                         |
|  9   | [Name](#UserForm.Name)                     | 返回或 指定控件或对象的名称，或与Font对象关联的字体对象的名称，只读 |
|  10  | [StartUpPostion](#UserForm.StartUpPostion) | 返回或者指定UserForm第一次出现的位置                                |
|  11  | [Top](#UserForm.Top)                       | 返回或 指定控件到包含它的窗体的上边缘的距离                         |
|  12  | [Width](#UserForm.Width)                   | 返回或 指定以像素点表示的对象的宽                                   |

## 成员方法

### UserForm.Controls

返回控件的窗体、子窗体、报告或者区块的控件合集

**语法**

**express.Controls ()**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回控件的窗体、子窗体、报告或者区块的控件合集

**示例**

``` JavaScript
/*通过button名获取button对象，并打印button名*/
function func1()
{
    var buttonName = UserForm1.CommandButton1.Name
    Console.log(buttonName)
    var controlSet = UserForm1.Controls(buttonName)
    Console.log(controlSet.Name)
}
```

### UserForm.Hide

隐藏UserForm对象但不卸载该对象

**语法**

**express.Hide ()**

\*express是一个代表 UserForm 对象的变量。

**说明**

隐藏UserForm对象但不卸载该对象

**示例**

``` JavaScript
/*隐藏该UserForm*/
function func1()
{
    UserForm1.Hide()
}
 
```

### UserForm.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 UserForm 对象的变量。

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
/*将窗体移动到窗口的左上角，并将移动后窗体的长宽设置为500*/
function func1()
{
    UserForm1.Move(100,100,500,500)
}
```

### UserForm.Show

显示一个用户窗体对象

**语法**

**express.Show ()**

\*express是一个代表 UserForm 对象的变量。

**说明**

显示一个用户窗体对象

**示例**

``` JavaScript
/*显示用户窗体*/
function func1()
{
    UserForm1.Show()
}
```

## 成员属性

### UserForm.BackColor

返回或指定当前对象的背景颜色

**语法**

**express.BackColor**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回或 指定当前对象的背景颜色

**示例**

``` JavaScript
/*设置UserForm的背景颜色为十六进制下550000表示的颜色（红棕色）*/
function func1()
{
    UserForm1.BackColor = 0x550000
}
```

### UserForm.BorderColor

返回或 当前对象的边框颜色，只读

**语法**

**express.BorderColor**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回或 当前对象的边框颜色，只读

### UserForm.Caption

返回或指定用来识别或描述对象的显示在对象上的描述性文字，只读

**语法**

**express.Caption**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回或指定用来识别或描述对象的显示在对象上的描述性文字，只读

### UserForm.Enabled

指定该控件是否能接收焦点或者对用户生成的事件进行响应，只读

**语法**

**express.Enabled**

\*express是一个代表 UserForm 对象的变量。

**说明**

指定该控件是否能接收焦点或者对用户生成的事件进行响应，只读

### UserForm.Font

返回或 指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回或 指定对象的字体 ，只读

### UserForm.ForeColor

返回或 指定对象的前景颜色

**语法**

**express.ForeColor**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回或 指定对象的前景颜色

**示例**

``` JavaScript
/*设置UserForm窗体的前景颜色为十六进程550000表示的颜色*/
function func1()
{
    UserForm1.ForeColor = 0x550000
}
```

### UserForm.Height

返回或 指定以像素点表示的对象的高

**语法**

**express.Height**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回或 指定以像素点表示的对象的高

**示例**

``` JavaScript
/*设置UserForm窗体的高为1000*/
function func1()
{
    UserForm1.Height = 1000
}
```

### UserForm.Left

返回或 指定控件到包含它的窗体的左边缘的距离

**语法**

**express.Left**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回或 指定控件到包含它的窗体的左边缘的距离

**示例**

``` JavaScript
/*设置UserForm窗体到窗口左侧边缘的距离为1000*/
function func1()
{
    UserForm1.Left = 1000
}
```

### UserForm.Name

返回或 指定控件或对象的名称，或与Font对象关联的字体对象的名称，只读

**语法**

**express.Name**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回或 指定控件或对象的名称，或与Font对象关联的字体对象的名称，只读

### UserForm.StartUpPostion

返回或者指定UserForm第一次出现的位置

**语法**

**express.StartUpPostion**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回或者指定UserForm第一次出现的位置

| 设置               | 值  | 描述                               |
|--------------------|-----|------------------------------------|
| **Manual**         | 0   | 不指定初始化设置                   |
| **CenterOwner**    | 1   | 出现在包含用户窗体的所在窗体的中间 |
| **CenterScreen**   | 2   | 出现在整个屏幕的中间               |
| **WindowsDefault** | 3   | 出现在屏幕的左上角                 |

**示例**

``` JavaScript
/*设置UserForm窗体第一次出现时在整个屏幕的中间*/
function func1()
{
    UserForm1.StartUpPostion = 2
}
```

### UserForm.Top

返回或 指定控件到包含它的窗体的上边缘的距离

**语法**

**express.Top**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回或 指定控件到包含它的窗体的上边缘的距离

**示例**

``` JavaScript
/*设置UserForm窗体到窗口上边缘的距离为1000*/
function func1()
{
    UserForm1.Top = 1000
}
```

### UserForm.Width

返回或 指定以像素点表示的对象的宽

**语法**

**express.Width**

\*express是一个代表 UserForm 对象的变量。

**说明**

返回或 指定以像素点表示的对象的宽

**示例**

``` JavaScript
/*设置UserForm窗体的宽为1000*/
function func1()
{
    UserForm1.Width = 1000
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/UserForm](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# AutoCorrect 对象

## 简介

包含 ET 的 AutoCorrect 属性（自动将日期名改为大写、自动更正连续两个大写字母、自动更正词条列表等等）。

## 说明

``` JavaScript
/*使用 AutoCorrect 属性可返回 AutoCorrect 对象。下例使 ET 更正以连续两个大写字母开头的单词。*/
function test()
{
    let correct = Application.AutoCorrect
    correct.TwoInitialCapitals = true
    correct.ReplaceText = true
}
```

## 方法

| 序号 | 名称                                                | 说明                                   |
|:----:|:----------------------------------------------------|:---------------------------------------|
|  1   | [AddReplacement](#AutoCorrect.AddReplacement)       | 向“自动更正”替换文本数组中添加项。     |
|  2   | [DeleteReplacement](#AutoCorrect.DeleteReplacement) | 删除“自动更正”替换数组中的一个输入项。 |

## 属性

| 序号 | 名称                                                                | 说明                                                                                                                                                                                                                                                                                                                                                                                      |
|:----:|:--------------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#AutoCorrect.Application)                             | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。                                                                                                                                                           |
|  2   | [AutoExpandListRange](#AutoCorrect.AutoExpandListRange)             | 表示 <span class="glossary" "appendpopup(this,'xldeflist_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">列表</span> 的自动扩展是否被启用的 Boolean 值。在列表旁的空行或空列键入时，如果启用了自动扩展功能，则列表将扩展为包含此行或此列。 Boolean 类型，可读写。 |
|  3   | [AutoFillFormulasInLists](#AutoCorrect.AutoFillFormulasInLists)     | 影响由自动填充列表创建的计算列的创建。可读/写 Boolean 类型。                                                                                                                                                                                                                                                                                                                              |
|  4   | [CapitalizeNamesOfDays](#AutoCorrect.CapitalizeNamesOfDays)         | 如果日期名称的第一个字母自动大写，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                                                                                                                             |
|  5   | [CorrectCapsLock](#AutoCorrect.CorrectCapsLock)                     | 如果 ET 自动更正无意中使用 Caps Lock 键，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                                                                                                                      |
|  6   | [CorrectSentenceCap](#AutoCorrect.CorrectSentenceCap)               | 如果 ET 自动更正句子（第一个单词）的大写，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                                                                                                                     |
|  7   | [Creator](#AutoCorrect.Creator)                                     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                                                                                                                                                                                |
|  8   | [DisplayAutoCorrectOptions](#AutoCorrect.DisplayAutoCorrectOptions) | 允许用户显示或隐藏 “自动更正选项” 按钮。默认值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                                                                                                                           |
|  9   | [Parent](#AutoCorrect.Parent)                                       | 返回指定对象的父对象。只读。                                                                                                                                                                                                                                                                                                                                                              |
|  10  | [ReplaceText](#AutoCorrect.ReplaceText)                             | 如果“自动更正”替换内容将被自动替换，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                                                                                                                           |
|  11  | [ReplacementList](#AutoCorrect.ReplacementList)                     | 返回“自动更正”替换内容的数组。                                                                                                                                                                                                                                                                                                                                                            |
|  12  | [TwoInitialCapitals](#AutoCorrect.TwoInitialCapitals)               | 如果自动更正以两个大写字母开头的词，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                                                                                                                                                                           |

## 成员方法

### AutoCorrect.AddReplacement

向“自动更正”替换文本数组中添加项。

**语法**

**express.AddReplacement ( *What* , *Replacement* )**

\*express是一个代表 AutoCorrect 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                       |
|---------------|-----------|----------|--------------------------------------------------------------------------------------------|
| *What*        | 必选      | String   | 要替换的文本。如果该字符串已存在于自动更正替换文本数组中，则现有的替换文本将被新文本替换。 |
| *Replacement* | 必选      | String   | 替换文本。                                                                                 |

**示例**

``` JavaScript
/*本示例在自动更正替换文字数组中，将单词“Paul”的替换单词指定为“Dreamer-Paul”。*/
Application.AutoCorrect.AddReplacement("Paul", "Dreamer-Paul")
```

### AutoCorrect.DeleteReplacement

删除“自动更正”替换数组中的一个输入项。

**语法**

**express.DeleteReplacement ( *What* )**

\*express是一个代表 AutoCorrect 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                   |
|--------|-----------|----------|------------------------------------------------------------------------------------------------------------------------|
| *What* | 必选      | String   | 要替换的文本（当文本出现在要从“自动更正”替换数组中删除的行上时）。如果该字符串不在“自动更正”替换数组中，此方法将失败。 |

**说明**

**示例**

``` JavaScript
/*本示例删除“自动更正”替换数组中的“Paul”词条。*/
Application.AutoCorrect.DeleteReplacement("Paul")
```

## 成员属性

### AutoCorrect.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 AutoCorrect 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
  if(Application.AutoCorrect.Application.Value == "ET"){
      alert("This is an ET Application object.")
  }
  else{
      alert("This is not an ET Application object.")
  }
}
```

### AutoCorrect.AutoExpandListRange

表示 <span class="glossary" "appendpopup(this,'xldeflist_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">列表</span> 的自动扩展是否被启用的 **Boolean** 值。在列表旁的空行或空列键入时，如果启用了自动扩展功能，则列表将扩展为包含此行或此列。 **Boolean** 类型，可读写。

**语法**

**express.AutoExpandListRange**

\*express是一个代表 AutoCorrect 对象的变量。

**示例**

``` JavaScript
/*下例在相邻的行或列键入时启用了列表的自动扩展功能。*/
function test(){
    Application.AutoCorrect.AutoExpandListRange = true
}
```

### AutoCorrect.AutoFillFormulasInLists

影响由自动填充列表创建的计算列的创建。可读/写 **Boolean** 类型。

**语法**

**express.AutoFillFormulasInLists**

\*express是一个代表 AutoCorrect 对象的变量。

**说明**

该属性不影响现有计算列或非自动填充创建的计算列。

### AutoCorrect.CapitalizeNamesOfDays

如果日期名称的第一个字母自动大写，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CapitalizeNamesOfDays**

\*express是一个代表 AutoCorrect 对象的变量。

**示例**

``` JavaScript
/*本示例设置 ET 将日期名称的第一个字母大写。*/
function test()
{
  let correct = Application.AutoCorrect
  correct.CapitalizeNamesOfDays = true
  correct.ReplaceText = true
}
```

### AutoCorrect.CorrectCapsLock

如果 ET 自动更正无意中使用 Caps Lock 键，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CorrectCapsLock**

\*express是一个代表 AutoCorrect 对象的变量。

**示例**

``` JavaScript
/*本示例对 Caps Lock 的处理进行设置，使 ET 可自动更正无意中使用 Caps Lock。*/
Application.AutoCorrect.CorrectCapsLock = true
```

### AutoCorrect.CorrectSentenceCap

如果 ET 自动更正句子（第一个单词）的大写，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CorrectSentenceCap**

\*express是一个代表 AutoCorrect 对象的变量。

**示例**

``` JavaScript
/*本示例对大写功能的处理进行设置，使 ET 可自动更正句子大写。*/
Application.AutoCorrect.CorrectSentenceCap = true
```

### AutoCorrect.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 AutoCorrect 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### AutoCorrect.DisplayAutoCorrectOptions

允许用户显示或隐藏 **“自动更正选项”** 按钮。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayAutoCorrectOptions**

\*express是一个代表 AutoCorrect 对象的变量。

**说明**

**DisplayAutoCorrectOptions** 属性是一个 WPS Office 范围的设置。在 ET 中更改该属性将影响其他的 Office 应用程序。

在 ET 中，只有在自动创建超链接时才会显示 **“自动更正选项”** 按钮。

**示例**

``` JavaScript
/*本示例确定是否显示“自动更正选项”按钮并通知用户。*/
function CheckDisplaySetting(){
    // Determine setting and notify user.
    if(Application.AutoCorrect.DisplayAutoCorrectOptions == true){
        alert("The AutoCorrect Options button can be displayed.")
    }
    else{
        alert("The AutoCorrect Options button cannot be displayed.")
    }
}
```

### AutoCorrect.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 AutoCorrect 对象的变量。

### AutoCorrect.ReplaceText

如果“自动更正”替换内容将被自动替换，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ReplaceText**

\*express是一个代表 AutoCorrect 对象的变量。

**示例**

``` JavaScript
/*本示例打开自动文本替换功能。*/
function test()
{
  let correct = Application.AutoCorrect
  correct.CapitalizeNamesOfDays = true
  correct.ReplaceText = false
}
```

### AutoCorrect.ReplacementList

返回“自动更正”替换内容的数组。

**语法**

**express.ReplacementList**

\*express是一个代表 AutoCorrect 对象的变量。

**说明**

参数

| **名称** | **必选/可选** | **数据类型** | **说明**                                                                                                                       |
|----------|---------------|--------------|--------------------------------------------------------------------------------------------------------------------------------|
| *Index*  | 可选          | **Variant**  | 要返回的“自动更正”替换内容数组的行索引。该行以二元一维数组的形式返回：第一个元素为第一列中的文本，第二个元素为第二列中的文本。 |

如果未指定 *Index* ，则本方法将返回一个二维数组。数组中每一行包含一对自动更正替换文本，如下表所示。

| 列  | 内容           |
|-----|----------------|
| 1   | 要被替换的文本 |
| 2   | 替换文本       |

使用 AddReplacement 方法可向替换列表添加自动更正项。

**示例**

``` JavaScript
/*本示例在替换列表中搜索“Temperature”的替换条目，如果存在，则显示该替换条目。*/
function test()
{
  let repl = Application.AutoCorrect.ReplacementList
  
  for(let i = 1; i <= repl.Count; i++){
      if(repl(i,1) == "Temperature")
      {
        alert(i,2)
      }
  }
}
```

### AutoCorrect.TwoInitialCapitals

如果自动更正以两个大写字母开头的词，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.TwoInitialCapitals**

\*express是一个代表 AutoCorrect 对象的变量。

**示例**

``` JavaScript
/*本示例设置 ET 自动更正以两个大写字母开头的词。*/
function test()
{
  let correct = Application.AutoCorrect
  correct.TwoInitialCapitals = true
  correct.ReplaceText = true
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/AutoCorrect](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Border 对象

## 简介

代表对象的边框。

## 说明

大多数具有边框的对象（除 Range 和 Style 对象外）都将边框作为单一实体处理，而不管边框有几个边。整个边框必须作为一个整体单位返回。例如，使用 TrendLine 对象的 Border 属性可返回此类对象的 Border 对象。

下例更改活动图表中趋势线的类型和线型。

``` JavaScript
function test() {  
   let trendlines = Application.ActiveChart.SeriesCollection(1).Trendlines(1)  
   trendlines.Type = xlLinear  trendlines.Border.LineStyle = xlDash
} 
```

Range 和 Style 对象具有四个分立的边框：左边框、右边框、顶部边框和底部边框，这四个边框既可单独返回，也可作为一个组同时返回。使用 Borders 属性可返回 Borders 集合，该集合包含所有四个边框，并将这些边框视为一个单位。下例向第一张工作表上的单元格 A1 添加双边框。

``` JavaScript
Application.Worksheets.Item(1).Range("A1").Borders.LineStyle = xlDouble
```

使用 Borders ( *index* )（其中 *index* 指定边框）可返回单个 Border 对象。下例设置单元格区域 A1:G1 的底部边框的颜色。

``` JavaScript
Application.Worksheets.Item("Sheet1").Range("A1:G1").Borders(xlEdgeBottom).Color = (255, 0, 0)
```

*Index* 可为以下 XlBordersIndex 常量之一： xlDiagonalDown 、 xlDiagonalUp 、 xlEdgeBottom 、 xlEdgeLeft 、 xlEdgeRight 、 xlEdgeTop 、 xlInsideHorizontal 或 xlInsideVertical 。

## 属性

| 序号 | 名称                                 | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Border.Application)   | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Color](#Border.Color)               | 返回或设置对象的主要颜色，如注释部分中的表格所示 。 Variant 型，可读写。                                                                                                                                                        |
|  3   | [ColorIndex](#Border.ColorIndex)     | 返回或设置一个 Variant 值，它代表边框的颜色。                                                                                                                                                                                   |
|  4   | [Creator](#Border.Creator)           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [LineStyle](#Border.LineStyle)       | 返回或设置边框的线型。 XlLineStyle 、 xlGray25 、 xlGray50 、 xlGray75 或 xlAutomatic 类型，可读写。                                                                                                                            |
|  6   | [Parent](#Border.Parent)             | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  7   | [ThemeColor](#Border.ThemeColor)     | 返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 Variant 类型。                                                                                                                                      |
|  8   | [TintAndShade](#Border.TintAndShade) | 返回或设置一个 Single ，使颜色变深或变浅。                                                                                                                                                                                      |
|  9   | [Weight](#Border.Weight)             | 返回或设置一个 XlBorderWeight 值，它代表边框的粗细。                                                                                                                                                                            |

## 成员属性

### Border.Application

如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Border 对象的变量。

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

### Border.Color

返回或设置对象的主要颜色，如注释部分中的表格所示 。 **Variant** 型，可读写。

**语法**

**express.Color**

\*express是一个代表 Border 对象的变量。

**说明**

| 对象         | 对应颜色                                                                            |
|--------------|-------------------------------------------------------------------------------------|
| **边框**     | 边框的颜色。                                                                        |
| **Borders**  | 一个区域的所有四条边的颜色。如果四边不是同一种颜色，则 **Color** 返回的是 0（零）。 |
| **Font**     | 字体的颜色。                                                                        |
| **Interior** | 单元格底纹的颜色或图形对象的填充颜色。                                              |
| **Tab**      | 选项卡的颜色。                                                                      |

**示例**

``` JavaScript
/* 此示例对 Chart1 中数值坐标轴的刻度线标志颜色进行设置*/
Application.Charts.Item("Chart1").Axes(xlValue).TickLabels.Font.Color = 0x0000ff00
```

### Border.ColorIndex

返回或设置一个 **Variant** 值，它代表边框的颜色。

**语法**

**express.ColorIndex**

\*express是一个代表 Border 对象的变量。

**说明**

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 XlColorIndex 常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

**示例**

``` JavaScript
/* 本示例为 Chart1 中数值坐标轴的主要网格线设置了颜色*/
function test() {
  let axes = Application.Charts.Item("Chart1").Axes(xlValue)
  if(axes.HasMajorGridlines){
    axes.MajorGridlines.Border.ColorIndex = 5    //set color to blue
  }
}
```

### Border.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Border 对象的变量。

### Border.LineStyle

返回或设置边框的线型。 **XlLineStyle** 、 **xlGray25** 、 **xlGray50** 、 **xlGray75** 或 **xlAutomatic** 类型，可读写。

**语法**

**express.LineStyle**

\*express是一个代表 Border 对象的变量。

**说明**

**xlDouble** 和 **xlSlantDashDot** 不适用于图表。

**示例**

``` JavaScript
/* 本示例为 Chart1 的图表区和绘图区域设置边框*/
function test() {
    let charts = Application.Charts.Item("Chart1")
    charts.ChartArea.Border.LineStyle = xlDashDot
    let border = charts.PlotArea.Border
    border.LineStyle = xlDashDotDot
    border.Weight = xlThick
}
```

### Border.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Border 对象的变量。

### Border.ThemeColor

返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。

**语法**

**express.ThemeColor**

\*express是一个代表 Border 对象的变量。

**说明**

如果对象当前未应用主题颜色，试图访问其主题颜色则会引起无效请求运行时错误。

### Border.TintAndShade

返回或设置一个 **Single** ，使颜色变深或变浅。

**语法**

**express.TintAndShade**

\*express是一个代表 Border 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

### Border.Weight

返回或设置一个 **XlBorderWeight** 值，它代表边框的粗细。

**语法**

**express.Weight**

\*express是一个代表 Border 对象的变量。

**示例**

``` JavaScript
/* 本示例对 Sheet1 上椭圆一的边框粗细进行设置*/
Application.Worksheets.Item("Sheet1").Ovals(1).Border.Weight = xlMedium
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Border](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CalculatedFields 对象

## 简介

由 PivotField 对象组成的集合，这些对象代表指定数据透视表中的所有计算字段。

## 方法

| 序号 | 名称                           | 说明                                     |
|:----:|:-------------------------------|:-----------------------------------------|
|  1   | [Add](#CalculatedFields.Add)   | 创建新的计算字段。返回 PivotField 对象。 |
|  2   | [Item](#CalculatedFields.Item) | 从集合中返回一个对象。                   |

## 属性

| 序号 | 名称                                         | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CalculatedFields.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#CalculatedFields.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#CalculatedFields.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#CalculatedFields.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### CalculatedFields.Add

创建新的计算字段。返回 **PivotField** 对象。

**语法**

**express.Add ( *Name* , *Formula* , *UseStandardFormula* )**

\*express是一个代表 CalculatedFields 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                                                                                                                                    |
|----------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Name*               | 必选      | String   | 字段的名称。                                                                                                                                            |
| *Formula*            | 必选      | String   | 字段的公式。                                                                                                                                            |
| *UseStandardFormula* | 可选      | Variant  | 为实现向上兼容，该值默认为 False。对于作为字段名称的任何参数中所含的字符串来说，该值为 True，并将被解释为已按标准美国英语进行了格式设置，而非本地设置。 |

**说明**

返回值 一个代表新计算字段的 **PivotField** 。

**示例**

``` JavaScript
/*此示例向第一张工作表上的第一张数据透视表添加计算字段。*/
Application.Worksheets.Item(1).PivotTables(1).CalculatedFields().Add("PxS", "= Product * Sales")
```

### CalculatedFields.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 CalculatedFields 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**说明**

返回值 包含在集合中的一个 **PivotField** 对象。

**示例**

``` JavaScript
/*此示例为第一个计算字段设置公式。*/
Application.Worksheets.Item(1).PivotTables(1).CalculatedFields().Item(1).Formula = "=Revenue - Cost"
```

## 成员属性

### CalculatedFields.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CalculatedFields 对象的变量。

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

### CalculatedFields.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 CalculatedFields 对象的变量。

### CalculatedFields.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CalculatedFields 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### CalculatedFields.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CalculatedFields 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/CalculatedFields](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ChartGroup 对象

## 简介

代表图表中用同一格式绘制的一个或多个数据系列。

## 说明

一张图表包含一个或多个图表组，每个图表组包含一个或多个 Series 对象，每个数据系列包含一个或多个 Points 对象。例如，单张图表可能既包含折线图图表组（其中包含所有用折线图格式绘制的数据系列），也包含条形图图表组（其中包含所有用条形图格式绘制的数据系列）。 ChartGroup 对象是 ChartGroups 集合的成员。

使用 ChartGroups ( *index* )（其中 *index* 是图表组的索引号）可以返回单个 ChartGroup 对象。

因为当特定图表组所用的图表格式更改时，该图表组的索引号可能会更改，所以使用命名图表组快捷方法之一来返回特定的图表组会更加容易。 PieGroups 方法返回图表中饼图图表组的集合， LineGroups 方法返回图表中折线图图表组的集合，依此类推。这些方法中的每一个都可以与索引号配合使用以返回单个 ChartGroup 对象，或不指定索引号而返回 ChartGroups 集合。

示例

以下示例向图表工作表 1 上的图表组 1 中添加垂直线。

``` JavaScript
Application.Charts.Item(1).ChartGroups.Item(1).HasDropLines = true
```

如果图表已被激活，就可使用 ActiveChart 属性。

``` JavaScript
Application.Charts.Item(1).Activate()
Application.ActiveChart.ChartGroups(1).HasDropLines = true
```

## 方法

| 序号 | 名称                                             | 说明                                                                                                     |
|:----:|:-------------------------------------------------|:---------------------------------------------------------------------------------------------------------|
|  1   | [SeriesCollection](#ChartGroup.SeriesCollection) | 返回一个对象，它代表图表或图表组中的一个系列（ Series 对象）或所有系列的集合（ SeriesCollection 集合）。 |

## 属性

| 序号 | 名称                                                   | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ChartGroup.Application)                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [AxisGroup](#ChartGroup.AxisGroup)                     | 返回或设置指定图表的组。可读/写。                                                                                                                                                                                               |
|  3   | [BubbleScale](#ChartGroup.BubbleScale)                 | 返回或设置指定图表组中气泡的缩放比例。可为从 0（零）到 300 之间的整数，表示相对于默认大小的百分率。仅适用于气泡图。 Long 类型，可读写。                                                                                         |
|  4   | [Creator](#ChartGroup.Creator)                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [DoughnutHoleSize](#ChartGroup.DoughnutHoleSize)       | 返回或设置圆环图图表组的内径大小。内径大小以图表大小的百分比表示，有效取值范围为 10% 到 90%。 Long 类型，可读写。                                                                                                               |
|  6   | [DownBars](#ChartGroup.DownBars)                       | 返回一个 DownBars 对象，该对象表示折线图中的跌柱线。仅应用于折线图。只读。                                                                                                                                                      |
|  7   | [DropLines](#ChartGroup.DropLines)                     | 返回一个 DropLines 对象，该对象表示折线图或面积图上数据系列的垂直线。仅适用于折线图或面积图。只读。                                                                                                                             |
|  8   | [FirstSliceAngle](#ChartGroup.FirstSliceAngle)         | 以度为单位返回或设置饼图或圆环图的第一扇区的角度（从垂直方向顺时针计算）。仅应用于饼图、三维饼图和圆环图。值的范围从 0 到 360。 Long 类型，可读写。                                                                             |
|  9   | [GapWidth](#ChartGroup.GapWidth)                       | 对于条形图和柱形图：以条形或柱形宽度百分数的形式返回或设置条形簇或柱形簇之间的距离。对于饼图中的扇形和饼图中的条形：返回或设置图表的第一节和第二节之间的间距。 Long 类型，可读写。                                              |
|  10  | [Has3DShading](#ChartGroup.Has3DShading)               | 如果图表组具有三维阴影，则该值为 True 。本属性仅适合曲面图，如果要将其设置给非曲面图，则将返回运行时错误。 Boolean 类型，可读写。                                                                                               |
|  11  | [HasDropLines](#ChartGroup.HasDropLines)               | 如果折线图或者面积图中有垂直线，则该值为 True 。仅应用于折线图和面积图。 Boolean 类型，可读写。                                                                                                                                 |
|  12  | [HasHiLoLines](#ChartGroup.HasHiLoLines)               | 如果折线图有高低点连线，则为 True 。仅应用于折线图。 Boolean 类型，可读写。                                                                                                                                                     |
|  13  | [Index](#ChartGroup.Index)                             | 返回 Long 值，它代表对象在其同类对象所组成的集合内的索引号。                                                                                                                                                                    |
|  14  | [Overlap](#ChartGroup.Overlap)                         | 指定放置条形和柱形的方式。取值范围在 -100 到 100 之间。仅应用于二维条形图和二维柱形图。 Long 类型，可读写。                                                                                                                     |
|  15  | [Parent](#ChartGroup.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  16  | [RadarAxisLabels](#ChartGroup.RadarAxisLabels)         | 返回一个 TickLabels 对象，该对象表示指定图表组的雷达图坐标轴标签。只读。                                                                                                                                                        |
|  17  | [SecondPlotSize](#ChartGroup.SecondPlotSize)           | 以主饼图大小的百分比形式返回或设置复合饼图或复合条饼图中第二部分的大小。可为 5 到 200 之间的值。 Long 类型，可读写。                                                                                                            |
|  18  | [SeriesLines](#ChartGroup.SeriesLines)                 |                                                                                                                                                                                                                                 |
|  19  | [ShowNegativeBubbles](#ChartGroup.ShowNegativeBubbles) | 如果在图表组中显示表示负值的气泡，则该值为 True 。仅对气泡图有效。 Boolean 类型，可读写。                                                                                                                                       |
|  20  | [SizeRepresents](#ChartGroup.SizeRepresents)           | 返回或设置气泡图中气泡的大小所表示的含义。可为以下 XlSizeRepresents 常量之一： xlSizeIsArea 或 xlSizeIsWidth 。 Long 类型，可读写。                                                                                             |
|  21  | [SplitType](#ChartGroup.SplitType)                     | 返回或设置在复合饼图或复合条饼图中拆分两部分的方式。 XlChartSplitType 类型，可读写。                                                                                                                                            |
|  22  | [SplitValue](#ChartGroup.SplitValue)                   | 返回或设置复合饼图或复合条饼图中分隔两部分的阈值。可读/写 Variant 类型。                                                                                                                                                        |
|  23  | [UpBars](#ChartGroup.UpBars)                           | 返回 UpBars 对象，该对象表示折线图上的涨柱线。仅应用于折线图。只读。                                                                                                                                                            |
|  24  | [VaryByCategories](#ChartGroup.VaryByCategories)       | 如果 ET 对每个数据标志指定不同的颜色或模式，则该值为 True 。图表中必须只包含一个数据系列。 Boolean 类型，可读写。                                                                                                               |

## 成员方法

### ChartGroup.SeriesCollection

返回一个对象，它代表图表或图表组中的一个系列（ **Series** 对象）或所有系列的集合（ **SeriesCollection** 集合）。

**语法**

**express.SeriesCollection ( *Index* )**

\*express是一个代表 ChartGroup 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                   |
|---------|-----------|----------|------------------------|
| *Index* | 可选      | Variant  | 数据系列的名称或编号。 |

**示例**

``` JavaScript
/*此示例打开 Chart1 中第一个数据系列的数据标签。*/
Application.Charts.Item("Chart1").SeriesCollection(1).HasDataLabels = true
```

## 成员属性

### ChartGroup.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    if(ActiveWorkbook.Application.Value == "ET"){
        alert("This is an ET Application object.")
    }
    else{
        alert("This is not an ET Application object.")
    }
}
```

### ChartGroup.AxisGroup

返回或设置指定图表的组。可读/写。

**语法**

**express.AxisGroup**

\*express是一个代表 ChartGroup 对象的变量。

**说明**

对于三维图表，仅 **xlPrimary** 有效。

**示例**

``` JavaScript
/*此示例在值轴处于辅助组时将其删除。*/
function test()
{
    if(myChart.Axes.Item(xlValue).AxisGroup == xlSecondary){
        myChart.Axes.Item(xlValue).Delete()
    }
}
```

### ChartGroup.BubbleScale

返回或设置指定图表组中气泡的缩放比例。可为从 0（零）到 300 之间的整数，表示相对于默认大小的百分率。仅适用于气泡图。 **Long** 类型，可读写。

**语法**

**express.BubbleScale**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例将第一个图表组中的气泡大小设置为默认大小的 200%。*/
function test()
{
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    chart.ChartGroups(1).BubbleScale = 200
}
```

### ChartGroup.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ChartGroup 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ChartGroup.DoughnutHoleSize

返回或设置圆环图图表组的内径大小。内径大小以图表大小的百分比表示，有效取值范围为 10% 到 90%。 **Long** 类型，可读写。

**语法**

**express.DoughnutHoleSize**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例对 Chart1 中第一个圆环图组的内径大小进行设置。本示例应在二维圆环图上运行。*/
Application.Charts.Item("Chart1").DoughnutGroups(1).DoughnutHoleSize = 10
```

### ChartGroup.DownBars

返回一个 **DownBars** 对象，该对象表示折线图中的跌柱线。仅应用于折线图。只读。

**语法**

**express.DownBars**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例显示 Chart1 的涨跌柱线，并对其颜色进行设置。本示例应在包含两组彼此间有一个或多个相交数据点的数据系列的二维折线图上运行。*/
function test()
{
    let chartgroup = Application.Charts.Item("Chart1").ChartGroups(1)
    chartgroup.HasUpDownBars = true
    chartgroup.DownBars.Interior.ColorIndex = 3
    chartgroup.UpBars.Interior.ColorIndex = 5
}
```

### ChartGroup.DropLines

返回一个 **DropLines** 对象，该对象表示折线图或面积图上数据系列的垂直线。仅适用于折线图或面积图。只读。

**语法**

**express.DropLines**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例显示 Chart1 中第一个图表组的垂直线，并对其线型、粗细和颜色进行设置。本示例应在包含一个数据系列的二维折线图上运行。*/
function test()
{
    let chartgroup = Application.Charts.Item("Chart1").ChartGroups(1)
    chartgroup.HasDropLines = true
    let border = chartgroup.DropLines.Border
        border.LineStyle = xlThin
        border.Weight = xlMedium
        border.ColorIndex = 3
}
```

### ChartGroup.FirstSliceAngle

以度为单位返回或设置饼图或圆环图的第一扇区的角度（从垂直方向顺时针计算）。仅应用于饼图、三维饼图和圆环图。值的范围从 0 到 360。 **Long** 类型，可读写。

**语法**

**express.FirstSliceAngle**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例设置 Chart1 中的第一个图表组中的第一个扇区的角度。本示例应在二维饼图上运行。*/
Application.Charts.Item("Chart1").ChartGroups(1).FirstSliceAngle = 15
```

### ChartGroup.GapWidth

对于条形图和柱形图：以条形或柱形宽度百分数的形式返回或设置条形簇或柱形簇之间的距离。对于饼图中的扇形和饼图中的条形：返回或设置图表的第一节和第二节之间的间距。 **Long** 类型，可读写。

**语法**

**express.GapWidth**

\*express是一个代表 ChartGroup 对象的变量。

**说明**

该属性的取值必须介于 0 到 500 之间。

**示例**

``` JavaScript
/*本示例将 Chart1 的列簇间距设置为列宽的 50%。*/
Application.Charts.Item("Chart1").ChartGroups(1).GapWidth = 50
```

### ChartGroup.Has3DShading

如果图表组具有三维阴影，则该值为 **True** 。本属性仅适合曲面图，如果要将其设置给非曲面图，则将返回运行时错误。 **Boolean** 类型，可读写。

**语法**

**express.Has3DShading**

\*express是一个代表 ChartGroup 对象的变量。

**说明**

注释：Setting Has3DShading = False has the effect of setting the chart material to "Flat", or removing the effect, whereas setting the Has3DShading = True sets the chart content to the default.

**示例**

``` JavaScript
/*本示例为图表组添加三维阴影。*/
Application.Charts.Item(1).ChartGroups(1).Has3DShading = true
```

### ChartGroup.HasDropLines

如果折线图或者面积图中有垂直线，则该值为 **True** 。仅应用于折线图和面积图。 **Boolean** 类型，可读写。

**语法**

**express.HasDropLines**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例显示 Chart1 中第一个图表组的垂直线，并对其线型、粗细和颜色进行设置。本示例应在包含一个数据系列的二维折线图上运行。*/
function test()
{
    let chartgroup = Application.Charts.Item("Chart1").ChartGroups(1)
    chartgroup.HasDropLines = true
    let border = chartgroup.DropLines.Border
        border.LineStyle = xlThin
        border.Weight = xlMedium
        border.ColorIndex = 3
}
```

### ChartGroup.HasHiLoLines

如果折线图有高低点连线，则为 **True** 。仅应用于折线图。 **Boolean** 类型，可读写。

**语法**

**express.HasHiLoLines**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例显示 Chart1 中第一个图表组的高低线，并对其线型、粗细和颜色进行设置。本示例应在具有三组类似股市行情（盘高－盘低－收盘）数据系列的二维折线图上运行。*/
function test()
{
    let chartgroup = Application.Charts.Item("Chart1").ChartGroups(1)
    chartgroup.HasHiLoLines = true
    let border = chartgroup.HiLoLines.Border
        border.LineStyle = xlThin
        border.Weight = xlMedium
        border.ColorIndex = 3
}
```

### ChartGroup.Index

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 ChartGroup 对象的变量。

### ChartGroup.Overlap

指定放置条形和柱形的方式。取值范围在 -100 到 100 之间。仅应用于二维条形图和二维柱形图。 **Long** 类型，可读写。

**语法**

**express.Overlap**

\*express是一个代表 ChartGroup 对象的变量。

**说明**

如果将该属性设置为 -100，条形之间有一个条形的间距。如果重叠比例为 0（零），条形之间没有间距（一个条形紧挨着另一条形）。如果重叠比例为 100，所有条形将完全重叠。

**示例**

``` JavaScript
/*本示例将第一个图表组的重叠比例设置为 -50。本示例应在包含两个或更多系列的二维柱形图上运行。*/
Application.Charts.Item("Chart1").ChartGroups(1).Overlap = -50
```

### ChartGroup.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ChartGroup 对象的变量。

### ChartGroup.RadarAxisLabels

返回一个 **TickLabels** 对象，该对象表示指定图表组的雷达图坐标轴标签。只读。

**语法**

**express.RadarAxisLabels**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例显示 Chart1 中第一个图表组的雷达图坐标轴标注，并对这些标注的颜色进行设置。本示例应在雷达图上运行。*/
function test()
{
    let chartgroup = Application.Charts.Item("Chart1").ChartGroups(1)
    chartgroup.HasRadarAxisLabels = true
    chartgroup.RadarAxisLabels.Font.ColorIndex = 3
}
```

### ChartGroup.SecondPlotSize

以主饼图大小的百分比形式返回或设置复合饼图或复合条饼图中第二部分的大小。可为 5 到 200 之间的值。 **Long** 类型，可读写。

**语法**

**express.SecondPlotSize**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例必须在复合饼图或复合条饼图中运行。本示例以数值拆分图表的两部分，在主饼图部分中合并所有小于 10 的数值，并在第二部分中显示这些数值。第二部分的大小为主条形图大小的百分之五十。*/
function test()
{
    let chartobjects = Application.Worksheets.Item(1).ChartObjects(1).Chart.ChartGroups(1)
    chartobjects.SplitType = xlSplitByValue
    chartobjects.SplitValue = 10
    chartobjects.VaryByCategories = true
    chartobjects.SecondPlotSize = 50
}
```

### ChartGroup.SeriesLines

**语法**

**express.SeriesLines**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例显示 Chart1 中图表组一的系列线，并对其线型、粗细和颜色进行设置。本示例应在有两个或多个数据系列的二维堆积柱形图上运行。*/
function test()
{
    let myChartGroups = Application.Charts.Item("Chart1").ChartGroups(1)
    myChartGroups.HasSeriesLines = true
    let myBorder = myChartGroups.SeriesLines.Border
        myBorder.LineStyle = xlThin
        myBorder.Weight = xlMedium
        myBorder.ColorIndex = 3
}
```

### ChartGroup.ShowNegativeBubbles

如果在图表组中显示表示负值的气泡，则该值为 **True** 。仅对气泡图有效。 **Boolean** 类型，可读写。

**语法**

**express.ShowNegativeBubbles**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
Application.Worksheets.Item(1).ChartObjects(1).Chart.ChartGroups(1).ShowNegativeBubbles = true
```

### ChartGroup.SizeRepresents

返回或设置气泡图中气泡的大小所表示的含义。可为以下 **XlSizeRepresents** 常量之一： **xlSizeIsArea** 或 **xlSizeIsWidth** 。 **Long** 类型，可读写。

**语法**

**express.SizeRepresents**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例设置图表组一中气泡的大小所表示的含义。*/
Application.Charts.Item(1).ChartGroups(1).SizeRepresents = xlSizeIsWidth
```

### ChartGroup.SplitType

返回或设置在复合饼图或复合条饼图中拆分两部分的方式。 **XlChartSplitType** 类型，可读写。

**语法**

**express.SplitType**

\*express是一个代表 ChartGroup 对象的变量。

**说明**

|                                                               |
|---------------------------------------------------------------|
| **XlChartSplitType** 可为以下 **XlChartSplitType** 常量之一。 |
| **xlSplitByCustomSplit**                                      |
| **xlSplitByPercentValue**                                     |
| **xlSplitByPosition**                                         |
| **xlSplitByValue**                                            |

**示例**

``` JavaScript
/*本示例必须在复合饼图或复合条饼图中运行。本示例以数值拆分图表的两部分，在主饼图部分中合并所有小于 10 的数值，并在第二部分中显示这些数值。*/
function test()
{
    let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart.ChartGroups(1)
    myChart.SplitType = xlSplitByValue
    myChart.SplitValue = 10
    myChart.VaryByCategories = true
}
```

### ChartGroup.SplitValue

返回或设置复合饼图或复合条饼图中分隔两部分的阈值。可读/写 **Variant** 类型。

**语法**

**express.SplitValue**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例必须在复合饼图或复合条饼图中运行。本示例以数值拆分图表的两部分，在主饼图部分中合并所有小于 10 的数值，并在第二部分中显示这些数值。*/
function test()
{
    let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart.ChartGroups(1)
    myChart.SplitType = xlSplitByValue
    myChart.SplitValue = 10
    myChart.VaryByCategories = true
}
```

### ChartGroup.UpBars

返回 **UpBars** 对象，该对象表示折线图上的涨柱线。仅应用于折线图。只读。

**语法**

**express.UpBars**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例显示 Chart1 中图表组一的涨跌柱线，并对其颜色进行设置。本示例应在二维折线图上运行，该折线图应包含两组彼此间有一个或多个相交数据点的系列。*/
function test()
{
    let myChart = Application.Charts.Item("Chart1").ChartGroups(1)
    myChart.HasUpDownBars = True
    myChart.DownBars.Interior.ColorIndex = 3
    myChart.UpBars.Interior.ColorIndex = 5
}
```

### ChartGroup.VaryByCategories

如果 ET 对每个数据标志指定不同的颜色或模式，则该值为 **True** 。图表中必须只包含一个数据系列。 **Boolean** 类型，可读写。

**语法**

**express.VaryByCategories**

\*express是一个代表 ChartGroup 对象的变量。

**示例**

``` JavaScript
/*本示例对第一个图表组中的每个数据标志指定一种不同的颜色或图案。本示例应在数据系列上有数据标志的二维折线图上运行。*/
Application.Charts.Item("Chart1").ChartGroups(1).VaryByCategories = true
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ChartGroup](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ConnectorFormat 对象

## 简介

包含应用于连接符的属性和方法。

## 说明

连接符是用于连接其他两个形状的线，所连接的位置叫做连接结点。如果重新排列已连接的形状，那么连接符的几何形状将自动调整，以使重新排列的形状仍保持连接。

连接结点通常按下表所示的规则进行编号。

| 形状类型                          | 连接结点标号方案                     |
|-----------------------------------|--------------------------------------|
| 自选形状、艺术字、图片和 OLE 对象 | 连接结点从顶部开始按逆时针进行编号。 |
| 任意多边形                        | 连接结点为顶点，与顶点编号相对应。   |

使用 ConnectorFormat 属性可返回 ConnectorFormat 对象。使用 BeginConnect 和 EndConnect 方法可将连接符的两端连到文档中的其他形状。使用 RerouteConnections 方法可自动查找通过连接符连接的两个形状间的最短路径。使用 Connector 属性可查看形状是否为连接符。

``` JavaScript
function test(){
let mainshape = Application.ActiveWindow.Selection.ShapeRange.Item(1)
    let bx = mainshape.Left + mainshape.Width + 50
    let by = mainshape.Top + mainshape.Height + 50

let mCount = mainshape.ConnectionSiteCount
    for(let j = 1;j <= mCount;j++) { 
        let myConnector = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 
                bx, by, bx + 50, by + 50)
            myConnector.ConnectorFormat.EndConnect(mainshape, j)
            myConnector.ConnectorFormat.Type = msoConnectorElbow
            myConnector.Line.ForeColor.RGB = (255, 0, 0)
            let l = myConnector.Left
            let t = myConnector.Top
       
        let myTextbox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 
                l, t, 36, 14)
            myTextbox.Fill.Visible = false
            myTextbox.Line.Visible = false
            myTextbox.TextFrame.Characters().Text = j 
    }
}
```

``` JavaScript
/*下例向 myDocument 中添加两个矩形并且用曲线连接符连接矩形。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let s = myDocument.Shapes
let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
let c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0)
    c.ConnectorFormat.BeginConnect(firstRect, 1)
    c.ConnectorFormat.EndConnect(secondRect, 1)
    c.RerouteConnections()
}
```

## 方法

| 序号 | 名称                                                | 说明                                                                                                                                                                                                                                                                               |
|:----:|:----------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [BeginConnect](#ConnectorFormat.BeginConnect)       | 将指定的连接符的起点连接到指定的形状上。如果在连接符的起点与其他形状之间已经有了连接，那么该已有的连接将中断。如果连接符的起点不在所需的连接结点上，本方法将把连接符的起点移到该连接结点，并对连接符的大小和位置作相应的调整。可用 EndConnect 方法将连接符的终点连接到某一形状上。 |
|  2   | [BeginDisconnect](#ConnectorFormat.BeginDisconnect) | 使指定的连接符的起点与其所连接的形状脱离。本方法并不修改连接符的尺寸和位置：连接符的起点仍保留在原来所连接的连接结点的位置，但与该连接结点之间不再有连接。可用 EndDisconnect 方法使连接符的终点与某一形状脱离。                                                                    |
|  3   | [EndConnect](#ConnectorFormat.EndConnect)           | 将指定的连接符的终点连接到指定的形状上。如果在连接符的终点与其他形状之间已经有了连接，那么该已有的连接将中断。如果连接符的终点不在所需的连接结点，本方法将把连接符的终点移到该连接结点，并对连接符的大小和位置作相应的调整。可用 BeginConnect 方法将连接符的起点连接到某一形状上。 |
|  4   | [EndDisconnect](#ConnectorFormat.EndDisconnect)     | 使指定的连接符的终点与其所连接的形状脱离。本方法并不更改连接符的尺寸和位置：连接符的终点仍保留在原来所连接的连接网站的位置，但与该连接网站之间不再有连接。可以使用 BeginDisconnect 方法使连接符的起点与某一形状分离。                                                              |

## 属性

| 序号 | 名称                                                        | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ConnectorFormat.Application)                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [BeginConnected](#ConnectorFormat.BeginConnected)           | 如果指定的连接符的起点已连接到了某一形状上，则该属性值为 True 。 MsoTriState 类型，只读。                                                                                                                                       |
|  3   | [BeginConnectedShape](#ConnectorFormat.BeginConnectedShape) | 返回一个 Shape 对象，该对象代表指定连接符的起点所连到的形状。只读。                                                                                                                                                             |
|  4   | [BeginConnectionSite](#ConnectorFormat.BeginConnectionSite) | 返回一个整数，该整数指定连接符起点的连接结点。 Long 类型，只读。                                                                                                                                                                |
|  5   | [Creator](#ConnectorFormat.Creator)                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  6   | [EndConnected](#ConnectorFormat.EndConnected)               | 如果该属性值为 msoTrue ，则指定的连接符的端点已连接到了某一形状上。 MsoTriState 类型，只读。                                                                                                                                    |
|  7   | [EndConnectedShape](#ConnectorFormat.EndConnectedShape)     | 返回一个 Shape 对象，该对象表示指定连接符的终点所连到的形状。只读。                                                                                                                                                             |
|  8   | [EndConnectionSite](#ConnectorFormat.EndConnectionSite)     | 返回一个整数，该整数指定连接符终点的连接结点。 Long 类型，只读。                                                                                                                                                                |
|  9   | [Parent](#ConnectorFormat.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  10  | [Type](#ConnectorFormat.Type)                               | 返回或设置一个 MsoConnectorType 值，它代表连接符格式类型。                                                                                                                                                                      |

## 成员方法

### ConnectorFormat.BeginConnect

将指定的连接符的起点连接到指定的形状上。如果在连接符的起点与其他形状之间已经有了连接，那么该已有的连接将中断。如果连接符的起点不在所需的连接结点上，本方法将把连接符的起点移到该连接结点，并对连接符的大小和位置作相应的调整。可用 **EndConnect** 方法将连接符的终点连接到某一形状上。

**语法**

**express.BeginConnect ( *ConnectedShape* , *ConnectionSite* )**

\*express是一个代表 ConnectorFormat 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                     |
|------------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ConnectedShape* | 必选      | Shape    | 要连接到连接符起点的形状。指定的 Shape 对象必须与连接符在同一个 Shapes 集合中。                                                                                                                                                                                          |
| *ConnectionSite* | 必选      | Long     | 由 ConnectedShape 指定的形状上的连接结点。必须是一个介于 1 和指定形状的 ConnectionSiteCount 属性的返回值之间的整数。如果希望连接符自动查找它所连接的两个形状间的最短路径，请为此参数指定任何一个有效整数，并且在连接符的两端连接到形状之后使用 RerouteConnections 方法。 |

**说明**

将连接符连接到某个对象以后，该连接符的大小和位置将在必要时进行自动调整。

**示例**

本示例向 ` myDocument ` 中添加了两个矩形，并用曲线连接符将这两个矩形连接起来。请注意，对 **RerouteConnections** 方法的调用使得在 **BeginConnect** 方法和 **EndConnect** 方法中所指定的 *ConnectionSite* 参数值不相关联。

``` JavaScript
function test(){
let myDocument = Application.Worksheets.Item(1)
let s = myDocument.Shapes
let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
let c = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100)
    c.ConnectorFormat.BeginConnect(firstRect, 1)
    c.ConnectorFormat.EndConnect(secondRect, 1)
    c.RerouteConnections()
}
```

### ConnectorFormat.BeginDisconnect

使指定的连接符的起点与其所连接的形状脱离。本方法并不修改连接符的尺寸和位置：连接符的起点仍保留在原来所连接的连接结点的位置，但与该连接结点之间不再有连接。可用 **EndDisconnect** 方法使连接符的终点与某一形状脱离。

**语法**

**express.BeginDisconnect ()**

\*express是一个代表 ConnectorFormat 对象的变量。

**示例**

本示例向 ` myDocument ` 中添加两个矩形，用一个连接符连接这两个矩形，并自动将连接符路径修改为最短路径，然后断开矩形间的连接符。

``` JavaScript
function test(){
let myDocument = Application.Worksheets.Item(1)
let s = myDocument.Shapes
let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
let c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0)
    c.ConnectorFormat.BeginConnect(firstRect, 1)
    c.ConnectorFormat.EndConnect(secondRect, 1)
    c.RerouteConnections()
    c.ConnectorFormat.BeginDisconnect()
    c.ConnectorFormat.EndDisconnect()
}
```

### ConnectorFormat.EndConnect

将指定的连接符的终点连接到指定的形状上。如果在连接符的终点与其他形状之间已经有了连接，那么该已有的连接将中断。如果连接符的终点不在所需的连接结点，本方法将把连接符的终点移到该连接结点，并对连接符的大小和位置作相应的调整。可用 **BeginConnect** 方法将连接符的起点连接到某一形状上。

**语法**

**express.EndConnect ( *ConnectedShape* , *ConnectionSite* )**

\*express是一个代表 ConnectorFormat 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                       |
|------------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ConnectedShape* | 必选      | Shape    | 要连接到连接符终点的形状。指定的 Shape 对象必须与连接符在同一个 Shapes 集合中。                                                                                                                                            |
| *ConnectionSite* | 必选      | Long     | 是一个介于 1 和指定形状的 ConnectionSiteCount 属性的返回值之间的整数。如果希望连接符自动查找它所连接的两个形状间的最短路径，请为此参数指定任何一个有效整数，并且在连接符的两端连接到形状之后使用 RerouteConnections 方法。 |

**说明**

将连接符连接到某个对象以后，该连接符的大小和位置将在必要时进行自动调整。

**示例**

本示例向 ` myDocument ` 中添加了两个矩形，并用曲线连接符将这两个矩形连接起来。请注意，对 **RerouteConnections** 方法的调用使得在 **BeginConnect** 方法和 **EndConnect** 方法中所指定的 *ConnectionSite* 参数值不相关联。

``` JavaScript
function test(){
let myDocument = Application.Worksheets.Item(1)
let s = myDocument.Shapes
let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
let c = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100)
    c.ConnectorFormat.BeginConnect(firstRect, 1)
    c.ConnectorFormat.EndConnect(secondRect, 1)
    c.RerouteConnections()
}
```

### ConnectorFormat.EndDisconnect

使指定的连接符的终点与其所连接的形状脱离。本方法并不更改连接符的尺寸和位置：连接符的终点仍保留在原来所连接的连接网站的位置，但与该连接网站之间不再有连接。可以使用 **BeginDisconnect** 方法使连接符的起点与某一形状分离。

**语法**

**express.EndDisconnect ()**

\*express是一个代表 ConnectorFormat 对象的变量。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加两个矩形，用一个连接符连接这两个矩形，并自动将连接符路径修改为最短路径，然后断开矩形间的连接符。*/
function test(){

let myDocument = Application.Worksheets.Item(1)
let s = myDocument.Shapes
let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
let c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0)
    c.ConnectorFormat.BeginConnect(firstRect, 1)
    c.ConnectorFormat.EndConnect(secondRect, 1)
    c.RerouteConnections()
    c.ConnectorFormat.BeginDisconnect()
    c.ConnectorFormat.EndDisconnect()
}
```

## 成员属性

### ConnectorFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ConnectorFormat 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
let myObject = Application.ActiveWorkbook
if(myObject.Application.Value == "ET") {
    MsgBox("This is an ET Application object.")
}
else {
    MsgBox("This is not an ET Application object.")
}
}
```

### ConnectorFormat.BeginConnected

如果指定的连接符的起点已连接到了某一形状上，则该属性值为 **True** 。 **MsoTriState** 类型，只读。

**语法**

**express.BeginConnected**

\*express是一个代表 ConnectorFormat 对象的变量。

**说明**

|                                                       |
|-------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue**                                          |
| **msoFalse**                                          |
| **msoTriStateMixed**                                  |
| **msoTriStateToggle**                                 |
| **msoTrue** ：指定连接符的起点与形状相连。            |

**示例**

如果 ` myDocument ` 上的第三个形状是连接符，且它的起点已连接到了某一形状上，本示例将把连接结点的编号存储到变量 ` oldBeginConnSite ` 中，把对所连接的形状的引用存储到对象变量 ` oldBeginConnShape ` 中，然后断开连接符的起点与形状的连接。

``` JavaScript
function test(){
let myDocument = Application.Worksheets.Item(1)
let myShape = myDocument.Shapes.Item(3)
    if(myShape.Connector) {
        let myFormat = myShape.ConnectorFormat
            if(myFormat.BeginConnected) {
                let oldBeginConnSite = myFormat.BeginConnectionSite
                let oldBeginConnShape = myFormat.BeginConnectedShape
                myFormat.BeginDisconnect()
            }
    }
}
```

### ConnectorFormat.BeginConnectedShape

返回一个 **Shape** 对象，该对象代表指定连接符的起点所连到的形状。只读。

**语法**

**express.BeginConnectedShape**

\*express是一个代表 ConnectorFormat 对象的变量。

**说明**

如果指定的连接符的起点并未连接到任何形状上，那么本属性将产生错误。

**示例**

本示例假定在 ` myDocument ` 上，有两个用连接符“Conn1To2”连接起来的形状。本示例的代码将向 ` myDocument ` 添加一个矩形和一个连接符。新添加的连接符的起点将连接到“Conn1To2”的起点所连接的同一连接结点上，而新添加的连接符的终点则连接到新添加的矩形的第一个连接结点上。

``` JavaScript
function test(){
let myDocument = Application.Worksheets.Item(1)
let myShape = myDocument.Shapes
    let r3 = myShape.AddShape(msoShapeRectangle, 450, 190, 200, 100)
    myShape.AddConnector(msoConnectorCurve, 0, 0, 10, 10).Name = "Conn1To3"
    let connFormat1 = myShape.Item("Conn1To2").ConnectorFormat
        let beginConnSite1 = connFormat1.BeginConnectionSite
        let beginConnShape1 = connFormat1.BeginConnectedShape
    let connFormat2 = myShape.Item("Conn1To3").ConnectorFormat
        connFormat2.BeginConnect(beginConnShape1, beginConnSite1)
        connFormat2.EndConnect(r3, 1)
}
```

### ConnectorFormat.BeginConnectionSite

返回一个整数，该整数指定连接符起点的连接结点。 **Long** 类型，只读。

**语法**

**express.BeginConnectionSite**

\*express是一个代表 ConnectorFormat 对象的变量。

**说明**

如果指定的连接符的起点并未连接到任何形状上，那么本属性将产生错误。

**示例**

本示例假定在 ` myDocument ` 上，有两个用连接符“Conn1To2”连接起来的形状。本示例的代码将向 ` myDocument ` 添加一个矩形和一个连接符。新添加的连接符的起点将连接到“Conn1To2”的起点所连接的同一连接结点上，而新添加的连接符的终点则连接到新添加的矩形的第一个连接结点上。

``` JavaScript
function test(){
let myDocument = Application.Worksheets.Item(1)
let myShape = myDocument.Shapes
    let r3 = myShape.AddShape(msoShapeRectangle, 450, 190, 200, 100)
    myShape.AddConnector(msoConnectorCurve, 0, 0, 10, 10).Name = "Conn1To3"
    let connFormat1 = myShape.Item("Conn1To2").ConnectorFormat
        let beginConnSite1 = connFormat1.BeginConnectionSite
        let beginConnShape1 = connFormat1.BeginConnectedShape
    let connFormat2 = myShape.Item("Conn1To3").ConnectorFormat
        connFormat2.BeginConnect(beginConnShape1, beginConnSite1)
        connFormat2.EndConnect(r3, 1)
}
```

### ConnectorFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ConnectorFormat 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ConnectorFormat.EndConnected

如果该属性值为 **msoTrue** ，则指定的连接符的端点已连接到了某一形状上。 **MsoTriState** 类型，只读。

**语法**

**express.EndConnected**

\*express是一个代表 ConnectorFormat 对象的变量。

**说明**

|                                                       |
|-------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue** 不应用于该属性。                         |
| **msoFalse** 指定连接符的端点未连接到某一形状上。     |
| **msoTriStateMixed** 不应用于该属性。                 |
| **msoTriStateToggle** 不应用于该属性。                |
| **msoTrue** 指定连接符的端点连接到形状上。            |

**示例**

如果 ` myDocument ` 上的第三个形状是连接符，并且它的终点已连接到了某一形状上，则本示例将把连接结点的编号存储到变量 ` oldEndConnSite ` 中，把对所连接的形状的引用存储到对象变量 ` oldEndConnShape ` 中，然后断开连接符的终点与形状的连接。

``` JavaScript
function test(){
let myDocument = Application.Worksheets.Item(1)
let myShape = myDocument.Shapes.Item(3)
    if(myShape.Connector) {
        let myFormat = myShape.ConnectorFormat
            if(myFormat.EndConnected) {
                let oldEndConnSite = myFormat.EndConnectionSite
                let oldEndConnShape = myFormat.EndConnectedShape
                myFormat.EndDisconnect()
            }
    }
}
```

### ConnectorFormat.EndConnectedShape

返回一个 **Shape** 对象，该对象表示指定连接符的终点所连到的形状。只读。

**语法**

**express.EndConnectedShape**

\*express是一个代表 ConnectorFormat 对象的变量。

**说明**

如果指定的连接符的终点没有连接到一个形状上，则该属性产生一个错误。

**示例**

本示例假定在 ` myDocument ` 上，有两个用连接符“Conn1To2”连接起来的形状。本示例的代码将向 ` myDocument ` 添加一个矩形和一个连接符。新添加的连接符的终点将连接到“Conn1To2”的终点所连接的同一连接结点上，而新添加的连接符的起点则连接到新添加的矩形的第一个连接结点上。

``` JavaScript
function test(){
let myDocument = Application.Worksheets.Item(1)
let myShape = myDocument.Shapes
    let r3 = myShape.AddShape(msoShapeRectangle, 100, 420, 200, 100)
    let connFormat1 = myShape.Item("Conn1To2").ConnectorFormat
        let endConnSite1 = connFormat1.EndConnectionSite
        let endConnShape1 = connFormat1.EndConnectedShape
    let connFormat2 = myShape.AddConnector(msoConnectorCurve, 
            0, 0, 10, 10).ConnectorFormat
        connFormat2.BeginConnect(r3, 1)
        connFormat2.EndConnect(endConnShape1, endConnSite1)
}
```

### ConnectorFormat.EndConnectionSite

返回一个整数，该整数指定连接符终点的连接结点。 **Long** 类型，只读。

**语法**

**express.EndConnectionSite**

\*express是一个代表 ConnectorFormat 对象的变量。

**说明**

如果指定的连接符的终点没有连接到一个形状上，则该属性产生一个错误。

**示例**

本示例假定在 ` myDocument ` 上，有两个用连接符“Conn1To2”连接起来的形状。本示例的代码将向 ` myDocument ` 添加一个矩形和一个连接符。新添加的连接符的终点将连接到“Conn1To2”的终点所连接的同一连接结点上，而新添加的连接符的起点则连接到新添加的矩形的第一个连接结点上。

``` JavaScript
function test(){
let myDocument = Application.Worksheets.Item(1)
let myShape = myDocument.Shapes
    let r3 = myShape.AddShape(msoShapeRectangle, 100, 420, 200, 100)
    let connFormat1 = myShape.Item("Conn1To2").ConnectorFormat
        let endConnSite1 = connFormat1.EndConnectionSite
        let endConnShape1 = connFormat1.EndConnectedShape
    let connFormat2 = myShape.AddConnector(msoConnectorCurve, 
            0, 0, 10, 10).ConnectorFormat
        connFormat2.BeginConnect(r3, 1)
        connFormat2.EndConnect(endConnShape1, endConnSite1)
}
```

### ConnectorFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ConnectorFormat 对象的变量。

### ConnectorFormat.Type

返回或设置一个 **MsoConnectorType** 值，它代表连接符格式类型。

**语法**

**express.Type**

\*express是一个代表 ConnectorFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ConnectorFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FreeformBuilder 对象

## 简介

代表任意多边形创建时的几何属性。

## 说明

使用 BuildFreeform 方法可返回 FreeformBuilder 对象。使用 AddNodes 方法可在任意多边形中添加节点。使用 ConvertToShape 方法可创建 FreeformBuilder 对象中定义的形状，并将它添加到 Shapes 集合中。

``` JavaScript
/*下例在 myDocument 中添加一个具有四段的任意多边形。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let shape1 = myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    shape1.AddNodes(msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300)
    shape1.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 2000)
    shape1.AddNodes(msoSegmentLine, msoEditingAuto, 480, 400)
    shape1.AddNodes(msoSegmentLine, msoEditingAuto, 360, 200)
    shape1.ConvertToShape()
}
```

## 方法

| 序号 | 名称                                              | 说明                                                                                     |
|:----:|:--------------------------------------------------|:-----------------------------------------------------------------------------------------|
|  1   | [AddNodes](#FreeformBuilder.AddNodes)             | 在当前形状中添加一点，然后绘制一个从当前节点到添加的最后一个节点的线条。                 |
|  2   | [ConvertToShape](#FreeformBuilder.ConvertToShape) | 创建一个具有指定 FreeformBuilder 对象的几何特性的形状。返回一个代表新形状的 Shape 对象。 |

## 属性

| 序号 | 名称                                        | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FreeformBuilder.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#FreeformBuilder.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Parent](#FreeformBuilder.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### FreeformBuilder.AddNodes

在当前形状中添加一点，然后绘制一个从当前节点到添加的最后一个节点的线条。

**语法**

**express.AddNodes ( *SegmentType* , *EditingType* , *X1* , *Y1* , *X2* , *Y2* , *X3* , *Y3* )**

\*express是一个代表 FreeformBuilder 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型       | 说明                                                                                                                                                                                                                              |
|---------------|-----------|----------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *SegmentType* | 必选      | MsoSegmentType | 要添加的线段的类型。                                                                                                                                                                                                              |
| *EditingType* | 必选      | MsoEditingType | 顶点的编辑属性。                                                                                                                                                                                                                  |
| *X1*          | 必选      | Single         | 如果新线段的 EditingType 为 msoEditingAuto，则此参数指定从文档左上角到新线段终点的水平距离（以磅为单位）。 如果新节点的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第一个控制点的水平距离（以磅为单位）。 |
| *Y1*          | 必选      | Single         | 如果新线段的 EditingType 为 msoEditingAuto，则此参数指定从文档左上角到新线段终点的水平距离（以磅为单位）。 如果新节点的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第一个控制点的水平距离（以磅为单位）。 |
| *X2*          | 可选      | Variant        | 如果新线段的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第二个控制点的水平距离（以磅为单位）。 如果新线段的 EditingType 为 msoEditingAuto，则不为该参数指定值。                                           |
| *Y2*          | 可选      | Variant        | 如果新线段的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第二个控制点的水平距离（以磅为单位）。 如果新线段的 EditingType 为 msoEditingAuto，则不为该参数指定值。                                           |
| *X3*          | 可选      | Variant        | 如果新线段的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第二个控制点的水平距离（以磅为单位）。 如果新线段的 EditingType 为 msoEditingAuto，则不为该参数指定值。                                           |
| *Y3*          | 可选      | Variant        | 如果新线段的 EditingType 为 msoEditingCorner，则此参数指定从文档左上角到新线段的第二个控制点的水平距离（以磅为单位）。 如果新线段的 EditingType 为 msoEditingAuto，则不为该参数指定值。                                           |

**说明**

|                                                             |
|-------------------------------------------------------------|
| **MsoSegmentType** 可以是下列 **MsoSegmentType** 常量之一。 |
| **msoSegmentLine**                                          |
| **msoSegmentCurve**                                         |

<table>
<colgroup>
<col style="width: 100%" />
</colgroup>
<tbody>
<tr class="odd">
<td><strong>MsoEditingType</strong> 可以是下列 <strong>MsoEditingType</strong> 常量之一。</td>
</tr>
<tr class="even">
<td><strong>msoEditingAuto</strong></td>
</tr>
<tr class="odd">
<td><strong>msoEditingCorner</strong></td>
</tr>
<tr class="even">
<td>不能是 <strong>msoEditingSmooth</strong> 或 <strong>msoEditingSymmetric</strong>
<p>如果 <em>SegmentType</em> 为 <strong>msoSegmentLine</strong> ，那么 <em>EditingType</em> 就必须是 <strong>msoEditingAuto</strong> 。</p></td>
</tr>
</tbody>
</table>

**示例**

``` JavaScript
/*本示例向 myDocument 中添加带有四条线段的任意多边形。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let shape1 = myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    shape1.AddNodes(msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300)
    shape1.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 200)
    shape1.AddNodes(msoSegmentLine, msoEditingAuto, 480, 400)
    shape1.AddNodes(msoSegmentLine, msoEditingAuto, 360, 200)
    shape1.ConvertToShape()
}
```

### FreeformBuilder.ConvertToShape

创建一个具有指定 **FreeformBuilder** 对象的几何特性的形状。返回一个代表新形状的 **Shape** 对象。

**语法**

**express.ConvertToShape ()**

\*express是一个代表 FreeformBuilder 对象的变量。

**返回值**

Shape

**说明**

对于 **FreeformBuilder** 对象，至少要应用一次 **AddNodes** 方法，然后才能使用 **ConvertToShape** 方法。

**示例**

``` JavaScript
/*本示例将一个有五个顶点的任意多边形添加到 myDocument 中。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let shape1 = myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    shape1.AddNodes(msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300)
    shape1.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 200)
    shape1.AddNodes(msoSegmentLine, msoEditingAuto, 480, 400)
    shape1.AddNodes(msoSegmentLine, msoEditingAuto, 360, 200)
    shape1.ConvertToShape()
}
```

## 成员属性

### FreeformBuilder.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 FreeformBuilder 对象的变量。

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

### FreeformBuilder.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FreeformBuilder 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### FreeformBuilder.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FreeformBuilder 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/FreeformBuilder](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ODBCConnection 对象

## 简介

代表 ODBC 连接。

## 说明

ODBC 连接可存储在 ET 工作簿中。当 Micrososft ET 打开此工作簿时， ET 会创建一个称为 ODBCConnection 对象的 ODBC 连接的内存副本。

ODBCConnection 对象包含与连接相关的信息，例如，要连接到的服务器的名称和要在该服务器上打开的对象的名称。（可选） ODBCConnection 对象也可以包括身份验证凭据信息或要传递到服务器并执行的命令（例如，要由 SQL Server 执行的 ` SELECT ` 语句）。

## 方法

| 序号 | 名称                                           | 说明                                             |
|:----:|:-----------------------------------------------|:-------------------------------------------------|
|  1   | [CancelRefresh](#ODBCConnection.CancelRefresh) | 对指定的 ODBC 连接取消正在进行的所有刷新操作。   |
|  2   | [Refresh](#ODBCConnection.Refresh)             | 刷新 ODBC 连接。                                 |
|  3   | [SaveAsODC](#ODBCConnection.SaveAsODC)         | 将 ODBC 连接保存为一个 WPS Office 数据连接文件。 |

## 属性

| 序号 | 名称                                                               | 说明                                                                                                                                                               |
|:----:|:-------------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AlwaysUseConnectionFile](#ODBCConnection.AlwaysUseConnectionFile) | 如果连接文件始终用于建立与数据源的连接，则为 True 。 Boolean 类型，可读写。                                                                                        |
|  2   | [Application](#ODBCConnection.Application)                         | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  3   | [BackgroundQuery](#ODBCConnection.BackgroundQuery)                 | 如果 ODBC 连接的查询是异步执行（在后台执行）的，则为 True 。可读/写 Boolean 类型。                                                                                 |
|  4   | [CommandText](#ODBCConnection.CommandText)                         | 返回或设置指定数据源的命令字符串。可读/写 Variant 类型。                                                                                                           |
|  5   | [CommandType](#ODBCConnection.CommandType)                         | 返回或设置 XlCmdType 常量之一。可读/写 XlCmdType 。                                                                                                                |
|  6   | [Connection](#ODBCConnection.Connection)                           | 返回或设置一个包含 ODBC 设置的字符串，这些设置使 ET 可以连接 ODBC 数据源。可读/写 Variant 类型。                                                                   |
|  7   | [Creator](#ODBCConnection.Creator)                                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  8   | [EnableRefresh](#ODBCConnection.EnableRefresh)                     | 如果用户可刷新连接，则为 True 。默认值是 True 。可读/写 Boolean 类型。                                                                                             |
|  9   | [Parent](#ODBCConnection.Parent)                                   | 返回指定对象的父对象。只读。                                                                                                                                       |
|  10  | [RefreshDate](#ODBCConnection.RefreshDate)                         | 返回上次刷新 ODBC 连接的日期。只读 Date 类型。                                                                                                                     |
|  11  | [Refreshing](#ODBCConnection.Refreshing)                           | 如果正在进行指定 ODBC 连接的后台 ODBC 查询，则为 True 。可读/写 Boolean 类型。                                                                                     |
|  12  | [RefreshOnFileOpen](#ODBCConnection.RefreshOnFileOpen)             | 如果每次打开工作簿都自动更新连接，则为 True 。默认值为 False 。可读/写 Boolean 类型。                                                                              |
|  13  | [RefreshPeriod](#ODBCConnection.RefreshPeriod)                     | 返回或设置两次刷新之间的时间间隔（单位为分钟）。可读/写 Long 类型。                                                                                                |
|  14  | [RobustConnect](#ODBCConnection.RobustConnect)                     | 返回或设置 ODBC 连接与其数据源连接的方式。可读/写 XlRobustConnect 。                                                                                               |
|  15  | [SavePassword](#ODBCConnection.SavePassword)                       | 如果将 ODBC 连接字符串中的密码信息保存在连接字符串中，则为 True 。如果删除密码，则为 False 。可读/写 Boolean 类型。                                                |
|  16  | [ServerCredentialsMethod](#ODBCConnection.ServerCredentialsMethod) | 返回或设置应该用于服务器验证的凭据的类型。可读/写                                                                                                                  |
|  17  | [ServerSSOApplicationID](#ODBCConnection.ServerSSOApplicationID)   | 返回或设置单一登录 (SSO) 应用程序标识符，该标识符用于在 SSO 数据库中查找凭据。可读/写 String 类型。                                                                |
|  18  | [SourceConnectionFile](#ODBCConnection.SourceConnectionFile)       | 返回或设置一个 String 类型的值，该值指示用于创建连接的 WPS Office 数据连接文件或类似文件。可读/写。                                                                |
|  19  | [SourceData](#ODBCConnection.SourceData)                           | 返回 ODBC 连接的数据源，如表中所示。可读/写 Variant 类型。                                                                                                         |
|  20  | [SourceDataFile](#ODBCConnection.SourceDataFile)                   | 返回或设置一个 String 类型的值，该值指示 ODBC 连接的源数据文件。可读/写。                                                                                          |

## 成员方法

### ODBCConnection.CancelRefresh

对指定的 ODBC 连接取消正在进行的所有刷新操作。

**语法**

**express.CancelRefresh ()**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

使用 **Refreshing** 属性可以确定当前是否正在进行刷新操作。

### ODBCConnection.Refresh

刷新 ODBC 连接。

**语法**

**express.Refresh ()**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

与 ODBC 数据源建立连接时，ET 使用 **Connection** 属性指定的连接字符串。如果指定的连接字符串缺少必需的值，将显示对话框，提示用户提供必需的信息。如果 **DisplayAlerts** 属性为 **False** ，则不显示对话框， **Refresh** 方法失败并导致“连接信息无效”异常。

在 ET 成功建立一个连接之后，将存储完整的连接字符串，以便以后在同一编辑会话中调用 **Refresh** 方法时将不再显示提示。通过检查 **Connection** 属性的值可以获得完整的连接字符串。

建立数据库连接后，将检查 SQL 查询的有效性。如果该查询无效， **Refresh** 方法将失败并导致“SQL 语法错误”异常。

如果查询需要参数，则必须在调用 **Refresh** 方法之前，用参数绑定信息初始化 **Parameters** 集合。如果未绑定足够的参数， **Refresh** 方法将失败并导致“参数错误”异常。如果将参数设置为提示用户输入参数值，则无论 **DisplayAlerts** 属性的设置如何，都会向用户显示对话框。如果用户取消参数对话框， **Refresh** 方法将停止并返回 **False** 。如果对 **Parameters** 集合绑定了额外的参数，则这些额外参数将被忽略。

如果成功地完成或启动查询，则 **Refresh** 方法将返回 **True** ；如果用户取消连接或参数对话框，则该方法返回 **False** 。

要查看取回的行数是否超过了工作表中的可用行数，请检查 **FetchedRowOverflow** 属性。每次调用 **Refresh** 方法之前，该属性都将被初始化。

### ODBCConnection.SaveAsODC

将 ODBC 连接保存为一个 WPS Office 数据连接文件。

**语法**

**express.SaveAsODC ( *ODCFileName* , *Description* , *Keywords* )**

\*express是一个代表 ODBCConnection 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                   |
|---------------|-----------|----------|----------------------------------------|
| *ODCFileName* | 必选      | String   | 保存文件的位置。                       |
| *Description* | 可选      | Variant  | 将保存在文件中的说明。                 |
| *Keywords*    | 可选      | Variant  | 可用于搜索该文件的以空格分隔的关键字。 |

**返回值**

无

**示例**

``` JavaScript
/*以下示例将连接保存为名为“ODCFile”的 ODC 文件。此示例假定 ODBC 连接位于活动工作表上。 */
function UseSaveAsODC(){
    Application.ActiveWorkbook.ODBCConnection.SaveAsODC ("ODCFile")
}
```

## 成员属性

### ODBCConnection.AlwaysUseConnectionFile

如果连接文件始终用于建立与数据源的连接，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AlwaysUseConnectionFile**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

当此属性为 **True** 时，连接文件将用于建立与数据源的连接。如果在工作簿中嵌入的连接不同于外部连接文件，则忽略嵌入的连接，且外部连接文件将是唯一考虑的版本。

### ODBCConnection.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### ODBCConnection.BackgroundQuery

如果 ODBC 连接的查询是异步执行（在后台执行）的，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.BackgroundQuery**

\*express是一个代表 ODBCConnection 对象的变量。

### ODBCConnection.CommandText

返回或设置指定数据源的命令字符串。可读/写 **Variant** 类型。

**语法**

**express.CommandText**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

应使用 **CommandText** 属性，而不应使用 **SQL** 属性，现在 SQL 属性的存在主要是为了与 ET 的早期版本兼容。如果您同时使用这两个属性，则 **CommandText** 属性的值优先。

**CommandType** 属性描述了 **CommandText** 属性的值。

**示例**

``` JavaScript
/*本示例为第一张查询表的 ODBC 数据源设置命令字符串。注意，该命令字符串是一个 SQL 语句。 */
function test(){
　　　　let qtQtrResults = Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)
　　　　qtQtrResults.CommandType = xlCmdSQL
　　　　qtQtrResults.CommandText = "Select ProductID From Products Where ProductID < 10"
　　　　qtQtrResults.Refresh()
}
```

### ODBCConnection.CommandType

返回或设置 **XlCmdType** 常量之一。可读/写 **XlCmdType** 。

**语法**

**express.CommandType**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

返回或设置的常量描述了 **CommandText** 属性的值。默认值是 **xlCmdSQL** 。

**示例**

``` JavaScript
/*本示例为第一张查询表的 ODBC 数据源设置命令字符串。该命令字符串是一个 SQL 语句。 */
function test(){
　　　　let qtQtrResults = Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)
　　　　qtQtrResults.CommandType = xlCmdSQL
　　　　qtQtrResults.CommandText = "Select ProductID From Products Where ProductID < 10"
　　　　qtQtrResults.Refresh()
}
```

### ODBCConnection.Connection

返回或设置一个包含 ODBC 设置的字符串，这些设置使 ET 可以连接 ODBC 数据源。可读/写 **Variant** 类型。

**语法**

**express.Connection**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

设置 **Connection** 属性并不会立即初始化到数据源的连接。必须使用 **Refresh** 方法建立连接并检索数据。

### ODBCConnection.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ODBCConnection.EnableRefresh

如果用户可刷新连接，则为 **True** 。默认值是 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.EnableRefresh**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

如果 **EnableRefresh** 属性设置为 **False** ，则 **RefreshOnFileOpen** 属性被忽略。对于 OLAP 数据源，如果将此属性设置为 **False** ，则会禁用更新。

### ODBCConnection.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ODBCConnection 对象的变量。

### ODBCConnection.RefreshDate

返回上次刷新 ODBC 连接的日期。只读 **Date** 类型。

**语法**

**express.RefreshDate**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

对于 OLAP 数据源，在每次查询后都会更新此属性。

### ODBCConnection.Refreshing

如果正在进行指定 ODBC 连接的后台 ODBC 查询，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.Refreshing**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

使用 **CancelRefresh** 方法可取消后台查询。

### ODBCConnection.RefreshOnFileOpen

如果每次打开工作簿都自动更新连接，则为 **True** 。默认值为 **False** 。可读/写 **Boolean** 类型。

**语法**

**express.RefreshOnFileOpen**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

在 Visual Basic 中使用 **Open** 方法打开工作簿时，不会自动刷新连接。打开工作簿后，可使用 **Refresh** 方法刷新数据。

### ODBCConnection.RefreshPeriod

返回或设置两次刷新之间的时间间隔（单位为分钟）。可读/写 **Long** 类型。

**语法**

**express.RefreshPeriod**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

将时间段设置为 0（零）会禁用定时自动刷新，等同于将此属性设置为 **Null** 。 **RefreshPeriod** 属性的值可以是从 0 到 32767 的整数。

### ODBCConnection.RobustConnect

返回或设置 ODBC 连接与其数据源连接的方式。可读/写 **XlRobustConnect** 。

**语法**

**express.RobustConnect**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

如果您使用可靠的连接，那么当 ET 无法使用工作簿连接信息建立连接时， ET 将检查工作簿连接是否具有外部连接文件的路径。如果有， ET 将打开外部连接文件并尝试使用外部连接文件中包含的信息进行连接。如果在使用外部连接文件之后能够建立连接，则 ET 将更新工作簿的连接对象。

在信息技术部门在一个中央位置维护和更新连接的情况下，这可以提供可靠的连接，每当上一版本的连接（缓存在工作簿中）失败时，允许用户的工作簿自动获取那些更新。

### ODBCConnection.SavePassword

如果将 ODBC 连接字符串中的密码信息保存在连接字符串中，则为 **True** 。如果删除密码，则为 **False** 。可读/写 **Boolean** 类型。

**语法**

**express.SavePassword**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

此属性由数据透视表和查询表在 ODBC 和 OLE DB 查询中使用。

### ODBCConnection.ServerCredentialsMethod

返回或设置应该用于服务器验证的凭据的类型。可读/写

**语法**

**express.ServerCredentialsMethod**

\*express是一个代表 ODBCConnection 对象的变量。

### ODBCConnection.ServerSSOApplicationID

返回或设置单一登录 (SSO) 应用程序标识符，该标识符用于在 SSO 数据库中查找凭据。可读/写 **String** 类型。

**语法**

**express.ServerSSOApplicationID**

\*express是一个代表 ODBCConnection 对象的变量。

### ODBCConnection.SourceConnectionFile

返回或设置一个 **String** 类型的值，该值指示用于创建连接的 WPS Office 数据连接文件或类似文件。可读/写。

**语法**

**express.SourceConnectionFile**

\*express是一个代表 ODBCConnection 对象的变量。

### ODBCConnection.SourceData

返回 ODBC 连接的数据源，如表中所示。可读/写 **Variant** 类型。

**语法**

**express.SourceData**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

| 数据源          | 返回值                                                                                     |
|-----------------|--------------------------------------------------------------------------------------------|
| ET 列表或数据库 | 文本形式的单元格引用。                                                                     |
| 外部数据源      | 数组。每行包含一个 SQL 连接字符串，该行的剩余元素作为查询字符串，以 255 个字符为单位分段。 |
| 多个合并区域    | 二维数组。每行由一个引用及与其相关联的页字段项组成。                                       |
| 其他数据透视表  | 以上三种信息之一。                                                                         |

### ODBCConnection.SourceDataFile

返回或设置一个 **String** 类型的值，该值指示 ODBC 连接的源数据文件。可读/写。

**语法**

**express.SourceDataFile**

\*express是一个代表 ODBCConnection 对象的变量。

**说明**

对于基于文件的数据源（如 Access）， **SourceDataFile** 属性包含源数据文件的完全限定路径。对于基于服务器的数据源（如 SQL Server），该属性为 null。如果通过编程方式更改 **Connection** 属性，则 **SourceDataFile** 属性设置为 null。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ODBCConnection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ODBCErrors 对象

## 简介

ODBCError 对象的集合。

## 说明

每一个 ODBCError 对象都代表一个由最近的 ODBC 查询返回的错误。如果指定的 ODBC 查询运行时没有出现错误，则 ODBCErrors 集合为空。集合中错误的索引次序与 ODBC 数据源生成它们时的次序相同。您不能给该集合添加成员。

## 方法

| 序号 | 名称                     | 说明                   |
|:----:|:-------------------------|:-----------------------|
|  1   | [Item](#ODBCErrors.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ODBCErrors.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#ODBCErrors.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#ODBCErrors.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#ODBCErrors.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### ODBCErrors.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 ODBCErrors 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *Index* | 必选      | Long     | 对象的索引号。 |

**返回值**

包含在集合中的一个 ODBCError 对象。

**示例**

``` JavaScript
/*本示例显示一个 ODBC 错误。*/
function test(){
　　　　let er = Application.ODBCErrors.Item(1)
　　　　MsgBox ("The following error occurred:" +
　　　　   er.ErrorString + " : " + er.SqlState)
}
```

## 成员属性

### ODBCErrors.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ODBCErrors 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
　　　　let myObject = ActiveWorkbook
    if (myObject.Application.Value == "ET"){
　　　　    MsgBox ("This is an ET Application object.")
　　　　}
　　　　else{
　　　　    MsgBox ("This is not an ET Application object.")
　　　　}
}
```

### ODBCErrors.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 ODBCErrors 对象的变量。

### ODBCErrors.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ODBCErrors 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ODBCErrors.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ODBCErrors 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ODBCErrors](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Pane 对象

## 简介

代表窗口中的窗格。

## 说明

Pane 对象只对于工作表和 ET 4.0 宏表存在。 Pane 对象是 Panes 集合的成员。 Panes 集合包含在一个窗口中显示的所有窗格。

使用 Panes ( *index* )（其中 *index* 是窗格索引号）可返回一个 Pane 对象。下例拆分显示工作表一的窗口，然后滚动左下角的窗格，直至第五行处于该窗格的顶端。

``` JavaScript
function test(){    
    Application.Worksheets.Item(1).Activate()
    Application.ActiveWindow.Split = true
    Application.ActiveWindow.Panes.Item(3).ScrollRow = 5
}
```

## 方法

| 序号 | 名称                                                 | 说明                                                                                                |
|:----:|:-----------------------------------------------------|:----------------------------------------------------------------------------------------------------|
|  1   | [Activate](#Pane.Activate)                           | 激活窗格。返回 Boolean值                                                                            |
|  2   | [LargeScroll](#Pane.LargeScroll)                     | 按页滚动窗口内容。                                                                                  |
|  3   | [PointsToScreenPixelsX](#Pane.PointsToScreenPixelsX) | 返回或设置屏幕上的像素点。返回Long值                                                                |
|  4   | [PointsToScreenPixelsY](#Pane.PointsToScreenPixelsY) | 返回或设置屏幕像素的位置。返回Long值                                                                |
|  5   | [ScrollIntoView](#Pane.ScrollIntoView)               | 滚动文档窗口，使指定矩形区域中的内容显示在文档窗口或窗格的左上角或右下角（取决于 *Start* 参数值）。 |
|  6   | [SmallScroll](#Pane.SmallScroll)                     | 按行或按列滚动窗口内容。返回Variant值                                                               |

## 属性

| 序号 | 名称                               | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Pane.Application)   | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#Pane.Creator)           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Index](#Pane.Index)               | 返回 Long 值，它代表对象在其同类对象所组成的集合内的索引号。                                                                                                                                                                    |
|  4   | [Parent](#Pane.Parent)             | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  5   | [ScrollColumn](#Pane.ScrollColumn) | 返回或设置指定窗格或窗口最左边的列号。 Long 型，可读写。                                                                                                                                                                        |
|  6   | [ScrollRow](#Pane.ScrollRow)       | 返回或设置指定窗格或窗口最上面显示的行号。 Long 型，可读写。                                                                                                                                                                    |
|  7   | [VisibleRange](#Pane.VisibleRange) | 返回一个 Range 对象，它代表显示在窗口或窗格中的单元格区域。如果列或行只显示了一部分，则说明它是包括在区域之内的。只读。                                                                                                         |

## 成员方法

### Pane.Activate

激活窗格。返回 Boolean值

**语法**

**express.Activate ()**

\*express是一个代表 Pane 对象的变量。

### Pane.LargeScroll

按页滚动窗口内容。

**语法**

**express.LargeScroll ( *Down* , *Up* , *ToRight* , *ToLeft* )**

\*express是一个代表 Pane 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                 |
|-----------|-----------|----------|----------------------|
| *Down*    | 可选      | Variant  | 向下滚动内容的页数。 |
| *Up*      | 可选      | Variant  | 向上滚动内容的页数。 |
| *ToRight* | 可选      | Variant  | 向右滚动内容的页数。 |
| *ToLeft*  | 可选      | Variant  | 向左滚动内容的页数。 |

**说明**

如果同时指定了 *Down* 和 *Up* ，窗口内容的滚动量由这两个参数的差值决定。例如，如果 *Down* 为 3， *Up* 为 6，则窗口向上滚动三页。

如果同时指定了 *ToLeft* 和 *ToRight* ，窗口内容的滚动量由这两个参数的差值决定。例如，如果 *ToLeft* 为 3， *ToRight* 为 6，则窗口向右滚动三页。

以上任何参数均可为负数。

### Pane.PointsToScreenPixelsX

返回或设置屏幕上的像素点。返回Long值

**语法**

**express.PointsToScreenPixelsX ( *Points* )**

\*express是一个代表 Pane 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明             |
|----------|-----------|----------|------------------|
| *Points* | 必选      | Long     | 屏幕像素的位置。 |

### Pane.PointsToScreenPixelsY

返回或设置屏幕像素的位置。返回Long值

**语法**

**express.PointsToScreenPixelsY ( *Points* )**

\*express是一个代表 Pane 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明           |
|----------|-----------|----------|----------------|
| *Points* | 必选      | Long     | 起始点的位置。 |

### Pane.ScrollIntoView

滚动文档窗口，使指定矩形区域中的内容显示在文档窗口或窗格的左上角或右下角（取决于 *Start* 参数值）。

**语法**

**express.ScrollIntoView ( *Left* , *Top* , *Width* , *Height* , *Start* )**

\*express是一个代表 Pane 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                                                               |
|----------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------|
| *Left*   | 必选      | Long     | 矩形距离文档窗口或窗格左边的水平位置（以磅为单位）。                                                                               |
| *Top*    | 必选      | Long     | 矩形距离文档窗口或窗格上边的垂直位置（以磅为单位）。                                                                               |
| *Width*  | 必选      | Long     | 矩形的宽度（以磅为单位）。                                                                                                         |
| *Height* | 必选      | Long     | 矩形的高度（以磅为单位）。                                                                                                         |
| *Start*  | 可选      | Variant  | 如果为 True，则使矩形的左上角位于文档窗口或窗格的左上角。如果为 False，则使矩形的右下角位于文档窗口或窗格的右下角。默认值是 True。 |

**说明**

当矩形比文档窗口或窗格大时， *Start* 参数对调整屏幕显示很有用处。

**示例**

``` JavaScript
/*此示例在当前活动文档窗口中定义一个 100x200 像素的矩形，其位置为距窗口顶部 20 像素，距窗口左边缘 50 像素。然后将文档向左、向上滚动，以使矩形的左上角与窗口的左上角对齐。*/
Application.ActiveWindow.ScrollIntoView(50,20,100,200)
```

### Pane.SmallScroll

按行或按列滚动窗口内容。返回Variant值

**语法**

**express.SmallScroll ( *Down* , *Up* , *ToRight* , *ToLeft* )**

\*express是一个代表 Pane 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                   |
|-----------|-----------|----------|------------------------|
| *Down*    | 可选      | Variant  | 将内容向下滚动的行数。 |
| *Up*      | 可选      | Variant  | 将内容向上滚动的行数。 |
| *ToRight* | 可选      | Variant  | 将内容向右滚动的列数。 |
| *ToLeft*  | 可选      | Variant  | 将内容向左滚动的列数。 |

**说明**

如果同时指定了 *Down* 和 *Up* ，窗口内容的滚动量由这两个参数的差值决定。例如，如果 *Down* 为 3， *Up* 为 6，则窗口向上滚动三行。

如果同时指定了 *ToLeft* 和 *ToRight* ，窗口内容的滚动量由这两个参数的差值决定。例如，如果 *ToLeft* 为 3， *ToRight* 为 6，则窗口内容向右滚动三列。

任何参数均可为负数。

## 成员属性

### Pane.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Pane 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
if(Application.ActiveWorkbook.Application.Value == "ET") {
    alert("This is an ET Application object.")
}
else {
    alert("This is not an ET Application object.")
}
```

### Pane.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Pane 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Pane.Index

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 Pane 对象的变量。

### Pane.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Pane 对象的变量。

### Pane.ScrollColumn

返回或设置指定窗格或窗口最左边的列号。 **Long** 型，可读写。

**语法**

**express.ScrollColumn**

\*express是一个代表 Pane 对象的变量。

**说明**

如果窗口处于拆分状态，则 **Window** 对象的 **ScrollColumn** 属性表示左上方的窗格。如果某些窗格处于冻结状态，则 **Window** 对象的 **ScrollColumn** 属性将不包括冻结的区域。

### Pane.ScrollRow

返回或设置指定窗格或窗口最上面显示的行号。 **Long** 型，可读写。

**语法**

**express.ScrollRow**

\*express是一个代表 Pane 对象的变量。

**说明**

如果窗口处于拆分状态，则 **Window** 对象的 **ScrollRow** 属性表示左上方的窗格。如果某些窗格处于冻结状态，则 **Window** 对象的 **ScrollRow** 属性将不包括冻结的区域。

**示例**

``` JavaScript
function test(){
  /*本示例将第十行移动到窗口顶部。*/
  Application.Worksheets.Item("Sheet1").Activate()
  Application.ActiveWindow.ScrollRow = 10
}
```

### Pane.VisibleRange

返回一个 **Range** 对象，它代表显示在窗口或窗格中的单元格区域。如果列或行只显示了一部分，则说明它是包括在区域之内的。只读。

**语法**

**express.VisibleRange**

\*express是一个代表 Pane 对象的变量。

**示例**

``` JavaScript
function test(){
  /*此示例显示工作表 sheet1 中可见单元格的数目。*/
  Application.Worksheets.Item("Sheet1").Activate()
  alert("There are " + Application.Windows.Item(1).VisibleRange.Cells.Count + " cells visible")
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Pane](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PictureFormat 对象

## 简介

包含应用于图片和 OLE 对象的属性和方法。

## 说明

LinkFormat 对象只包含应用于链接 OLE 对象的属性和方法。 OLEFormat 对象包含应用于 OLE 对象的属性和方法，无论这些对象是不是链接对象。

## 方法

| 序号 | 名称                                                      | 说明                                                                   |
|:----:|:----------------------------------------------------------|:-----------------------------------------------------------------------|
|  1   | [IncrementBrightness](#PictureFormat.IncrementBrightness) | 按指定数值改变图片亮度。可以使用 Brightness 属性设置图片的绝对亮度。   |
|  2   | [IncrementContrast](#PictureFormat.IncrementContrast)     | 按指定值改变图片的对比度。可以使用 Contrast 属性设置图片的绝对对比度。 |

## 属性

| 序号 | 名称                                                          | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#PictureFormat.Application)                     | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Brightness](#PictureFormat.Brightness)                       | 返回或设置指定图片或 OLE 对象的亮度。该属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数。 Single 型，可读写。                                                                                                                   |
|  3   | [ColorType](#PictureFormat.ColorType)                         | 返回或设置应用于指定图片或 OLE 对象的颜色转换类型。 MsoPictureColorType 类型，可读写。                                                                                                                                          |
|  4   | [Contrast](#PictureFormat.Contrast)                           | 返回或设置指定图片或 OLE 对象的对比度。该属性的值必须是 0.0（最低对比度）到 1.0（最高对比度）的数。 Single 型，可读写。                                                                                                         |
|  5   | [Creator](#PictureFormat.Creator)                             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  6   | [Crop](#PictureFormat.Crop)                                   | 返回一个 Crop 对象，该对象代表指定的 PictureFormat 对象的裁剪设置。                                                                                                                                                             |
|  7   | [CropBottom](#PictureFormat.CropBottom)                       | 返回或设置从指定图片或 OLE 对象的底部所裁剪下的磅数。 Single 型，可读写。                                                                                                                                                       |
|  8   | [CropLeft](#PictureFormat.CropLeft)                           | 返回或设置从指定图片或 OLE 对象的左边所裁剪下的磅数。 Single 型，可读写。                                                                                                                                                       |
|  9   | [CropRight](#PictureFormat.CropRight)                         | 返回或设置从指定图片或 OLE 对象的右边所裁剪下的磅数。 Single 型，可读写。                                                                                                                                                       |
|  10  | [CropTop](#PictureFormat.CropTop)                             | 返回或设置从指定图片或 OLE 对象的顶部所裁剪下的磅数。 Single 型，可读写。                                                                                                                                                       |
|  11  | [Parent](#PictureFormat.Parent)                               | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  12  | [TransparencyColor](#PictureFormat.TransparencyColor)         | 返回或设置指定图片的透明色，以红-绿-蓝 (RGB) 值表示。必须将 TransparentBackground 属性设置为 True ，该属性才有效。仅应用于位图。 Long 类型，可读写。                                                                            |
|  13  | [TransparentBackground](#PictureFormat.TransparentBackground) | 使用 TransparencyColor 属性可设置透明颜色。仅应用于位图。MsoTriState 类型，可读写。                                                                                                                                             |

## 成员方法

### PictureFormat.IncrementBrightness

按指定数值改变图片亮度。可以使用 **Brightness** 属性设置图片的绝对亮度。

**语法**

**express.IncrementBrightness ( *Increment* )**

\*express是一个代表 PictureFormat 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                       |
|-------------|-----------|----------|----------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定图片的 Brightness 属性值的改变量。正值表示图片变亮，负值表示图片变暗。 |

**说明**

调整图片亮度不能超出 **Brightness** 属性的界限。例如，如果 **Brightness** 属性初始设为 0.9 并将 *Increment* 参数指定为 0.3，则结果亮度为 1.0（ **Brightness** 属性的上限）而不是 1.2。

### PictureFormat.IncrementContrast

按指定值改变图片的对比度。可以使用 **Contrast** 属性设置图片的绝对对比度。

**语法**

**express.IncrementContrast ( *Increment* )**

\*express是一个代表 PictureFormat 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                         |
|-------------|-----------|----------|------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定图片的 Contrast 属性值的改变量。正值表示对比度增大，负值表示对比度减小。 |

**说明**

调整图片的对比度不能超出 **Contrast** 属性的上下限。例如，如果 **Contrast** 属性初始设为 0.9 并将 *Increment* 参数指定为 0.3，则结果对比度为 1.0（ **Contrast** 属性的上限）而不是 1.2。

**示例**

``` JavaScript
/*本示例增大 myDocument 上所有对比度未到最大值的图片的对比度。*/
function test(){
　　　　let shape = Worksheets.Item(1).Shapes
       for(let i =1; i<=shape.Count; i++) {
　　　　    if(shape.Item(i).Type == msoPicture || shape.Item(i).Type == msoLinkedPicture) {
  　　　　      shape.Item(i).PictureFormat.IncrementContrast(0.1)
　　　　    }
　　　　}
}
```

## 成员属性

### PictureFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PictureFormat 对象的变量。

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

### PictureFormat.Brightness

返回或设置指定图片或 OLE 对象的亮度。该属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数。 **Single** 型，可读写。

**语法**

**express.Brightness**

\*express是一个代表 PictureFormat 对象的变量。

**示例**

``` JavaScript
/*此示例设置了 myDocument 中第一个形状的亮度。第一个形状必须是图片或 OLE 对象。*/
function test(){
let myDocument = Worksheets.Item(1)
myDocument.Shapes.Item(1).PictureFormat.Brightness = 0.3
}
```

### PictureFormat.ColorType

返回或设置应用于指定图片或 OLE 对象的颜色转换类型。 **MsoPictureColorType** 类型，可读写。

**语法**

**express.ColorType**

\*express是一个代表 PictureFormat 对象的变量。

**示例**

``` JavaScript
/*本示例通过 MsoPictureColorType 将 myDocument 上的第一个形状的颜色变换类型设置为灰度。第一个形状必须是图片或 OLE 对象。*/
function test(){
　　　　let myDocument = Worksheets.Item(1)
　　　　myDocument.Shapes.Item(1).PictureFormat.ColorType = msoPictureGrayscale
}
```

### PictureFormat.Contrast

返回或设置指定图片或 OLE 对象的对比度。该属性的值必须是 0.0（最低对比度）到 1.0（最高对比度）的数。 **Single** 型，可读写。

**语法**

**express.Contrast**

\*express是一个代表 PictureFormat 对象的变量。

**示例**

``` JavaScript
/*此示例设置了 myDocument 中第一个形状的对比度。第一个形状必须是图片或 OLE 对象。*/
function test(){
let myDocument = Worksheets.Item(1)
myDocument.Shapes.Item(1).PictureFormat.Contrast = 0.8
}
```

### PictureFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PictureFormat 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PictureFormat.Crop

返回一个 **Crop** 对象，该对象代表指定的 **PictureFormat** 对象的裁剪设置。

**语法**

**express.Crop**

\*express是一个代表 PictureFormat 对象的变量。

### PictureFormat.CropBottom

返回或设置从指定图片或 OLE 对象的底部所裁剪下的磅数。 **Single** 型，可读写。

**语法**

**express.CropBottom**

\*express是一个代表 PictureFormat 对象的变量。

**说明**

裁剪的计算是同图像的原始尺寸相关的。例如，如果插入一幅原始高度为 100 磅的图像，将其放大为 200 磅高，然后设置 **CropBottom** 属性为 50, 那么从图像的底部切除的将是 100 磅（而不是 50 磅）。

**示例**

``` JavaScript
/*此示例在 myDocument 中第三个形状的底部裁剪了 20 磅。要使此示例执行，第三个形状必须是图片或 OLE 对象。*/
function test(){
　　　　let myDocument = Worksheets.Item(1)
　　　　myDocument.Shapes.Item(3).PictureFormat.CropBottom = 20
}
```

``` JavaScript
/*使用此示例可指定从选定的形状底部裁去的百分比，不管是否已调整了该形状的大小。为使此示例正常运行，选定的形状必须是图片或 OLE 对象。*/
function test(){
　　　　let percentToCrop = prompt( "What percentage do you want to crop off " + " the bottom of this picture?")
　　　　let shapeToCrop = ActiveWindow.Selection.ShapeRange.Item(1)
　　　　let du = shapeToCrop.Duplicate()
　　　　du.ScaleHeight(1, true,du.Height)
　　　　du.Delete()
　　　　let cropPoints = origHeight * percentToCrop / 100
　　　　shapeToCrop.PictureFormat.CropBottom = cropPoints
}
```

### PictureFormat.CropLeft

返回或设置从指定图片或 OLE 对象的左边所裁剪下的磅数。 **Single** 型，可读写。

**语法**

**express.CropLeft**

\*express是一个代表 PictureFormat 对象的变量。

**说明**

裁剪的计算是同图像的原始尺寸相关的。例如，如果插入一幅原始宽度为 100 磅的图像，将其放大为 200 磅宽，然后设置 **CropLeft** 属性为 50, 那么 100 磅（不是 50 磅）将从图像的左边被切除。

**示例**

``` JavaScript
/*此示例在 myDocument 中第三个形状的左侧裁剪了 20 磅。要使此示例执行，第三个形状必须是图片或 OLE 对象。*/
function test(){
　　　　let myDocument = Worksheets.Item(1)
　　　　myDocument.Shapes.Item(3).PictureFormat.CropLeft = 20
}
```

``` JavaScript
/*使用此示例可指定从选定的形状左边裁去的百分比，不管是否已调整了该形状的大小。为使此示例正常运行，选定的形状必须是图片或 OLE 对象。*/
function test(){
　　　　let percentToCrop = prompt( "What percentage do you want to crop off " + " the left of this picture?")
　　　　let shapeToCrop = ActiveWindow.Selection.ShapeRange.Item(1)
　　　　let du = shapeToCrop.Duplicate()
　　　　du.ScaleHeight(1, true,du.Width)
　　　　du.Delete()
　　　　let cropPoints = origHeight * percentToCrop / 100
　　　　shapeToCrop.PictureFormat.CropBottom = cropPoints
}
```

### PictureFormat.CropRight

返回或设置从指定图片或 OLE 对象的右边所裁剪下的磅数。 **Single** 型，可读写。

**语法**

**express.CropRight**

\*express是一个代表 PictureFormat 对象的变量。

**说明**

裁剪的计算是同图像的原始尺寸相关的。例如，如果插入一幅原始宽度为 100 磅的图像，将其放大为 200 磅宽，然后设置 **CropRight** 属性为 50, 那么 100 磅（不是 50 磅）将从图像的右边被切除。

**示例**

``` JavaScript
/*此示例在 myDocument 中第三个形状的右侧裁剪了 20 磅。要使此示例执行，第三个形状必须是图片或 OLE 对象。*/
function test(){
　　　　let myDocument = Worksheets.Item(1)
　　　　myDocument.Shapes.Item(3).PictureFormat.CropRight = 20
}
```

``` JavaScript
/*使用此示例可指定从选定的形状右边裁去的百分比，不管是否已调整了该形状的大小。为使此示例正常运行，选定的形状必须是图片或 OLE 对象。*/
function test(){
　　　　let percentToCrop = prompt( "What percentage do you want to crop  " + " off the right of this picture?")
　　　　let shapeToCrop = ActiveWindow.Selection.ShapeRange.Item(1)
　　　　let du = shapeToCrop.Duplicate()
　　　　du.ScaleHeight(1, true,du.Width)
　　　　du.Delete()
　　　　let cropPoints = origHeight * percentToCrop / 100
　　　　shapeToCrop.PictureFormat.CropBottom = cropPoints
}
```

### PictureFormat.CropTop

返回或设置从指定图片或 OLE 对象的顶部所裁剪下的磅数。 **Single** 型，可读写。

**语法**

**express.CropTop**

\*express是一个代表 PictureFormat 对象的变量。

**说明**

裁剪的计算是同图像的原始尺寸相关的。例如，如果插入一幅原始高度为 100 磅的图像，将其放大为 200 磅高，然后设置 **CropTop** 属性为 50, 那么 100 磅（不是 50 磅）将从图像的顶部被切除。

**示例**

``` JavaScript
/*此示例在 myDocument 中第三个形状的顶部裁剪了 20 磅。要使此示例执行，第三个形状必须是图片或 OLE 对象。*/
function test(){
　　　　let myDocument = Worksheets.Item(1)
　　　　myDocument.Shapes.Item(3).PictureFormat.CropTop = 20
}
```

``` JavaScript
/*使用此示例可指定从选定的形状顶部裁去的百分比，不管是否已调整了该形状的大小。为使此示例正常运行，选定的形状必须是图片或 OLE 对象。*/
function test(){
　　　　let percentToCrop = prompt( "What percentage do you want to crop " + "  off the top of this picture?")
　　　　let shapeToCrop = ActiveWindow.Selection.ShapeRange.Item(1)
　　　　let du = shapeToCrop.Duplicate()
　　　　du.ScaleHeight(1, true,du.Height)
　　　　du.Delete()
　　　　let cropPoints = origHeight * percentToCrop / 100
　　　　shapeToCrop.PictureFormat.CropBottom = cropPoints
}
```

### PictureFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PictureFormat 对象的变量。

### PictureFormat.TransparencyColor

返回或设置指定图片的透明色，以红-绿-蓝 (RGB) 值表示。必须将 **TransparentBackground** 属性设置为 **True** ，该属性才有效。仅应用于位图。 **Long** 类型，可读写。

**语法**

**express.TransparencyColor**

\*express是一个代表 PictureFormat 对象的变量。

**说明**

如果想透过图片的透明部分始终看到图片后面的对象，必须将图片的 **FillFormat** 对象的 **Visible** 属性设为 **False** 。如果图片有透明颜色且 **FillFormat** 对象的 **Visible** 属性被设为 **True** ，则图片的填充可以透过透明颜色看到，但图片后面的对象将难以看到。

**示例**

``` JavaScript
/*本例将 myDocument 的第一个形状中带有 RGB 值的颜色设置为透明色，该颜色由 RGB(0, 0, 255) 函数返回。要使本示例执行，第一个形状必须是位图。*/
function test(){
　　　　let blueScreen = (0,0,255)
　　　　let myDocument = Worksheets.Item(1)
　　　　let myshape = myDocument.Shapes.Item(1)
  　　　let mypicture = myshape.PictureFormat
     　　　　   mypicture.TransparentBackground = true
     　　　　   mypicture.TransparencyColor = blueScreen
 　　　　   myshape.Fill.Visible = false
}
```

### PictureFormat.TransparentBackground

使用 **TransparencyColor** 属性可设置透明颜色。仅应用于位图。MsoTriState 类型，可读写。

**语法**

**express.TransparentBackground**

\*express是一个代表 PictureFormat 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。    |
|--------------------------------------------------|
| msoCTrue                                         |
| msoFalse                                         |
| msoTriStateMixed                                 |
| msoTriStateToggle                                |
| msoTrue 图片中定义为透明颜色的部分显示为透明的。 |

如果想透过图片的透明部分始终看到图片后面的对象，必须将图片的 **FillFormat** 对象的 **Visible** 属性设为 **False** 。如果图片有透明颜色且 **FillFormat** 对象的 **Visible** 属性被设为 **True** ，则图片的填充可以透过透明颜色看到，但图片后面的对象将难以看到。

**示例**

``` JavaScript
/*本例将 myDocument 的第一个形状中带有 RGB 值的颜色设置为透明色，该颜色由 RGB(0, 24, 240) 函数返回。要使本示例执行，第一个形状必须是位图。 */
function test(){
　　　　let blueScreen = (0,0,255)
　　　　let myDocument = Worksheets.Item(1)
　　　　let myshape = myDocument.Shapes.Item(1)
  　　　let mypicture = myshape.PictureFormat
      　　　　　mypicture.TransparentBackground = true
        　　　　mypicture.TransparencyColor = blueScreen
    　　　　myshape.Fill.Visible = false
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PictureFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotItem 对象

## 简介

代表数据透视表字段中的项目。

## 说明

这些项目是某个字段类型中的各个数据项。 PivotItem 对象是 PivotItems 集合的成员。 PivotItems 集合包含某个 PivotField 对象中所有的项目。使用 PivotItems ( *index* )（其中 *index* 是项目索引号或名称）可返回一个 PivotItem 对象。下例隐藏“Sheet3”上第一张数据透视表中“Year”字段中包含“1998”的所有数据项。

``` JavaScript
Worksheets.Item("sheet3").PivotTables(1).PivotFields("year").PivotItems("1998").Visible = false
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
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.Delete">Delete</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>删除对象。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.DrillTo">DrillTo</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>DrillTo 方法支持从 PivotItem 中深化到指定的透视字段。</p></td>
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
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 <span>Application</span> 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.Caption">Caption</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 String 值，它代表数据透视项的标签文本。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.ChildItems">ChildItems</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个对象，它代表一个数据透视表数据项（ <span>PivotItem</span> 对象）或所有数据项的集合（ <span>PivotItems</span> 对象），该集合可以由指定字段中的子项组成，也可以由指定数据项中的子项组成。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.DataRange">DataRange</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>Range</span> 对象，如下表所示。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.DrilledDown">DrilledDown</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定数据透视表字段或数据透视表项的标志设置为“drilled”（展开或可见），则为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.Formula">Formula</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 String 值，它代表 A1 样式表示法和宏语言中的对象的公式。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.IsCalculated">IsCalculated</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>如果数据透视表字段或数据透视表项为计算字段或计算项，则此属性为 True 。 Boolean 类型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.LabelRange">LabelRange</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 <span>Range</span> 对象，它表示数据透视表中所有包含数据项的单元格。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.Name">Name</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 String 值，它代表对象的名称。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.ParentItem">ParentItem</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 PivotItem 对象，该对象表示父 <span>PivotField</span> 对象中的父数据透视表项（必须对该字段分组使其有一个父级）。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.ParentShowDetail">ParentShowDetail</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定数据项是由于其某个父数据项正在显示明细数据而进行显示的，则该值为 True 。如果由于其父数据项隐藏明细数据而使该数据项未显示，则该值为 False 。本属性仅对已分组的数据项有效。 Boolean 类型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.Position">Position</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Long 值，它代表当前正在显示数据项时，该数据项字段中数据项的位置。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.RecordCount">RecordCount</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回数据透视表缓存中的记录数目或包含指定内容的缓存记录数目。 Long 型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">16</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.ShowDetail">ShowDetail</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>如果扩展了指定区域的分级显示（从而行或列的明细数据可见），则为 True 。指定区域必须为分级显示中的单个汇总列或汇总行。 Variant 型，可读写。对于 PivotItem 对象（如果该区域在数据透视表中，则为 Range 对象），当数据项显示明细数据时，此属性设为 True 。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">17</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.SourceName">SourceName</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Variant 值，它代表指定的对象出现在指定的数据透视表的原始源数据中时的名称。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">18</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.SourceNameStandard">SourceNameStandard</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 String 类型的数值，该数值表示以标准英文（美国）格式设置的数据透视表项目的源名称。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">19</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.StandardFormula">StandardFormula</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 String 值，该值指定使用标准英语（美国）格式的公式。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">20</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.Value">Value</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 String 值，它代表数据透视表中指定的数据项的名称。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">21</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotItem.Visible">Visible</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 Boolean 值，它确定对象是否可见。可读写。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PivotItem.Delete

table { word-break:break-all; }

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 PivotItem 对象的变量。

### PivotItem.DrillTo

table { word-break:break-all; }

**DrillTo** 方法支持从 PivotItem 中深化到指定的透视字段。

**语法**

**express.DrillTo ( *PivotItemName* )**

\*express是一个代表 PivotItem 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                          |
|-----------------|-----------|----------|-------------------------------|
| *PivotItemName* | 必选      | String   | 要深化到的 PivotItem 的名称。 |

**说明**

table { word-break:break-all; }

对于 OLAP 数据源，要深化到的透视字段必须与正在深化的 PivotItem 位于同一层次结构中，或者，如果多个属性层次结构彼此相邻地放置在行或列上，则要深化到的透视字段必须是彼此相邻的属性层次结构之一；不可以将用户层次结构放置在正在深化的 PivotItem 的透视字段与要深化到的透视字段之间。如果不符合这些条件，则会返回一个运行时错误。

## 成员属性

### PivotItem.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotItem 对象的变量。

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

### PivotItem.Caption

table { word-break:break-all; }

返回一个 **String** 值，它代表数据透视项的标签文本。

**语法**

**express.Caption**

\*express是一个代表 PivotItem 对象的变量。

**说明**

table { word-break:break-all; }

下表显示 **Caption** 属性及其相关属性的示例值，假设存在一个具有唯一名称为“\[Europe\].\[France\].\[Paris\]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，并且存在一个包含名为“Paris”项的非 OLAP 数据源。

| 属性           | 值（OLAP 数据源）                       | 值（非 OLAP 数据源）        |
|----------------|-----------------------------------------|-----------------------------|
| **Caption**    | Paris                                   | Paris                       |
| **Name**       | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |
| **SourceName** | \[Europe\].\[France\].\[Paris\]（只读） | （与 SQL 属性值相同；只读） |
| **Value**      | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |

在将一个索引值指定到 **PivotItems** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源）                                      | 语法（非 OLAP 数据源）         |
|----------------------------------------------------------|--------------------------------|
| expression.PivotItems("\[Europe\].\[France\].\[Paris\]") | expression.PivotItems("Paris") |

在使用 **Item** 属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源）             | 名称（非 OLAP 数据源） |
|---------------------------------|------------------------|
| \[Europe\].\[France\].\[Paris\] | Paris                  |

### PivotItem.ChildItems

返回一个对象，它代表一个数据透视表数据项（ **PivotItem** 对象）或所有数据项的集合（ **PivotItems** 对象），该集合可以由指定字段中的子项组成，也可以由指定数据项中的子项组成。只读。

**语法**

**express.ChildItems**

\*express是一个代表 PivotItem 对象的变量。

**说明**

此属性对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

此示例将“vegetables”数据项的所有子数据项添加到一张新工作表上的数据清单中。

``` JavaScript
let nwSheet = Worksheets.Add()
nwSheet.Activate()
let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable
let rw = 0
for(let i=1;i <= pvtTable.PivotFields("product").PivotItems("vegetables").ChildItems.Count;i++) {
    rw++
    nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("product").PivotItems("vegetables").ChildItems.Item(i).Name
}
```

### PivotItem.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotItem 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotItem.DataRange

返回一个 **Range** 对象，如下表所示。只读。

**语法**

**express.DataRange**

\*express是一个代表 PivotItem 对象的变量。

**说明**

| 对象             | 数据区域               |
|------------------|------------------------|
| “数据”字段       | 字段中包含的数据       |
| 行、列或者页字段 | 字段中的数据项         |
| 数据项           | 数据项规范所限定的数据 |

**示例**

此示例选定名为“REGION”字段中的数据透视表数据项。

``` JavaScript
let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
Worksheets.Item("Sheet1").Activate()
pvtTable.PivotFields("REGION").DataRange.Select()
```

### PivotItem.DrilledDown

如果指定数据透视表字段或数据透视表项的标志设置为“drilled”（展开或可见），则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DrilledDown**

\*express是一个代表 PivotItem 对象的变量。

**说明**

此属性仅对 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源有效。

如果字段或项被隐藏，则不可设置此属性。

**示例**

此示例将活动工作表上第三个数据透视表中状态字段的所有项的标记设置为“not drilled”。

``` JavaScript
ActiveSheet.PivotTables("PivotTable3").PivotFields("state").DrilledDown = false
```

### PivotItem.Formula

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

**语法**

**express.Formula**

\*express是一个代表 PivotItem 对象的变量。

**说明**

table { word-break:break-all; }

此属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回空字符串。如果单元格包含公式， **Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

### PivotItem.IsCalculated

table { word-break:break-all; }

如果数据透视表字段或数据透视表项为计算字段或计算项，则此属性为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.IsCalculated**

\*express是一个代表 PivotItem 对象的变量。

**说明**

table { word-break:break-all; }

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性始终返回 **False** 。

### PivotItem.LabelRange

table { word-break:break-all; }

返回一个 **Range** 对象，它表示数据透视表中所有包含数据项的单元格。只读。

**语法**

**express.LabelRange**

\*express是一个代表 PivotItem 对象的变量。

### PivotItem.Name

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 PivotItem 对象的变量。

**说明**

table { word-break:break-all; }

下表显示了 **Name** 属性和相关属性的示例值，同时假设有一个具有唯一名称为“\[Europe\].\[France\].\[Paris\]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，有一个包含项名称为“Paris”的非 OLAP 数据源。

| 属性           | 值（OLAP 数据源）                       | 值（非 OLAP 数据源）        |
|----------------|-----------------------------------------|-----------------------------|
| **Caption**    | Paris                                   | Paris                       |
| **Name**       | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |
| **SourceName** | \[Europe\].\[France\].\[Paris\]（只读） | （与 SQL 属性值相同，只读） |
| **Value**      | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |

在将一个索引值指定到 **PivotItems** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源）                                      | 语法（非 OLAP 数据源）         |
|----------------------------------------------------------|--------------------------------|
| expression.PivotItems("\[Europe\].\[France\].\[Paris\]") | expression.PivotItems("Paris") |

在使用 **Item** 属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源）             | 名称（非 OLAP 数据源） |
|---------------------------------|------------------------|
| \[Europe\].\[France\].\[Paris\] | Paris                  |

### PivotItem.Parent

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotItem 对象的变量。

### PivotItem.ParentItem

返回一个 **PivotItem** 对象，该对象表示父 **PivotField** 对象中的父数据透视表项（必须对该字段分组使其有一个父级）。只读。

**语法**

**express.ParentItem**

\*express是一个代表 PivotItem 对象的变量。

**说明**

该属性不可用于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。

**示例**

本示例显示包含活动单元格的项的父项名称。

``` JavaScript
Worksheets.Item("Sheet1").Activate()
MsgBox("This item is a subitem of " + ActiveCell.PivotItem.ParentItem.Name)
```

### PivotItem.ParentShowDetail

如果指定数据项是由于其某个父数据项正在显示明细数据而进行显示的，则该值为 **True** 。如果由于其父数据项隐藏明细数据而使该数据项未显示，则该值为 **False** 。本属性仅对已分组的数据项有效。 **Boolean** 类型，只读。

**语法**

**express.ParentShowDetail**

\*express是一个代表 PivotItem 对象的变量。

**说明**

该属性不可用于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。

**示例**

如果包含活动单元格的项由于其父项正在显示明细数据而可见，则本示例将显示一条消息。

``` JavaScript
Worksheets.Item("Sheet1").Activate()
let pvtItem = ActiveCell.PivotItem
if(pvtItem.ParentShowDetail == true) {
    MsgBox("Parent item is showing detail")
}
```

### PivotItem.Position

返回或设置一个 **Long** 值，它代表当前正在显示数据项时，该数据项字段中数据项的位置。

**语法**

**express.Position**

\*express是一个代表 PivotItem 对象的变量。

**示例**

此示例显示包含活动单元格的数据透视表项的位置数字。

``` JavaScript
Worksheets.Item("Sheet1").Activate()
MsgBox("The active item is in position number " + ActiveCell.PivotItem.Position)
```

### PivotItem.RecordCount

返回数据透视表缓存中的记录数目或包含指定内容的缓存记录数目。 **Long** 型，只读。

**语法**

**express.RecordCount**

\*express是一个代表 PivotItem 对象的变量。

**说明**

此属性反映了查询高速缓存时的瞬时状态。该高速缓存可对不同查询而更改。

**示例**

此示例显示“Products”字段中包含“Kiwi”的缓存记录数目。

``` JavaScript
MsgBox(Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Product").PivotItems("Kiwi").RecordCount)
```

### PivotItem.ShowDetail

table { word-break:break-all; }

如果扩展了指定区域的分级显示（从而行或列的明细数据可见），则为 **True** 。指定区域必须为分级显示中的单个汇总列或汇总行。 **Variant** 型，可读写。对于 **PivotItem** 对象（如果该区域在数据透视表中，则为 **Range** 对象），当数据项显示明细数据时，此属性设为 **True** 。

**语法**

**express.ShowDetail**

\*express是一个代表 PivotItem 对象的变量。

**说明**

table { word-break:break-all; }

此属性不可用于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。

如果指定区域不在数据透视表中，则下列声明为真：

- 指定区域必须在单个汇总行或汇总列中。
- 如果指定行或列的任意子级处于隐藏状态，则此属性返回 **False** 。
- 将此属性设置为 **True** 相当于显示指定汇总行或汇总列的所有子级内容。
- 将此属性设置为 **False** 相当于隐藏汇总行或汇总列的所有子级内容。

如果指定区域为数据透视表，则当该区域连续时，就可以同时对多个单元格设置此属性。仅当指定区域为单个单元格时，才能返回此属性的值。

### PivotItem.SourceName

返回一个 **Variant** 值，它代表指定的对象出现在指定的数据透视表的原始源数据中时的名称。

**语法**

**express.SourceName**

\*express是一个代表 PivotItem 对象的变量。

**说明**

如果用户在创建数据透视表之后重命名数据项，此属性的值可能会与当前数据项名称不同。

下表显示 **SourceName** 属性及其相关属性的示例值，假设存在唯一名称为“\[Europe\].\[France\].\[Paris\]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，并且存在包含名为“Paris”的项的非 OLAP 数据源。

| 属性           | 值（OLAP 数据源）                       | 值（非 OLAP 数据源）        |
|----------------|-----------------------------------------|-----------------------------|
| **Caption**    | Paris                                   | Paris                       |
| **Name**       | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |
| **SourceName** | \[Europe\].\[France\].\[Paris\]（只读） | （与 SQL 属性值相同，只读） |
| **Value**      | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |

在将一个索引值指定到 **PivotItems** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源）                                      | 语法（非 OLAP 数据源）         |
|----------------------------------------------------------|--------------------------------|
| expression.PivotItems("\[Europe\].\[France\].\[Paris\]") | expression.PivotItems("Paris") |

在使用 **Item** 属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源）             | 名称（非 OLAP 数据源） |
|---------------------------------|------------------------|
| \[Europe\].\[France\].\[Paris\] | Paris                  |

**示例**

本示例显示包含活动单元格的数据项的原始名称（在源数据库中的名称）。

``` JavaScript
Worksheets.Item("Sheet1").Activate()
ActiveSheet.PivotTables(1).PivotSelect("1998", xlDataAndLabel)
MsgBox("The original item name is " + ActiveCell.PivotItem.SourceName)
```

### PivotItem.SourceNameStandard

返回一个 **String** 类型的数值，该数值表示以标准英文（美国）格式设置的数据透视表项目的源名称。只读。

**语法**

**express.SourceNameStandard**

\*express是一个代表 PivotItem 对象的变量。

**说明**

当某一项具有本地化版本，并且它的 **SourceNameStandard** 属性值不同于 **SourceName** 属性值（例如带日期格式）时，使用该属性。

**示例**

本示例显示活动数据透视表的第五个字段上第六个项目的源名称。本示例假定数据透视表位于活动工作表上，并且数据源至少包括五个字段，且每个字段至少包括六个项目。

``` JavaScript
function CheckSourceNameStandard() {
    let pvtTable = ActiveSheet.PivotTables(1)
    let pvtField = pvtTable.PivotFields(5)
    let pvtItem = pvtField.PivotItems(6)
    // Display source name.
    MsgBox("The source name is: " + pvtItem.SourceNameStandard)
}
```

### PivotItem.StandardFormula

返回或设置一个 **String** 值，该值指定使用标准英语（美国）格式的公式。可读写。

**语法**

**express.StandardFormula**

\*express是一个代表 PivotItem 对象的变量。

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

### PivotItem.Value

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表数据透视表中指定的数据项的名称。

**语法**

**express.Value**

\*express是一个代表 PivotItem 对象的变量。

### PivotItem.Visible

table { word-break:break-all; }

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 PivotItem 对象的变量。

**说明**

table { word-break:break-all; }

如果数据透视表中的某一项在当前是可见的，则该项的 **Visible** 属性为 **True** 。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotItem](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# RoutingSlip 对象

## 简介

代表工作簿的传送名单。传送名单用于在电子邮件系统中发送工作簿。

## 说明

RoutingSlip 对象不存在，无法返回该对象，除非工作簿的 HasRoutingSlip 属性是 True 。

## 方法

| 序号 | 名称                        | 说明                                                                                                                                                         |
|:----:|:----------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Reset](#RoutingSlip.Reset) | 重新设置传送名单，从而可用同一名单初始化一个新的传送过程（使用同一收件人列表和收件人地址信息）。使用本方法之前，该传送过程必须已完成。否则本方法将导致错误。 |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#RoutingSlip.Application)       | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#RoutingSlip.Creator)               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Delivery](#RoutingSlip.Delivery)             | 返回或设置邮件传递的方法。可以是 XlRoutingSlipDelivery 常量之一。 Long 类型，可读写。                                                                                                                                           |
|  4   | [Message](#RoutingSlip.Message)               | 返回或设置传送名单的消息文字。这些文字将作为传送指定工作簿时的邮件消息内容。 String 类型，可读写。                                                                                                                              |
|  5   | [Parent](#RoutingSlip.Parent)                 | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  6   | [Recipients](#RoutingSlip.Recipients)         | 返回或设置传送名单上的收件人。                                                                                                                                                                                                  |
|  7   | [ReturnWhenDone](#RoutingSlip.ReturnWhenDone) | 如果传送过程结束后，工作簿返回给发件人，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                             |
|  8   | [Status](#RoutingSlip.Status)                 | 指示传送名单的状态。 XlRoutingSlipStatus 类型，只读。                                                                                                                                                                           |
|  9   | [Subject](#RoutingSlip.Subject)               | 返回或设置邮件或传送名单的主题。 String 型，可读写。                                                                                                                                                                            |
|  10  | [TrackStatus](#RoutingSlip.TrackStatus)       | 如果允许对传送名单进行状态追踪，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                     |

## 成员方法

### RoutingSlip.Reset

重新设置传送名单，从而可用同一名单初始化一个新的传送过程（使用同一收件人列表和收件人地址信息）。使用本方法之前，该传送过程必须已完成。否则本方法将导致错误。

**语法**

**express.Reset ()**

\*express是一个代表 RoutingSlip 对象的变量。

**返回值**

Variant

**示例**

``` JavaScript
/*如果工作簿 Book1.xls 的传送过程已完成，本示例将重新设置传送名单。*/
function test(){
　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip
       if(slip.Status == xlRoutingComplete) {
　　　　    slip.Reset()
　　　　}
　　　　else {
　　　　    MsgBox("Cannot reset routing; not yet complete")
　　　　}
}
```

## 成员属性

### RoutingSlip.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 RoutingSlip 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
　　　　let myObject = ActiveWorkbook
       if(myObject.Application.Value == "ET") {
 　　　　   MsgBox("This is an ET Application object.")
　　　　}
　　　　else{
　　　　    MsgBox("This is not an ET Application object.")
　　　　}　　　　
}
```

### RoutingSlip.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 RoutingSlip 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### RoutingSlip.Delivery

返回或设置邮件传递的方法。可以是 **XlRoutingSlipDelivery** 常量之一。 **Long** 类型，可读写。

**语法**

**express.Delivery**

\*express是一个代表 RoutingSlip 对象的变量。

**示例**

``` JavaScript
/*本示例将 Book1.xls 依次发送给三个收件人。*/
function test(){
　　　　Workbooks.Item("BOOK1.XLS").HasRoutingSlip = true
　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip
　　　　slip.Delivery = xlOneAfterAnother
　　　　slip.Recipients = ["Adam Bendel","Jean Selva", "Bernard Gabor"]
　　　　slip.Subject = "Here is BOOK1.XLS"
　　　　slip.Message = "Here is the workbook. What do you think?"
　　　　Workbooks.Item("BOOK1.XLS").Route()
}
```

### RoutingSlip.Message

返回或设置传送名单的消息文字。这些文字将作为传送指定工作簿时的邮件消息内容。 **String** 类型，可读写。

**语法**

**express.Message**

\*express是一个代表 RoutingSlip 对象的变量。

**示例**

``` JavaScript
/*本示例将 Book1.xls 依次发送给三个收件人。*/
function test(){
　　　　Workbooks.Item("BOOK1.XLS").HasRoutingSlip = true
　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip
　　　　slip.Delivery = xlOneAfterAnother
　　　　slip.Recipients = ["Adam Bendel","Jean Selva","Bernard Gabor"]
　　　　slip.Subject = "Here is BOOK1.XLS"
　　　　slip.Message = "Here is the workbook. What do you think?"
　　　　Workbooks.Item("BOOK1.XLS").Route()
}
```

### RoutingSlip.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 RoutingSlip 对象的变量。

### RoutingSlip.Recipients

返回或设置传送名单上的收件人。

**语法**

**express.Recipients**

\*express是一个代表 RoutingSlip 对象的变量。

**说明**

如果传送过程的递送选项设置为 **xlOneAfterAnother** ，则收件人列表的顺序就定义了传递的顺序。

如果某个传送名单正在进行中，则只会返回或设置那些还没有收到和传送该文档的收件人。

**示例**

``` JavaScript
/*本示例将打开的工作簿顺次发送给三个收件人。*/
function test(){
　　　　let book= ThisWorkbook
　　　　book.HasRoutingSlip = true
　　　　let slip= book.RoutingSlip
　　　　slip.Delivery = xlOneAfterAnother
　　　　slip.Recipients = ["Adam Bendel","Jean Selva","Bernard Gabor"]
　　　　slip.Subject = "Here is the workbook"
　　　　slip.Message = "Here is the workbook. What do you think?"
　　　　slip.ReturnWhenDone = true
　　　　book.Route()
}
```

### RoutingSlip.ReturnWhenDone

如果传送过程结束后，工作簿返回给发件人，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ReturnWhenDone**

\*express是一个代表 RoutingSlip 对象的变量。

**说明**

如果正在进行传送，则无法设置该属性。

**示例**

``` JavaScript
/*本示例将工作簿 Book1.xls 按次序发送给三个收件人，传送过程结束后，再将工作簿返回给发件人。*/
function test(){
　　　　Workbooks.Item("BOOK1.XLS").HasRoutingSlip = true
　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip
　　　　slip.Delivery = xlOneAfterAnother
　　　　slip.Recipients = ["Adam Bendel","Jean Selva","Bernard Gabor"]
　　　　slip.Subject = "Here is BOOK1.XLS"
　　　　slip.Message = "Here is the workbook. What do you think?"
　　　　slip.ReturnWhenDone = true
　　　　Workbooks.Item("BOOK1.XLS").Route()
}
```

### RoutingSlip.Status

指示传送名单的状态。 **XlRoutingSlipStatus** 类型，只读。

**语法**

**express.Status**

\*express是一个代表 RoutingSlip 对象的变量。

**说明**

| XlRoutingSlipStatus 可为以下 XlRoutingSlipStatus 常量之一。 |
|-------------------------------------------------------------|
| xlNotYetRouted                                              |
| xlRoutingComplete                                           |
| xlRoutingInProgress                                         |

**示例**

``` JavaScript
/*如果工作簿 Book1.xls 的传送过程已完成，本示例将重新设置传送名单。*/
function test(){
　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip
       if(slip.Status == xlRoutingComplete) {
   　　　　 slip.Reset()
　　　　}
　　　　else {
 　　　　   MsgBox("Cannot reset routing; not yet complete.")
　　　　}
}
```

### RoutingSlip.Subject

返回或设置邮件或传送名单的主题。 **String** 型，可读写。

**语法**

**express.Subject**

\*express是一个代表 RoutingSlip 对象的变量。

**说明**

**RoutingSlip** 对象的主题被用作传送工作簿时的邮件信息的主题。

**示例**

``` JavaScript
/*此示例设置打开的工作簿中传送名单的主题。此示例只能在已安装了 Microsoft Exchange 的情况下才能运行。*/
function test(){
　　　　let book= ThisWorkbook
　　　　book.HasRoutingSlip = true
　　　　let slip=book.RoutingSlip
　　　　slip.Delivery = xlOneAfterAnother
　　　　slip.Recipients = ["Adam Bendel","Jean Selva","Bernard Gabor"]
　　　　slip.Subject = "Here is the workbook"
　　　　slip.Message = "Here is the workbook. What do you think?"
　　　　slip.ReturnWhenDone = true
　　　　book.Route()
}
```

### RoutingSlip.TrackStatus

如果允许对传送名单进行状态追踪，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.TrackStatus**

\*express是一个代表 RoutingSlip 对象的变量。

**说明**

如果正在进行传送，则无法设置该属性。

**示例**

``` JavaScript
/*本示例将工作簿 Book1.xls 发送给三个收件人，并允许进行状态追踪。*/
function test(){
　　　　Workbooks.Item("BOOK1.XLS").HasRoutingSlip = true
　　　　let slip= Workbooks.Item("BOOK1.XLS").RoutingSlip
　　　　slip.Delivery = xlOneAfterAnother
　　　　slip.Recipients = ["Adam Bendel","Jean Selva", "Bernard Gabor"]
　　　　slip.Subject = "Here is BOOK1.XLS"
　　　　slip.Message = "Here is the workbook. What do you think?"
　　　　slip.ReturnWhenDone = true
　　　　slip.TrackStatus = true
　　　　Workbooks.Item("BOOK1.XLS").Route()
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/RoutingSlip](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Scenarios 对象

## 简介

指定工作表上所有 Scenario 对象的集合。

## 说明

方案是一组经过命名和保存的输入值（被称为“可变单元格”）。

## 方法

| 序号 | 名称                                      | 说明                                                                        |
|:----:|:------------------------------------------|:----------------------------------------------------------------------------|
|  1   | [Add](#Scenarios.Add)                     | 新建方案并将其添加到当前工作表可用的方案列表中。                            |
|  2   | [CreateSummary](#Scenarios.CreateSummary) | 新建一个工作表，并在其中显示指定工作表中所有方案的摘要报表。 Variant 类型。 |
|  3   | [Item](#Scenarios.Item)                   | 从集合中返回一个对象。                                                      |
|  4   | [Merge](#Scenarios.Merge)                 | 将另一张工作表中的方案合并到 Scenarios 集合中。                             |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Scenarios.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Scenarios.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Scenarios.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Scenarios.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Scenarios.Add

新建方案并将其添加到当前工作表可用的方案列表中。

**语法**

**express.Add ( *Name* , *ChangingCells* , *Values* , *Comment* , *Locked* , *Hidden* )**

\*express是一个代表 Scenarios 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                     |
|-----------------|-----------|----------|----------------------------------------------------------------------------------------------------------|
| *Name*          | 必选      | String   | 方案名。                                                                                                 |
| *ChangingCells* | 必选      | Variant  | 一个 Range 对象，它引用方案的可变单元格。                                                                |
| *Values*        | 可选      | Variant  | 包含 ChangingCells 中单元格方案值的数组。如果省略此参数，就假定方案值是 ChangingCells 中单元格的当前值。 |
| *Comment*       | 可选      | Variant  | 一个指定方案批注文本的字符串。添加新方案时，作者的姓名和日期自动添加在批注文本的开始部分。               |
| *Locked*        | 可选      | Variant  | 如果为 True，则锁定方案以防更改。默认值为 True。                                                         |
| *Hidden*        | 可选      | Variant  | 如果为 True，则隐藏方案。默认值为 False。                                                                |

**返回值**

一个代表新方案的 Scenario 对象。

**说明**

方案名称必须是唯一的；如果试图用已经在使用的名称创建方案，ET 会产生错误。

**示例**

``` JavaScript
//本示例向工作表 Sheet1 添加新方案。
Application.Worksheets.Item("Sheet1").Scenarios.Add("Best Case",Worksheets("Sheet1").Range("A1:A4"), [23, 5, 6, 21],"Most favorable outcome.")
```

### Scenarios.CreateSummary

新建一个工作表，并在其中显示指定工作表中所有方案的摘要报表。 **Variant** 类型。

**语法**

**express.CreateSummary ( *ReportType* , *ResultCells* )**

\*express是一个代表 Scenarios 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型            | 说明                                                                                                                                                                                                                              |
|---------------|-----------|---------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ReportType*  | 可选      | XlSummaryReportType | 指定摘要报表是数据透视表还是标准摘要。                                                                                                                                                                                            |
| *ResultCells* | 可选      | Variant             | 一个代表指定工作表中计算结果单元格的 Range 对象。一般而言，此区域指一个或多个包含公式的单元格，即显示特定方案计算结果的单元格；这些单元格中的公式的值由模型中的可变单元格值计算而得。如果省略此参数，报表中将没有任何结果单元格。 |

**返回值**

Variant

**示例**

``` JavaScript
//本示例用工作表 Sheet1 上单元格区域 C4:C9 的计算结果创建 Sheet1 上方案的摘要报表。
Application.Worksheets.Item("Sheet1").Scenarios.CreateSummary(undefined, Worksheets.Item("Sheet1").Range("C4:C9"))
```

### Scenarios.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Scenarios 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 Scenario 对象。

**示例**

``` JavaScript
//本示例在名为“Options”的工作表上显示名为“Typical”的方案。
Application.Worksheets.Item("options").Scenarios.Item("typical").Show()
```

### Scenarios.Merge

将另一张工作表中的方案合并到 **Scenarios** 集合中。

**语法**

**express.Merge ( *Source* )**

\*express是一个代表 Scenarios 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                          |
|----------|-----------|----------|---------------------------------------------------------------|
| *Source* | 必选      | Variant  | 待合并方案所在工作表的名称，或代表该工作表的 Worksheet 对象。 |

**返回值**

Variant

**说明**

合并区域的值在该区域左上角的单元格中指定。

## 成员属性

### Scenarios.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Scenarios 对象的变量。

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

### Scenarios.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Scenarios 对象的变量。

### Scenarios.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Scenarios 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Scenarios.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Scenarios 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Scenarios](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SortField 对象

## 简介

SortField 对象包含 Worksheet 、 Lists 和 AutoFilter 对象的所有排序信息。

## 说明

开发人员可以使用 BeforeSort 事件重写 ET 的默认行为，将自己的排序算法写入应用程序中。

## 方法

| 序号 | 名称                              | 说明                                            |
|:----:|:----------------------------------|:------------------------------------------------|
|  1   | [Delete](#SortField.Delete)       | 从 SortFields 集合中删除指定的 SortField 对象。 |
|  2   | [ModifyKey](#SortField.ModifyKey) | 修改字段中按其排序的键值。                      |
|  3   | [SetIcon](#SortField.SetIcon)     | 为 SortField 对象设置图标。                     |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                               |
|:----:|:--------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#SortField.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#SortField.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [CustomOrder](#SortField.CustomOrder) | 指定对字段进行排序的自定义次序。可读/写 Variant 类型。                                                                                                             |
|  4   | [DataOption](#SortField.DataOption)   | 指定 SortField 对象所指定的区域中的文本排序方式。可读/写 XlSortDataOption 类型。                                                                                   |
|  5   | [Key](#SortField.Key)                 | 指定排序字段，该字段确定要排序的值。只读。                                                                                                                         |
|  6   | [Order](#SortField.Order)             | 确定关键字所指定的值的排序次序。可读/写。                                                                                                                          |
|  7   | [Parent](#SortField.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  8   | [Priority](#SortField.Priority)       | 指定排序字段的优先级。可读/写。                                                                                                                                    |
|  9   | [SortOn](#SortField.SortOn)           | 返回或设置要排序的单元格属性。可读/写 XlSortOn 类型。                                                                                                              |
|  10  | [SortOnValue](#SortField.SortOnValue) | 返回指定 SortField 对象执行排序的值。只读。                                                                                                                        |

## 成员方法

### SortField.Delete

从 **SortFields** 集合中删除指定的 **SortField** 对象。

**语法**

**express.Delete ()**

\*express是一个代表 SortField 对象的变量。

### SortField.ModifyKey

修改字段中按其排序的键值。

**语法**

**express.ModifyKey ( *Key* )**

\*express是一个代表 SortField 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明             |
|-------|-----------|----------|------------------|
| *Key* | 必选      | Range    | 指定要修改的键。 |

### SortField.SetIcon

为 **SortField** 对象设置图标。

**语法**

**express.SetIcon ( *Icon* )**

\*express是一个代表 SortField 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明           |
|--------|-----------|----------|----------------|
| *Icon* | 必选      | Icon     | 要设置的图标。 |

## 成员属性

### SortField.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 SortField 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### SortField.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 SortField 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### SortField.CustomOrder

指定对字段进行排序的自定义次序。可读/写 **Variant** 类型。

**语法**

**express.CustomOrder**

\*express是一个代表 SortField 对象的变量。

### SortField.DataOption

指定 **SortField** 对象所指定的区域中的文本排序方式。可读/写 XlSortDataOption 类型。

**语法**

**express.DataOption**

\*express是一个代表 SortField 对象的变量。

**说明**

**xlSortDataOption** 可以是 **xlSortNormal** 或 **xlSortTextAsNumbers** 。

### SortField.Key

指定排序字段，该字段确定要排序的值。只读。

**语法**

**express.Key**

\*express是一个代表 SortField 对象的变量。

**说明**

关键字可以是区域名称（字符串）或 **Range** 对象。

### SortField.Order

确定关键字所指定的值的排序次序。可读/写。

**语法**

**express.Order**

\*express是一个代表 SortField 对象的变量。

### SortField.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SortField 对象的变量。

### SortField.Priority

指定排序字段的优先级。可读/写。

**语法**

**express.Priority**

\*express是一个代表 SortField 对象的变量。

### SortField.SortOn

返回或设置要排序的单元格属性。可读/写 **XlSortOn** 类型。

**语法**

**express.SortOn**

\*express是一个代表 SortField 对象的变量。

**说明**

|                                             |
|---------------------------------------------|
| XlSortOn 可以是下列 **XlSortOn** 常量之一： |
| **SortOnCellColor**                         |
| **SortOnFontColor**                         |
| **SortOnIcon**                              |
| **SortOnValues**                            |

### SortField.SortOnValue

返回指定 **SortField** 对象执行排序的值。只读。

**语法**

**express.SortOnValue**

\*express是一个代表 SortField 对象的变量。

**说明**

**SortOnValue** 可与值、单元格颜色、字体颜色和单元格图标一起使用。

**示例**

``` JavaScript
function test(){
  /*本示例对 B 列和 Sheet1 中的数据按字体颜色以升序排序*/
  Application.ActiveWorkbook.Worksheets.Item("Sheet1").Sort.SortFields.Clear()
  Application.ActiveWorkbook.Worksheets.Item("Sheet1").Sort.SortFields.Add(Range("B1:B25"), xlSortOnFontColor, xlAscending, " ", xlSortNormal).SortOnValue.Color.RGB = (0, 0, 0)
                                    
  let sor = Application.ActiveWorkbook.Worksheets.Item("Sheet1").Sort
  sor.SetRange(Range("A1:B25"))
  sor.Header = xlGuess
  sor.MatchCase = false
  sor.Orientation = xlTopToBottom
  sor.SortMethod = xlPinYin
  sor.Apply()
  
  /*单元格颜色*/
  let SortOn = xlSortOnCellColor
  SortOnValue.Color.RGB = (255, 255, 0)
  
  /*字体颜色*/
  SortOn = xlSortOnFontColor
  SortOnValue.Color.RGB = (255, 255, 0)
  
  /*图标*/
  let SortOn = xlSortOnIcon
  SortOnValue.Color.RGB = (255, 255, 0)
  SortField.SetIcon(ActiveWorkbook.IconSets(1).Item(3))
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/SortField](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TableStyle 对象

## 简介

代表可应用于表格或切片器的单个样式。

## 说明

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，列是表格的元素。表格样式可以规定使用交替格式（也称为条带或条纹）对表格中的列进行格式设置。

## 方法

| 序号 | 名称                               | 说明                                             |
|:----:|:-----------------------------------|:-------------------------------------------------|
|  1   | [Delete](#TableStyle.Delete)       | 删除 TableStyle 对象。                           |
|  2   | [Duplicate](#TableStyle.Duplicate) | 复制 TableStyle 对象，并返回对新复制对象的引用。 |

## 属性

| 序号 | 名称                                                                         | 说明                                                                                                                                                               |
|:----:|:-----------------------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#TableStyle.Application)                                       | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [BuiltIn](#TableStyle.BuiltIn)                                               | 如果样式为内置样式，则为 True 。只读 Boolean 类型。                                                                                                                |
|  3   | [Creator](#TableStyle.Creator)                                               | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Name](#TableStyle.Name)                                                     | 返回对象的名称。只读 String 类型。                                                                                                                                 |
|  5   | [NameLocal](#TableStyle.NameLocal)                                           | 以用户语言返回或设置对象的名称。只读 String 类型。                                                                                                                 |
|  6   | [Parent](#TableStyle.Parent)                                                 | 返回指定对象的父对象。只读。                                                                                                                                       |
|  7   | [ShowAsAvailablePivotTableStyle](#TableStyle.ShowAsAvailablePivotTableStyle) | 设置或返回一个值，该值表示是否在数据透视表样式库中显示样式，可读/写 Boolean 类型。                                                                                 |
|  8   | [ShowAsAvailableSlicerStyle](#TableStyle.ShowAsAvailableSlicerStyle)         | 返回或设置指定的表样式是否作为可用的表样式在切片器样式库中显示。可读/写                                                                                            |
|  9   | [ShowAsAvailableTableStyle](#TableStyle.ShowAsAvailableTableStyle)           | 返回或设置在表样式库中显示为可用的表样式。可读/写 Boolean 类型。                                                                                                   |
|  10  | [TableStyleElements](#TableStyle.TableStyleElements)                         | 返回 TableStyleElements 对象，只读。                                                                                                                               |

## 成员方法

### TableStyle.Delete

删除 **TableStyle** 对象。

**语法**

**express.Delete ()**

\*express是一个代表 TableStyle 对象的变量。

**返回值**

无

### TableStyle.Duplicate

复制 **TableStyle** 对象，并返回对新复制对象的引用。

**语法**

**express.Duplicate ( *NewTableStyleName* )**

\*express是一个代表 TableStyle 对象的变量。

**参数**

| 名称                | 必选/可选 | 数据类型 | 说明             |
|---------------------|-----------|----------|------------------|
| *NewTableStyleName* | 必选      | Variant  | 新表样式的名称。 |

**返回值**

TableStyle

**说明**

如果未指定名称，则 **Duplicate** 方法将使用与用户界面默认名称相同的默认名称。

## 成员属性

### TableStyle.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 TableStyle 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### TableStyle.BuiltIn

如果样式为内置样式，则为 **True** 。只读 **Boolean** 类型。

**语法**

**express.BuiltIn**

\*express是一个代表 TableStyle 对象的变量。

### TableStyle.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 TableStyle 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### TableStyle.Name

返回对象的名称。只读 **String** 类型。

**语法**

**express.Name**

\*express是一个代表 TableStyle 对象的变量。

### TableStyle.NameLocal

以用户语言返回或设置对象的名称。只读 **String** 类型。

**语法**

**express.NameLocal**

\*express是一个代表 TableStyle 对象的变量。

**说明**

如果样式为内置样式，则此属性以当前系统所使用的地区语言返回样式的名称。

### TableStyle.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 TableStyle 对象的变量。

### TableStyle.ShowAsAvailablePivotTableStyle

设置或返回一个值，该值表示是否在数据透视表样式库中显示样式，可读/写 **Boolean** 类型。

**语法**

**express.ShowAsAvailablePivotTableStyle**

\*express是一个代表 TableStyle 对象的变量。

**说明**

如果在数据透视表样式库中显示样式，则返回 **True** 。

| 注释                                                                                                                                                                                                                                           |
|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 即使已将样式应用于表格或数据透视表，用户仍可以将 **ShowAsAvailableTableStyle** 或 **ShowAsAvailablePivotTableStyle** 属性设置为 **False** 。在这种情况下，当活动单元格位于表格或数据透视表中时，库将不会显示样式并且没有样式显示为选中的样式。 |

### TableStyle.ShowAsAvailableSlicerStyle

返回或设置指定的表样式是否作为可用的表样式在切片器样式库中显示。可读/写

**语法**

**express.ShowAsAvailableSlicerStyle**

\*express是一个代表 TableStyle 对象的变量。

### TableStyle.ShowAsAvailableTableStyle

返回或设置在表样式库中显示为可用的表样式。可读/写 **Boolean** 类型。

**语法**

**express.ShowAsAvailableTableStyle**

\*express是一个代表 TableStyle 对象的变量。

**说明**

如果为 **True** ，则此样式显示在表样式库中。

即使样式已应用于表，也可以将此属性设置为 **False** 。这种情况下，库中将不会显示该样式，活动单元格在该表中时，样式状态不会显示为选中。

### TableStyle.TableStyleElements

返回 **TableStyleElements** 对象，只读。

**语法**

**express.TableStyleElements**

\*express是一个代表 TableStyle 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/TableStyle](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# AddIns 对象

## 简介

卸载所有已加载的加载项，并根据 *RemoveFromList* 参数值删除 AddIns 集合中相应的加载项。

## 方法

| 序号 | 名称                     | 说明                                                                                   |
|:----:|:-------------------------|:---------------------------------------------------------------------------------------|
|  1   | [Add](#AddIns.Add)       | 返回一个 AddIn 对象，该对象表示一个添加到可用加载项列表中的加载项。                    |
|  2   | [Item](#AddIns.Item)     | 返回集合中的单个对象。                                                                 |
|  3   | [Unload](#AddIns.Unload) | 卸载所有已加载的加载项，并根据 *RemoveFromList* 参数值删除 AddIns 集合中相应的加载项。 |

## 属性

| 序号 | 名称                               | 说明                                                                                |
|:----:|:-----------------------------------|:------------------------------------------------------------------------------------|
|  1   | [Application](#AddIns.Application) | 返回一个 Application 对象，该对象代表 WPS 应用程序。                                |
|  2   | [Count](#AddIns.Count)             | 返回 AddIns 集合中 AddIn 对象的数目。 Long 类型，只读。                             |
|  3   | [Creator](#AddIns.Creator)         | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。        |
|  4   | [Parent](#AddIns.Parent)           | 返回一个 Object 类型的值，该值代表 Addins 集合中的父对象。通常为 Application 对象。 |

## 成员方法

### AddIns.Add

返回一个 **AddIn** 对象，该对象表示一个添加到可用加载项列表中的加载项。

**语法**

**express.Add ( *FileName* , *Install* )**

\*express是一个代表 AddIns 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                 |
|------------|-----------|----------|------------------------------------------------------------------------------------------------------|
| *FileName* | 必选      | String   | 模板或 WLL 的路径。                                                                                  |
| *Install*  | 可选      | Variant  | 如果为 True，则安装加载项；如果为 False，则将加载项添加到加载项列表中，但不进行安装。默认值为 True。 |

**说明**

可以使用加载项的 **Installed** 属性查看该项是否已经安装。

**示例**

``` JavaScript
/*本示例安装名为 MyFax.dot 的模板并将其添加至“模板和加载项”对话框中的加载项列表中。 */
function test()
{
 /* For this example to work correctly, verify that the
       path is correct and the file exists. */

    Application.AddIns.Add("C:\\Program Files\\Microsoft Office\\Templates\\Letters Faxes\\MyFax.dot",true)

}
```

### AddIns.Item

返回集合中的单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 AddIns 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                   |
|---------|-----------|----------|------------------------------------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个加载项。Index 可以是代表加载项在集合中的排序位置的 Long 类型值，或者是代表单个加载项名称的 String 类型值。 |

### AddIns.Unload

卸载所有已加载的加载项，并根据 *RemoveFromList* 参数值删除 **AddIns** 集合中相应的加载项。

**语法**

**express.Unload ( *RemoveFromList* )**

\*express是一个代表 AddIns 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型       | 说明                                                                                                                                                                                                                                                           |
|------------------|-----------|----------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *RemoveFromList* | 必选      | RemoveFromList | 如果为 True，则从 AddIns 集合中删除已卸载的加载项（从“模板和加载项”对话框中删除名称）。如果为 False，则保留集合中已卸载的加载项。如果已卸载的加载项的 Autoload 属性返回 True，则 Unload 无法从 AddIns 集合中删除该加载项，而不考虑 RemoveFromList 的值是什么。 |

**说明**

要卸载单个模板或 WLL，请将 **AddIn** 对象的 **Installed** 属性设置为 **False** 。要从 **AddIns** 集合中删除单个模板或 WLL，请将 **Delete** 方法应用于 **AddIn** 对象。

**示例**

``` JavaScript
/*此示例卸载“模板和加载项”对话框中列出的所有加载项。而加载项的名称还保留在 AddIns 集合中。*/
function test() {
    if (Application.AddIns.Count > 0) {
        Application.AddIns.Unload(false)
    }
}
```

## 成员属性

### AddIns.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 AddIns 对象的变量。

### AddIns.Count

返回 **AddIns** 集合中 **AddIn** 对象的数目。 **Long** 类型，只读。

**语法**

**express.Count**

\*express是一个代表 AddIns 对象的变量。

### AddIns.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 AddIns 对象的变量。

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

### AddIns.Parent

返回一个 **Object** 类型的值，该值代表 **Addins** 集合中的父对象。通常为 **Application** 对象。

**语法**

**express.Parent**

\*express是一个代表 AddIns 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/AddIns](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Break 对象

## 简介

代表页面中的单个分页符、分栏符和分节符。使用 Break 对象及相关方法和属性 可通过编程方式定义文档的页面版式。

## 说明

使用 Breaks 集合的 Item 方法返回特定 Break 对象。以下示例返回活动文档第一页中的第一个分隔符。

``` JavaScript
Application.ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1).Breaks.Item(1)
```

## 属性

| 序号 | 名称                              | 说明                                                                         |
|:----:|:----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Break.Application) | 返回一个 Application 对象，该对象代表 WPS 应用程序。                         |
|  2   | [Creator](#Break.Creator)         | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。 |
|  3   | [PageIndex](#Break.PageIndex)     | 返回一个 Long 类型的值，该值代表指定分隔符出现的页码。                       |
|  4   | [Parent](#Break.Parent)           | 返回一个 Object 类型的值，该值代表指定 Break 对象的父对象。                  |
|  5   | [Range](#Break.Range)             | 返回一个 Range 对象，该对象代表指定对象所含的部分文档。                      |

## 成员属性

### Break.Application

返回一个 Application 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Break 对象的变量。

### Break.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Break 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

### Break.PageIndex

返回一个 **Long** 类型的值，该值代表指定分隔符出现的页码。

**语法**

**express.PageIndex**

\*express是一个代表 Break 对象的变量。

**示例**

``` JavaScript
/*下列代码返回指定分隔符出现的页码。*/
Application.ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1).Breaks.Item(1).PageIndex
```

### Break.Parent

返回一个 **Object** 类型的值，该值代表指定 **Break** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Break 对象的变量。

### Break.Range

返回一个 Range 对象，该对象代表指定对象所含的部分文档。

**语法**

**express.Range**

\*express是一个代表 Break 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Break](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Cell 对象

## 简介

代表单个表格单元格。Cell 对象是 Cells 集合的成员。 Cells 集合代表指定对象中所有的单元格。

## 说明

若要返回 Cell 对象，应使用 Table 对象的 Cell ( *row* , *column* )（其中 *row* 是行号， *column* 是列号）；或使用 Cells ( *index* )（其中 *index* 是索引编号）。以下示例对第一行的第二个单元格应用底纹。

``` JavaScript
function test() {
    let myCell = Application.ActiveDocument.Tables.Item(1).Cell(1,2)
    myCell.Shading.Texture = wdTexture20Percent
}
```

以下示例对第一行的第一个单元格应用底纹。

``` JavaScript
Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Item(1).Shading.Texture = wdTexture20Percent
```

使用 Add 方法将 Cell 对象添加到 Cells 集合。还可以使用 Selection 对象的 InsertCells 方法插入新的单元格。以下示例在 ` myTable ` 的第一个单元格之前添加一个单元格。

``` JavaScript
function test() {
    let myTable = Application.ActiveDocument.Tables.Item(1)
    myTable.Range.Cells.Add(myTable.Cell(1, 1))
}
```

以下示例设置引用第一个表格的头两个单元格的区域 ( *myRange* )。设置此区域后，通过 Merge 方法合并单元格。

``` JavaScript
function test() {
    let myTable = Application.ActiveDocument.Tables.Item(1)
    let myRange = Application.ActiveDocument.Range(myTable.Cell(1, 1).Range.Start, myTable.Cell(1, 2).Range.End)
    myRange.Cells.Merge()
}
```

可将 Add 方法与 Rows 或 Columns 集合配合使用，添加一行或一列单元格。

可将 Information 属性与 Selection 对象配合使用，返回当前行号和列号。以下示例更改选定内容中第一个单元格的宽度，然后显示该单元格的行号和列号。

``` JavaScript
function test() {
    if(Application.Selection.Information(wdWithInTable) == true) {             Application.Selection.Cells.Item(1).Width = 22
        alert("Cell " + Application.Selection.Information(wdStartOfRangeRowNumber) + "," + Application.Selection.Information(wdStartOfRangeColumnNumber))
    }
}
```

## 方法

| 序号 | 名称                         | 说明                                                                                                |
|:----:|:-----------------------------|:----------------------------------------------------------------------------------------------------|
|  1   | [AutoSum](#Cell.AutoSum)     | 插入一个 = (Formula) 域，以计算和显示表达式指定的位于表格某一单元格上方或左侧的各单元格的数值之和。 |
|  2   | [Delete](#Cell.Delete)       | 删除一个或多个表格单元格并可选择控制如何移动剩余的单元格。                                          |
|  3   | [Formula](#Cell.Formula)     | 将包含指定公式的 = (Formula) 域插入到单元格中。                                                     |
|  4   | [Merge](#Cell.Merge)         | 将指定单元格与另一表格单元格合并，成为一个单独的表格单元格。                                        |
|  5   | [Select](#Cell.Select)       | 选择指定的对象。                                                                                    |
|  6   | [SetHeight](#Cell.SetHeight) | 设置表格单元格的高度。                                                                              |
|  7   | [SetWidth](#Cell.SetWidth)   | 设置表格列或单元格的宽度                                                                            |
|  8   | [Split](#Cell.Split)         | 将一个表格单元格拆分为多个。                                                                        |

## 属性

| 序号 | 名称                                           | 说明                                                                                                   |
|:----:|:-----------------------------------------------|:-------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Cell.Application)               | 返回一个 Application 对象，该对象代表 WPS 应用程序。                                                   |
|  2   | [Borders](#Cell.Borders)                       | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                                  |
|  3   | [BottomPadding](#Cell.BottomPadding)           | 返回或设置在表格中单独的单元格或所有单元格内容的下方添加的间距（以磅为单位）。 Single 类型，可读写。   |
|  4   | [Column](#Cell.Column)                         | 返回一个 Column 对象，该对象代表包含指定单元格的表格列。只读。                                         |
|  5   | [ColumnIndex](#Cell.ColumnIndex)               | 返回包含指定单元格的表格列数。 Long 类型，只读。                                                       |
|  6   | [Creator](#Cell.Creator)                       | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 Long 类型，只读。                           |
|  7   | [FitText](#Cell.FitText)                       | 如果 WPS 会减小键入单元格中的文字的可视大小，使其适应列宽，则该属性值为 True 。 Boolean 类型，可读写。 |
|  8   | [Height](#Cell.Height)                         | 返回或设置指定表格单元格的高度。                                                                       |
|  9   | [HeightRule](#Cell.HeightRule)                 | 返回或设置一个 WdRowHeightRule 常量，该常量代表确定指定单元格或行高度的规则。可读写。                  |
|  10  | [Id](#Cell.Id)                                 | 在当前文档保存为网页时，返回或设置指定对象的识别标签。 String 类型，可读写。                           |
|  11  | [LeftPadding](#Cell.LeftPadding)               | 返回或设置在表格的单个单元格或所有单元格的内容左侧的间距（以磅为单位）。 Single 类型，可读写。         |
|  12  | [NestingLevel](#Cell.NestingLevel)             | 返回指定单元格的嵌套层。 Long 类型，只读。                                                             |
|  13  | [Next](#Cell.Next)                             | 返回一个 Cell 对象，该对象代表 Cells 集合中的下一个表格单元格。只读。                                  |
|  14  | [Parent](#Cell.Parent)                         | 返回一个 Object 类型的值，该值代表指定 Cell 对象的父对象。                                             |
|  15  | [PreferredWidth](#Cell.PreferredWidth)         | 返回或设置指定单元格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。 Single 类型，可读写。           |
|  16  | [PreferredWidthType](#Cell.PreferredWidthType) | 返回或设置用于指定单元格的宽度的首选度量单位。 WdPreferredWidthType 类型，只读。                       |
|  17  | [Previous](#Cell.Previous)                     | 返回一个 Cell 对象，该对象代表 Cells 集合中的前一个表格单元格。只读。                                  |
|  18  | [Range](#Cell.Range)                           | 返回一个 Range 对象，该对象代表指定对象所含的部分文档。                                                |
|  19  | [RightPadding](#Cell.RightPadding)             | 返回或设置表格中某个单元格或全部单元格内容右侧的空格量（以磅为单位）。 Single 类型，可读写。           |
|  20  | [Row](#Cell.Row)                               | 返回一个 Row 对象，该对象代表指定单元格所在的行。                                                      |
|  21  | [RowIndex](#Cell.RowIndex)                     | 返回指定单元格所在行的行数。 Long 类型，只读。                                                         |
|  22  | [Shading](#Cell.Shading)                       | 返回一个 Shading 对象，该对象指明指定对象的底纹格式。                                                  |
|  23  | [Tables](#Cell.Tables)                         | 返回一个 Tables 集合，该集合代表指定表格单元格中的所有嵌套表格。只读。                                 |
|  24  | [TopPadding](#Cell.TopPadding)                 | 返回或设置表格中单独的单元格或全部单元格的内容上方要增加的间距（以磅为单位）。 Single 类型，可读写。   |
|  25  | [VerticalAlignment](#Cell.VerticalAlignment)   | 返回或设置表格中一个或多个单元格中文字的垂直对齐方式。 WdCellVerticalAlignment 类型，可读写。          |
|  26  | [Width](#Cell.Width)                           | 返回或设置指定表格单元格的宽度（以磅为单位）。 Long 类型，可读写。                                     |
|  27  | [WordWrap](#Cell.WordWrap)                     | 如果为 True ，则 WPS 使文本自动换行并加长单元格，使单元格的宽度保持相同。 Boolean 类型，可读写。       |

## 成员方法

### Cell.AutoSum

插入一个 = (Formula) 域，以计算和显示表达式指定的位于表格某一单元格上方或左侧的各单元格的数值之和。

**语法**

**express.AutoSum ()**

\*express是一个代表 Cell 对象的变量。

**说明**

有关 WPS 如何确定添加哪些值的信息，请参阅 **Formula** 方法。

**示例**

### Cell.Delete

删除一个或多个表格单元格并可选择控制如何移动剩余的单元格。

**语法**

**express.Delete ( *ShiftCells* )**

\*express是一个代表 Cell 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                  |
|--------------|-----------|----------|-------------------------------------------------------------------------------------------------------|
| *ShiftCells* | 可选      | Variant  | 剩余单元格移动的方向。可以是任意 WdDeleteCells 常量。如果省略，最后删除的单元格的右侧单元格向左移动。 |

**示例**

``` JavaScript
/* 本示例删除活动文档中第一个表格中的第一个单元格 */
Application.ActiveDocument.Tables.Item(1).Cell(1,1).Delete()
```

### Cell.Formula

将包含指定公式的 = (Formula) 域插入到单元格中。

**语法**

**express.Formula ( *Formula* , *NumFormat* )**

\*express是一个代表 Cell 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                        |
|-------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Formula*   | 可选      | Variant  | 希望 = (Formula) 域求值的数学公式。可以使用与电子表格类似的方式引用表格中的单元格。例如，“=SUM(A4:C4)”指定第四行中的前三个值。有关 = (Formula) 域的详细内容，请参阅域代码：= (Formula) 域。 |
| *NumFormat* | 可选      | Variant  | = (Formula) 域的结果的格式。有关可应用的格式类型的信息，请参阅“数字图片 (\\)”域开关。                                                                                                       |

**说明**

如果插入点所在单元格的上面或左面至少有一个单元格包含数值，则 **Formula** 是可选的。如果插入点上面的单元格含有数值，则插入的域为 {=SUM(ABOVE)}；如果插入点左面的单元格含有数值，则插入的域为 {=SUM(LEFT)}；如果在插入点上面和左面的单元格中都含有数值，则 WPS 使用以下规则来确定插入哪一个 SUM 函数：

- 如果紧邻插入点上面的单元格中含有数值，则 WPS 插入 {=SUM(ABOVE)}。
- 如果紧邻插入点上面的单元格中不包含数值，而紧邻插入点左面的单元格中含有数值，则 WPS 插入 {=SUM(LEFT)}。
- 如果紧邻单元格中均不包含数值，则 WPS 插入 {=SUM(ABOVE)}。
- 如果未指定 **Formula** ，而且插入点上面和左面的所有单元格均为空，则域的结果将会出错。

**示例**

``` JavaScript
/*本示例在活动文档的开始处创建一个 3x3 表格，然后计算第一列的平均值。*/
function test()
{
let myRange = ActiveDocument.Range(0, 0)
let myTable = ActiveDocument.Tables.Add(myRange, 3, 3)
myTable.Cell(1,1).Range.InsertAfter("100")
myTable.Cell(2,1).Range.InsertAfter("50")
myTable.Cell(3,1).Formula("=Average(Above)")
}
```

``` JavaScript
/*本示例在插入点插入一个公式，该公式将决定选定单元格上面的单元格中的最大数值。*/
function test()
{
Selection.Collapse(wdCollapseStart)
if(Selection.Information(wdWithInTable) == true) {
    Selection.Cells.Item(1).Formula("=Max(Above)")
}
else {
    alert("The insertion point is not in a table.")
}
}
```

### Cell.Merge

将指定单元格与另一表格单元格合并，成为一个单独的表格单元格。

**语法**

**express.Merge ( *MergeTo* )**

\*express是一个代表 Cell 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型    | 说明             |
|-----------|-----------|-------------|------------------|
| *MergeTo* | 必选      | Cell object | 要合并的单元格。 |

**示例**

``` JavaScript
/*本示例将活动文档的表 1 的前两个单元格合并，然后删除表格的边框。*/
function test() {
    if (Application.ActiveDocument.Tables.Count >= 1) {
        let Tab = Application.ActiveDocument.Tables.Item(1)
        Tab.Cell(1, 1).Merge(Tab.Cell(1, 2))
        Tab.Borders.Enable = false
    }
}
```

### Cell.Select

选择指定的对象。

**语法**

**express.Select ()**

\*express是一个代表 Cell 对象的变量。

**说明**

使用该方法后，可用 **Selection** 对象来处理选定项目。有关详细信息，请参阅 处理 Selection 对象。

### Cell.SetHeight

设置表格单元格的高度。

**语法**

**express.SetHeight ( *RowHeight* , *HeightRule* )**

\*express是一个代表 Cell 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型        | 说明                           |
|--------------|-----------|-----------------|--------------------------------|
| *RowHeight*  | 必选      | Variant         | 一行或多行的高度，以磅为单位。 |
| *HeightRule* | 必选      | WdRowHeightRule | 用于确定指定单元格高度的方法。 |

**说明**

设置 **Cell** 对象的 **SetHeight** 属性可自动为整行设置该属性。

**示例**

``` JavaScript
/*本示例将选定单元格的行高设置为不小于 18 磅。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        Application.Selection.Cells.SetHeight(18, wdRowHeightAtLeast)
    }
    else {
        alert("The insertion point is not in a table.")
    }
}
```

### Cell.SetWidth

设置表格列或单元格的宽度

**语法**

**express.SetWidth ( *ColumnWidth* , *RulerStyle* )**

\*express是一个代表 Cell 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型     | 说明                            |
|---------------|-----------|--------------|---------------------------------|
| *ColumnWidth* | 必选      | Single       | 指定列的宽度，以磅为单位。      |
| *RulerStyle*  | 必选      | WdRulerStyle | 控制 WPS 调整单元格宽度的方式。 |

**说明**

上述 **WdRulerStyle** 行为应用于左对齐的表格。 **WdRulerStyle** 行为应用于中对齐和右对齐的表格时可能会出现未知效果，因此应谨慎使用 **SetWidth** 方法。

**示例**

``` JavaScript
/*本示例在新文档中创建一张表格，设置第二行第一个单元格宽度为 1.5 英寸。本示例保持表格中其他单元格的宽度。*/
function test() {
    let newDoc = Application.Documents.Add()
    let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)
    myTable.Cell(2, 1).SetWidth(InchesToPoints(1.5), wdAdjustNone)
}
```

``` JavaScript
/*本示例设置包含插入点的单元格宽度为 36 磅。本示例缩小第一列的宽度以保持表格的右边界位置。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        Application.Selection.Cells.Item(1).SetWidth(36, wdAdjustFirstColumn)
    }
    else {
        alert("The insertion point is not in a table.")
    }
}
```

### Cell.Split

将一个表格单元格拆分为多个。

**语法**

**express.Split ( *NumRows* , *NumColumns* )**

\*express是一个代表 Cell 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                 |
|--------------|-----------|----------|--------------------------------------|
| *NumRows*    | 可选      | Variant  | 一个单元格或多个单元格拆分成的行数。 |
| *NumColumns* | 可选      | Variant  | 一个单元格或多个单元格拆分成的列数。 |

**示例**

``` JavaScript
/* 本示例将第一张表格的第一个单元格拆分为两个单元格 */
Application.ActiveDocument.Tables.Item(1).Cell(1, 1).Split(1, 2)
```

## 成员属性

### Cell.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 Cell 对象的变量。

### Cell.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Cell 对象的变量。

### Cell.BottomPadding

返回或设置在表格中单独的单元格或所有单元格内容的下方添加的间距（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.BottomPadding**

\*express是一个代表 Cell 对象的变量。

**说明**

单独单元格的 **BottomPadding** 属性的设置值会覆盖整个表格的 **BottomPadding** 属性的设置值。

**示例**

``` JavaScript
/*本示例将活动文档中第一个表格的底边距设为 40 像素。*/
Application.ActiveDocument.Tables.Item(1).BottomPadding = PixelsToPoints(40, true)
```

### Cell.Column

返回一个 **Column** 对象，该对象代表包含指定单元格的表格列。只读。

**语法**

**express.Column**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*本示例创建一个 3x5 表格并为偶数列设置底纹。*/
function test() {
    Application.Selection.Collapse(wdCollapseStart)
    let tableNew = Application.ActiveDocument.Tables.Add(Selection.Range, 3, 5)
    for (let i = 1; i <= tableNew.Rows.Item(1).Cells.Count; i++) {
        if ((tableNew.Rows.Item(1).Cells.Item(i).ColumnIndex) % 2 == 0) {
            tableNew.Rows.Item(1).Cells.Item(i).Column.Shading.Texture = wdTexture10Percent
        }
    }
}
```

### Cell.ColumnIndex

返回包含指定单元格的表格列数。 **Long** 类型，只读。

**语法**

**express.ColumnIndex**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*本示例在新文档中创建表格，选择第一行中的每个单元格，然后返回包含所选单元格的列数。*/
function test() {
    let docNew = Application.Documents.Add()
    let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)
    for (let i = 1; i <= tableNew.Rows.Item(1).Cells.Count; i++) {
        tableNew.Rows.Item(1).Cells.Item(i).Select()
        alert("This is column " + tableNew.Rows.Item(1).Cells.Item(i).ColumnIndex)
    }
}
```

``` JavaScript
/*本示例返回插入点所在的单元格的列数。*/
alert(Application.Selection.Cells.Item(1).ColumnIndex)
```

### Cell.Creator

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 Cell 对象的变量。

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

### Cell.FitText

如果 WPS 会减小键入单元格中的文字的可视大小，使其适应列宽，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.FitText**

\*express是一个代表 Cell 对象的变量。

**说明**

如果将 **FitText** 设置为 **True** ，则虽然字体尺寸没有改变，但将调整字符的可视宽度，使单元格可容纳所有键入的文字。

**示例**

``` JavaScript
/*本示例实现的功能是：设定选定内容的第一单元格在其宽度内自动调整键入的文字大小。*/
Application.Selection.Cells.Item(1).FitText = true
```

### Cell.Height

返回或设置指定表格单元格的高度。

**语法**

**express.Height**

\*express是一个代表 Cell 对象的变量。

**说明**

如果指定行的 **HeightRule** 属性为 **wdRowHeightAuto** ，则 **Height** 返回 **wdUndefined** ；如果设置 **Height** 属性，系统将 **HeightRule** 设置为 **wdRowHeightAtLeast** 。可读/写 **Single** 类型。

### Cell.HeightRule

返回或设置一个 **WdRowHeightRule** 常量，该常量代表确定指定单元格或行高度的规则。可读写。

**语法**

**express.HeightRule**

\*express是一个代表 Cell 对象的变量。

**说明**

设置 **Cell** 对象的 **HeightRule** 属性将自动设置整行的高度。

### Cell.Id

在当前文档保存为网页时，返回或设置指定对象的识别标签。 **String** 类型，可读写。

**语法**

**express.Id**

\*express是一个代表 Cell 对象的变量。

**说明**

可以将标签作为引用其他网页或当前文档中内容的超链接。

### Cell.LeftPadding

返回或设置在表格的单个单元格或所有单元格的内容左侧的间距（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.LeftPadding**

\*express是一个代表 Cell 对象的变量。

**说明**

单个单元格的 **LeftPadding** 属性设置将覆盖整个表格的 **LeftPadding** 属性的设置。

**示例**

``` JavaScript
/*本示例将活动文档中第一个表格的第一个单元格的左边距设为 40 像素。*/
Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Item(1).LeftPadding = PixelsToPoints(40, false)
```

### Cell.NestingLevel

返回指定单元格的嵌套层。 **Long** 类型，只读。

**语法**

**express.NestingLevel**

\*express是一个代表 Cell 对象的变量。

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
    Rng1.Cells.Item(1).Range.Text = Rng1.Cells.Item(1).NestingLevel
    Rng1.Cells.Item(5).Range.PasteAsNestedTable()

    let Rng2 = Rng1.Cells.Item(5).Tables.Item(1).Range
    Rng2.Cells.Item(1).Range.Text = Rng2.Cells.Item(1).NestingLevel
    Rng2.Cells.Item(5).Range.PasteAsNestedTable()

    let Rng3 = Rng2.Cells.Item(5).Tables.Item(1).Range
    Rng3.Cells.Item(1).Range.Text = Rng3.Cells.Item(1).NestingLevel
}
```

### Cell.Next

返回一个 **Cell** 对象，该对象代表 **Cells** 集合中的下一个表格单元格。只读。

**语法**

**express.Next**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*如果选定内容位于表格之中，则本示例选定表格中下一个单元格的内容。
*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        Application.Selection.Cells.Item(1).Next.Select()
    }
}
```

### Cell.Parent

返回一个 **Object** 类型的值，该值代表指定 **Cell** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*本示例为活动文档中第一个表格的第一个单元格设置变量，将单元格的宽度更改为 36 磅，并删除表格的边框。*/
function test() {
    let objCell = Application.ActiveDocument.Tables.Item(1).Cell(1, 1)
    objCell.SetWidth(36, wdAdjustNone)
    objCell.Parent.Borders.Enable = false
}
```

### Cell.PreferredWidth

返回或设置指定单元格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。 **Single** 类型，可读写。

**语法**

**express.PreferredWidth**

\*express是一个代表 Cell 对象的变量。

**说明**

如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPoints** ，则 **PreferredWidth** 属性返回或设置宽度（以磅为单位）。如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPercent** ，则 **PreferredWidth** 属性返回或设置宽度（以窗口宽度的百分比表示）。

### Cell.PreferredWidthType

返回或设置用于指定单元格的宽度的首选度量单位。 **WdPreferredWidthType** 类型，只读。

**语法**

**express.PreferredWidthType**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*本示例将 WPS 设置为以窗口宽度的百分比来表示宽度，然后将文档中第一张表格的宽度设置为窗口宽度的 50%。*/
function test() {
    let rng = Application.ActiveDocument.Tables.Item(1)
    rng.PreferredWidthType = wdPreferredWidthPercent
    rng.PreferredWidth = 50
}
```

### Cell.Previous

返回一个 **Cell** 对象，该对象代表 **Cells** 集合中的前一个表格单元格。只读。

**语法**

**express.Previous**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*如果选定内容位于表格中，则本示例选定前一单元格的内容。*/
function test()
{
if(Selection.Information(wdWithInTable) == true) {
    Selection.Rows.Item(1).Cells.Item(3).Previous.Select()
}
}
```

### Cell.Range

返回一个 **Range** 对象，该对象代表指定对象所含的部分文档。

**语法**

**express.Range**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/* 本示例复制第一个表格首行中第一个单元格中的内容 */
Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Item(1).Range.Copy()
```

### Cell.RightPadding

返回或设置表格中某个单元格或全部单元格内容右侧的空格量（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.RightPadding**

\*express是一个代表 Cell 对象的变量。

**说明**

某个单元格的 **RightPadding** 属性的设置将覆盖整个表格的 **RightPadding** 属性的设置。

**示例**

``` JavaScript
/*本示例将活动文档第一个表格首行中第一个单元格的右边距设置为 40 像素。*/
Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Item(1).RightPadding = PixelsToPoints(40, false)
```

### Cell.Row

返回一个 **Row** 对象，该对象代表指定单元格所在的行。

**语法**

**express.Row**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*本示例为表格中插入点所在的行应用底纹。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        Application.Selection.Cells.Item(1).Row.Shading.Texture = wdTexture10Percent
    }
    else {
        alert("The insertion point is not in a table.")
    }
}
```

### Cell.RowIndex

返回指定单元格所在行的行数。 **Long** 类型，只读。

**语法**

**express.RowIndex**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*本示例在新文档中创建一个 3x3 表格，选定第一列的每个单元格，然后显示包含所选单元格的行的行号。*/
function test() {
    let newDoc = Application.Documents.Add()
    let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)
    for (let i = 1; i <= myTable.Columns.Item(1).Cells.Count; i++) {
        myTable.Columns.Item(1).Cells.Item(i).Select()
        alert("This is row " + myTable.Columns.Item(1).Cells.Item(i).RowIndex)
    }
}
```

``` JavaScript
/*本示例显示选定内容分第一行的行号。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        alert(Application.Selection.Cells.Item(1).RowIndex)
    }
}
```

### Cell.Shading

返回一个 **Shading** 对象，该对象指明指定对象的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*本示例对第一个表格首行中的第一个单元格应用水平线底纹。*/
function test() {
    if (Application.ActiveDocument.Tables.Count >= 1) {
        Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Item(1).Shading.Texture = wdTextureHorizontal
    }
}
```

### Cell.Tables

返回一个 **Tables** 集合，该集合代表指定表格单元格中的所有嵌套表格。只读。

**语法**

**express.Tables**

\*express是一个代表 Cell 对象的变量。

### Cell.TopPadding

返回或设置表格中单独的单元格或全部单元格的内容上方要增加的间距（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.TopPadding**

\*express是一个代表 Cell 对象的变量。

**说明**

单独单元格的 **TopPadding** 属性设置将覆盖整个表格的 **TopPadding** 属性设置。

**示例**

``` JavaScript
/*本示例将活动文档中第一张表格的顶部边距设为 40 像素。*/
Application.ActiveDocument.Tables.Item(1).TopPadding = PixelsToPoints(40, true)
```

### Cell.VerticalAlignment

返回或设置表格中一个或多个单元格中文字的垂直对齐方式。 **WdCellVerticalAlignment** 类型，可读写。

**语法**

**express.VerticalAlignment**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*本示例在一个新文档中创建一个 3x3 表格，并为表格中的每个单元格分配连续的单元格号。然后将第一行的高度设置为 20 磅，并在单元格的顶端垂直对齐文本。*/
function test() {
    let newDoc = Application.Documents.Add()
    let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)
    i = 1
    for (let j = 1; j <= myTable.Range.Cells.Count; j++) {
        myTable.Range.Cells.Item(j).Range.InsertAfter("cell" + i)
        i++
    }
    let fpx = myTable.Rows.Item(1)
    fpx.Height = 20
    fpx.Cells.VerticalAlignment = wdAlignVerticalTop
}
```

### Cell.Width

返回或设置指定表格单元格的宽度（以磅为单位）。 **Long** 类型，可读写。

**语法**

**express.Width**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*本示例返回插入点所在单元格的宽度（以英寸为单位）。*/
function test() {
    if (Application.Selection.Information(wdWithInTable) == true) {
        alert(Application.PointsToInches(Selection.Cells.Item(1).Width))
    }
}
```

### Cell.WordWrap

如果为 **True** ，则 WPS 使文本自动换行并加长单元格，使单元格的宽度保持相同。 **Boolean** 类型，可读写。

**语法**

**express.WordWrap**

\*express是一个代表 Cell 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 对第一张表格的第三个单元格中的文本自动换行，以使单元格的宽度保持相同。*/
Application.ActiveDocument.Tables.Item(1).Range.Cells.Item(3).WordWrap = true
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Cell](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Editors 对象

## 简介

Editor 对象的集合代表已被授予特定权限可编辑文档部分的用户或用户组集合。

## 方法

| 序号 | 名称                  | 说明                                                                           |
|:----:|:----------------------|:-------------------------------------------------------------------------------|
|  1   | [Add](#Editors.Add)   | 返回一个 Editor 对象，该对象代表指定用户用于修改文档中区域或选定内容的新权限。 |
|  2   | [Item](#Editors.Item) | 返回一个 Editor 对象，该对象代表已分配给权限来编辑部分文档的特定用户或用户组。 |

## 属性

| 序号 | 名称                                | 说明                                                                                                            |
|:----:|:------------------------------------|:----------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Editors.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                                                  |
|  2   | [Code](#Editors.Code)               | 返回一个代表域代码的 Range 对象。可读写。                                                                       |
|  3   | [Count](#Editors.Count)             | 返回一个 Long 类型值，该值代表集合中 Editor 对象的数目。只读。                                                  |
|  4   | [Creator](#Editors.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                    |
|  5   | [Data](#Editors.Data)               | 返回或设置 ADDIN 域中的数据。可读写 String 类型。                                                               |
|  6   | [Index](#Editors.Index)             | 返回一个 Long 类型的值，该值代表项目在集合中的位置。只读。                                                      |
|  7   | [InlineShape](#Editors.InlineShape) | 返回一个 InlineShape 对象，该对象可以代表图片、OLE 对象以及作为 INCLUDEPICTURE 或 EMBED 域结果的 ActiveX 控件。 |
|  8   | [Kind](#Editors.Kind)               | 返回 Field 对象的链接类型。只读 WdFieldKind 类型。                                                              |
|  9   | [LinkFormat](#Editors.LinkFormat)   | 返回一个 LinkFormat 对象，该对象代表指定域的链接选项。只读。                                                    |
|  10  | [Locked](#Editors.Locked)           | 如果指定域已被锁定，则该属性值为 True 。可读写 Boolean 类型。                                                   |
|  11  | [Next](#Editors.Next)               | 返回集合中的下一个对象。只读。                                                                                  |
|  12  | [OLEFormat](#Editors.OLEFormat)     | 返回一个 OLEFormat 对象，该对象代表指定域的 OLE 特征（链接除外）。只读。                                        |
|  13  | [Parent](#Editors.Parent)           | 返回一个 Object 类型值，该值代表指定 Editors 对象的父对象。                                                     |
|  14  | [Previous](#Editors.Previous)       | 返回集合中的上一个对象。只读。                                                                                  |
|  15  | [Result](#Editors.Result)           | 返回一个 Range 对象，该对象代表域的结果。可读写。                                                               |
|  16  | [ShowCodes](#Editors.ShowCodes)     | 如果显示指定域的域代码而不是域结果，则该属性值为 True 。 Boolean 类型，可读写。                                 |
|  17  | [Type](#Editors.Type)               | 返回域的类型。只读 WdFieldType 类型。                                                                           |

## 成员方法

### Editors.Add

返回一个 **Editor** 对象，该对象代表指定用户用于修改文档中区域或选定内容的新权限。

**语法**

**express.Add ( *EditorID* )**

\*express是一个代表 Editors 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                         |
|------------|-----------|----------|--------------------------------------------------------------------------------------------------------------|
| *EditorID* | 必选      | Variant  | 可为代表用户电子邮件别名（如果在相同的域中）和电子邮件地址的 String 字符串或代表用户组的 WdEditorType 常量。 |

**示例**

``` JavaScript
/*以下示例将选定文本的编辑权限分配给当前用户。*/
let objEditor = Selection.Editors.Add(wdEditorCurrent)
```

### Editors.Item

返回一个 **Editor** 对象，该对象代表已分配给权限来编辑部分文档的特定用户或用户组。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Editors 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                 |
|---------|-----------|----------|----------------------------------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 可以是代表用户电子邮件别名（如果在相同的域中）和电子邮件地址的 String 类型值，或者是代表用户组的 WdEditorType 常量。 |

**返回值**

Editor

## 成员属性

### Editors.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Editors 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Editors.Code

返回一个代表域代码的 **Range** 对象。可读写。

**语法**

**express.Code**

\*express是一个代表 Editors 对象的变量。

**说明**

域代码是包含在域字符 ( **{ }** ) 中的所有内容，包括前导和尾部空格。无需更改域结果的显示方式即可访问域代码。

**示例**

``` JavaScript
/*本示例显示活动文档中每一个域的域代码。
*/
function test()
{
for (let i=1;i<= ActiveDocument.Fields.Count;i++) {
    MsgBox("\"" + ActiveDocument.Fields.Item(i).Code.Text + "\"")
}
}
```

``` JavaScript
/*本示例将活动文档中第一个域的域代码改为 CREATEDATE。*/
function test()
{
let rngTemp = ActiveDocument.Fields.Item(1).Code
    rngTemp.Text = " CREATEDATE "
    ActiveDocument.Fields.Item(1).Update()
}
```

``` JavaScript
/*本示例判断活动文档中是否包含一个名为“Title”的邮件合并域。*/
function test()
{
for(let i=1;i<= ActiveDocument.MailMerge.Fields.Count;i++) {
    let str=ActiveDocument.MailMerge.Fields.Item(i).Code.Text
    if(str.indexOf("Title",1)) {
        MsgBox("A Title merge field is in this document")
    }
}
}
```

### Editors.Count

返回一个 **Long** 类型值，该值代表集合中 Editor 对象的数目。只读。

**语法**

**express.Count**

\*express是一个代表 Editors 对象的变量。

### Editors.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Editors 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Editors.Data

返回或设置 ADDIN 域中的数据。可读写 **String** 类型。

**语法**

**express.Data**

\*express是一个代表 Editors 对象的变量。

**说明**

该数据在域代码或结果中是不可见的，只有通过返回 **Data** 属性的值才能进行访问。如果该域不是 ADDIN 域，该属性将发生错误。

**示例**

``` JavaScript
/*以下示例在活动文档的插入点插入一个 ADDIN 域，然后设置该域的数据。*/
function test()
{
Selection.Collapse(wdCollapseStart)
let fldTemp =ActiveDocument.Fields.Add(Selection.Range,wdFieldAddin)
fldTemp.Data = "Hidden information"
}
```

### Editors.Index

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Editors 对象的变量。

**示例**

``` JavaScript
/*本示例返回选定域在 Fields 集合中的位置。*/
let num = Selection.Fields.Item(1).Index
```

### Editors.InlineShape

返回一个 **InlineShape** 对象，该对象可以代表图片、OLE 对象以及作为 INCLUDEPICTURE 或 EMBED 域结果的 ActiveX 控件。

**语法**

**express.InlineShape**

\*express是一个代表 Editors 对象的变量。

**说明**

可将 **InlineShape** 对象作为字符处理，并作为字符放置在一行文本中。

**示例**

``` JavaScript
/*本示例返回活动文档中与第一个域有关的嵌入式图形的宽度。若要使此示例运行，此域必须是 INCLUDEPICTURE 域。*/
function test()
{
if( ActiveDocument.Fields.Item(1).Type == wdFieldIncludePicture) {
    MsgBox( ActiveDocument.Fields.Item(1).InlineShape.Width)
}
}
```

### Editors.Kind

返回 **Field** 对象的链接类型。只读 **WdFieldKind** 类型。

**语法**

**express.Kind**

\*express是一个代表 Editors 对象的变量。

**示例**

``` JavaScript
/*以下示例更新活动文档中的所有热链接域。*/
function test()
{
let aField = ActiveDocument.Fields
for(let i=1; i<=aField.Count; i++) {
    if(aField.Item(i).Kind == wdFieldKindWarm) {
        aField.Item(i).Update()
    }
}
}
```

### Editors.LinkFormat

返回一个 **LinkFormat** 对象，该对象代表指定域的链接选项。只读。

**语法**

**express.LinkFormat**

\*express是一个代表 Editors 对象的变量。

**示例**

``` JavaScript
/*以下示例更新活动文档中尚未自动更新的域。*/
function test()
{
for(let i=1;i<= ActiveDocument.Fields.Count;i++) {
    if(ActiveDocument.Fields.Item(i).LinkFormat.AutoUpdate == false) {  
        ActiveDocument.Fields.Item(i).LinkFormat.Update()
    }
}
}
```

### Editors.Locked

如果指定域已被锁定，则该属性值为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.Locked**

\*express是一个代表 Editors 对象的变量。

**说明**

当域处于锁定状态时，无法更新该域的结果。

**示例**

``` JavaScript
/*以下示例在所选内容的开头插入一个 DATE 域，然后锁定该域。*/
function test()
{
Selection.Collapse(wdCollapseStart)
let myField = ActiveDocument.Fields.Add(Selection.Range,wdFieldDate)
myField.Locked = true
}
```

### Editors.Next

返回集合中的下一个对象。只读。

**语法**

**express.Next**

\*express是一个代表 Editors 对象的变量。

**示例**

``` JavaScript
/*如果 Next 方法返回一个 Field 对象，并且该域不是 FILLIN 域，则以下示例更新活动文档的第一节中的域。*/
function test()
{
let myField
if (ActiveDocument.Sections.Item(1).Range.Fields.Count >= 1) {
    myField = ActiveDocument.Fields.Item(1)
    while(myField!=null) {
        if(myField.Type != wdFieldFillIn ) {
            myField.Update()
            myField = myField.Next
        }
    }
}
}
```

### Editors.OLEFormat

返回一个 **OLEFormat** 对象，该对象代表指定域的 OLE 特征（链接除外）。只读。

**语法**

**express.OLEFormat**

\*express是一个代表 Editors 对象的变量。

### Editors.Parent

返回一个 **Object** 类型值，该值代表指定 **Editors** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Editors 对象的变量。

### Editors.Previous

返回集合中的上一个对象。只读。

**语法**

**express.Previous**

\*express是一个代表 Editors 对象的变量。

**示例**

``` JavaScript
/*以下示例显示活动文档中倒数第二个域的域代码。*/
function test()
{
let aField = ActiveDocument.Fields.Item(ActiveDocument.Fields.Count).Previous
MsgBox( "Field code = " + aField.Code)
}
```

### Editors.Result

返回一个 **Range** 对象，该对象代表域的结果。可读写。

**语法**

**express.Result**

\*express是一个代表 Editors 对象的变量。

**说明**

无需改变域代码的显示方式即可获得一个域结果。使用 **Text** 属性可返回 **Range** 对象中的文本。

**示例**

``` JavaScript
/*以下示例对所选内容中的第一个域应用加粗格式。*/
function test()
{
if(Selection.Fields.Count >= 1) { 
    let myRange = Selection.Fields.Item(1).Result
    myRange.Bold = true
}
}
```

### Editors.ShowCodes

如果显示指定域的域代码而不是域结果，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowCodes**

\*express是一个代表 Editors 对象的变量。

**示例**

``` JavaScript
/*本示例选定下一个域，并显示其域代码。*/
function test()
{
Selection.GoTo(wdGoToField)
Selection.Expand(wdWord)
if( Selection.Fields.Count == 1 ) {
    Selection.Fields.Item(1).ShowCodes = true
}
}
```

``` JavaScript
/*本示例更新并显示活动文档中第一个域的结果。*/
function test()
{
if( ActiveDocument.Fields.Count >= 1) {
    ActiveDocument.Fields.Item(1).Update()
    ActiveDocument.Fields.Item(1).ShowCodes = false
}
}
```

### Editors.Type

返回域的类型。只读 **WdFieldType** 类型。

**语法**

**express.Type**

\*express是一个代表 Editors 对象的变量。

**说明**

**安全性** ??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Editors](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FootnoteOptions 对象

## 简介

代表分配给文档中脚注的范围或所选内容的属性。

## 说明

使用 Range 或 Selection 对象可返回 FootnoteOptions 对象。使用 FootnoteOptions 对象可为文档的不同区域分配不同的脚注属性。例如，您可能希望长文档说明部分的脚注以小写字母显示，而文档中其余部分的脚注以星号显示。以下示例使用 NumberingRule 、 NumberStyle 和 StartingNumber 属性设置活动文档的第一节中的尾注格式。

``` JavaScript
function BookIntro(){
    //Sets the range as section one of the active document
    let rngIntro = Application.ActiveDocument.Sections.Item(1).Range
    //Formats the EndnoteOptions properties
    rngIntro.FootnoteOptions.NumberingRule = wdRestartPage
    rngIntro.FootnoteOptions.NumberStyle = wdNoteNumberStyleLowercaseLetter
    rngIntro.FootnoteOptions.StartingNumber = 1
}
```

## 属性

| 序号 | 名称                                              | 说明                                                                                |
|:----:|:--------------------------------------------------|:------------------------------------------------------------------------------------|
|  1   | [Application](#FootnoteOptions.Application)       | 返回一个代表 WPS 应用程序的 Application 对象。                                      |
|  2   | [Creator](#FootnoteOptions.Creator)               | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。        |
|  3   | [Location](#FootnoteOptions.Location)             | 返回或设置所有脚注的位置。 WdFootnoteLocation 类型，可读写。                        |
|  4   | [NumberStyle](#FootnoteOptions.NumberStyle)       | 返回或设置脚注的编号样式。 WdNoteNumberStyle 类型，可读写。                         |
|  5   | [NumberingRule](#FootnoteOptions.NumberingRule)   | 返回或设置在分页符或分节符之后的脚注或尾注的编号方式。可读写 WdNumberingRule 类型。 |
|  6   | [Parent](#FootnoteOptions.Parent)                 | 返回一个 Object 类型值，该值代表指定 FootnoteOptions 对象的父对象。                 |
|  7   | [StartingNumber](#FootnoteOptions.StartingNumber) | 返回或设置脚注的起始编号。 Long 类型，可读写。                                      |

## 成员属性

### FootnoteOptions.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 FootnoteOptions 对象的变量。

### FootnoteOptions.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FootnoteOptions 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### FootnoteOptions.Location

返回或设置所有脚注的位置。 **WdFootnoteLocation** 类型，可读写。

**语法**

**express.Location**

\*express是一个代表 FootnoteOptions 对象的变量。

**示例**

``` JavaScript
/*本示例将脚注置于每页的底端。*/
Application.ActiveDocument.Footnotes.Location = wdBottomOfPage
```

### FootnoteOptions.NumberStyle

返回或设置脚注的编号样式。 **WdNoteNumberStyle** 类型，可读写。

**语法**

**express.NumberStyle**

\*express是一个代表 FootnoteOptions 对象的变量。

**说明**

某些 **WdNoteNumberStyle** 常量可能不可用，具体取决于所选择或安装的语言支持（例如，美国英语）。

### FootnoteOptions.NumberingRule

返回或设置在分页符或分节符之后的脚注或尾注的编号方式。可读写 **WdNumberingRule** 类型。

**语法**

**express.NumberingRule**

\*express是一个代表 FootnoteOptions 对象的变量。

**示例**

``` JavaScript
/*如果将第一节中的脚注编号设置为在每个分节符后重新开始，则以下示例设置每页都重新开始编号。*/
function test() {
  let myRange = Application.ActiveDocument.Sections.Item(1).Range
  if(myRange.Footnotes.NumberingRule == wdRestartSection) {
      myRange.Footnotes.NumberingRule = wdRestartPage
  }
}
```

### FootnoteOptions.Parent

返回一个 **Object** 类型值，该值代表指定 **FootnoteOptions** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 FootnoteOptions 对象的变量。

### FootnoteOptions.StartingNumber

返回或设置脚注的起始编号。 **Long** 类型，可读写。

**语法**

**express.StartingNumber**

\*express是一个代表 FootnoteOptions 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/FootnoteOptions](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Footnotes 对象

## 简介

以后的版本中将提供关于此成员的说明。

## 方法

| 序号 | 名称                                                                | 说明                                                                                                    |
|:----:|:--------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------|
|  1   | [Add](#Footnotes.Add)                                               | 返回一个 Footnote 对象，该对象代表添加到区域中的脚注。                                                  |
|  2   | [Convert](#Footnotes.Convert)                                       | 将尾注转换为脚注，或将脚注转换为尾注。                                                                  |
|  3   | [Item](#Footnotes.Item)                                             | 返回集合中的单个 Footnote 对象。                                                                        |
|  4   | [ResetContinuationNotice](#Footnotes.ResetContinuationNotice)       | 将脚注或尾注延续标记重新设置为默认标记。                                                                |
|  5   | [ResetContinuationSeparator](#Footnotes.ResetContinuationSeparator) | 将脚注或尾注延续分隔符重新设置为默认分隔符。                                                            |
|  6   | [ResetSeparator](#Footnotes.ResetSeparator)                         | 将脚注分隔符重新设置为默认分隔符。                                                                      |
|  7   | [SwapWithEndnotes](#Footnotes.SwapWithEndnotes)                     | 将一篇文档中的所有尾注转换成脚注或将脚注转换为尾注。要将某一区域的尾注转换为脚注，请使用 Convert 方法。 |

## 属性

| 序号 | 名称                            | 说明                       |
|:----:|:--------------------------------|:---------------------------|
|  1   | [Count](#Footnotes.Count)       | 返回指定集合中的项目数。   |
|  2   | [Location](#Footnotes.Location) | 返回或设置所有脚注的位置。 |

## 成员方法

### Footnotes.Add

返回一个 **Footnote** 对象，该对象代表添加到区域中的脚注。

**语法**

**express.Add ( *Range* , *Reference* , *Text* )**

\*express是一个代表 Footnotes 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型     | 说明                                                                 |
|-------------|-----------|--------------|----------------------------------------------------------------------|
| *Range*     | 必选      | Range object | 标记用于尾注或脚注的区域。可以是折叠区域。                           |
| *Reference* | 可选      | Variant      | 自定义引用标记的文字。如果省略该参数，WPS 将插入自动编号的引用标记。 |
| *Text*      | 可选      | Variant      | 尾注或脚注文本。                                                     |

**说明**

要为 *Reference* 参数指定一个符号，可用 ` {FontName CharNum} ` 语法。FontName 是包含该符号的字体名称。修饰字体的名称显示在 **“插入”** 菜单 **“符号”** 对话框的 **“字体”** 框中。CharNum 是要插入符号的相应位置序数与 31 之和，在符号表中从左向右计数。例如，如果指定的符号是“Symbol”字体的 omega (ω)，它在符号表格中的位置为 56，该参数就应为“{Symbol 87}”。

**示例**

``` JavaScript
/*以下代码示例在选定内容的末尾添加一条自动编号的脚注。*/
Application.ActiveDocument.Footnotes.Add(Selection.Range,"The Willow Tree, (Lone Creek Press, 1996).")

/*以下代码示例为引用标记添加一条使用自定义符号的脚注。*/
Application.ActiveDocument.Footnotes.Add(Selection.Range ,"More information in the full report.","{Symbol -3998}")
```

### Footnotes.Convert

将尾注转换为脚注，或将脚注转换为尾注。

**语法**

**express.Convert ()**

\*express是一个代表 Footnotes 对象的变量。

**示例**

``` JavaScript
/*本示例将选定内容的脚注转换为尾注。*/
function test()
{
if( Selection.Footnotes.Count > 0 ) {
    Selection.Footnotes.Convert()
}
}
```

### Footnotes.Item

返回集合中的单个 **Footnote** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Footnotes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

### Footnotes.ResetContinuationNotice

将脚注或尾注延续标记重新设置为默认标记。

**语法**

**express.ResetContinuationNotice ()**

\*express是一个代表 Footnotes 对象的变量。

**说明**

默认标记为空（无文本）。

**示例**

``` JavaScript
/*以下示例在 Sales.doc 中重新设置脚注延续标记并将脚注引用标记的起始编号设置为 2。*/
function test()
{
let note= Documents.Item("Sales.doc").Sections.Item(1).Range.Footnotes
note.ResetContinuationNotice()
note.NumberingRule = wdRestartContinuous
note.StartingNumber = 2
}
```

### Footnotes.ResetContinuationSeparator

将脚注或尾注延续分隔符重新设置为默认分隔符。

**语法**

**express.ResetContinuationSeparator ()**

\*express是一个代表 Footnotes 对象的变量。

**说明**

默认分隔符是一条长横线，该横线分隔文档文字与上接前一页的注释。

**示例**

``` JavaScript
/*以下示例将脚注延续分隔符重新设置为默认分隔线。*/
ActiveDocument.Footnotes.ResetContinuationSeparator()
```

### Footnotes.ResetSeparator

将脚注分隔符重新设置为默认分隔符。

**语法**

**express.ResetSeparator ()**

\*express是一个代表 Footnotes 对象的变量。

**说明**

默认分隔符为一条短横线，该横线分隔文档文字和注释。

**示例**

``` JavaScript
/*以下示例将脚注分隔符重新设置为默认分隔线。*/
ActiveDocument.Footnotes.ResetSeparator()
```

### Footnotes.SwapWithEndnotes

将一篇文档中的所有尾注转换成脚注或将脚注转换为尾注。要将某一区域的尾注转换为脚注，请使用 **Convert** 方法。

**语法**

**express.SwapWithEndnotes ()**

\*express是一个代表 Footnotes 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中的脚注转换为尾注，将尾注转换为脚注。*/
ActiveDocument.Footnotes.SwapWithEndnotes()
```

## 成员属性

### Footnotes.Count

返回指定集合中的项目数。

**语法**

**express.Count**

\*express是一个代表 Footnotes 对象的变量。

### Footnotes.Location

返回或设置所有脚注的位置。

**语法**

**express.Location**

\*express是一个代表 Footnotes 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Footnotes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# HeaderFooter 对象

## 简介

代表单一页眉或页脚。 HeaderFooter 对象是 HeadersFooters 集合的一个成员。 HeadersFooters 集合包含指定文档节中的所有页眉和页脚。

## 说明

使用 Headers ( *Index* ) 或 Footers ( *Index* ) 可返回一个 HeaderFooter 对象，其中 index 为 WdHeaderFooterIndex 常量（ wdHeaderFooterEvenPages 、 wdHeaderFooterFirstPage 或 wdHeaderFooterPrimary ）之一。以下示例更改活动文档第一节中主页眉和主页脚的文字。

``` JavaScript
function test() {
    let mySection = Application.ActiveDocument.Sections.Item(1)
        mySection.Headers.Item(wdHeaderFooterPrimary).Range.Text = "Header text"
        mySection.Footers.Item(wdHeaderFooterPrimary).Range.Text = "Footer text"
}
```

您也可以使用 HeaderFooter 属性和 Selection 对象来返回单个 HeaderFooter 对象。

使用 DifferentFirstPageHeaderFooter 属性和 PageSetup 对象可指定不同的首页。以下示例在活动文档的首页页脚中插入文字。

``` JavaScript
function test() {
    Application.ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = true
    Application.ActiveDocument.Sections.Item(1).Footers.Item(wdHeaderFooterFirstPage).Range.InsertBefore("Written by Joe Smith")
}
```

使用 OddAndEvenPagesHeaderFooter 属性和 PageSetup 对象可为奇数页和偶数页设置不同的页眉和页脚。如果 OddAndEvenPagesHeaderFooter 属性为 True ，则可使用 wdHeaderFooterPrimary 返回奇数页的页眉或页脚，使用 wdHeaderFooterEvenPages 返回偶数页的页眉和页脚。

使用 Add 方法和 PageNumbers 对象可在页眉或页脚中添加页码。以下示例在活动文档第一节的主页脚中添加页码。

``` JavaScript
Application.ActiveDocument.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).PageNumbers.Add()
```

## 属性

| 序号 | 名称                                           | 说明                                                                                          |
|:----:|:-----------------------------------------------|:----------------------------------------------------------------------------------------------|
|  1   | [Application](#HeaderFooter.Application)       | 返回一个代表 WPS 应用程序的 Application 对象。                                                |
|  2   | [Creator](#HeaderFooter.Creator)               | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                  |
|  3   | [Exists](#HeaderFooter.Exists)                 | 如果存在指定的 HeaderFooter 对象，则该属性值为 True 。 Boolean 类型，可读写。                 |
|  4   | [Index](#HeaderFooter.Index)                   | 返回 WdHeaderFooterIndex ，代表文档或节中指定的页眉或页脚。只读。                             |
|  5   | [IsHeader](#HeaderFooter.IsHeader)             | 如果指定的 HeaderFooter 对象是页眉，则该属性值为 True 。 Boolean 类型，只读。                 |
|  6   | [LinkToPrevious](#HeaderFooter.LinkToPrevious) | 如果指定页眉或页脚链接至前一节中相应的页眉或页脚，则该属性值为 True 。 Boolean 类型，可读写。 |
|  7   | [PageNumbers](#HeaderFooter.PageNumbers)       | 返回一个 PageNumbers 集合，该集合代表指定页眉或页脚中包含的所有页码域。                       |
|  8   | [Parent](#HeaderFooter.Parent)                 | 返回一个 Object 类型值，该值代表指定 HeaderFooter 对象的父对象。                              |
|  9   | [Range](#HeaderFooter.Range)                   | 返回一个 Range 对象，该对象代表指包含在定页眉或页脚中的文档部分。                             |
|  10  | [Shapes](#HeaderFooter.Shapes)                 | 返回一个 Shapes 集合，该集合代表页眉或页脚中的所有的 Shape 对象。只读。                       |

## 成员属性

### HeaderFooter.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 HeaderFooter 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### HeaderFooter.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 HeaderFooter 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### HeaderFooter.Exists

如果存在指定的 **HeaderFooter** 对象，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Exists**

\*express是一个代表 HeaderFooter 对象的变量。

**说明**

默认情况下，主页眉和页脚存在于所有新文档中。使用此方法可以确定首页或奇数页是否使用页眉或页脚。还可以使用 **DifferentFirstPageHeaderFooter** 或 **OddAndEvenPagesHeaderFooter** 属性返回或设置指定文档或节中的页眉和页脚数目。

**示例**

``` JavaScript
/*如果第一节包含首页页眉，则本示例设置页眉的文字。*/
function test() {
  let secTemp = Application.ActiveDocument.Sections.Item(1)
  if(secTemp.Headers.Item(wdHeaderFooterFirstPage).Exists == true){
     secTemp.Headers.Item(wdHeaderFooterFirstPage).Range.Text = "First Page"
  }
}
```

### HeaderFooter.Index

返回 **WdHeaderFooterIndex** ，代表文档或节中指定的页眉或页脚。只读。

**语法**

**express.Index**

\*express是一个代表 HeaderFooter 对象的变量。

**示例**

``` JavaScript
/*如果指定的变量引用第一个页眉，则本示例将一个图形添至活动文档中的第一个页眉。*/
function test() {
    let hdrFirstPage = Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterFirstPage)
    if(hdrFirstPage.Index == wdHeaderFooterFirstPage){
       hdrFirstPage.Shapes.AddShape(msoShapeHeart, 36, 36, 36, 36).Fill.ForeColor.RGB = (255, 0, 0)
    }
}
```

### HeaderFooter.IsHeader

如果指定的 **HeaderFooter** 对象是页眉，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.IsHeader**

\*express是一个代表 HeaderFooter 对象的变量。

**示例**

``` JavaScript
/*本示例选择页脚并添加页码。*/
function test() {
  let view = Application.ActiveDocument.ActiveWindow.ActivePane.View
  view.Type = wdPrintView
  view.SeekView = wdSeekCurrentPageHeader
  
  if(Application.Selection.HeaderFooter.IsHeader == true){
      Application.ActiveDocument.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
  }
  
  Application.Selection.HeaderFooter.PageNumbers.Add()
}
```

### HeaderFooter.LinkToPrevious

如果指定页眉或页脚链接至前一节中相应的页眉或页脚，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.LinkToPrevious**

\*express是一个代表 HeaderFooter 对象的变量。

**说明**

链接页眉或页脚时，其内容与前一个页眉或页脚中的内容相同。因为默认情况下， **LinkToPrevious** 属性设为 **True** ，所以可以通过设置第一节的页眉、页脚和页码来为整个文档添加页眉、页脚和页码。例如，下面的示例将页码添加至活动文档所有节中全部页的页眉中。

``` JavaScript
Apploication.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).PageNumbers.Add() 
```

**LinkToPrevious** 属性可以分别应用于每个页眉或页脚。例如，可将偶数页的页眉的 **LinkToPrevious** 属性设置为 **True** ，而将偶数页的页脚设置为 **False** 。

**示例**

``` JavaScript
/*本示例的第一部分创建一个有两节的新文档，第二部分为新文档的第一和第二节的偶数页和奇数页创建不同的页眉。*/
function test() {
  Application.Documents.Add()
  
  for(let j = 1;j <= 4;j++){
      Application.Selection.TypeParagraph()
      Application.Selection.InsertBreak()
      Application.Selection.TypeParagraph()
  }
  
  Application.ActiveDocument.Paragraphs.Item(5).Range.InsertBreak(wdSectionBreakNextPage)
  Application.ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = true
  
  let oddHeader = Application.ActiveDocument.Sections.Item(2).Headers.Item(wdHeaderFooterPrimary)
      oddHeader.LinkToPrevious = false
      oddHeader.Range.InsertBefore("Section 2 Odd Header")
  
  let evenHeader = Application.ActiveDocument.Sections.Item(2).Headers.Item(wdHeaderFooterEvenPages)
      evenHeader.LinkToPrevious = false
      evenHeader.Range.InsertBefore("Section 2 Even Header")
  
  Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Range.InsertBefore("Section 1 Odd Header")
  Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterEvenPages).Range.InsertBefore("Section 1 Even Header")
}
```

### HeaderFooter.PageNumbers

返回一个 **PageNumbers** 集合，该集合代表指定页眉或页脚中包含的所有页码域。

**语法**

**express.PageNumbers**

\*express是一个代表 HeaderFooter 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例创建一个新文档，并向页脚中添加页码。*/
function test() {
  let myDoc = Application.Documents.Add()
  myDoc.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).PageNumbers.Add(wdAlignPageNumberCenter)
}
```

### HeaderFooter.Parent

返回一个 **Object** 类型值，该值代表指定 **HeaderFooter** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 HeaderFooter 对象的变量。

### HeaderFooter.Range

返回一个 **Range** 对象，该对象代表指包含在定页眉或页脚中的文档部分。

**语法**

**express.Range**

\*express是一个代表 HeaderFooter 对象的变量。

### HeaderFooter.Shapes

返回一个 **Shapes** 集合，该集合代表页眉或页脚中的所有的 **Shape** 对象。只读。

**语法**

**express.Shapes**

\*express是一个代表 HeaderFooter 对象的变量。

**说明**

该集合可以包含绘图、形状、图片、OLE 对象、ActiveX 控件、文本对象和标注。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

当 **Shapes** 属性应用于文档时，该属性将返回正文中的所有 **Shape** 对象，不包括页眉和页脚。当将其应用到 **HeaderFooter** 对象时， **Shapes** 属性返回文档中所有页眉和页脚中的 **Shape** 对象。

**示例**

``` JavaScript
/*以下示例显示活动文档第一节的主页眉和页脚中所有形状的总数。*/
alert(Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Shapes.Count)
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/HeaderFooter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Lists 对象

## 简介

List 对象的集合，这些对象代表指定文档中的所有列表。

## 方法

| 序号 | 名称                | 说明 |
|:----:|:--------------------|:-----|
|  1   | [Item](#Lists.Item) |      |

## 属性

| 序号 | 名称                              | 说明                                                       |
|:----:|:----------------------------------|:-----------------------------------------------------------|
|  1   | [Application](#Lists.Application) | 返回一个代表 WPS 应用程序的 Application 对象               |
|  2   | [Count](#Lists.Count)             | 返回一个 Number 类型的值，该值代表集合中列表的数目。只读。 |

## 成员方法

### Lists.Item

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Lists 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                           |
|---------|-----------|----------|----------------------------------------------------------------|
| *Index* | 必选      | Number   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

## 成员属性

### Lists.Application

返回一个代表 WPS 应用程序的 **Application** 对象

**语法**

**express.Application**

\*express是一个代表 Lists 对象的变量。

### Lists.Count

返回一个 **Number** 类型的值，该值代表集合中列表的数目。只读。

**语法**

**express.Count**

\*express是一个代表 Lists 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Lists](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Range 对象

## 简介

代表文档中的一个连续区域。每个 Range 对象由一个起始字符位置和一个终止字符位置定义

## 说明

与书签在文档中的使用方法类似， Range 对象在 Visual Basic 过程中用来标识文档的特定部分。但与书签不同的是， Range 对象只在定义该对象的过程运行时才存在。 Range 对象独立于所选内容。也就是说，您可以定义和处理一个范围而无需更改所选内容。还可以在文档中定义多个范围，但每个窗格中只能有一个所选内容。

## 方法

| 序号 | 名称                                      | 说明                                                                           |
|:----:|:------------------------------------------|:-------------------------------------------------------------------------------|
|  1   | [Collapse](#Range.Collapse)               | 将某一区域或所选内容折叠到起始位置或结束位置。折叠之后起始位置和结束位置相同。 |
|  2   | [Copy](#Range.Copy)                       | 将指定范围复制到剪贴板。                                                       |
|  3   | [Cut](#Range.Cut)                         | 将指定对象从文档中移到剪贴板上                                                 |
|  4   | [Delete](#Range.Delete)                   | 删除指定数量的字符或单词                                                       |
|  5   | [Expand](#Range.Expand)                   | 扩展指定区域或选定内容。返回添至该区域或选定内容的字符数。 Long 类型。         |
|  6   | [GoTo](#Range.GoTo)                       | 返回一个 Range 对象，该对象代表指定项（如页、书签或域）的起始位置。            |
|  7   | [InsertAfter](#Range.InsertAfter)         | 在范围的末尾插入指定文本。                                                     |
|  8   | [InsertBefore](#Range.InsertBefore)       | 在指定的范围前插入指定文本。                                                   |
|  9   | [InsertBreak](#Range.InsertBreak)         | 插入分页符、分栏符或分节符。                                                   |
|  10  | [InsertFile](#Range.InsertFile)           | 插入指定文件的全部或一部分。                                                   |
|  11  | [InsertParagraph](#Range.InsertParagraph) | 用新段落替换指定范围。                                                         |
|  12  | [Paste](#Range.Paste)                     | 将“剪贴板”中的内容插入指定范围。                                               |
|  13  | [Range](#Range.Range)                     | 将“剪贴板”中的内容插入指定范围。                                               |
|  14  | [Select](#Range.Select)                   | 选择指定的范围。                                                               |
|  15  | [XML](#Range.XML)                         | 返回一个 String 类型的值，该值代表指定对象中的 XML 文本。                      |

## 属性

| 序号 | 名称                                                          | 说明                                                                                                                                                                    |
|:----:|:--------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Range.Application)                             | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                          |
|  2   | [Bold](#Range.Bold)                                           | 如果选定区域中字体的格式为加粗，则该属性值为 True 。 Long 类型，可读写。                                                                                                |
|  3   | [BoldBi](#Range.BoldBi)                                       | 如果字体或区域为加粗格式，则为 True 。该属性值返回 True 、 False 或 wdUndefined （用于加粗和非加粗混合文本）。可设置为 True 、 False 或 wdToggle 。 Long 类型，可读写。 |
|  4   | [BookmarkID](#Range.BookmarkID)                               | 返回位于指定区域开始位置的书签编号。如果没有相应的书签，则返回 0（零）。 Long 类型，只读。                                                                              |
|  5   | [Bookmarks](#Range.Bookmarks)                                 | 返回一个 Bookmarks 集合。该集合代表某一文档、区域或选定内容中的所有书签。只读。                                                                                         |
|  6   | [Borders](#Range.Borders)                                     | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                                                                                                   |
|  7   | [Case](#Range.Case)                                           | 返回或设置一个 WdCharacterCase 常量，该常量代表指定区域中文字的大小写。可读写。                                                                                         |
|  8   | [Cells](#Range.Cells)                                         | 返回一个 Cells 集合，该集合代表在某区域中的表格单元格。只读。                                                                                                           |
|  9   | [Characters](#Range.Characters)                               | 返回一个 Characters 集合，该集合代表区域中的字符。只读。                                                                                                                |
|  10  | [CharacterStyle](#Range.CharacterStyle)                       | 返回一个 Variant 类型的值，该值代表用于设置一个或多个字符格式的样式。只读。                                                                                             |
|  11  | [CharacterWidth](#Range.CharacterWidth)                       | 返回或设置指定区域的字符宽度。 WdCharacterWidth 类型，可读写。                                                                                                          |
|  12  | [Columns](#Range.Columns)                                     | 返回一个 Columns 集合，该集合代表区域中的所有表格列。只读。                                                                                                             |
|  13  | [CombineCharacters](#Range.CombineCharacters)                 | 如果指定区域包含合并字符，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                   |
|  14  | [Comments](#Range.Comments)                                   | 返回一个 Comments 集合，该集合代表指定文档、选定内容或区域中的所有批注。只读。                                                                                          |
|  15  | [Conflicts](#Range.Conflicts)                                 | 返回 Conflicts 集合对象，该对象包含范围中的所有冲突对象。只读。                                                                                                         |
|  16  | [ContentControls](#Range.ContentControls)                     | 返回一个 ContentControls 集合，该集合代表一个区域中包含的内容控件。只读。                                                                                               |
|  17  | [Creator](#Range.Creator)                                     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                            |
|  18  | [DisableCharacterSpaceGrid](#Range.DisableCharacterSpaceGrid) | 如果 WPS 忽略相应 Range 对象的每行字符数，则该属性值为 True 。 Boolean 类型，可读写。                                                                                   |
|  19  | [Document](#Range.Document)                                   | 返回与指定区域相关的 Document 对象。只读。                                                                                                                              |
|  20  | [Duplicate](#Range.Duplicate)                                 | 返回一个只读 Range 对象，该对象代表指定区域的所有属性。                                                                                                                 |
|  21  | [Editors](#Range.Editors)                                     | 返回一个 Editors 对象，该对象代表已授权修改文档中选定内容或区域的所有用户。                                                                                             |
|  22  | [EmphasisMark](#Range.EmphasisMark)                           | 返回或设置字符或指定的字符串的着重号。 WdEmphasisMark 类型，可读写。                                                                                                    |
|  23  | [End](#Range.End)                                             | 返回或设置某区域中结束字符的位置。可读/写 Long 类型。                                                                                                                   |
|  24  | [EndnoteOptions](#Range.EndnoteOptions)                       | 返回一个 EndnoteOptions 对象，该对象代表区域中的尾注。                                                                                                                  |
|  25  | [Endnotes](#Range.Endnotes)                                   | 返回一个 Endnotes 集合，该集合代表区域中的所有尾注。只读。                                                                                                              |
|  26  | [EnhMetaFileBits](#Range.EnhMetaFileBits)                     | 返回一个 Variant 类型的值，该值代表文本区域的显示方式的图片代表形式。                                                                                                   |
|  27  | [Fields](#Range.Fields)                                       | 返回一个 Fields 集合，该集合代表区域中的所有域。只读。                                                                                                                  |
|  28  | [Find](#Range.Find)                                           | 返回一个 Find 对象，该对象包含查找操作所需的条件。只读。                                                                                                                |
|  29  | [FitTextWidth](#Range.FitTextWidth)                           | 该属性返回或设置 WPS 在当前选定内容或区域中填入文字的宽度（使用当前的度量单位）。 Single 类型，可读写。                                                                 |
|  30  | [Font](#Range.Font)                                           | 返回或设置 Font 对象，该对象代表指定对象的字符格式。 Font 类型，可读写。                                                                                                |
|  31  | [FootnoteOptions](#Range.FootnoteOptions)                     | 返回 FootnoteOptions 对象，该对象代表选定内容或区域的脚注。                                                                                                             |
|  32  | [Footnotes](#Range.Footnotes)                                 | 返回一个 Footnotes 集合，该集合代表区域中的所有脚注。只读。                                                                                                             |
|  33  | [FormattedText](#Range.FormattedText)                         | 返回或设置一个 Range 对象，该对象包含指定区域或选定内容中进行过格式编排的文字。可读写。                                                                                 |
|  34  | [FormFields](#Range.FormFields)                               | 返回一个 FormFields 集合，该集合代表区域中的所有窗体域。只读。                                                                                                          |
|  35  | [Frames](#Range.Frames)                                       | 返回一个 Frames 集合，该集合代表区域中的所有图文框。只读。                                                                                                              |
|  36  | [GrammarChecked](#Range.GrammarChecked)                       | 如果已经检查了指定范围或文档的语法，则该属性值为 True 。 Boolean 类型，可读写。                                                                                         |
|  37  | [GrammaticalErrors](#Range.GrammaticalErrors)                 | 返回一个 ProofreadingErrors 集合，该集合代表指定文档或区域中有语法检查错误的句子。只读。                                                                                |
|  38  | [HighlightColorIndex](#Range.HighlightColorIndex)             | 返回或设置指定区域的突出显示颜色。 WdColorIndex 类型，可读写。                                                                                                          |
|  39  | [HorizontalInVertical](#Range.HorizontalInVertical)           | 返回或设置位于垂直排列文字中的水平排列文字的格式。 WdHorizontalInVerticalType 类型，可读写。                                                                            |
|  40  | [HTMLDivisions](#Range.HTMLDivisions)                         | 返回一个 HTMLDivisions 对象，该对象代表 Web 文档中的 HTML 划分。                                                                                                        |
|  41  | [Hyperlinks](#Range.Hyperlinks)                               | 返回一个 Hyperlinks 集合，该集合代表指定范围内的所有超链接。只读。                                                                                                      |
|  42  | [ID](#Range.ID)                                               | 返回或设置特定范围的标识名称。可读写 String 类型。                                                                                                                      |
|  43  | [Information](#Range.Information)                             | 返回有关指定范围的信息。只读 Variant 类型。                                                                                                                             |
|  44  | [InlineShapes](#Range.InlineShapes)                           | 返回一个 InlineShape 集合，该集合代表范围中的所有 InlineShapes 对象。只读。                                                                                             |
|  45  | [IsEndOfRowMark](#Range.IsEndOfRowMark)                       | 如果指定范围被折叠且位于表格中的行尾标志处，则该属性值为 True 。只读 Boolean 类型。                                                                                     |
|  46  | [Italic](#Range.Italic)                                       | 如果将字体或范围设置为倾斜格式，则该属性值为 True 。可读写 Long 类型。                                                                                                  |
|  47  | [ItalicBi](#Range.ItalicBi)                                   | 如果将字体或范围设置为倾斜格式，则该属性值为 True 。可读写 Long 类型。                                                                                                  |
|  48  | [Kana](#Range.Kana)                                           | 返回或设置日文文本的指定区域是平假名还是片假名。 WdKana 类型，可读写。                                                                                                  |
|  49  | [LanguageDetected](#Range.LanguageDetected)                   | 返回或设置一个值，该值指定 WPS 是否已经检测过指定文本的语言。可读/写 Boolean 类型。                                                                                     |
|  50  | [LanguageID](#Range.LanguageID)                               | 返回或设置一个 WdLanguageID 常量，该常量代表指定范围的语言。可读写。                                                                                                    |
|  51  | [LanguageIDFarEast](#Range.LanguageIDFarEast)                 | 返回或设置指定对象的东亚语言。可读写 WdLanguageID 类型。                                                                                                                |
|  52  | [LanguageIDOther](#Range.LanguageIDOther)                     | 返回或设置指定范围的语言。可读写 WdLanguageID 类型。                                                                                                                    |
|  53  | [ListFormat](#Range.ListFormat)                               | 返回一个 ListFormat 对象，该对象代表某区域中所有的列表格式特征。只读。                                                                                                  |
|  54  | [ListParagraphs](#Range.ListParagraphs)                       | 返回一个 ListParagraphs 集合，该集合代表范围中的所有编号段落。只读。                                                                                                    |
|  55  | [ListStyle](#Range.ListStyle)                                 | 返回一个 Variant 类型的值，该值代表用于设置项目符号列表或编号列表的样式。只读。                                                                                         |
|  56  | [Locks](#Range.Locks)                                         | 返回 CoAuthLocks 集合对象，该对象代表范围中的所有锁。只读。                                                                                                             |
|  57  | [NextStoryRange](#Range.NextStoryRange)                       | 返回一个 Range 对象，该对象代表下一个 文章 。 Range 类型，只读。                                                                                                        |
|  58  | [NoProofing](#Range.NoProofing)                               | 如果拼写和语法检查程序忽略指定文本，则该属性值为 True 。可读写 Long 类型。                                                                                              |
|  59  | [OMaths](#Range.OMaths)                                       | 返回一个 OMaths 集合，该集合代表指定区域内的 OMath 对象。只读。                                                                                                         |
|  60  | [Orientation](#Range.Orientation)                             | 在启用了“文字方向”功能时返回或设置范围中文字的方向。可读写 WdTextOrientation 类型。                                                                                     |
|  61  | [PageSetup](#Range.PageSetup)                                 | 返回一个 PageSetup 对象，该对象与指定范围相关联。                                                                                                                       |
|  62  | [ParagraphFormat](#Range.ParagraphFormat)                     | 返回或设置一个 ParagraphFormat 对象，该对象代表指定范围的段落设置。可读/写。                                                                                            |
|  63  | [Paragraphs](#Range.Paragraphs)                               | 返回一个 Paragraphs 集合，该集合代表指定范围中的所有段落。只读。                                                                                                        |
|  64  | [ParagraphStyle](#Range.ParagraphStyle)                       | 返回一个 Variant 类型的值，该值代表用于设置段落格式的样式。只读。                                                                                                       |
|  65  | [Parent](#Range.Parent)                                       | 返回一个 Object 类型值，该值代表指定 Range 对象的父对象。                                                                                                               |
|  66  | [ParentContentControl](#Range.ParentContentControl)           | 返回一个 ContentControl 对象，该对象代表指定区域的父内容控件。只读。                                                                                                    |
|  67  | [PreviousBookmarkID](#Range.PreviousBookmarkID)               | 返回最后一个书签的编号，该书签从指定范围的前面或与指定范围相同的位置开始。只读 Long 类型。                                                                              |
|  68  | [ReadabilityStatistics](#Range.ReadabilityStatistics)         | 返回一个 ReadabilityStatistics 集合，该集合代表指定文档或范围的可读性统计信息。只读。                                                                                   |
|  69  | [Revisions](#Range.Revisions)                                 | 返回一个 Revisions 集合，该集合代表范围中的修订。只读。                                                                                                                 |
|  70  | [Rows](#Range.Rows)                                           | 返回一个 Rows 集合，该集合代表范围中的所有表格行。只读。                                                                                                                |
|  71  | [Scripts](#Range.Scripts)                                     | 返回一个 Scripts 集合，该集合代表指定对象中 HTML 脚本的集合。                                                                                                           |
|  72  | [Sections](#Range.Sections)                                   | 返回一个 Sections 集合，该集合代表指定范围中的各节。只读。                                                                                                              |
|  73  | [Sentences](#Range.Sentences)                                 | 返回一个 Sentences 集合，该集合代表范围中的所有句子。只读。                                                                                                             |
|  74  | [Shading](#Range.Shading)                                     | 返回一个 Shading 对象，该对象代表指定对象的底纹格式。                                                                                                                   |
|  75  | [ShapeRange](#Range.ShapeRange)                               | 返回一个 ShapeRange 集合，该集合代表指定范围中的所有 Shape 对象。只读。                                                                                                 |
|  76  | [ShowAll](#Range.ShowAll)                                     | 如果显示所有非打印字符（如隐藏文字、制表符、空格和段落标记），则该属性值为 True 。可读写 Boolean 类型。                                                                 |
|  77  | [SpellingChecked](#Range.SpellingChecked)                     | 如果已对指定的区域或文档完成拼写检查，则该属性值为 True 。如果所有或部分区域或文档尚未进行拼写检查，则该属性值为 False 。 Boolean 类型，可读写。                        |
|  78  | [SpellingErrors](#Range.SpellingErrors)                       | 返回一个 ProofreadingErrors 集合，该集合代表指定范围中标识为拼写错误的单词。只读。                                                                                      |
|  79  | [Start](#Range.Start)                                         | 返回或设置某区域中起始字符的位置。 Long 类型，可读写。                                                                                                                  |
|  80  | [StoryLength](#Range.StoryLength)                             | 返回包含指定区域的 文字部分 中的字符数。 Long 类型，只读。                                                                                                              |
|  81  | [StoryType](#Range.StoryType)                                 | 返回指定范围、所选内容或书签的 文字部分 类型。只读 WdStoryType 类型。                                                                                                   |
|  82  | [Style](#Range.Style)                                         | 返回或设置指定对象的样式。可读写 Variant 类型。                                                                                                                         |
|  83  | [Subdocuments](#Range.Subdocuments)                           | 返回一个 Subdocuments 集合，该集合代表指定范围或文档中的所有子文档。只读。                                                                                              |
|  84  | [SynonymInfo](#Range.SynonymInfo)                             | 返回一个 SynonymInfo 对象，该对象包含同义词库中有关某范围的内容的同义词、反义词或相关单词和表达方式的信息。                                                             |
|  85  | [Tables](#Range.Tables)                                       | 返回一个 Tables 集合，该集合代表指定范围内的所有表格。只读。                                                                                                            |
|  86  | [TableStyle](#Range.TableStyle)                               | 返回一个 Variant 类型的值，该值代表用于设置表格格式的样式。只读。                                                                                                       |
|  87  | [Text](#Range.Text)                                           | 返回或设置指定区域或选定内容中的文本。 String 类型，可读写。                                                                                                            |
|  88  | [TextRetrievalMode](#Range.TextRetrievalMode)                 | 返回一个 TextRetrievalMode 对象，该对象控制从指定 区域 检索文字的方式。可读写。                                                                                         |
|  89  | [TopLevelTables](#Range.TopLevelTables)                       | 返回一个 Tables 集合，该集合代表当前范围最外部嵌套层上的表格。只读。                                                                                                    |
|  90  | [TwoLinesInOne](#Range.TwoLinesInOne)                         | 返回或设置 WPS 是否将两行文本合并为一行，并指定括住文本的字符（如果有）。 WdTwoLinesInOneType 类型，可读写。                                                            |
|  91  | [Underline](#Range.Underline)                                 | 返回或设置应用于范围的下划线的类型。可读写 WdUnderline 类型。                                                                                                           |
|  92  | [Updates](#Range.Updates)                                     | 返回 CoAuthUpdates 集合对象，该对象代表范围中的所有可用更新。只读。                                                                                                     |
|  93  | [WordOpenXML](#Range.WordOpenXML)                             | 返回一个 String 类型的值，该值以 WPS Open XML 格式表示区域中包含的 XML。只读。                                                                                          |
|  94  | [Words](#Range.Words)                                         | 返回一个 Words 集合，该集合代表范围中的所有单词。只读。                                                                                                                 |
|  95  | [XMLNodes](#Range.XMLNodes)                                   | 返回一个 XMLNodes 集合，该集合代表指定区域中的 XML 元素（包括任何只是部分属于该区域的元素）。只读。                                                                     |
|  96  | [XMLParentNode](#Range.XMLParentNode)                         | 返回一个 XMLNode 对象，该对象代表区域的父级 XML 节点。只读。                                                                                                            |

## 成员方法

### Range.Collapse

将某一区域或所选内容折叠到起始位置或结束位置。折叠之后起始位置和结束位置相同。

**语法**

**express.Collapse ()**

\*express是一个代表 Range 对象的变量。

**说明**

如果用 **wdCollapseEnd** 折叠一个代表完整段落的区域，则该区域将定位于段落结束标记之后（即下段开头）。但是，在该区域折叠后，可以用 **MoveEnd** 方法将区域回移一个字符

**示例**

``` JavaScript
/*光标定位第一段段位然后再最后插入内容*/
function test()
{
  let mgRange = Application.ActiveDocument.Paragraphs.Item(1).Range
  mgRange.Collapse(wdCollapseEnd)
  mgRange.Text = "111"
}
```

### Range.Copy

将指定范围复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例复制活动文档的第一段，并将该段落粘贴到文档的末尾。*/

function test()
{
  let doc = Application.ActiveDocument
  doc.Paragraphs.Item(1).Range.Copy()
  let myRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
  myRange.Paste()
}
```

### Range.Cut

将指定对象从文档中移到剪贴板上

**语法**

**express.Cut ()**

\*express是一个代表 Range 对象的变量。

**说明**

将区域的内容移动到剪贴板上，但折叠区域仍保留在文档中。

**示例**

``` JavaScript
/*
本示例剪切活动文档的第一段，并将该段落粘贴到文档的末尾。
*/
function test()
{
    let doc = Application.ActiveDocument
    doc.Paragraphs.Item(1).Range.Cut()
    let myRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    myRange.Paste()
}
```

### Range.Delete

删除指定数量的字符或单词

**语法**

**express.Delete ( *Unit* , *Count* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                 |
|---------|-----------|----------|----------------------------------------------------------------------------------------------------------------------|
| *Unit*  | 可选      | Variant  | 删除折叠区域时所基于的单位。可以是 WdUnits 常量之一。                                                                |
| *Count* | 可选      | Variant  | 要删除的单位数。要删除该区域后的单位，请折叠该区域并使用一个正值。要删除该区域前的单位，请折叠该区域并使用一个负值。 |

**说明**

要删除的单位数。要删除该区域后的单位，请折叠该区域并使用一个正值。要删除该区域前的单位，请折叠该区域并使用一个负值。

**示例**

``` JavaScript
/*
本示例选择并删除活动文档中的内容。
*/
function test()
{
    Application.ActiveDocument.Content.Select()
    Application.Selection.Range.Delete()
}
```

### Range.Expand

扩展指定区域或选定内容。返回添至该区域或选定内容的字符数。 **Long** 类型。

**语法**

**express.Expand ( *Unit* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                                                                                        |
|--------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Unit* | 可选      | Variant  | 扩展区域时所基于的单位。可以是下列 WdUnits 常量之一：wdCharacter、wdWord、wdSentence、wdParagraph、wdSection、wdStory、wdCell、wdColumn、wdRow 或 wdTable。 |

**示例**

``` JavaScript
/*本示例创建指向活动文档第一个单词的区域，然后将该区域扩展至文档首段。*/
function test() {
  let myRange = Application.ActiveDocument.Words.Item(1)
  myRange.Expand(wdParagraph)
}

/*本示例先将选定内容的首个字符设为大写，再将选定内容扩展至整句。*/
function test() {
  let characters = Application.Selection
  characters.Characters.Item(1).Case = wdTitleSentence
  characters.Expand(wdSentence)
}
```

### Range.GoTo

返回一个 **Range** 对象，该对象代表指定项（如页、书签或域）的起始位置。

**语法**

**express.GoTo ( *What* , *Which* , *Count* , *Name* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                        |
|---------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *What*  | 可选      | Variant  | 范围要移动到的项的类别。可以是 WdGoToItem 常量之一。                                                                                                                                                        |
| *Which* | 可选      | Variant  | 范围要移动到的项。可以是 WdGoToDirection 常量之一。                                                                                                                                                         |
| *Count* | 可选      | Variant  | 文档中的项数。默认值为 1。只有正值有效。若要指定一个位于该范围之前的项，可将 wdGoToPrevious 用作 Which 参数，并指定一个 Count 值。                                                                          |
| *Name*  | 可选      | Variant  | 如果 What 参数为 wdGoToBookmark、wdGoToComment、wdGoToField 或 wdGoToObject，则此参数指定一个名称。只有正值有效。若要指定一个位于该范围之前的项，可将 wdGoToPrevious 用作 Which 参数，并指定一个 Count 值。 |

**说明**

以下示例将范围向上移动两行。

``` JavaScript
Application.ActiveDocument.Range().GoTo(wdGoToLine, wdGoToPrevious, 2)
```

以下示例将所选内容移至下一个 DATE 域。

``` JavaScript
Application.ActiveDocument.Range().GoTo(wdGoToField, null, null, "Date") 
```

以下示例将范围移至文档中的第四行。

``` JavaScript
Application.ActiveDocument.Range().GoTo(wdGoToLine, wdGoToAbsolute, 4)
```

下列示例的功能等效，都将范围移至文档中的第一个标题处。

``` JavaScript
Application.ActiveDocument.Range().GoTo(wdGoToHeading, wdGoToFirst)
Application.ActiveDocument.Range().GoTo(wdGoToHeading, wdGoToAbsolute, 1)
```

将 **GoTo** 方法与 **wdGoToGrammaticalError** 、 **wdGoToProofreadingError** 或 **wdGoToSpellingError** 常量一起使用时，返回的 **Range** 对象中包括所有含语法或拼写错误的文本。

**示例**

``` JavaScript
/*以下示例将插入点移至活动文档中第五个尾注引用标记的前面。*/
function test() {
  if(Application.ActiveDocument.Endnotes.Count >= 5){
      Application.ActiveDocument.Range().GoTo(wdGoToEndnote, wdGoToAbsolute, 5)
  }
}

/*以下示例将 R1 设置为等于活动文档中第一个脚注引用标记。*/
function test() {
  if(Application.ActiveDocument.Footnotes.Count >= 1){
      let R1 = Application.ActiveDocument.Range().GoTo(wdGoToFootnote, wdGoToFirst)
      R1.Expand(wdCharacter)
  }
}
```

### Range.InsertAfter

在范围的末尾插入指定文本。

**语法**

**express.InsertAfter ( *Text* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明           |
|--------|-----------|----------|----------------|
| *Text* | 可选      | String   | 要插入的文本。 |

**说明**

应用此方法之后，该范围将扩展，以包含新文本。

使用 Visual Basic **Chr** 函数和 **InsertAfter** 方法，可以插入引号、制表符和不间断连字符等字符。还可以使用下列 Visual Basic 常量： **vbCr** 、 **vbLf** 、 **vbCrLf** 和 **vbTab** 。

如果对引用整个段落的范围使用此方法，则在末段标记之后插入文本（插入的文本将出现在下一段的开头）。要在段尾插入文本，请先确定终点，再从该位置减去 1（因为段落标记是一个字符），如以下示例所示。

``` JavaScript
function test() {
    let doc = Application.ActiveDocument
    let rngRange = doc.Range(doc.Paragraphs.Item(1).Start, doc.Paragraphs.Item(1).End - 1)
    rngRange.InsertAfter(" This is now the last sentence in paragraph one.")
}
```

然而，如果该范围以一个段落标记结尾，而该段落标记正好又是文档的末尾，则 WPS 在末段标记前插入文本，而不是在文档末尾创建一个新段落。

同样，如果该范围是书签， WPS 将插入指定的文本，但不会扩展范围或书签以包含新文本。

**示例**

``` JavaScript
/*以下示例在活动文档的末尾插入文本。Content 属性返回一个 Range 对象。*/
Application.ActiveDocument.Content.InsertAfter("end of document")

/*以下示例将输入框中的文本作为活动文档的第二段插入到文档中。*/
function test() {
  let response = prompt("Type some text")
  let range2 = Application.ActiveDocument.Paragraphs.Item(1).Range
  range2.InsertAfter("1." + "\t" + response)
  range2.InsertParagraphAfter()
}
```

### Range.InsertBefore

在指定的范围前插入指定文本。

**语法**

**express.InsertBefore ( *Text* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明           |
|--------|-----------|----------|----------------|
| *Text* | 必选      | String   | 要插入的文本。 |

**说明**

在插入文本之后，该范围将扩展，以包含新文本。如果该范围是书签，则书签也会扩展，以包含新文本。

使用 Visual Basic **Chr** 函数和 **InsertBefore** 方法，可以插入引号、制表符和不间断连字符等字符。还可以使用下列 Visual Basic 常量： **vbCr** 、 **vbLf** 、 **vbCrLf** 和 **vbTab** 。

**示例**

``` JavaScript
/*以下示例在活动文档的开头插入文本“Introduction”，并将其作为一个单独段落。*/
function test() {
  let content = Application.ActiveDocument.Content
  content.InsertParagraphBefore()
  content.InsertBefore("Introduction")
}
```

### Range.InsertBreak

插入分页符、分栏符或分节符。

**语法**

**express.InsertBreak ( *Type* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                                                      |
|--------|-----------|----------|-------------------------------------------------------------------------------------------|
| *Type* | 必选      | Variant  | 要插入的分隔符类型。可以是 WdBreakType 常量之一。如果省略该参数，则默认值为 wdPageBreak。 |

**说明**

当插入分页符或分栏符时，范围将被插入的分隔符所替换。如果不希望替换该范围，可在使用 **InsertBreak** 方法之前使用 **Collapse** 方法。在插入分节符时，此分节符紧接在 **Range** 之前插入。

根据您所选择或安装的语言支持（例如，美国英语），上述部分常量可能不可用。

**示例**

``` JavaScript
/*以下示例紧接在活动文档第二段之后插入一个分页符。*/
function test() {
  let myRange = Application.ActiveDocument.Paragraphs.Item(2).Range
  myRange.Collapse(wdCollapseEnd)
  myRange.InsertBreak(wdPageBreak)
}
```

### Range.InsertFile

插入指定文件的全部或一部分。

**语法**

**express.InsertFile ( *FileName* , *Range* , *ConfirmConversions* , *Link* , *Attachment* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                                                                                                                        |
|----------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------|
| *FileName*           | 必选      | String   | 要插入的文件的路径及文件名。如果没有指定路径，则 WPS 假定文件位于当前文件夹中。                                                             |
| *Range*              | 可选      | Variant  | 如果指定的文件是 WPS 文档，则此参数代表书签。如果该文件为其他类型（例如，ET 工作表），则此参数代表命名区域或单元格区域（例如，R1C1:R3C4）。 |
| *ConfirmConversions* | 可选      | Variant  | 如果该参数值为 True，则在插入非 WPS 文档格式的文件时， WPS 将提示您确认转换。                                                               |
| *Link*               | 可选      | Variant  | 如果该参数值为 True，则使用 INCLUDETEXT 域插入该文件。                                                                                      |
| *Attachment*         | 可选      | Variant  | 如果该参数值为 True，则将该文件以附件形式插入电子邮件中。                                                                                   |

**示例**

``` JavaScript
/* 在当前文档的最后插入文件"MyDoc.doc"的内容 */
function test() {
    let end = Application.ActiveDocument.Range().End;
    Application.Selection.Start = end;
    Application.Selection.End = end;    
    Application.Selection.Range.InsertFile("C:\\MyDoc.doc");
}
```

### Range.InsertParagraph

用新段落替换指定范围。

**语法**

**express.InsertParagraph ()**

\*express是一个代表 Range 对象的变量。

**说明**

使用此方法后，该范围将成为一个新段落。

如果您不希望替换该范围，可在使用此方法之前先使用 **Collapse** 方法。 **InsertParagraphAfter** 方法可在 **Range** 对象后插入一个新段落。

**示例**

``` JavaScript
/*以下示例在活动文档的开头插入一个新段落。*/
function test() {
  let myRange = Application.ActiveDocument.Range(0, 0)
  myRange.InsertParagraph()
  myRange.InsertBefore("Dear Sirs,")
}
```

### Range.Paste

将“剪贴板”中的内容插入指定范围。

**语法**

**express.Paste ()**

\*express是一个代表 Range 对象的变量。

**说明**

如果不希望替换该范围的内容，可在使用此方法之前先使用 **Collapse** 方法。

如果您将此方法与 **Range** 对象一起使用，该范围将扩展以包含“剪贴板”中的内容。

**示例**

``` JavaScript
/*以下示例复制活动文档中第一个表格，并将其粘贴至新文档。*/
function test() {
  if(Application.ActiveDocument.Tables.Count >= 1){
      Application.ActiveDocument.Tables.Item(1).Range.Copy()
      Application.Documents.Add().Content.Paste()
  }
}

/*以下示例复制所选内容并将其粘贴到文档末尾。*/
function test() {
  if(Application.Selection.Type != wdSelectionIP){
      Application.Selection.Copy()
      let Range2 = Application.ActiveDocument.Content
      Range2.Collapse(wdCollapseEnd)
      Range2.Paste()
  }
}
```

### Range.Range

将“剪贴板”中的内容插入指定范围。

**语法**

**express.Range ()**

\*express是一个代表 Range 对象的变量。

**说明**

如果不希望替换该范围的内容，可在使用此方法之前先使用 **Collapse** 方法。

如果您将此方法与 **Range** 对象一起使用，该范围将扩展以包含“剪贴板”中的内容。

**示例**

``` JavaScript
/*
复制活动文档中第一个表格，并将其粘贴至新文档
*/
function test()
{
    let doc = Application.ActiveDocument
    if (doc.Tables.Count)
        doc.Tables.Item(1).Range.Copy()
    Application.Documents.Add().Content.Paste();
}
```

### Range.Select

选择指定的范围。

**语法**

**express.Select ()**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例选择活动文档中的第一个段落。*/
function test(){
  Application.ActiveDocument.Paragraphs.Item(1).Range.Select()
  Application.Selection.Font.Bold = true
}
```

### Range.XML

返回一个 **String** 类型的值，该值代表指定对象中的 XML 文本。

**语法**

**express.XML ( *DataOnly* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                          |
|------------|-----------|----------|-------------------------------------------------------------------------------|
| *DataOnly* | 可选      | Boolean  | 如果该属性值为 True，则返回不带 WPS XML 标记的 XML 的文本。默认设置为 False。 |

## 成员属性

### Range.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Range 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Range.Bold

如果选定区域中字体的格式为加粗，则该属性值为 **True** 。 **Long** 类型，可读写。

**语法**

**express.Bold**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*
把第一段的内容属性加粗
*/
Application.ActiveDocument.Paragraphs.Item(1).Range.Bold = true
```

### Range.BoldBi

如果字体或区域为加粗格式，则为 **True** 。该属性值返回 **True** 、 **False** 或 **wdUndefined** （用于加粗和非加粗混合文本）。可设置为 **True** 、 **False** 或 **wdToggle** 。 **Long** 类型，可读写。

**语法**

**express.BoldBi**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将从右向左语言的活动文档中第一段的格式设为加粗。*/
ActiveDocument.Paragraphs.Item(1).Range.BoldBi = true
```

### Range.BookmarkID

返回位于指定区域开始位置的书签编号。如果没有相应的书签，则返回 0（零）。 **Long** 类型，只读。

**语法**

**express.BookmarkID**

\*express是一个代表 Range 对象的变量。

**说明**

编号或书签对应于书签在文档中的位置：1 对应于第一个书签，2 对应于第二个，以此类推。

**示例**

``` JavaScript
/*如果文档的开始位置没有设置书签，则本示例添加一个名为“temp”的书签。*/
function test()
{
let myRange = ActiveDocument.Content
myRange.Collapse(wdCollapseStart)
if(myRange.BookmarkID == 0){
    ActiveDocument.Bookmarks.Add("temp", myRange)
}
}
```

### Range.Bookmarks

返回一个 **Bookmarks** 集合。该集合代表某一文档、区域或选定内容中的所有书签。只读。

**语法**

**express.Bookmarks**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.Case

返回或设置一个 **WdCharacterCase** 常量，该常量代表指定区域中文字的大小写。可读写。

**语法**

**express.Case**

\*express是一个代表 Range 对象的变量。

**说明**

由于您选择或安装的语言支持不同（例如美国英语），上述部分常量可能无法使用。

### Range.Cells

返回一个 **Cells** 集合，该集合代表在某区域中的表格单元格。只读。

**语法**

**express.Cells**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例创建一个 3x3 表格并给表格中的每个单元格依次分配一个单元格编号。*/
function test()
{
let newDoc = Documents.Add()
let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)
let i = 1
for(let c = 1; c <= myTable.Range.Cells.Count; c++){
    myTable.Range.Cells.Item(c).Range.InsertAfter("Cell " + i)
    i = i + 1
}
}
```

### Range.Characters

返回一个 **Characters** 集合，该集合代表区域中的字符。只读。

**语法**

**express.Characters**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.CharacterStyle

返回一个 **Variant** 类型的值，该值代表用于设置一个或多个字符格式的样式。只读。

**语法**

**express.CharacterStyle**

\*express是一个代表 Range 对象的变量。

### Range.CharacterWidth

返回或设置指定区域的字符宽度。 **WdCharacterWidth** 类型，可读写。

**语法**

**express.CharacterWidth**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将当前选定内容转换为半角字符。*/
Selection.Range.CharacterWidth = wdWidthHalfWidth
```

### Range.Columns

返回一个 **Columns** 集合，该集合代表区域中的所有表格列。只读。

**语法**

**express.Columns**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示活动文档的第一个表格中的列数。*/
function test()
{
if(ActiveDocument.Tables.Count >= 1){
    MsgBox(ActiveDocument.Tables.Item(1).Columns.Count)
}
}
```

``` JavaScript
/*本示例将当前列的宽度设置为 1 英寸。*/
function test()
{
if(Selection.Information(wdWithInTable) == true) {
    Selection.Columns.SetWidth(InchesToPoints(1), wdAdjustProportional)
}
}
```

### Range.CombineCharacters

如果指定区域包含合并字符，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CombineCharacters**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例合并选定区域中的字符。*/
Selection.Range.CombineCharacters = true
```

### Range.Comments

返回一个 **Comments** 集合，该集合代表指定文档、选定内容或区域中的所有批注。只读。

**语法**

**express.Comments**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.Conflicts

返回 Conflicts 集合对象，该对象包含范围中的所有冲突对象。只读。

**语法**

**express.Conflicts**

\*express是一个代表 Range 对象的变量。

**说明**

使用 **Conflicts** 属性可以返回文档的 Conflicts 集合对象。使用 Conflicts( *Index* )（其中 *Index* 为冲突索引号）可以返回一个 Conflict 对象。

| 注释                                                                                             |
|--------------------------------------------------------------------------------------------------|
| 该属性仅可用于支持共同创作的文档。如果尝试在不支持共同创作的文档中访问该属性，将导致运行时错误。 |

**示例**

``` JavaScript
/*下面的代码示例显示活动文档的第一个段落中的冲突数量。*/
MsgBox(ActiveDocument.Paragraphs.Item(1).Range.Conflicts.Count)
```

### Range.ContentControls

返回一个 **ContentControls** 集合，该集合代表一个区域中包含的内容控件。只读。

**语法**

**express.ContentControls**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*下面的示例将一个下拉列表内容控件插入到活动文档中的指定位置。*/
function test()
{
let objRange = ActiveDocument.Range(200, 200)
let objCC = objRange.ContentControls.Add(wdContentControlDropdownList)

//List entries
objCC.DropdownListEntries.Add("Cat")
objCC.DropdownListEntries.Add("Dog")
objCC.DropdownListEntries.Add("Horse")
objCC.DropdownListEntries.Add("Monkey")
objCC.DropdownListEntries.Add("Snake")
objCC.DropdownListEntries.Add("Other")
}
```

### Range.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Range 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Range.DisableCharacterSpaceGrid

如果 WPS 忽略相应 **Range** 对象的每行字符数，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisableCharacterSpaceGrid**

\*express是一个代表 Range 对象的变量。

**说明**

如果仅将部分指定字体或区域的 **DisableCharacterSpaceGrid** 属性设置为 **True** ，则返回 **wdUndefined** 。

### Range.Document

返回与指定区域相关的 **Document** 对象。只读。

**语法**

**express.Document**

\*express是一个代表 Range 对象的变量。

### Range.Duplicate

返回一个只读 **Range** 对象，该对象代表指定区域的所有属性。

**语法**

**express.Duplicate**

\*express是一个代表 Range 对象的变量。

**说明**

通过复制 **Range** 对象，可更改所复制的区域的开始或结尾字符的位置，而不会更改源区域。

### Range.Editors

返回一个 **Editors** 对象，该对象代表已授权修改文档中选定内容或区域的所有用户。

**语法**

**express.Editors**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例给当前用户分配编辑权限以修改活动的选定内容。*/
let objEditor = Selection.Editors.Add(wdEditorCurrent)
```

### Range.EmphasisMark

返回或设置字符或指定的字符串的着重号。 **WdEmphasisMark** 类型，可读写。

**语法**

**express.EmphasisMark**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：在活动文档的第四个单词上面设置着重号：逗号。*/
ActiveDocument.Words.Item(4).EmphasisMark = wdEmphasisMarkOverComma
```

### Range.End

返回或设置某区域中结束字符的位置。可读/写 **Long** 类型。

**语法**

**express.End**

\*express是一个代表 Range 对象的变量。

**说明**

**Range** 对象均包含开始位置和结束位置。结束位置是距文档开头部分最远的点。如果该属性的设置值小于 **Start** 属性值，则 **Start** 属性将设为同一值（即 **Start** 与 **End** 属性值相等）。

该属性返回结束字符相对于文档开头部分的位置。文档主体部分 ( **wdMainTextStory** ) 的起始字符位置为 0（零）。设置该属性可以改变选定内容、区域或者书签的大小。

**示例**

``` JavaScript
/*本示例将 myRange 的结束位置移动一个字符。*/
function test() {
  let myRange = Application.ActiveDocument.Paragraphs.Item(1).Range
  myRange.End = myRange.End - 1
}
```

### Range.EndnoteOptions

返回一个 **EndnoteOptions** 对象，该对象代表区域中的尾注。

**语法**

**express.EndnoteOptions**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*如果活动文档第二节中尾注的起始号码不是 1，则本示例将其设为 1。*/
function SetEndnoteOptionsRange() {
    let Range2 = ActiveDocument.Sections.Item(2).Range.EndnoteOptions
    if(Range2.StartingNumber != 1) {
        Range2.StartingNumber = 1
    }
}
```

### Range.Endnotes

返回一个 **Endnotes** 集合，该集合代表区域中的所有尾注。只读。

**语法**

**express.Endnotes**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.EnhMetaFileBits

返回一个 **Variant** 类型的值，该值代表文本区域的显示方式的图片代表形式。

**语法**

**express.EnhMetaFileBits**

\*express是一个代表 Range 对象的变量。

**说明**

**EnhMetaFileBits** 属性返回一个字节数组，该数组可在 Visual Basic 或 Microsoft C++ 开发环境中通过 Microsoft Windows 32 应用程序编程接口来使用。

### Range.Fields

返回一个 **Fields** 集合，该集合代表区域中的所有域。只读。

**语法**

**express.Fields**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例删除活动文档正文和页脚中的所有域。*/
function test()
{
for(let aField = 1; aField <= ActiveDocument.Fields.Count; aField++) {
    ActiveDocument.Fields.Item(aField).Delete()
}
let myRange = ActiveDocument.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).Range
for(let a = 1; a <= myRange.Fields.Count; a++) {
    myRange.Fields.Item(a).Delete()
}
}
```

### Range.Find

返回一个 **Find** 对象，该对象包含查找操作所需的条件。只读。

**语法**

**express.Find**

\*express是一个代表 Range 对象的变量。

### Range.FitTextWidth

该属性返回或设置 WPS 在当前选定内容或区域中填入文字的宽度（使用当前的度量单位）。 **Single** 类型，可读写。

**语法**

**express.FitTextWidth**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将当前选定内容填入 5 厘米宽的空间。*/
Selection.FitTextWidth = CentimetersToPoints(5)
```

### Range.Font

返回或设置 **Font** 对象，该对象代表指定对象的字符格式。 **Font** 类型，可读写。

**语法**

**express.Font**

\*express是一个代表 Range 对象的变量。

**说明**

要设置该属性，需指定一个返回 **Font** 对象的表达式。

**示例**

``` JavaScript
/*本示例将取消活动文档的“标题 1”样式中的加粗格式。*/
Application.ActiveDocument.Styles.Item(wdStyleHeading1).Font.Bold = false

/*本示例在 Arial 和 Times New Roman 之间切换活动文档中第二段的字体。*/
function test() {
  let myRange = Application.ActiveDocument.Paragraphs.Item(2).Range
  if(myRange.Font.Name == "Times New Roman"){
      myRange.Font.Name = "Arial"
  }
  else{
      myRange.Font.Name = "Times New Roman"
  }
}
```

### Range.FootnoteOptions

返回 **FootnoteOptions** 对象，该对象代表选定内容或区域的脚注。

**语法**

**express.FootnoteOptions**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例设置第二节中的编号规则以便在新节开始时重新编号。*/
function SetFootnoteOptionsRange() {
    ActiveDocument.Sections.Item(2).Range.FootnoteOptions
        .NumberingRule = wdRestartSection
}
```

### Range.Footnotes

返回一个 **Footnotes** 集合，该集合代表区域中的所有脚注。只读。

**语法**

**express.Footnotes**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.FormattedText

返回或设置一个 **Range** 对象，该对象包含指定区域或选定内容中进行过格式编排的文字。可读写。

**语法**

**express.FormattedText**

\*express是一个代表 Range 对象的变量。

**说明**

此属性返回 **Range** 对象，以及指定的区域或所选内容中的字符格式和文本。如果在区域或所选内容中有一个段落标记，则 **Range** 对象中包含段落格式。

在设置此属性时，区域中的文本会被格式文本替换。如果不想替换现有的文本，则在使用此属性之前使用 **Collapse** 方法（见第一个示例）。

### Range.FormFields

返回一个 **FormFields** 集合，该集合代表区域中的所有窗体域。只读。

**语法**

**express.FormFields**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例检索第二节第一个窗体域的类型。*/
function test()
{
let myType = ActiveDocument.Sections.Item(2).Range.FormFields.Item(1).Type
switch(myType){
    case wdFieldFormTextInput: thetype = "TextBox"
        break
    case wdFieldFormDropDown : thetype = "DropDown"
        break
    case wdFieldFormCheckBox : thetype = "CheckBox"
        break
}
}
```

### Range.Frames

返回一个 **Frames** 集合，该集合代表区域中的所有图文框。只读。

**语法**

**express.Frames**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例使活动文档第一节中的文字环绕图文框。*/
function test()
{
let rng = ActiveDocument.Sections.Item(1).Range
for(let aFrame = 1; aFrame <= rng.Frames.Count; aFrame++) {
    rng.Frames.Item(aFrame).TextWrap = true
}
}
```

### Range.GrammarChecked

如果已经检查了指定范围或文档的语法，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.GrammarChecked**

\*express是一个代表 Range 对象的变量。

**说明**

如果指定的区域或文档的部分或全部未进行语法检查，则该属性值为 **False** 。要重新检查区域或文档的语法，可将 **GrammarChecked** 属性设置为 **False** 。

### Range.GrammaticalErrors

返回一个 **ProofreadingErrors** 集合，该集合代表指定文档或区域中有语法检查错误的句子。只读。

**语法**

**express.GrammaticalErrors**

\*express是一个代表 Range 对象的变量。

**说明**

在每个句子中可有多个错误。如果文档中没有语法错误，则由 **GrammaticalErrors** 属性返回的 **ProofreadingErrors** 对象的 **Count** 属性的返回值为 0（零）。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例检查活动文档的第三段的语法错误，并显示包含一个或多个错误的每一句。*/
function test()
{
let myErrors = ActiveDocument.Paragraphs.Item(3).Range.GrammaticalErrors
for(let myerr = 1; myerr <= myErrors.Count; myerr++) {
    MsgBox(myErrors.Item(myerr).Text)
}
}
```

### Range.HighlightColorIndex

返回或设置指定区域的突出显示颜色。 **WdColorIndex** 类型，可读写。

**语法**

**express.HighlightColorIndex**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例删除选定内容中的突出显示格式。*/
Selection.Range.HighlightColorIndex = wdNoHighlight
```

``` JavaScript
/*本示例用黄色来突出显示活动文档中的每个书签。*/
function test()
{
let myBookmarks = ActiveDocument.Bookmarks
for(let abookmark = 1; abookmark <= myBookmarks.Count; abookmark++) {
    myBookmarks.Item(abookmark).Range.HighlightColorIndex = wdYellow
}
}
```

### Range.HorizontalInVertical

返回或设置位于垂直排列文字中的水平排列文字的格式。 **WdHorizontalInVerticalType** 类型，可读写。

**语法**

**express.HorizontalInVertical**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将垂直排列文字中的局部所选内容设置为水平排列的文字格式，并调整文字以适应垂直排列文字的行宽。*/
Selection.Range.HorizontalInVertical = wdHorizontalInVerticalFitInLine
```

### Range.HTMLDivisions

返回一个 **HTMLDivisions** 对象，该对象代表 Web 文档中的 HTML 划分。

**语法**

**express.HTMLDivisions**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例设置活动文档中三个嵌套划分的格式。本示例假定该活动文档为具有至少三个划分的 HTML 文档。*/
function test()
{
let htmlDiv1 = ActiveDocument.Range().HTMLDivisions.Item(1)
    htmlDiv1.Borders.Item(wdBorderLeft).Color = wdColorRed
    htmlDiv1.Borders.Item(wdBorderLeft).LineStyle = wdLineStyleSingle
    htmlDiv1.Borders.Item(wdBorderRight).Color = wdColorRed
    htmlDiv1.Borders.Item(wdBorderRight).LineStyle = wdLineStyleSingle
    
    let htmlDiv2 = htmlDiv1.HTMLDivisions.Item(1)
    htmlDiv2.LeftIndent = InchesToPoints.Item(1)
    htmlDiv2.RightIndent = InchesToPoints.Item(1)
    htmlDiv2.Borders.Item(wdBorderTop).Color = wdColorBlue
    htmlDiv2.Borders.Item(wdBorderTop).LineStyle = wdLineStyleDouble
    htmlDiv2.Borders.Item(wdBorderBottom).Color = wdColorBlue
    htmlDiv2.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleDouble
        
    let htmlDiv3 = htmlDiv2.HTMLDivisions.Item(1)
    htmlDiv3.LeftIndent = InchesToPoints(1)
    htmlDiv3.RightIndent = InchesToPoints(1)
    htmlDiv3.Borders.Item(wdBorderLeft).LineStyle = wdLineStyleDot
    htmlDiv3.Borders.Item(wdBorderRight).LineStyle = wdLineStyleDot
    htmlDiv3.Borders.Item(wdBorderTop).LineStyle = wdLineStyleDot
    htmlDiv3.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleDot

}
```

### Range.Hyperlinks

返回一个 **Hyperlinks** 集合，该集合代表指定范围内的所有超链接。只读。

**语法**

**express.Hyperlinks**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例显示活动文档中前十个段落中的每个超链接的名称。*/
function test()
{
let objRange = ActiveDocument.Range(
    ActiveDocument.Paragraphs.Item(1).Range.Start,
    ActiveDocument.Paragraphs.Item(10).Range.End)

let hplink = objRange.Hyperlinks   
for(let objLink = 1; objLink <= hplink.Count; objLink++) {
    if((hplink.Item(objLink).Address.toLowerCase().search("microsoft")) != -1) {
        MsgBox(hplink.Item(objLink).Name)
    }
}
}
```

### Range.ID

返回或设置特定范围的标识名称。可读写 **String** 类型。

**语法**

**express.ID**

\*express是一个代表 Range 对象的变量。

### Range.Information

返回有关指定范围的信息。只读 **Variant** 类型。

**语法**

**express.Information**

\*express是一个代表 Range 对象的变量。

**说明**

| **名称** | **必选/可选** | **数据类型**      | **说明**   |
|----------|---------------|-------------------|------------|
| *Type*   | 必选          | **WdInformation** | 消息类型。 |

**示例**

``` JavaScript
/*如果第十个单词位于某个表格中，则以下示例选定该表格。*/
function test() {
  if(Application.ActiveDocument.Words.Item(10).Information(wdWithInTable)) {
      Application.ActiveDocument.Words.Item(10).Tables.Item(1).Select()
  }
}
```

### Range.InlineShapes

返回一个 **InlineShape** 集合，该集合代表范围中的所有 **InlineShapes** 对象。只读。

**语法**

**express.InlineShapes**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例显示活动文档中的形状和内嵌形状的数目。*/
function test() {
  let doc = Application.ActiveDocument
  alert("InlineShape = " + doc.InlineShapes.Count + "\r" + "Shapes = " + doc.Shapes.Count)
}
```

### Range.IsEndOfRowMark

如果指定范围被折叠且位于表格中的行尾标志处，则该属性值为 **True** 。只读 **Boolean** 类型。

**语法**

**express.IsEndOfRowMark**

\*express是一个代表 Range 对象的变量。

**说明**

该属性与下面的表达式等效：

``` JavaScript
ActiveDocument.Range().Information(wdAtEndOfRowMarker)
```

**示例**

``` JavaScript
/*以下示例实现的功能是：折叠所选内容，如果插入点位于行尾（刚好在行结束标记之前），则选定当前行。*/
function test()
{
ActiveDocument.Range().Collapse(wdCollapseEnd)
if(ActiveDocument.Range().IsEndOfRowMark == true) {
    ActiveDocument.Range().Rows.Item(1).Select()
}
}
```

### Range.Italic

如果将字体或范围设置为倾斜格式，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.Italic**

\*express是一个代表 Range 对象的变量。

**说明**

此属性返回 **True** 、 **False** 或 **wdUndefined** （ **True** 和 **False** 混合组成），并且可设置为 **True** 、 **False** 或 **wdToggle** 。

**示例**

``` JavaScript
/*以下示例将活动文档中第一个单词的格式设置为倾斜。*/
ActiveDocument.Words.Item(1).Italic = true
```

### Range.ItalicBi

如果将字体或范围设置为倾斜格式，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.ItalicBi**

\*express是一个代表 Range 对象的变量。

**说明**

该属性返回 **True** 、 **False** 或 **wdUndefined** （对于倾斜和非倾斜的混合文本）。可设置为 **True** 、 **False** 或 **wdToggle** 。

| 注释                                        |
|---------------------------------------------|
| **ItalicBi** 属性应用于从右向左语言的文本。 |

**示例**

``` JavaScript
/*以下示例将从右向左语言的活动文档中的第一段设置为倾斜格式。*/
ActiveDocument.Paragraphs.Item(1).Range.ItalicBi = true
```

### Range.Kana

返回或设置日文文本的指定区域是平假名还是片假名。 **WdKana** 类型，可读写。

**语法**

**express.Kana**

\*express是一个代表 Range 对象的变量。

**说明**

如果指定区域含有平假名和片假名的混合文本，或包含其他非日文文本，本属性返回 **wdUndefined** 。如果将 **Kana** 属性设为 **wdUndefined** ，则会产生一个错误。

**示例**

``` JavaScript
/*本示例显示当前选定内容中所含的日文文本的类型。*/
function test()
{
myKana = Selection.Range.Kana
switch(myKana) {
    case wdKanaHiragana:
        MsgBox("This text is hiragana.")
        break
    case wdKanaKatakana:
        MsgBox("This text is katakana.")
        break
    case wdUndefined:
        MsgBox("This text is a mix of " 
            + "hiragana and katakana.")
        break
}
}
```

### Range.LanguageDetected

返回或设置一个值，该值指定 WPS 是否已经检测过指定文本的语言。可读/写 **Boolean** 类型。

**语法**

**express.LanguageDetected**

\*express是一个代表 Range 对象的变量。

**说明**

检查以前所有语言检测结果的 **LanguageID** 属性。

调用 **DetectLanguage** 方法时， **LanguageDetected** 属性被设置为 **True** 。要重新检测指定文本的语言，必须先将 **LanguageDetected** 属性设置为 **False** 。

**示例**

``` JavaScript
/*本示例检查活动文档以确定用于编写它的语言，然后显示结果。*/
function test()
{
let rng = ActiveDocument.Range
    if(rng.LanguageDetected == true) {
        let x = MsgBox("This document has already " 
            + "been checked. Do you want to check " 
            + "it again?",jsYesNo)
        if(x == jsResultYes) {
            rng.LanguageDetected = false
            rng.DetectLanguage()
        }
    }
    else {
        rng.DetectLanguage()
    }
    if(rng.Range.LanguageID == wdEnglishUS) {
        MsgBox("This is a U.S. English document.")
    }
    else {
        MsgBox("This is not a U.S. English document.")
    }
}
```

### Range.LanguageID

返回或设置一个 **WdLanguageID** 常量，该常量代表指定范围的语言。可读写。

**语法**

**express.LanguageID**

\*express是一个代表 Range 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **WdLanguageID** 常量可能不可用。

**示例**

``` JavaScript
/*以下示例将活动文档的第二段格式设置为法语格式，然后添加新的自定义词典，此词典用于查阅法语文本。*/
function test()
{
ActiveDocument.Paragraphs.Item(2).Range.LanguageID = wdFrench
let myDictionary = CustomDictionaries.Add("French.dic")
    myDictionary.LanguageSpecific = true
    myDictionary.LanguageID = wdFrench
}
```

### Range.LanguageIDFarEast

返回或设置指定对象的东亚语言。可读写 **WdLanguageID** 类型。

**语法**

**express.LanguageIDFarEast**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中第一段的语言设置为朝鲜语。*/
ActiveDocument.Paragraphs.Item(1).Range.LanguageIDFarEast = wdKorean
```

### Range.LanguageIDOther

返回或设置指定范围的语言。可读写 **WdLanguageID** 类型。

**语法**

**express.LanguageIDOther**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例将所选内容的语言设置为法语。*/
Selection.Range.LanguageIDOther = wdFrench
```

### Range.ListFormat

返回一个 **ListFormat** 对象，该对象代表某区域中所有的列表格式特征。只读。

**语法**

**express.ListFormat**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中 3 到 6 段中的一个区域赋给变量 myDoc。然后，根据此区域中是否存在列表格式，来设置默认的多级符号列表格式或清除已有编号格式。*/
function test()
{
let myDoc = ActiveDocument
let myRange = 
    myDoc.Range(myDoc.Paragraphs.Item(3).Range.Start, 
    myDoc.Paragraphs.Item(6).Range.End)
myRange.ListFormat.ApplyOutlineNumberDefault()
}
```

``` JavaScript
/*本示例对选定内容中所有段落应用“项目符号和编号”对话框“编号”选项卡中的第二个列表模板。*/
function test()
{
Selection.Range.ListFormat.ApplyListTemplate(
    ListGalleries.Item(wdNumberGallery).ListTemplates.Item(2))
}
```

### Range.ListParagraphs

返回一个 **ListParagraphs** 集合，该集合代表范围中的所有编号段落。只读。

**语法**

**express.ListParagraphs**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Range.ListStyle

返回一个 **Variant** 类型的值，该值代表用于设置项目符号列表或编号列表的样式。只读。

**语法**

**express.ListStyle**

\*express是一个代表 Range 对象的变量。

### Range.Locks

返回 CoAuthLocks 集合对象，该对象代表范围中的所有锁。只读。

**语法**

**express.Locks**

\*express是一个代表 Range 对象的变量。

**说明**

使用 **Locks** 属性可返回 CoAuthLocks 集合。

| 注释                                                                                             |
|--------------------------------------------------------------------------------------------------|
| 该属性仅可用于支持共同创作的文档。如果尝试在不支持共同创作的文档中访问该属性，将导致运行时错误。 |

**示例**

``` JavaScript
/*下面的代码示例显示活动文档的第一个段落中的锁数量。*/
MsgBox(ActiveDocument.Paragraphs.Item(1).Range.Locks.Count)
```

### Range.NextStoryRange

返回一个 **Range** 对象，该对象代表下一个 文章 。 **Range** 类型，只读。

**语法**

**express.NextStoryRange**

\*express是一个代表 Range 对象的变量。

**说明**

下表列出了依据文章的类型而返回的区域。

| 文章的类型                                                                                                                                                                   | 用 NextStoryRange 方法返回的项目 |
|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|----------------------------------|
| **wdMainTextStory** 、 **wdFootnotesStory** 、 **wdEndnotesStory** 和 **wdCommentsStory**                                                                                    | 总是返回 **Nothing**             |
| **wdTextFrameStory**                                                                                                                                                         | 下一组链接文本框的文章           |
| **wdEvenPagesHeaderStory** 、 **wdPrimaryHeaderStory** 、 **wdEvenPagesFooterStory** 、 **wdPrimaryFooterStory** 、 **wdFirstPageHeaderStory** 、 **wdFirstPageFooterStory** | 下一节中相同类型的文章           |

### Range.NoProofing

如果拼写和语法检查程序忽略指定文本，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.NoProofing**

\*express是一个代表 Range 对象的变量。

**说明**

如果只将某些指定文本的 **NoProofing** 属性设置为 **True** ，则该属性返回 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例实现的功能是：标记当前所选内容，以便拼写和语法检查程序将其忽略。*/
Selection.Range.NoProofing = true
```

### Range.OMaths

返回一个 **OMaths** 集合，该集合代表指定区域内的 **OMath** 对象。只读。

**语法**

**express.OMaths**

\*express是一个代表 Range 对象的变量。

### Range.Orientation

在启用了“文字方向”功能时返回或设置范围中文字的方向。可读写 **WdTextOrientation** 类型。

**语法**

**express.Orientation**

\*express是一个代表 Range 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **WdTextOrientation** 常量可能不可用。

您可以设置文本框架或者恰好位于文本框架内的范围的方向。有关文本框架和文本框之间的区别的信息，请参阅 **TextFrame** 对象。

### Range.PageSetup

返回一个 **PageSetup** 对象，该对象与指定范围相关联。

**语法**

**express.PageSetup**

\*express是一个代表 Range 对象的变量。

### Range.ParagraphFormat

返回或设置一个 **ParagraphFormat** 对象，该对象代表指定范围的段落设置。可读/写。

**语法**

**express.ParagraphFormat**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例对包含 MyDoc.doc 所有内容的有关范围设置段落格式：2 倍行距，并且在 0.25 英寸的位置设置一个自定义制表位。*/
function test()
{
let myRange = Documents.Item("MyDoc.doc").Content
let myPFormat = myRange.ParagraphFormat
    myPFormat.Space2()
    myPFormat.TabStops.Add(InchesToPoints(.25))
}
```

### Range.Paragraphs

返回一个 **Paragraphs** 集合，该集合代表指定范围中的所有段落。只读。

**语法**

**express.Paragraphs**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例将活动文档中第一节中所有段落的集合的行距设置为单倍行距。*/
function test()
{
ActiveDocument.Sections.Item(1).Range.Paragraphs.LineSpacingRule = 
    wdLineSpaceSingle
}
```

### Range.ParagraphStyle

返回一个 **Variant** 类型的值，该值代表用于设置段落格式的样式。只读。

**语法**

**express.ParagraphStyle**

\*express是一个代表 Range 对象的变量。

### Range.Parent

返回一个 **Object** 类型值，该值代表指定 **Range** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Range 对象的变量。

### Range.ParentContentControl

返回一个 **ContentControl** 对象，该对象代表指定区域的父内容控件。只读。

**语法**

**express.ParentContentControl**

\*express是一个代表 Range 对象的变量。

**说明**

如果指定区域没有父内容控件，则此属性返回 **Nothing** 。

### Range.PreviousBookmarkID

返回最后一个书签的编号，该书签从指定范围的前面或与指定范围相同的位置开始。只读 **Long** 类型。

**语法**

**express.PreviousBookmarkID**

\*express是一个代表 Range 对象的变量。

**说明**

如果没有相应的书签，则该属性返回 0（零）。

**示例**

``` JavaScript
/*以下示例显示第二段前面的书签的名称。*/
function test()
{
let num = ActiveDocument.Paragraphs.Item(2).Range.PreviousBookmarkID
if(num != 0) {
    MsgBox(ActiveDocument.Content.Bookmarks.Item(num).Name)
}
}
```

### Range.ReadabilityStatistics

返回一个 **ReadabilityStatistics** 集合，该集合代表指定文档或范围的可读性统计信息。只读。

**语法**

**express.ReadabilityStatistics**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

**示例**

``` JavaScript
/*以下示例显示文档 1 的每一种可读性统计信息及其值。*/
function test()
{
let RStatistics = Documents.Item(1).ReadabilityStatistics
for(let rs = 1; rs <= RStatistics.Count; rs++) {
    MsgBox(RStatistics.Item(rs).Name + " - " + rs.Value)
}
}
```

### Range.Revisions

返回一个 **Revisions** 集合，该集合代表范围中的修订。只读。

**语法**

**express.Revisions**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

**示例**

``` JavaScript
/*以下示例显示活动文档中第一节中修订的数目。*/
MsgBox(ActiveDocument.Sections.Item(1).Range.Revisions.Count)
```

``` JavaScript
/*以下示例接受所选内容的第一段中的所有修订。*/
function test()
{
let myRange = Selection.Paragraphs.Item(1).Range
myRange.Revisions.AcceptAll()
}
```

### Range.Rows

返回一个 **Rows** 集合，该集合代表范围中的所有表格行。只读。

**语法**

**express.Rows**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

### Range.Scripts

返回一个 **Scripts** 集合，该集合代表指定对象中 HTML 脚本的集合。

**语法**

**express.Scripts**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例测试指定范围中的第二个 Script 对象以确定其语言。*/
function test()
{
let lage = Selection.Range.Scripts.Item(2).Language
switch(lage) {
    case msoScriptLanguageASP:
        MsgBox("Active Server Pages")
        break
    case msoScriptLanguageVisualBasic:
        MsgBox("VBScript")
        break
    case msoScriptLanguageJava:
        MsgBox("JavaScript")
        break
    case msoScriptLanguageOther:
        MsgBox("Unknown type of script")
}
}
```

### Range.Sections

返回一个 **Sections** 集合，该集合代表指定范围中的各节。只读。

**语法**

**express.Sections**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

### Range.Sentences

返回一个 **Sentences** 集合，该集合代表范围中的所有句子。只读。

**语法**

**express.Sentences**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

**示例**

``` JavaScript
/*以下示例显示活动文档中第一段的句子数。*/
function test()
{
MsgBox(ActiveDocument.Paragraphs.Item(1).Range 
    .Sentences.Count + " sentences")
}
```

### Range.Shading

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例对所选内容的第一段应用黄色底纹。*/
function test()
{
let myShading = Selection.Paragraphs.Item(1).Shading
    myShading.Texture = wdTexture12Pt5Percent
    myShading.BackgroundPatternColorIndex = wdYellow
    myShading.ForegroundPatternColorIndex = wdBlack
}
```

``` JavaScript
/*以下示例将活动文档中第一个单词的底纹设置为 10%。*/
ActiveDocument.Words.Item(1).Shading.Texture = wdTexture10Percent
```

### Range.ShapeRange

返回一个 **ShapeRange** 集合，该集合代表指定范围中的所有 **Shape** 对象。只读。

**语法**

**express.ShapeRange**

\*express是一个代表 Range 对象的变量。

**说明**

图形范围可以包含绘图、图形、图片、OLE 对象、ActiveX 控件、文本对象和标注。

**示例**

``` JavaScript
/*以下示例将活动文档中所有图形的填充前景色设置为紫色。*/
Application.ActiveDocument.Content.ShapeRange.Fill.ForeColor.RGB = (255, 0, 255)
```

### Range.ShowAll

如果显示所有非打印字符（如隐藏文字、制表符、空格和段落标记），则该属性值为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.ShowAll**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动文档中显示所有非打印字符。*/
ActiveDocument.Range.ShowAll = true
```

### Range.SpellingChecked

如果已对指定的区域或文档完成拼写检查，则该属性值为 **True** 。如果所有或部分区域或文档尚未进行拼写检查，则该属性值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.SpellingChecked**

\*express是一个代表 Range 对象的变量。

**说明**

若要重新检查某一区域或文档的拼写，请将 **SpellingChecked** 属性设置为 **False** 。

若要查看该区域或文档中是否含有拼写错误，请使用 **SpellingErrors** 属性。

**示例**

``` JavaScript
/*以下示例判定是否已经检查过活动文档第一节中的拼写。如果没有，则开始拼写检查。*/
function test()
{
let myRange = ActiveDocument.Sections.Item(1).Range
let isChecked = myRange.SpellingChecked
if(isChecked == false) {
    myRange.CheckSpelling()
}
else {
    MsgBox("Spelling has already been checked in the range.")
}
}
```

### Range.SpellingErrors

返回一个 **ProofreadingErrors** 集合，该集合代表指定范围中标识为拼写错误的单词。只读。

**语法**

**express.SpellingErrors**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

**示例**

``` JavaScript
/*以下示例检查指定范围是否存在拼写错误，并显示找到的每个错误。*/
function test()
{
let myErrors = ActiveDocument.Paragraphs.Item(3).Range.SpellingErrors
if(myErrors.Count == 0) {
    MsgBox("No spelling errors found.")
}
else {
    for(let myErr = 1;myErr <= myErrors.Count;myErr++) {
        Msgbox(myErrors.Item(myErr).Text)
    }
}
}
```

### Range.Start

返回或设置某区域中起始字符的位置。 **Long** 类型，可读写。

**语法**

**express.Start**

\*express是一个代表 Range 对象的变量。

**说明**

**Range** 对象包括起始字符和结束字符位置。起始字符位置是指距文档开头部分最近的字符位置。如果将该属性的值设置为大于 **End** 属性的值，则 **End** 属性值会设置为与 **Start** 属性值相同。

该属性返回起始字符相对于文档开头部分的位置。文本主体部分 ( **wdMainTextStory** ) 的起始字符位置为 0（零）。通过设置该属性可以更改选定内容、区域或书签的大小。

**示例**

``` JavaScript
/*本示例返回活动文档第二段的起始位置和第四段的结束位置。这些字符位置用于创建区域 myRange。*/
function test() {
  let pos = Application.ActiveDocument.Paragraphs.Item(2).Range.Start
  let pos2 = Application.ActiveDocument.Paragraphs.Item(4).Range.End
  let myRange = Application.ActiveDocument.Range(pos,pos2)
}

/*本示例将 myRange 起始字符的位置向右移动一个字符（这会使该区域缩小一个字符）。*/
function test() {
  let myRange = Application.Selection.Range
  myRange.SetRange(myRange.Start + 1, myRange.End)
}
```

### Range.StoryLength

返回包含指定区域的 文字部分 中的字符数。 **Long** 类型，只读。

**语法**

**express.StoryLength**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例确定活动文档的页眉部分是否为空。如果页眉部分不为空，则在消息框中显示页眉的内容。如果为空，则 StoryLength 属性返回 1 作为最后的段落标记。*/
function test()
{
let myRange = ActiveDocument.Sections.Item(1) 
    .Headers.Item(wdHeaderFooterPrimary).Range
if(myRange.StoryLength > 1) {
    MsgBox(myRange.Text)
}
}
```

``` JavaScript
/*如果文档为空，则以下示例在关闭文档时不保存所做的更改。*/
function test()
{
if(ActiveDocument.Content.StoryLength == 1) {
    ActiveDocument.Close(wdDoNotSaveChanges)
}
}
```

### Range.StoryType

返回指定范围、所选内容或书签的 文字部分 类型。只读 **WdStoryType** 类型。

**语法**

**express.StoryType**

\*express是一个代表 Range 对象的变量。

### Range.Style

返回或设置指定对象的样式。可读写 **Variant** 类型。

**语法**

**express.Style**

\*express是一个代表 Range 对象的变量。

**说明**

要设置该属性，请指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个代表样式的对象。如果返回包含多个样式的范围的样式，则只返回第一个字符样式或段落样式。

**示例**

| 注释                                                   |
|--------------------------------------------------------|
| **Characters** 集合的每个元素都是一个 **Range** 对象。 |

``` JavaScript
function test()
{
for(let c = 1; c <= Selection.Characters.Count; c++) {
    MsgBox(Selection.Characters.Item(c).Style)
}
}
```

### Range.Subdocuments

返回一个 **Subdocuments** 集合，该集合代表指定范围或文档中的所有子文档。只读。

**语法**

**express.Subdocuments**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例显示嵌入在活动文档中的子文档的数量。*/
MsgBox(ActiveDocument.Range.Subdocuments.Count)
```

### Range.SynonymInfo

返回一个 **SynonymInfo** 对象，该对象包含同义词库中有关某范围的内容的同义词、反义词或相关单词和表达方式的信息。

**语法**

**express.SynonymInfo**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例返回所选内容的第一个含义的同义词列表。*/
function test()
{
let Slist = Selection.Range.SynonymInfo.SynonymList(1)
for(let i = 0;i <= Slist.length - 1;i++) {
    MsgBox(Slist[i])
}
}
```

### Range.Tables

返回一个 **Tables** 集合，该集合代表指定范围内的所有表格。只读。

**语法**

**express.Tables**

\*express是一个代表 Range 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例在活动文档中创建一个 5x5 表格，然后对其应用预定义格式。*/
function test()
{
Selection.Collapse(wdCollapseStart)
let myTable = ActiveDocument.Tables.Add(Selection.Range,5,5)
myTable.AutoFormat(wdTableFormatClassic2)
}
```

``` JavaScript
/*以下示例在活动文档中第一个表格的首列中插入数值和文本。*/
function test()
{
let num = 90
let myCells = ActiveDocument.Tables.Item(1).Columns.Item(1).Cells
for(let acell = 1; acell <= myCells.Count; acell++) {
    myCells.Item(acell).Range.Text = num + " Sales"
    num = num + 1
}
}
```

### Range.TableStyle

返回一个 **Variant** 类型的值，该值代表用于设置表格格式的样式。只读。

**语法**

**express.TableStyle**

\*express是一个代表 Range 对象的变量。

### Range.Text

返回或设置指定区域或选定内容中的文本。 **String** 类型，可读写。

**语法**

**express.Text**

\*express是一个代表 Range 对象的变量。

**说明**

**Text** 属性返回该区域的无格式纯文本。如果设置该属性，则将替换该区域中的现有文本。

**示例**

``` JavaScript
/*本示例用“Dear”替换活动文档的第一个词。*/
function test() {
  let myRange = Application.ActiveDocument.Words.Item(1)
  myRange.Text = "Dear "
}
```

### Range.TextRetrievalMode

返回一个 **TextRetrievalMode** 对象，该对象控制从指定 **区域** 检索文字的方式。可读写。

**语法**

**express.TextRetrievalMode**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例检索选定文字（排除隐藏文字），并将其插入活动文档第三段的开始处。*/
function test()
{
if(Selection.Type == wdSelectionNormal) {
    let Range1 = Selection.Range
    Range1.TextRetrievalMode.IncludeHiddenText = false
    let Range2 = ActiveDocument.Paragraphs.Item(2).Range
    Range2.InsertAfter(Range1.Text)
}
}
```

``` JavaScript
/*本示例在大纲视图中检索并显示前三个段落。*/
function test()
{
let myRange = ActiveDocument.Range(ActiveDocument
    .Paragraphs.Item(1).Range.Start, 
    ActiveDocument.Paragraphs.Item(3).Range.End)
myRange.TextRetrievalMode.ViewType = wdOutlineView
MsgBox(myRange.Text)
}
```

``` JavaScript
/*本示例先排除引用选定文字的区域中的域代码和隐藏文字，然后在一个消息框中显示文字。*/
function test()
{
if(Selection.Type == wdSelectionNormal) {
    let aRange = Selection.Range
    aRange.TextRetrievalMode.IncludeHiddenText = false
    aRange.TextRetrievalMode.IncludeFieldCodes = false
    MsgBox(aRange.Text)
}
}
```

### Range.TopLevelTables

返回一个 **Tables** 集合，该集合代表当前范围最外部嵌套层上的表格。只读。

**语法**

**express.TopLevelTables**

\*express是一个代表 Range 对象的变量。

**说明**

此方法返回一个集合，该集合仅包含当前范围的上下文中最外部嵌套层上的表格。这些表格可能不在整套嵌套表格的最外嵌套层中。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

以下示例新建一个文档，创建一个三层嵌套表格，并在每张表格的第一个单元格中填入该表格所在的嵌套层数。接着选定第二层表格的第二列，然后选定所选内容中的顶层表格的第一列。尽管最里面的表格在整套嵌套表格的上下文关系中并非顶层表格，但仍会被选定。

``` JavaScript
function test()
{
Documents.Add()
ActiveDocument.Tables.Add(Selection.Range,
    3, 3, wdWord9TableBehavior, wdAutoFitContent)
let aRange = ActiveDocument.Tables.Item(1).Range
    aRange.Copy()
    aRange.Cells.Item(1).Range.Text = aRange.Cells.Item(1).NestingLevel
    aRange.Cells.Item(5).Range.PasteAsNestedTable()
    let acRange = aRange.Cells.Item(5).Tables.Item(1).Range
        acRange.Cells.Item(1).Range.Text = acRange.Cells.Item(1).NestingLevel
        acRange.Cells.Item(5).Range.PasteAsNestedTable()
        let accRange = acRange.Cells.Item(5).Tables.Item(1).Range
            accRange.Cells.Item(1).Range.Text = 
                accRange.Cells.Item(1).NestingLevel       
        acRange.Columns.Item(2).Select()
        Selection.Range.TopLevelTables(1).Select()
}
```

### Range.TwoLinesInOne

返回或设置 WPS 是否将两行文本合并为一行，并指定括住文本的字符（如果有）。 **WdTwoLinesInOneType** 类型，可读写。

**语法**

**express.TwoLinesInOne**

\*express是一个代表 Range 对象的变量。

**说明**

将 **TwoLinesInOne** 属性设为 **wdTwoLinesInOneNoBrackets** 可将两行文本合并为一行，且不将文本括在任何字符中。将 **TwoLinesInOne** 属性设为 **wdTwoLinesInOneNone** 可将合并为一行的文本还原为两行。

**示例**

``` JavaScript
/*本示例设置当前选定内容的格式，使两行文本合并为一行，并将其括在括号中。*/
Selection.Range.TwoLinesInOne = wdTwoLinesInOneParentheses
```

### Range.Underline

返回或设置应用于范围的下划线的类型。可读写 **WdUnderline** 类型。

**语法**

**express.Underline**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例对活动文档中的第四个单词应用双下划线。*/
Application.ActiveDocument.Words.Item(4).Underline = wdUnderlineDouble
```

### Range.Updates

返回 CoAuthUpdates 集合对象，该对象代表范围中的所有可用更新。只读。

**语法**

**express.Updates**

\*express是一个代表 Range 对象的变量。

**说明**

使用 **Updates** 属性可返回 CoAuthUpdates 集合。

| 注释                                                                                             |
|--------------------------------------------------------------------------------------------------|
| 该属性仅可用于支持共同创作的文档。如果尝试在不支持共同创作的文档中访问该属性，将导致运行时错误。 |

**示例**

``` JavaScript
/*下面的代码示例显示活动文档的第一个段落中的可用更新数量。*/
function test()
{
let countOfUpdates = ActiveDocument.Paragraphs.Item(1).Range.Updates.Count

MsgBox("The number of updates is " + countOfUpdates)
}
```

### Range.WordOpenXML

返回一个 **String** 类型的值，该值以 WPS Open XML 格式表示区域中包含的 XML。只读。

**语法**

**express.WordOpenXML**

\*express是一个代表 Range 对象的变量。

**说明**

此属性只返回文档中为表示指定区域所需的 XML。

### Range.Words

返回一个 **Words** 集合，该集合代表范围中的所有单词。只读。

**语法**

**express.Words**

\*express是一个代表 Range 对象的变量。

**说明**

文档中的标点符号和段落标记包括在 **Words** 集合中。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例显示所选内容的字数。段落标记、部分单词和标点符号都统计在内。*/
MsgBox("There are " + Selection.Words.Count + " words.")
```

``` JavaScript
/*以下示例遍历 myRange（范围为从活动文档的开始到所选内容的末尾）中的单词，并删除该范围中出现的单词“Franklin”（包含尾部空格）。*/
function test()
{
let myRange = ActiveDocument.Range(0,Selection.End)
for(let aWord = 1; aWord <= myRange.Words.Count; aWord++) {
    if(myRange.Words.Item(aWord).Text == "Franklin ") {
        myRange.Words.Item(aWord).Delete()
    }
}
}
```

### Range.XMLNodes

返回一个 **XMLNodes** 集合，该集合代表指定区域中的 XML 元素（包括任何只是部分属于该区域的元素）。只读。

**语法**

**express.XMLNodes**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例访问活动文档的开头 200 个字符中包含的 XML 元素。*/
function test()
{
let objRange = ActiveDocument.Range(1, 200)

let objNode = objRange.XMLNodes.Item(1)
}
```

### Range.XMLParentNode

返回一个 **XMLNode** 对象，该对象代表区域的父级 XML 节点。只读。

**语法**

**express.XMLParentNode**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*以下示例返回指定区域的父节点。*/
function test()
{
let objRange = ActiveDocument.Range(1, 200)
let objNode = objRange.XMLParentNode
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Range](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Table 对象

## 简介

代表单个表格。 Table 对象是 Tables 集合的一个成员。 Tables 集合包含指定的所选内容、范围或文档中的所有表格。

## 说明

使用 Tables ( *Index* ) 可返回一个 Table 对象，其中 *Index* 为索引号。索引号代表表格在所选内容、范围或文档中的位置。以下示例将活动文档中的第一个表格转换为文本。

``` JavaScript
Application.ActiveDocument.Tables.Item(1).ConvertToText(wdSeparateByTabs)
```

使用 Add 方法可在指定范围添加表格。以下示例在活动文档的开头添加一个 3x4 的表格。

``` JavaScript
let myRange = Application.ActiveDocument.Range(0, 0)
Application.ActiveDocument.Tables.Add(myRange, 3, 4)
```

## 方法

| 序号 | 名称                                                            | 说明                                                                                        |
|:----:|:----------------------------------------------------------------|:--------------------------------------------------------------------------------------------|
|  1   | [ApplyStyleDirectFormatting](#Table.ApplyStyleDirectFormatting) | 应用指定的样式，但保留用户直接应用的所有格式。                                              |
|  2   | [AutoFitBehavior](#Table.AutoFitBehavior)                       | 确定 WPS 如何使用自动调整功能来调整表格的大小。                                             |
|  3   | [AutoFormat](#Table.AutoFormat)                                 | 将预定义外观应用于表格。                                                                    |
|  4   | [Cell](#Table.Cell)                                             | 返回一个 Cell 对象，该对象代表表格中的一个单元格。                                          |
|  5   | [Converttotext](#Table.Converttotext)                           | 将表格转换为文本并返回一个 Range 对象，该对象代表带分隔符的文本。                           |
|  6   | [Delete](#Table.Delete)                                         | 删除指定的表格。                                                                            |
|  7   | [Select](#Table.Select)                                         | 选择指定的表格。                                                                            |
|  8   | [Sort](#Table.Sort)                                             | 对指定的表格进行排序。                                                                      |
|  9   | [SortAscending](#Table.SortAscending)                           | 按字母数字升序对段落或表格行进行排序。                                                      |
|  10  | [SortDescending](#Table.SortDescending)                         | 以字母数字降序排列表格行。                                                                  |
|  11  | [Split](#Table.Split)                                           | 在表格中紧靠指定行的上面插入一空段落，并且返回一个 Table 对象，此对象包含指定行及其下一行。 |
|  12  | [UpdateAutoFormat](#Table.UpdateAutoFormat)                     | 使用预定义表格式对表格式进行更新。                                                          |

## 属性

| 序号 | 名称                                                  | 说明                                                                                                         |
|:----:|:------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------|
|  1   | [Allowautofit](#Table.Allowautofit)                   | 使 WPS 可以自动调整表格中的单元格的大小以适应内容。 Boolean 类型，可读写。                                   |
|  2   | [AllowPageBreaks](#Table.AllowPageBreaks)             | 允许 WPS 对指定表格进行跨页断行。 Boolean 类型，可读写。                                                     |
|  3   | [Application](#Table.Application)                     | 返回一个代表 WPS 应用程序的 Application 对象。                                                               |
|  4   | [ApplyStyleColumnBands](#Table.ApplyStyleColumnBands) | 返回或设置 Boolean 值，该值表示如果所应用的预置表样式为列提供了样式带，则是否对表中的列应用样式带。可读/写。 |
|  5   | [ApplyStyleFirstColumn](#Table.ApplyStyleFirstColumn) | 如果为 True ，则 WPS 对指定表格的第一列应用第一列的格式。 Boolean 类型，可读写。                             |
|  6   | [ApplyStyleHeadingRows](#Table.ApplyStyleHeadingRows) | 如果为 True ，则 WPS 为选定表格的第一行应用标题行的格式。 Boolean 类型，可读写。                             |
|  7   | [ApplyStyleLastColumn](#Table.ApplyStyleLastColumn)   | 如果为 True ，则 WPS 对指定表格的最后一列应用最后一列的格式。 Boolean 类型，可读写。                         |
|  8   | [ApplyStyleLastRow](#Table.ApplyStyleLastRow)         | 如果为 True ，则 WPS 对指定表格的最后一行应用最后一行的格式。 Boolean 类型，可读写。                         |
|  9   | [ApplyStyleRowBands](#Table.ApplyStyleRowBands)       | 返回或设置 Boolean 值，该值表示如果所应用的预置表样式为行提供了样式带，则是否对表中的行应用样式带。可读/写。 |
|  10  | [AutoFormatType](#Table.AutoFormatType)               | 返回指定表格所应用的自动套用格式类型。 Long 类型，只读。                                                     |
|  11  | [Borders](#Table.Borders)                             | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                                        |
|  12  | [BottomPadding](#Table.BottomPadding)                 | 返回或设置在表格中单独的单元格或所有单元格内容的下方添加的间距（以磅为单位）。 Single 类型，可读写。         |
|  13  | [Columns](#Table.Columns)                             | 返回一个 Columns 集合，该集合代表表格中的所有列。只读。                                                      |
|  14  | [Creator](#Table.Creator)                             | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                 |
|  15  | [Descr](#Table.Descr)                                 | 返回或设置包含指定表格的说明的 String 类型值。可读/写。                                                      |
|  16  | [ID](#Table.ID)                                       | 在将文档另存为网页时，返回或设置指定的表的标识标签。可读/写 String 类型。                                    |
|  17  | [LeftPadding](#Table.LeftPadding)                     | 返回或设置在表格中所有单元格的内容左侧要增加的间距（以磅为单位）。可读写 Single 类型。                       |
|  18  | [Nestinglevel](#Table.Nestinglevel)                   | 返回指定表格的嵌套层。只读 Long 类型。                                                                       |
|  19  | [Parent](#Table.Parent)                               | 返回一个 Object 类型值，该值代表指定 Table 对象的父对象。                                                    |
|  20  | [Preferredwidth](#Table.Preferredwidth)               | 返回或设置指定表格的指定宽度（以磅为单位或以窗口宽度的百分比表示）。可读写 Single 类型。                     |
|  21  | [Preferredwidthtype](#Table.Preferredwidthtype)       | 返回指定表格所应用的自动套用格式类型。 Long 类型，只读。                                                     |
|  22  | [Range](#Table.Range)                                 | 返回一个 Range 对象，该对象代表指定表格中所含的文档部分。                                                    |
|  23  | [RightPadding](#Table.RightPadding)                   | 返回或设置在表格的所有单元格的内容右侧要增加的间距（以磅为单位）。可读写 Single 类型。                       |
|  24  | [Rows](#Table.Rows)                                   | 返回一个 Rows 集合，该集合代表表格中所有的表格行。只读。                                                     |
|  25  | [Shading](#Table.Shading)                             | 返回一个 Shading 对象，该对象代表指定对象的底纹格式。                                                        |
|  26  | [Spacing](#Table.Spacing)                             | 返回或设置表格中单元格的间距（以磅为单位）。可读写 Single 类型。                                             |
|  27  | [Style](#Table.Style)                                 | 返回或设置指定表格的样式。可读写 Variant 类型。                                                              |
|  28  | [TableDirection](#Table.TableDirection)               | 返回或设置 WPS 对指定表格中的单元格进行排序的方向。可读写 WdTableDirection 类型。                            |
|  29  | [Tables](#Table.Tables)                               | 返回一个 Tables 集合，该集合代表指定表格中的所有嵌套表格。只读。                                             |
|  30  | [Title](#Table.Title)                                 | 返回或设置包含指定表格的标题的 String 类型值。可读/写。                                                      |
|  31  | [TopPadding](#Table.TopPadding)                       | 返回或设置表格中所有单元格的内容上方要增加的间距（以磅为单位）。可读写 Single 类型。                         |
|  32  | [Uniform](#Table.Uniform)                             | 如果表格中所有的行具有相同的列数，则该属性值为 True 。 Boolean 类型，只读。                                  |

## 成员方法

### Table.ApplyStyleDirectFormatting

应用指定的样式，但保留用户直接应用的所有格式。

**语法**

**express.ApplyStyleDirectFormatting ( *StyleName* )**

\*express是一个代表 Table 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                 |
|-------------|-----------|----------|----------------------|
| *StyleName* | 必选      | String   | 要应用的样式的名称。 |

### Table.AutoFitBehavior

确定 WPS 如何使用自动调整功能来调整表格的大小。

**语法**

**express.AutoFitBehavior ( *Behavior* )**

\*express是一个代表 Table 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型          | 说明                                                    |
|------------|-----------|-------------------|---------------------------------------------------------|
| *Behavior* | 必选      | WdAutoFitBehavior | 指定 WPS 如何使用“自动调整”功能重新调整指定表格的大小。 |

**说明**

WPS 可以根据表格单元格的内容或文档窗口的宽度重新调整表格的大小。也可使用本方法来关闭“自动调整”功能。这样，表格的大小是固定值，而不随单元格内容或窗口宽度而改变。

将 **AutoFitBehavior** 属性设置为 **wdAutoFitContent** 或 **wdAutoFitWindow** 会将 **AllowAutoFit** 属性设置为 **True** （如果该属性当前为 **False** ）。同样，将 **AutoFitBehavior** 属性设置为 **wdAutoFitFixed** 会将 **AllowAutoFit** 属性设置为 **False** （如果该属性当前为 **True** ）。

**示例**

``` JavaScript
/*本示例实现的功能是：设定活动文档中第一张表格的“自动调整”操作，使其可根据文档窗口的宽度自动调整大小。*/
Application.ActiveDocument.Tables.Item(1).AutoFitBehavior(wdAutoFitWindow)
```

### Table.AutoFormat

将预定义外观应用于表格。

**语法**

**express.AutoFormat ( *Format* , *ApplyBorders* , *ApplyShading* , *ApplyFont* , *ApplyColor* , *ApplyHeadingRows* , *ApplyLastRow* , *ApplyFirstColumn* , *ApplyLastColumn* , *AutoFit* )**

\*express是一个代表 Table 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                                                                                             |
|--------------------|-----------|----------|--------------------------------------------------------------------------------------------------|
| *Format*           | 可选      | Variant  | 要应用的格式。此参数可以是 WdTableFormat 常量、WdTableFormatApply 常量或 TableStyle 对象。       |
| *ApplyBorders*     | 可选      | Variant  | 如果该参数值为 True，则应用指定格式的边框属性。默认值为 True。                                   |
| *ApplyShading*     | 可选      | Variant  | 如果该参数值为 True，则应用指定格式的底纹属性。默认值为 True。                                   |
| *ApplyFont*        | 可选      | Variant  | 如果该参数值为 True，则应用指定格式的字体属性。默认值为 True。                                   |
| *ApplyColor*       | 可选      | Variant  | 如果该参数值为 True，则应用指定格式的颜色属性。默认值为 True。                                   |
| *ApplyHeadingRows* | 可选      | Variant  | Variant 如果该参数值为 True，则应用指定格式的标题行属性。默认值为 True。                         |
| *ApplyLastRow*     | 可选      | Variant  | 如果该参数值为 True，则应用指定格式的最后一行的属性。默认值为 False。                            |
| *ApplyFirstColumn* | 可选      | Variant  | 如果该参数值为 True，则应用指定格式的第一列的属性。默认值为 True。                               |
| *ApplyLastColumn*  | 可选      | Variant  | 如果该参数值为 True，则应用指定格式的最后一列的属性。默认值为 False。                            |
| *AutoFit*          | 可选      | Variant  | 如果该参数值为 True，则在不改变单元格内文字的换行方式的前提下尽可能缩小表格列宽。默认值为 True。 |

**说明**

此方法的参数对应于 **“表格自动套用格式”** 对话框中的选项。

**示例**

``` JavaScript
/*以下示例在新文档中创建一个 5x5 的表格，并将“彩色型 2”格式的所有属性应用于此表格。*/
function test()
{
let newDoc = Documents.Add()
let myTable = newDoc.Tables.Add(Selection.Range, 5, 5)
myTable.AutoFormat(wdTableFormatColorful2)
}
```

``` JavaScript
/*以下示例将“古典型 2”格式的所有属性应用于插入点所在表格。如果插入点不在表格中，则显示一个消息框。*/
function test()
{
Selection.Collapse(wdCollapseStart)

if(Selection.Information(wdWithInTable) == true){
    Selection.Tables.Item(1).AutoFormat(wdTableFormatClassic2)
}
else{
    MsgBox("The insertion point is not in a table.")
}
}
```

### Table.Cell

返回一个 **Cell** 对象，该对象代表表格中的一个单元格。

**语法**

**express.Cell ( *Row* , *Column* )**

\*express是一个代表 Table 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                      |
|----------|-----------|----------|-----------------------------------------------------------|
| *Row*    | 必选      | Long     | 要返回的表格行数。可以是介于 1 和表格行数之间的任意整数。 |
| *Column* | 必选      | Long     | 要返回的表格列数。可以是介于 1 和表格列数之间的任意整数。 |

**示例**

``` JavaScript
/*以下示例在新文档中创建一个 3x3 表格，并在表格的第一个和最后一个单元格中插入文本*/
function test() {
    let docNew = Application.Documents.Add()
    let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)
        
    tableNew.Cell(1,1).Range.InsertAfter("First cell")
    tableNew.Cell(tableNew.Rows.Count, tableNew.Columns.Count).Range.InsertAfter("Last Cell")
}

/*以下示例删除活动文档的第一个表格中的第一个单元格*/
function test() {
    if(Application.ActiveDocument.Tables.Count >= 1){
        Application.ActiveDocument.Tables.Item(1).Cell(1, 1).Delete()
    }
}
```

### Table.Converttotext

将表格转换为文本并返回一个 **Range** 对象，该对象代表带分隔符的文本。

**语法**

**express.Converttotext ( *Separator* , *NestedTables* )**

\*express是一个代表 Table 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                                                                   |
|----------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------|
| *Separator*    | 可选      | Variant  | 用以分隔被转换列的分隔符（被转换行由段落标记分隔）。可以是下列 WdTableFieldSeparator 常量之一。                        |
| *NestedTables* | 可选      | Variant  | 如果要将嵌套的表格转换为文本，则为 True 。如果 Separator 不是 wdSeparateByParagraphs ，则将忽略此参数。默认值为 True。 |

**说明**

将 **ConvertToText** 方法应用于一个 **Table** 对象时，该对象将被删除。如果要保留对已转换的表格内容的引用，就必须为 **ConvertToText** 方法返回的 **Range** 对象赋予新的对象变量。在下面示例中，将活动文档第一个表格转换为文本，并将其格式设为项目符号列表。

``` JavaScript
function test() {
    let tableTemp = Application.ActiveDocument.Tables.Item(1)
    let rngTemp = tableTemp.ConvertToText(wdSeparateByParagraphs)

    rngTemp.ListFormat.ApplyListTemplate(ListGalleries.Item(wdBulletGallery).ListTemplates.Item(1))
}
```

**示例**

``` JavaScript
/*本示例创建一张表格，然后将其转换为文本，以制表符作为分隔字符*/
function test() {
    let docNew = Application.Documents.Add()
    let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)
    
    for(let i = 1; i <= tableNew.Range.Cells.Count; i++){
        tableNew.Range.Cells.Item(i).Range.InsertAfter("Cell" + i)
    }
    MsgBox("Click OK to convert table to text.")
    let rngTemp = tableNew.ConvertToText(wdSeparateByTabs)
}

/*本示例将包含选定内容的表格转换为文本，各列之间用空格分隔*/
function test() {
    if(Application.Selection.Information(wdWithInTable) == true){
            Application.Selection.Tables.Item(1).ConvertToText(" ")
    }
    else{
        MsgBox("The insertion point is not in a table.")
    }
}
```

### Table.Delete

删除指定的表格。

**语法**

**express.Delete ()**

\*express是一个代表 Table 对象的变量。

### Table.Select

选择指定的表格。

**语法**

**express.Select ()**

\*express是一个代表 Table 对象的变量。

**说明**

使用此方法后，可使用 **Selection** 对象处理所选项目。有关详细信息，请参阅 处理 Selection 对象。

### Table.Sort

对指定的表格进行排序。

**语法**

**express.Sort ( *ExcludeHeader* , *FieldNumber* , *SortFieldType* , *SortOrder* , *FieldNumber2* , *SortFieldType2* , *SortOrder2* , *FieldNumber3* , *SortFieldType3* , *SortOrder3* , *CaseSensitive* , *BidiSort* , *IgnoreThe* , *IgnoreKashida* , *IgnoreDiacritics* , *IgnoreHe* , *LanguageID* )**

\*express是一个代表 Table 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明                                                                                                                                                                |
|--------------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ExcludeHeader*    | 可选      | Variant  | 如果该参数值为 True ，则不对首行进行排序。默认值为 False 。                                                                                                         |
| *FieldNumber*      | 可选      | Variant  | 用于排序的第一个域。WPS 依次按 FieldNumber 、 FieldNumber2、 FieldNumber3 排序。                                                                                    |
| *SortFieldType*    | 可选      | Variant  | FieldNumber 的排序类型。可以是 WdSortFieldType 常量之一。根据您选择或安装的语言支持（例如，美国英语），以上某些常量可能不可用。默认值为 wdSortFieldAlphanumeric 。  |
| *SortOrder*        | 可选      | Variant  | 对 FieldNumber 进行排序时使用的排序顺序。可以是 WdSortOrder 常量。                                                                                                  |
| *FieldNumber2*     | 可选      | Variant  | 用于排序的第二个域。                                                                                                                                                |
| *SortFieldType2*   | 可选      | Variant  | FieldNumber2 的排序类型。可以是 WdSortFieldType 常量之一。根据您选择或安装的语言支持（例如，美国英语），以上某些常量可能不可用。默认值为 wdSortFieldAlphanumeric 。 |
| *SortOrder2*       | 可选      | Variant  | 对 FieldNumber2 进行排序时使用的排序顺序。可以是一个 WdSortOrder 常量。                                                                                             |
| *FieldNumber3*     | 可选      | Variant  | 用于排序的第三个域。                                                                                                                                                |
| *SortFieldType3*   | 可选      | Variant  | FieldNumber3 的排序类型。可以是 WdSortFieldType 常量之一。根据您选择或安装的语言支持（例如，美国英语），以上某些常量可能不可用。默认值为 wdSortFieldAlphanumeric 。 |
| *SortOrder3*       | 可选      | Variant  | 对 FieldNumber3 进行排序时使用的排序顺序。可以是一个 WdSortOrder 常量。                                                                                             |
| *CaseSensitive*    | 可选      | Variant  | 如果该参数值为 True ，则排序时区分大小写。默认值为 False 。                                                                                                         |
| *BidiSort*         | 可选      | Variant  | 如果该参数值为 True，则基于从右向左语言的规则进行排序。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。                                           |
| *IgnoreThe*        | 可选      | Variant  | 如果该参数值为 True，则对从右向左的语言文本进行排序时忽略阿拉伯字符 alef lam。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。                    |
| *IgnoreKashida*    | 可选      | Variant  | 如果该参数值为 True ，则对从右向左的语言文本进行排序时忽略 Kashidas 。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。                            |
| *IgnoreDiacritics* | 可选      | Variant  | 如果该参数值为 True ，则对从右向左的语言文本进行排序时忽略双向控制字符。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。                          |
| *IgnoreHe*         | 可选      | Variant  | 如果该参数值为 True ，则对从右向左的语言文本进行排序时忽略希伯来字符 he 。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。                        |
| *LanguageID*       | 可选      | Variant  | 指定排序语言。可以是 WdLanguageID 常量之一。要获得 WdLanguageID 常量的列表，请参阅“对象浏览器”。                                                                    |

**示例**

``` JavaScript
/*以下示例对活动文档中的第一个表格进行排序，首行除外*/
function NewTableSort(){
    Application.ActiveDocument.Tables.Item(1).Sort(true)
}
```

### Table.SortAscending

按字母数字升序对段落或表格行进行排序。

**语法**

**express.SortAscending ()**

\*express是一个代表 Table 对象的变量。

**说明**

将表格第一行视为域名记录，不进行排序。可用 **Sort** 方法将首行包括在排序中。此方法提供针对包含数据列的邮件合并数据源进行排序的一种简化形式。 对于大多数排序任务，请使用 **Sort** 方法。

**示例**

``` JavaScript
/*以下示例以升序排列包含所选内容的表格。*/
function test()
{
if(Selection.Information(wdWithInTable) == true){
    Selection.Tables.Item(1).SortAscending()
}
else{
    MsgBox("The insertion point is not in a table.")
}
}
```

### Table.SortDescending

以字母数字降序排列表格行。

**语法**

**express.SortDescending ()**

\*express是一个代表 Table 对象的变量。

**说明**

将表格第一行视为域名记录，不进行排序。使用 **Sort** 方法将域名记录包含在排序中。

此方法提供针对包含数据列的邮件合并数据源进行排序的一种简化形式。对于大多数排序任务，请使用 **Sort** 方法。

**示例**

``` JavaScript
/*以下示例在新文档中创建一张 5x5 表格，将文本插入每个单元格，然后以字母数字降序排列表格。*/
function test()
{
let newDoc = Documents.Add()
let myTable = newDoc.Tables.Add(Selection.Range, 5, 5)

for(let i = 1; i <= myTable.Rows.Count; i++){
    for(let j = 1; j <= myTable.Columns.Count; j++){
        let MyRange = myTable.Rows.Item(i).Cells.Item(j).Range
        MyRange.InsertAfter("Cell " + i + "," + j)
    }
}
MsgBox("Click OK to sort in descending order.")
myTable.SortDescending()
}
```

``` JavaScript
/*以下示例以字母数字降序排列插入点所在的表格。*/
function test()
{
if(Selection.Information(wdWithInTable) == true){
    Selection.Tables.Item(1).SortDescending()
}
else{
    MsgBox("The insertion point is not in a table.")
}
}
```

### Table.Split

在表格中紧靠指定行的上面插入一空段落，并且返回一个 **Table** 对象，此对象包含指定行及其下一行。

**语法**

**express.Split ( *BeforeRow* )**

\*express是一个代表 Table 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                            |
|-------------|-----------|----------|-------------------------------------------------|
| *BeforeRow* | 必选      | Variant  | 将要拆分的表格的前一行。可以是行号或 Row 对象。 |

### Table.UpdateAutoFormat

使用预定义表格式对表格式进行更新。

**语法**

**express.UpdateAutoFormat ()**

\*express是一个代表 Table 对象的变量。

**说明**

下面举例说明本方法的作用：如果使用 **AutoFormat** 来应用一种表格式，然后插入行和列，那么表格的外观可能不再像预定义的那样。 **UpdateAutoFormat** 会恢复该格式。

**示例**

``` JavaScript
/*本示例创建一个表格并应用预定义格式，再添加一行，然后重新应用预定义格式。*/
function test()
{
let docNew = Documents.Add()
let tableNew = docNew.Tables.Add(Selection.Range, 5, 5)

tableNew.AutoFormat(wdTableFormatColumns1)
tableNew.Rows.Add(tableNew.Rows.Item(1))

MsgBox("Click OK to reapply autoformatting.")
tableNew.UpdateAutoFormat()
}
```

``` JavaScript
/*本示例恢复插入点所在表格的预定义格式。*/
function test()
{
if(Selection.Information(wdWithInTable) == true){
    Selection.Tables.Item(1).UpdateAutoFormat()
}
else{
    MsgBox("The insertion point is not in a table.")
}
}
```

## 成员属性

### Table.Allowautofit

使 WPS 可以自动调整表格中的单元格的大小以适应内容。 **Boolean** 类型，可读写。

**语法**

**express.Allowautofit**

\*express是一个代表 Table 对象的变量。

**示例**

``` JavaScript
/*本示例设置活动文档的第一张表格可自动调整大小以适应内容*/
function AllowFit(){
    Applicaiton.ActiveDocument.Tables.Item(1).AllowAutoFit = true
}
```

### Table.AllowPageBreaks

允许 WPS 对指定表格进行跨页断行。 **Boolean** 类型，可读写。

**语法**

**express.AllowPageBreaks**

\*express是一个代表 Table 对象的变量。

**示例**

``` JavaScript
/*本示例设置活动文档的第二个表格跨页断行。*/
function test(){
    ActiveDocument.Tables.Item(2).AllowPageBreaks = true
}
```

### Table.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Table 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Table.ApplyStyleColumnBands

返回或设置 **Boolean** 值，该值表示如果所应用的预置表样式为列提供了样式带，则是否对表中的列应用样式带。可读/写。

**语法**

**express.ApplyStyleColumnBands**

\*express是一个代表 Table 对象的变量。

### Table.ApplyStyleFirstColumn

如果为 **True** ，则 WPS 对指定表格的第一列应用第一列的格式。 **Boolean** 类型，可读写。

**语法**

**express.ApplyStyleFirstColumn**

\*express是一个代表 Table 对象的变量。

**说明**

指定的表格样式必须包含第一列的格式，才能将此格式应用于表格。

**示例**

``` JavaScript
/*本示例用“Table Style 1”的表格样式设置活动文档中的第二个表格的格式，并删除第一行和最后一行、第一列和最后一列的格式。本示例假定名为“Table Style 1”的表格样式存在，并且包含第一列的格式。*/
function test(){
    let table = ActiveDocument.Tables.Item(2)

    table.Style = "Table Style 1"
    table.ApplyStyleFirstColumn = false
    table.ApplyStyleHeadingRows = false
    table.ApplyStyleLastColumn = false
    table.ApplyStyleLastRow = false
}
```

### Table.ApplyStyleHeadingRows

如果为 **True** ，则 WPS 为选定表格的第一行应用标题行的格式。 **Boolean** 类型，可读写。

**语法**

**express.ApplyStyleHeadingRows**

\*express是一个代表 Table 对象的变量。

**说明**

指定的表格样式必须包含标题行的格式，才能将此格式应用于表格。

**示例**

``` JavaScript
/*本示例用“Table Style 1”的表格样式设置活动文档中的第二个表格的格式，并删除第一行和最后一行、第一列和最后一列的格式。本示例假定名为“Table Style 1”的表格样式存在，并且包含标题行的格式。*/
function TableStyles(){
    let table = ActiveDocument.Tables.Item(2)

    table.Style = "Table Style 1"
    table.ApplyStyleFirstColumn = false
    table.ApplyStyleHeadingRows = false
    table.ApplyStyleLastColumn = false
    table.ApplyStyleLastRow = false
}
```

### Table.ApplyStyleLastColumn

如果为 **True** ，则 WPS 对指定表格的最后一列应用最后一列的格式。 **Boolean** 类型，可读写。

**语法**

**express.ApplyStyleLastColumn**

\*express是一个代表 Table 对象的变量。

**说明**

指定的表格样式必须包含最后一列的格式，才能将此格式应用于表格。

**示例**

``` JavaScript
/*本示例用“Table Style 1”的表格样式设置活动文档中的第二个表格的格式，并删除第一行和最后一行、第一列和最后一列的格式。本示例假定名为“Table Style 1”的表格样式存在，并且其包含最后一列的格式。*/
function TableStyles(){
    let table = ActiveDocument.Tables.Item(2)

    table.Style = "Table Style 1"
    table.ApplyStyleFirstColumn = false
    table.ApplyStyleHeadingRows = false
    table.ApplyStyleLastColumn = false
    table.ApplyStyleLastRow = false
}
```

### Table.ApplyStyleLastRow

如果为 **True** ，则 WPS 对指定表格的最后一行应用最后一行的格式。 **Boolean** 类型，可读写。

**语法**

**express.ApplyStyleLastRow**

\*express是一个代表 Table 对象的变量。

**说明**

指定的表格样式必须包含最后一行的格式，才能将此格式应用于表格。

**示例**

``` JavaScript
/*本示例用“Table Style 1”的表格样式设置活动文档中的第二个表格的格式，并删除第一行和最后一行、第一列和最后一列的格式。本示例假定名为“Table Style 1”的表格样式存在，并且其包含最后一行的格式。*/
function test(){
    let table = ActiveDocument.Tables.Item(2)

    table.Style = "Table Style 1"
    table.ApplyStyleFirstColumn = false
    table.ApplyStyleHeadingRows = false
    table.ApplyStyleLastColumn = false
    table.ApplyStyleLastRow = false
}
```

### Table.ApplyStyleRowBands

返回或设置 **Boolean** 值，该值表示如果所应用的预置表样式为行提供了样式带，则是否对表中的行应用样式带。可读/写。

**语法**

**express.ApplyStyleRowBands**

\*express是一个代表 Table 对象的变量。

### Table.AutoFormatType

返回指定表格所应用的自动套用格式类型。 **Long** 类型，只读。

**语法**

**express.AutoFormatType**

\*express是一个代表 Table 对象的变量。

**说明**

该属性可以是 **WdTableFormat** 常量之一。可使用 **AutoFormat** 方法为表格设置自动套用格式。

**示例**

``` JavaScript
/*如果活动文档的第一个表格的当前格式是“简明型 1”、“简明型 2”或“简明型 3”，那么本示例为该表格自动套用“古典型 1”格式。*/
function test()
{
if(ActiveDocument.Tables.Count >= 1){
    if(ActiveDocument.Tables.Item(1).AutoFormatType <= wdTableFormatSimple3){
        ActiveDocument.Tables.Item(1).AutoFormat(wdTableFormatClassic1)
    }
}
}
```

### Table.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Table 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例对活动文档中的第一个表格应用内部和外部边框*/
function test(){
    let myTable = Application.ActiveDocument.Tables.Item(1)
    myTable.Borders.InsideLineStyle = wdLineStyleSingle
    myTable.Borders.OutsideLineStyle = wdLineStyleDouble
}
```

### Table.BottomPadding

返回或设置在表格中单独的单元格或所有单元格内容的下方添加的间距（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.BottomPadding**

\*express是一个代表 Table 对象的变量。

**说明**

单独单元格的 **BottomPadding** 属性的设置值会覆盖整个表格的 **BottomPadding** 属性的设置值。

**示例**

``` JavaScript
/*本示例将活动文档中第一个表格的底边距设为 40 像素。*/
ActiveDocument.Tables.Item(1).BottomPadding = PixelsToPoints(40, true)
```

### Table.Columns

返回一个 **Columns** 集合，该集合代表表格中的所有列。只读。

**语法**

**express.Columns**

\*express是一个代表 Table 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例显示活动文档的第一个表格中的列数*/
if(Application.ActiveDocument.Tables.Count >= 1){
    alart(Application.ActiveDocument.Tables.Item(1).Columns.Count)
}
```

### Table.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Table 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Table.Descr

返回或设置包含指定表格的说明的 **String** 类型值。可读/写。

**语法**

**express.Descr**

\*express是一个代表 Table 对象的变量。

**说明**

使用 **Descr** 属性可提供表格的可选文字说明。该属性将文本添加到 WPS 2015 中 **“表格属性”** 对话框的 **“可选文字”** 选项卡上的 **“说明”** 文本框中。

| 注释                                                                                                                       |
|----------------------------------------------------------------------------------------------------------------------------|
| Web 浏览器在加载表格的过程中或表格丢失时显示可选文字。Web 搜索引擎利用可选文字帮助查找网页。可选文字也可用来帮助残障人士。 |

**示例**

``` JavaScript
/*下面的代码示例向活动文档的第一个表格中添加可选文字表格说明。*/
function test()
{
let doc = ActiveDocument
let tbl = doc.Tables.Item(1)
tbl.Descr = "This is a table description."
}
```

### Table.ID

在将文档另存为网页时，返回或设置指定的表的标识标签。可读/写 **String** 类型。

**语法**

**express.ID**

\*express是一个代表 Table 对象的变量。

### Table.LeftPadding

返回或设置在表格中所有单元格的内容左侧要增加的间距（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.LeftPadding**

\*express是一个代表 Table 对象的变量。

**说明**

单个单元格的 **LeftPadding** 属性设置将覆盖整个表格的 **LeftPadding** 属性设置。

**示例**

``` JavaScript
/*以下示例将活动文档中第一个表格的左边距设为 40 像素。*/
ActiveDocument.Tables.Item(1).LeftPadding = PixelsToPoints(40, false)
```

### Table.Nestinglevel

返回指定表格的嵌套层。只读 **Long** 类型。

**语法**

**express.Nestinglevel**

\*express是一个代表 Table 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

### Table.Parent

返回一个 **Object** 类型值，该值代表指定 **Table** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Table 对象的变量。

### Table.Preferredwidth

返回或设置指定表格的指定宽度（以磅为单位或以窗口宽度的百分比表示）。可读写 **Single** 类型。

**语法**

**express.Preferredwidth**

\*express是一个代表 Table 对象的变量。

**说明**

如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPoints** ，则 **PreferredWidth** 属性返回或设置宽度（以磅为单位）。如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPercent** ，则 **PreferredWidth** 属性返回或设置宽度（以窗口宽度的百分比表示）。

**示例**

``` JavaScript
/*以下示例将 WPS 设置为以窗口宽度的百分比来表示指定宽度，然后将文档中第一张表格的指定宽度设置为窗口宽度的 50%*/
Application.ActiveDocument.Tables.Item(1).PreferredWidthType = wdPreferredWidthPercent
Application.ActiveDocument.Tables.Item(1).PreferredWidth = 50
```

### Table.Preferredwidthtype

返回指定表格所应用的自动套用格式类型。 **Long** 类型，只读。

**语法**

**express.Preferredwidthtype**

\*express是一个代表 Table 对象的变量。

**说明**

该属性可以是 **WdTableFormat** 常量之一。可使用 **AutoFormat** 方法为表格设置自动套用格式。

**示例**

``` JavaScript
/*如果活动文档的第一个表格的当前格式是“简明型 1”、“简明型 2”或“简明型 3”，那么本示例为该表格自动套用“古典型 1”格式*/
if(Application.ActiveDocument.Tables.Count >= 1){
    if(Application.ActiveDocument.Tables.Item(1).AutoFormatType <= wdTableFormatSimple3){
        Application.ActiveDocument.Tables.Item(1).AutoFormat(wdTableFormatClassic1)
    }
}
```

### Table.Range

返回一个 **Range** 对象，该对象代表指定表格中所含的文档部分。

**语法**

**express.Range**

\*express是一个代表 Table 对象的变量。

### Table.RightPadding

返回或设置在表格的所有单元格的内容右侧要增加的间距（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.RightPadding**

\*express是一个代表 Table 对象的变量。

**说明**

单个单元格的 **RightPadding** 属性设置将覆盖整个表格的 **RightPadding** 属性设置。

**示例**

``` JavaScript
/*以下示例将活动文档中第一个表格的右边距设置为 40 像素。*/
ActiveDocument.Tables.Item(1).RightPadding = PixelsToPoints(40, false)
```

### Table.Rows

返回一个 **Rows** 集合，该集合代表表格中所有的表格行。只读。

**语法**

**express.Rows**

\*express是一个代表 Table 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象

**示例**

``` JavaScript
/*以下示例删除活动文档第一个表格的第二行*/
Application.ActiveDocument.Tables.Item(1).Rows.Item(2).Delete()
```

### Table.Shading

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Table 对象的变量。

**示例**

``` JavaScript
/*以下示例对活动文档中第一个表格应用横线纹理。*/
if(ActiveDocument.Tables.Count >= 1){
    ActiveDocument.Tables.Item(1).Shading.Texture = wdTextureHorizontal
}
```

### Table.Spacing

返回或设置表格中单元格的间距（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.Spacing**

\*express是一个代表 Table 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档的第一个表格中的单元格的间距设置为 9 磅。*/
ActiveDocument.Tables.Item(1).Spacing = 9
```

### Table.Style

返回或设置指定表格的样式。可读写 **Variant** 类型。

**语法**

**express.Style**

\*express是一个代表 Table 对象的变量。

**说明**

要设置该属性，您可以指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个包含所需表格格式的 **TableStyle** 对象。

### Table.TableDirection

返回或设置 WPS 对指定表格中的单元格进行排序的方向。可读写 **WdTableDirection** 类型。

**语法**

**express.TableDirection**

\*express是一个代表 Table 对象的变量。

**说明**

如果将 **TableDirection** 属性设为 **wdTableDirectionLtr** ，则选定的行按照最左侧的第一列进行排列。如果将 **TableDirection** 属性设为 **wdTableDirectionRtl** ，则选定的行按照最右侧的第一列进行排列。

**示例**

``` JavaScript
/*以下示例设置 WPS 从右向左对选定行中的单元格进行排序。*/
Selection.Rows.TableDirection = wdTableDirectionRtl
```

### Table.Tables

返回一个 **Tables** 集合，该集合代表指定表格中的所有嵌套表格。只读。

**语法**

**express.Tables**

\*express是一个代表 Table 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Table.Title

返回或设置包含指定表格的标题的 **String** 类型值。可读/写。

**语法**

**express.Title**

\*express是一个代表 Table 对象的变量。

**说明**

使用 **Title** 属性可为表格提供可选文字标题。该属性将标题文本添加到 WPS 2015 中 **“表格属性”** 对话框的 **“可选文字”** 选项卡上的 **“标题”** 文本框中。

| 注释                                                                                                                       |
|----------------------------------------------------------------------------------------------------------------------------|
| Web 浏览器在加载表格的过程中或表格丢失时显示可选文字。Web 搜索引擎利用可选文字帮助查找网页。可选文字也可用来帮助残障人士。 |

**示例**

``` JavaScript
/*下面的代码示例向活动文档中的第一个表格中添加可选文字表格标题。*/
ActiveDocument.Tables.Item(1).Title = "Table 1."
```

### Table.TopPadding

返回或设置表格中所有单元格的内容上方要增加的间距（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.TopPadding**

\*express是一个代表 Table 对象的变量。

**说明**

单个单元格的 **TopPadding** 属性设置将覆盖整个表格的 **TopPadding** 属性设置。

**示例**

``` JavaScript
/*以下示例将活动文档中第一个表格的上边距设为 40 像素。*/
ActiveDocument.Tables.Item(1).TopPadding = PixelsToPoints(40, true)
```

### Table.Uniform

如果表格中所有的行具有相同的列数，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.Uniform**

\*express是一个代表 Table 对象的变量。

**示例**

``` JavaScript
/*本示例创建一个包含拆分单元格的表格，然后显示一个消息框，确认并非表格中的每一行都具有相同的列数。*/
function test()
{
let newDoc = Documents.Add()
let myTable = newDoc.Tables.Add(Selection.Range, 5, 5)
myTable.Cell(3, 3).Split(1, 2)

if(!myTable.Uniform){
    MsgBox("Table is not uniform")
}
}
```

``` JavaScript
/*本示例确定包含选定内容的表格中的各行是否具有相同的列数。*/
function test()
{
if(Selection.Information(wdWithInTable) == true){
    MsgBox(Selection.Tables.Item(1).Uniform)
}
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Table](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Zooms 对象

## 简介

Zoom 对象的集合，该集合代表每个视图（如大纲视图、普通视图或页面视图）的缩放选项。

## 方法

| 序号 | 名称                | 说明                              |
|:----:|:--------------------|:----------------------------------|
|  1   | [Item](#Zooms.Item) | 返回指定的 Zoom 对象。返回 Zoom值 |

## 属性

| 序号 | 名称                              | 说明                                                                         |
|:----:|:----------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Zooms.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Creator](#Zooms.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  3   | [Parent](#Zooms.Parent)           | 返回一个 Object 类型值，该值代表指定 Zooms 对象的父对象。                    |

## 成员方法

### Zooms.Item

返回指定的 **Zoom** 对象。返回 Zoom值

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Zooms 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型   | 说明                 |
|---------|-----------|------------|----------------------|
| *Index* | 可选      | WdViewType | 指定的显示比例类型。 |

## 成员属性

### Zooms.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 Zooms 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Zooms.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Zooms 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 Creator 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Zooms.Parent

返回一个 **Object** 类型值，该值代表指定 **Zooms** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Zooms 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Zooms](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ColorSchemes 对象

## 简介

指定的演示文稿中所有 ColorScheme 对象的集合。每个 ColorScheme 对象代表一种配色方案，即同时用于幻灯片的一组颜色。

## 方法

| 序号 | 名称                       | 说明                                                                                |
|:----:|:---------------------------|:------------------------------------------------------------------------------------|
|  1   | [Add](#ColorSchemes.Add)   | 将一种配色方案添加到可用方案集合中。返回一个代表所添加配色方案的 ColorScheme 对象。 |
|  2   | [Item](#ColorSchemes.Item) | 从指定集合中返回单个对象。                                                          |

## 属性

| 序号 | 名称                                     | 说明                                                    |
|:----:|:-----------------------------------------|:--------------------------------------------------------|
|  1   | [Application](#ColorSchemes.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#ColorSchemes.Count)             | 返回指定集合中的对象数目。                              |
|  3   | [Parent](#ColorSchemes.Parent)           | 返回指定对象的父对象。                                  |

## 成员方法

### ColorSchemes.Add

将一种配色方案添加到可用方案集合中。返回一个代表所添加配色方案的 **ColorScheme** 对象。

**语法**

**express.Add ( *Scheme* )**

\*express是一个代表 ColorSchemes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型    | 说明                                                                                                                                                                                                                                                   |
|----------|-----------|-------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Scheme* | 必选      | ColorScheme | ColorScheme 对象。要添加的配色方案。可以是任意幻灯片或母版中的 ColorScheme 对象，或者是任何打开的演示文稿的 ColorSchemes 集合中的一项。如果省略此参数，则使用所指定的演示文稿的 ColorScheme 集合中的第一个 ColorSchemes 对象（即第一个标准配色方案）。 |

**返回值**

ColorScheme

**说明**

新配色方案基于指定幻灯片或母版使用的颜色，或者一个打开的演示文稿中指定配色方案中的颜色。

**ColorSchemes** 集合最多可以包含 16 种配色方案。如果需要向集合中添加另一配色方案但 **ColorSchemes** 集合已满，可以使用 **Delete** 方法删除某个现有的配色方案。

请注意，WPP 在用户界面上添加配色方案之前会自动检测重复的配色方案；但是在 Visual Basic 的过程中添加配色方案时，不会进行这样的检测。过程必须进行自检以避免添加多余的配色方案。

### ColorSchemes.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 ColorSchemes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                             |
|---------|-----------|----------|----------------------------------|
| *Index* | 必选      | Long     | 集合中要返回的单个对象的索引号。 |

**返回值**

ColorScheme

## 成员属性

### ColorSchemes.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 ColorSchemes 对象的变量。

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

### ColorSchemes.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 ColorSchemes 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭第二个窗口。*/
function test(){
　　　　let w = Application.Windows
    for(let i = 2; i <= w.Count; i++) {
　　　　    w.Item(2).Close()
　　　　}
}
```

### ColorSchemes.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 ColorSchemes 对象的变量。

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

> WPS 开发文档： [WPS 基础接口/演示 API 参考/ColorSchemes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# HeaderFooter 对象

## 简介

代表幻灯片或母版上的页眉、页脚、日期和时间、幻灯片编号或页码。幻灯片或母版的所有 HeaderFooter 对象都包含在 HeadersFooters 对象中。

## 说明

使用下表中列出的属性之一返回 HeaderFooter 对象。

| 使用此属性  | 返回                                                                         |
|-------------|------------------------------------------------------------------------------|
| DateAndTime | 代表幻灯片的日期和时间的 HeaderFooter 对象。                                 |
| Footer      | 代表幻灯片页脚的 HeaderFooter 对象。                                         |
| Header      | 代表幻灯片页眉的 HeaderFooter 对象。仅对备注页和讲义有效，不用于幻灯片。     |
| SlideNumber | 代表幻灯片编号（在幻灯片中）或页码（在备注页或讲义中）的 HeaderFooter 对象。 |

## 属性

| 序号 | 名称                                     | 说明                                                                                                                                                                       |
|:----:|:-----------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#HeaderFooter.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                    |
|  2   | [Format](#HeaderFooter.Format)           | 返回或设置自动更新的日期和时间的格式。仅适用于 HeaderFooter 对象，该对象代表时间和日期（由 DateAndTime 属性从 HeadersFooters 集合中返回）。可读/写 PpDateTimeFormat 类型。 |
|  3   | [Parent](#HeaderFooter.Parent)           | 返回指定对象的父对象。                                                                                                                                                     |
|  4   | [Text](#HeaderFooter.Text)               | 返回或设置一个 String 类型值，该值代表指定对象中包含的文本。可读/写。                                                                                                      |
|  5   | [UseFormat](#HeaderFooter.UseFormat)     | 决定日期和时间对象是否包含自动更新的信息。可读/写 MsoTriState 类型。                                                                                                       |
|  6   | [Visible](#HeaderFooter.Visible)         | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 MsoTriState 类型。                                                                                             |

## 成员属性

### HeaderFooter.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 HeaderFooter 对象的变量。

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

### HeaderFooter.Format

返回或设置自动更新的日期和时间的格式。仅适用于 **HeaderFooter** 对象，该对象代表时间和日期（由 **DateAndTime** 属性从 **HeadersFooters** 集合中返回）。可读/写 **PpDateTimeFormat** 类型。

**语法**

**express.Format**

\*express是一个代表 HeaderFooter 对象的变量。

**说明**

通过将 **UseFormat** 属性设置为 **True** ，确保将日期和时间设置为自动更新（不作为固定文本显示）

| PpDateTimeFormat 可以是下列 PpDateTimeFormat 常量之一。 |
|---------------------------------------------------------|
| ppDateTimeddddMMMMddyyyy                                |
| ppDateTimedMMMMyyyy                                     |
| ppDateTimedMMMyy                                        |
| ppDateTimeFormatMixed                                   |
| ppDateTimeHmm                                           |
| ppDateTimehmmAMPM                                       |
| ppDateTimeHmmss                                         |
| ppDateTimehmmssAMPM                                     |
| ppDateTimeMdyy                                          |
| ppDateTimeMMddyyHmm                                     |
| ppDateTimeMMddyyhmmAMPM                                 |
| ppDateTimeMMMMdyyyy                                     |
| ppDateTimeMMMMyy                                        |
| ppDateTimeMMyy                                          |

**示例**

``` JavaScript
/*本示例将当前演示文稿的幻灯片母版中的日期和时间设为自动更新，并将日期和时间格式设为显示小时、分钟和秒。*/
function test(){
　　　　let da = Application.ActivePresentation.SlideMaster.HeadersFooters.DateAndTime
　　　　da.UseFormat = true
　　　　da.Format = ppDateTimeHmmss
}
```

### HeaderFooter.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 HeaderFooter 对象的变量。

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

### HeaderFooter.Text

返回或设置一个 **String** 类型值，该值代表指定对象中包含的文本。可读/写。

**语法**

**express.Text**

\*express是一个代表 HeaderFooter 对象的变量。

### HeaderFooter.UseFormat

决定日期和时间对象是否包含自动更新的信息。可读/写 **MsoTriState** 类型。

**语法**

**express.UseFormat**

\*express是一个代表 HeaderFooter 对象的变量。

**说明**

此属性仅适用于 **HeaderFooter** 对象，该对象代表日期和时间（由 **DateAndTime** 属性返回）。如果希望使用 **Format** 属性来设置或返回日期合时间的格式，请将日期和时间的 **HeaderFooter** 对象的 **UseFormat** 属性设为 **True** 。如果希望设置或返回固定日期和时间的文本字符串，请将 **UseFormat** 属性设为 **msoFalse** 。

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue                                      |
| msoFalse 日期和时间对象是固定字符串。         |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue 日期和时间对象包含自动更新的信息。    |

**示例**

``` JavaScript
/*本示例将当前演示文稿的幻灯片母版中的日期和时间设为自动更新，并将日期和时间格式设为显示小时、分钟和秒。*/
function test(){
　　　　let da = Application.ActivePresentation.SlideMaster.HeadersFooters.DateAndTime
　　　　da.UseFormat = msoTrue
　　　　da.Format = ppDateTimeHmmss
}
```

### HeaderFooter.Visible

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 **MsoTriState** 类型。

**语法**

**express.Visible**

\*express是一个代表 HeaderFooter 对象的变量。

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
　　　　let sh= ActivePresentation.Slides.Item(1).Shapes.Item(3).Shadow
　　　　sh.Visible = msoTrue
　　　　sh.OffsetX = 5
　　　　sh.OffsetY = -3
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/HeaderFooter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PageSetup 对象

## 简介

包含演示文稿中有关幻灯片、备注页、讲义及大纲的页面设置的信息。

## 说明

使用 PageSetup 属性返回 PageSetup 对象。以下示例将活动演示文稿中所有幻灯片设为宽 11 英寸、高 8.5 英寸，并将演示文稿的幻灯片编号设为从 17 开始。

``` JavaScript
function test() {
    let page = Application.ActivePresentation.PageSetup
    page.SlideWidth = 11 * 72
    page.SlideHeight = 8.5 * 72
    page.FirstSlideNumber = 2
}
```

## 属性

| 序号 | 名称                                            | 说明                                                                                          |
|:----:|:------------------------------------------------|:----------------------------------------------------------------------------------------------|
|  1   | [Application](#PageSetup.Application)           | 返回一个 Application 对象，该对象表示指定对象的创建者。                                       |
|  2   | [FirstSlideNumber](#PageSetup.FirstSlideNumber) | 返回或设置演示文稿中第一张幻灯片的编号。可读/写 Long 类型。                                   |
|  3   | [NotesOrientation](#PageSetup.NotesOrientation) | 返回或设置指定演示文稿的备注页、讲义和大纲的屏幕显示和打印方向。可读/写 MsoOrientation 类型。 |
|  4   | [Parent](#PageSetup.Parent)                     | 返回指定对象的父对象。                                                                        |
|  5   | [SlideHeight](#PageSetup.SlideHeight)           | 以磅为单位返回或设置幻灯片的高度。可读/写 Single 类型。                                       |
|  6   | [SlideOrientation](#PageSetup.SlideOrientation) | 返回或设置指定演示文稿中幻灯片的屏幕显示和打印方向。可读/写 MsoOrientation 类型。             |
|  7   | [SlideSize](#PageSetup.SlideSize)               | 返回或设置指定演示文稿的幻灯片大小。可读/写 PpSlideSizeType 类型。                            |
|  8   | [SlideWidth](#PageSetup.SlideWidth)             | 以磅为单位返回或设置幻灯片的宽度。可读/写 Single 类型。                                       |

## 成员属性

### PageSetup.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(pptPres){
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test(){
  for(let i = 1 ; i <= Application.ActivePresentation.Slides.Item(1).Shapes.Count; i++){
      if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject){
          alert(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)
      }
  }
}
```

### PageSetup.FirstSlideNumber

返回或设置演示文稿中第一张幻灯片的编号。可读/写 **Long** 类型。

**语法**

**express.FirstSlideNumber**

\*express是一个代表 PageSetup 对象的变量。

**说明**

幻灯片编号是在显示幻灯片编号时出现在幻灯片右下角的实际编号。此编号是由演示文稿中幻灯片编号（顺序，即 **SlideIndex** 属性值）和演示文稿的起始幻灯片编号（ **FirstSlideNumber** 属性值）所决定的。幻灯片编号始终等于起始幻灯片编号加上幻灯片索引号再减去一。 **SlideNumber** 属性返回幻灯片编号。

**示例**

``` JavaScript
/*本示例显示更改活动演示文稿中第一张幻灯片的编号。*/
function test(){
  let tion = Application.ActivePresentation
  tion.PageSetup.FirstSlideNumber = 1         //starts slide numbering at 1
  alert(tion.Slides.Item(2).SlideNumber)     //returns 2
  tion.PageSetup.FirstSlideNumber = 10        //starts slide numbering at 10
  alert(tion.Slides.Item(2).SlideNumber)     //returns 11
}
```

### PageSetup.NotesOrientation

返回或设置指定演示文稿的备注页、讲义和大纲的屏幕显示和打印方向。可读/写 **MsoOrientation** 类型。

**语法**

**express.NotesOrientation**

\*express是一个代表 PageSetup 对象的变量。

**说明**

|                                                     |
|-----------------------------------------------------|
| MsoOrientation 可以是下列 MsoOrientation 常量之一。 |
| **msoOrientationHorizontal**                        |
| **msoOrientationMixed**                             |
| **msoOrientationVertical**                          |

**示例**

``` JavaScript
//本示例将活动演示文稿中所有备注页、讲义和大纲的方向设为水平（横向）。
Application.ActivePresentation.PageSetup.NotesOrientation = msoOrientationHorizontal
```

### PageSetup.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
//以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。
function test() {
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

### PageSetup.SlideHeight

以磅为单位返回或设置幻灯片的高度。可读/写 **Single** 类型。

**语法**

**express.SlideHeight**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*以下示例会将当前演示文稿幻灯片的高度设为 8.5 英寸，宽度设为 11 英寸。*/
function test(){
  let page = Application.ActivePresentation.PageSetup
  page.SlideWidth = 11 * 72
  page.SlideHeight = 8.5 * 72
}
```

### PageSetup.SlideOrientation

返回或设置指定演示文稿中幻灯片的屏幕显示和打印方向。可读/写 **MsoOrientation** 类型。

**语法**

**express.SlideOrientation**

\*express是一个代表 PageSetup 对象的变量。

**说明**

|                                                     |
|-----------------------------------------------------|
| MsoOrientation 可以是下列 MsoOrientation 常量之一。 |
| **msoOrientationHorizontal**                        |
| **msoOrientationMixed**                             |
| **msoOrientationVertical**                          |

**示例**

``` JavaScript
/*本示例将活动演示文稿所有幻灯片的方向设为垂直（纵向）。*/
Application.ActivePresentation.PageSetup.SlideOrientation = msoOrientationVertical
```

### PageSetup.SlideSize

返回或设置指定演示文稿的幻灯片大小。可读/写 **PpSlideSizeType** 类型。

**语法**

**express.SlideSize**

\*express是一个代表 PageSetup 对象的变量。

**说明**

|                                                       |
|-------------------------------------------------------|
| PpSlideSizeType 可以是下列 PpSlideSizeType 常量之一。 |
| **ppSlideSize35MM**                                   |
| **ppSlideSizeA3Paper**                                |
| **ppSlideSizeA4Paper**                                |
| **ppSlideSizeB4ISOPaper**                             |
| **ppSlideSizeB4JISPaper**                             |
| **ppSlideSizeB5ISOPaper**                             |
| **ppSlideSizeB5JISPaper**                             |
| **ppSlideSizeBanner**                                 |
| **ppSlideSizeCustom**                                 |
| **ppSlideSizeHagakiCard**                             |
| **ppSlideSizeLedgerPaper**                            |
| **ppSlideSizeLetterPaper**                            |
| **ppSlideSizeOnScreen**                               |
| **ppSlideSizeOverhead**                               |

**示例**

``` JavaScript
本示例将活动演示文稿的幻灯片设为投影机幻灯片大小。
Application.ActivePresentation.PageSetup.SlideSize = ppSlideSizeOverhead
```

### PageSetup.SlideWidth

以磅为单位返回或设置幻灯片的宽度。可读/写 **Single** 类型。

**语法**

**express.SlideWidth**

\*express是一个代表 PageSetup 对象的变量。

**示例**

``` JavaScript
/*以下示例会将当前演示文稿幻灯片的高度设为 8.5 英寸，宽度设为 11 英寸。*/
function test(){
  let page = Application.ActivePresentation.PageSetup
  page.SlideWidth = 11 * 72
  page.SlideHeight = 8.5 * 72
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/PageSetup](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PrintOptions 对象

## 简介

包含演示文稿的打印选项。

注释： 指定 PrintOut 方法的可选参数 From 、 To 、 Copies 和 Collate 将设置 PrintOptions 对象的相应属性。

## 说明

使用 PrintOptions 属性返回 PrintOptions 对象。以下示例以非逐份方式打印活动演示文稿所有幻灯片（无论可见或隐藏）的两个彩色副本。该示例还调整每个幻灯片的大小以适应打印页，并给每个幻灯片加细边框。

``` JavaScript
function test(){
    let active = Application.ActivePresentation
        let print = active.PrintOptions
            print.NumberOfCopies = 2
            print.Collate = false
            print.PrintColorType = ppPrintColor
            print.PrintHiddenSlides = true
            print.FitToPage = true
            print.FrameSlides = true
            print.OutputType = ppPrintOutputSlides
        active.PrintOut()
}
```

使用 RangeType 属性指定打印整个演示文稿或仅打印其中的指定部分。如果想要只打印特定幻灯片，请将 RangeType 属性设为 ppPrintSlideRange ，并使用 Ranges 属性指定要打印的页。以下示例打印活动演示文稿中的第一、四、五、六张幻灯片 。

``` JavaScript
function test(){
    let active = Application.ActivePresentation
        let print = active.PrintOptions
            print.RangeType = ppPrintSlideRange
            let ranges = print.Ranges
                ranges.Add(1, 1)
                ranges.Add(4, 6)
        active.PrintOut()
}
```

## 属性

| 序号 | 名称                                                       | 说明                                                                                                                                                                                                                                                                                     |
|:----:|:-----------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ActivePrinter](#PrintOptions.ActivePrinter)               | 返回当前打印机的名称。只读。 String 类型。                                                                                                                                                                                                                                               |
|  2   | [Application](#PrintOptions.Application)                   | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                                                                  |
|  3   | [Collate](#PrintOptions.Collate)                           | 决定是否在打印指定演示文稿的完整副本之后才打印其下一个副本的第一页。可读/写 MsoTriState 类型。                                                                                                                                                                                           |
|  4   | [FitToPage](#PrintOptions.FitToPage)                       | 决定是否调整幻灯片大小以适应打印页面。 MsoTriState 类型，可读/写。                                                                                                                                                                                                                       |
|  5   | [FrameSlides](#PrintOptions.FrameSlides)                   | 决定是否在打印的幻灯片的边框周围放置窄框架。可读/写 MsoTriState 类型。适用于打印的幻灯片、讲义和备注页。                                                                                                                                                                                 |
|  6   | [HandoutOrder](#PrintOptions.HandoutOrder)                 | 返回或设置在打印的讲义上显示幻灯片的页面版式顺序，其中可在一个页面上显示多张幻灯片。可读/写 PpPrintHandoutOrder 类型。                                                                                                                                                                   |
|  7   | [NumberOfCopies](#PrintOptions.NumberOfCopies)             | 返回或设置要打印的演示文稿份数。默认值为 1。可读/写 Long 类型。                                                                                                                                                                                                                          |
|  8   | [OutputType](#PrintOptions.OutputType)                     | 返回或设置一个值，该值指示要打印演示文稿的哪些组件（幻灯片、讲义、备注页或大纲）。可读/写 PpPrintOutputType 类型。                                                                                                                                                                       |
|  9   | [Parent](#PrintOptions.Parent)                             | 返回指定对象的父对象。                                                                                                                                                                                                                                                                   |
|  10  | [PrintColorType](#PrintOptions.PrintColorType)             | 返回或设置指定文档的打印方式：黑白方式、纯黑色和纯白色（亦称为高对比度）方式或彩色方式。默认值由打印机设置。可读/写 PpPrintColorType 类型。                                                                                                                                              |
|  11  | [PrintComments](#PrintOptions.PrintComments)               | 设置或返回是否打印批注。可读/写 MsoTriState 类型。                                                                                                                                                                                                                                       |
|  12  | [PrintFontsAsGraphics](#PrintOptions.PrintFontsAsGraphics) | 决定是否将 TrueType 字体作为图形打印。可读/写 MsoTriState 类型。                                                                                                                                                                                                                         |
|  13  | [PrintHiddenSlides](#PrintOptions.PrintHiddenSlides)       | 决定是否打印指定演示文稿中的隐藏幻灯片。可读/写 MsoTriState 类型。                                                                                                                                                                                                                       |
|  14  | [PrintInBackground](#PrintOptions.PrintInBackground)       | 决定是否在后台打印指定的演示文稿。可读/写 MsoTriState 类型。                                                                                                                                                                                                                             |
|  15  | [Ranges](#PrintOptions.Ranges)                             | 返回一个 PrintRanges 对象，该对象代表演示文稿中要打印的幻灯片范围。只读。                                                                                                                                                                                                                |
|  16  | [RangeType](#PrintOptions.RangeType)                       | 返回或设置演示文稿打印范围的类型。可读/写 PpPrintRangeType 类型。                                                                                                                                                                                                                        |
|  17  | [SlideShowName](#PrintOptions.SlideShowName)               | 返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（ ActionSetting 对象）；返回或设置自定义幻灯片放映的名称以便打印（ PrintOptions 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（ PublishObject 对象）。 String 类型，可读/写。 |

## 成员属性

### PrintOptions.ActivePrinter

返回当前打印机的名称。只读。 **String** 类型。

**语法**

**express.ActivePrinter**

\*express是一个代表 PrintOptions 对象的变量。

**示例**

``` JavaScript
/*以下示例显示当前打印机的名称。*/
alert("The name of the active printer is " + Application.ActivePrinter)
```

### PrintOptions.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 PrintOptions 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test() {
  let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
  for(let i = 1; i <= myShapes.Count; i++){
      if(myShapes.Item(i).Type == msoLinkedOLEObject){
          alert(myShapes.Item(i).OLEFormat.Application.Name)
      }
  }
}
```

### PrintOptions.Collate

决定是否在打印指定演示文稿的完整副本之后才打印其下一个副本的第一页。可读/写 **MsoTriState** 类型。

**语法**

**express.Collate**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

指定 **PrintOut** 方法的 **Collate** 参数值将设置该属性的值。

|                                                                                  |
|----------------------------------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。                                    |
| **msoCTrue**                                                                     |
| **msoFalse**                                                                     |
| **msoTriStateMixed**                                                             |
| **msoTriStateToggle**                                                            |
| **msoTrue** 默认值。在打印指定演示文稿的完整副本之后才打印其下一个副本的第一页。 |

**示例**

``` JavaScript
//本示例将活动演示文稿逐份打印三份。
function CheckOutPresentation(strPresentation) {
    let print = Application.ActivePresentation.PrintOptions
    print.NumberOfCopies = 3
    print.Collate = msoTrue
    print.Parent.PrintOut()
}
```

### PrintOptions.FitToPage

决定是否调整幻灯片大小以适应打印页面。 **MsoTriState** 类型，可读/写。

**语法**

**express.FitToPage**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

|                                                                                                                    |
|--------------------------------------------------------------------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。                                                              |
| **msoCTrue**                                                                                                       |
| **msoFalse** 默认值。幻灯片将使用“页面设置”对话框中指定的尺寸，不论该尺寸是否与打印页面相适应。                    |
| **msoTriStateMixed**                                                                                               |
| **msoTriStateToggle**                                                                                              |
| **msoTrue** 调整指定幻灯片的大小以适应打印页面，而不考虑“页面设置”对话框（“文件”菜单）中“高度”框和“宽度”框中的值。 |

**示例**

``` JavaScript
/*以下示例将打印当前演示文稿并调整每张幻灯片的大小以适应打印页面。*/
function test() {
  let active = Application.ActivePresentation
      active.PrintOptions.FitToPage = msoTrue
      active.PrintOut()
}
```

### PrintOptions.FrameSlides

决定是否在打印的幻灯片的边框周围放置窄框架。可读/写 **MsoTriState** 类型。适用于打印的幻灯片、讲义和备注页。

**语法**

**express.FrameSlides**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

|                                                  |
|--------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。    |
| **msoCTrue**                                     |
| **msoFalse** 默认值。                            |
| **msoTriStateMixed**                             |
| **msoTriStateToggle**                            |
| **msoTrue** 在打印的幻灯片的边框周围放置窄框架。 |

**示例**

``` JavaScript
//本示例打印活动演示文稿，其中每张幻灯片周围都有一个框架。
function test() {
    let active = Application.ActivePresentation
    active.PrintOptions.FrameSlides = msoTrue
    active.PrintOut()
}
```

### PrintOptions.HandoutOrder

返回或设置在打印的讲义上显示幻灯片的页面版式顺序，其中可在一个页面上显示多张幻灯片。可读/写 **PpPrintHandoutOrder** 类型。

**语法**

**express.HandoutOrder**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

|                                                                                                                                                                              |
|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| PpPrintHandoutOrder 可以是下列 PpPrintHandoutOrder 常量之一。                                                                                                                |
| **ppPrintHandoutHorizontalFirst** 幻灯片水平排序，第一张幻灯片位于左上角，第二张位于其右侧。如果语言设置指定的是从右向左的语言，则第一张幻灯片位于右上角，第二张位于其左侧。 |
| **ppPrintHandoutVerticalFirst** 幻灯片垂直排序，第一张幻灯片位于左上角，第二张位于其下。如果语言设置指定的是从右向左的语言，则第一张幻灯片位于右上角，第二张位于其下。       |

**示例**

``` JavaScript
//本示例设置当前演示文稿的讲义使每个页面包含六张幻灯片，并在讲义上水平排序，然后打印。
function test() {
    let active = Application.ActivePresentation
    active.PrintOptions.OutputType = ppPrintOutputSixSlideHandouts
    active.PrintOptions.HandoutOrder = ppPrintHandoutHorizontalFirst
    active.PrintOut()
}
```

### PrintOptions.NumberOfCopies

返回或设置要打印的演示文稿份数。默认值为 1。可读/写 **Long** 类型。

**语法**

**express.NumberOfCopies**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

指定 **PrintOut** 方法的 **Copies** 参数将设置此属性的值。

**示例**

``` JavaScript
/*本示例将活动演示文稿逐份打印三份。*/
function test() {
  let print = Application.ActivePresentation.PrintOptions
      print.NumberOfCopies = 3
      print.Collate = true
      print.Parent.PrintOut()
}
```

### PrintOptions.OutputType

返回或设置一个值，该值指示要打印演示文稿的哪些组件（幻灯片、讲义、备注页或大纲）。可读/写 **PpPrintOutputType** 类型。

**语法**

**express.OutputType**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

|                                                           |
|-----------------------------------------------------------|
| PpPrintOutputType 可以是下列 PpPrintOutputType 常量之一。 |
| **ppPrintOutputBuildSlides**                              |
| **ppPrintOutputFourSlideHandouts**                        |
| **ppPrintOutputNineSlideHandouts**                        |
| **ppPrintOutputNotesPages**                               |
| **ppPrintOutputOneSlideHandouts**                         |
| **ppPrintOutputOutline**                                  |
| **ppPrintOutputSixSlideHandouts**                         |
| **ppPrintOutputSlides**                                   |
| **ppPrintOutputThreeSlideHandouts**                       |
| **ppPrintOutputTwoSlideHandouts**                         |

**示例**

``` JavaScript
//本示例打印活动演示文稿的讲义，每页打印六张幻灯片。
function test() {
    let active = Application.ActivePresentation
    active.PrintOptions.OutputType = ppPrintOutputSixSlideHandouts
    active.PrintOut()
}
```

### PrintOptions.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 PrintOptions 对象的变量。

**示例**

``` JavaScript
//以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。
function test() {
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

### PrintOptions.PrintColorType

返回或设置指定文档的打印方式：黑白方式、纯黑色和纯白色（亦称为高对比度）方式或彩色方式。默认值由打印机设置。可读/写 **PpPrintColorType** 类型。

**语法**

**express.PrintColorType**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

|                                                             |
|-------------------------------------------------------------|
| PpPrintColorType 可以是下列 PpPrintColorType 类型常数之一。 |
| **ppPrintBlackAndWhite**                                    |
| **ppPrintColor**                                            |
| **ppPrintPureBlackAndWhite**                                |

**示例**

``` JavaScript
本示例以彩色方式打印活动演示文稿的幻灯片。
function test() {
  let appli = Application.ActivePresentation
      appli.PrintOptions.PrintColorType = ppPrintColor
      appli.PrintOut()
}
```

### PrintOptions.PrintComments

设置或返回是否打印批注。可读/写 **MsoTriState** 类型。

**语法**

**express.PrintComments**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

|                                                       |
|-------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| **msoCTrue** 不适用于此属性。                         |
| **msoFalse** 默认值。不打印批注。                     |
| **msoTriStateMixed** 不适用于此属性。                 |
| **msoTriStateToggle** 不适用于此属性。                |
| **msoTrue** 打印批注。                                |

**示例**

``` JavaScript
/*本示例指示 WPP 打印批注。*/
function test(){
    Application.ActivePresentation.PrintOptions.PrintComments = msoTrue
}
```

### PrintOptions.PrintFontsAsGraphics

决定是否将 TrueType 字体作为图形打印。可读/写 **MsoTriState** 类型。

**语法**

**express.PrintFontsAsGraphics**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 将 TrueType 字体作为图形打印。    |

**示例**

``` JavaScript
//本示例指定将活动演示文稿中的 TrueType 字体作为图形打印。
Application.ActivePresentation.PrintOptions.PrintFontsAsGraphics = msoTrue
```

### PrintOptions.PrintHiddenSlides

决定是否打印指定演示文稿中的隐藏幻灯片。可读/写 **MsoTriState** 类型。

**语法**

**express.PrintHiddenSlides**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse** 默认值。                         |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 打印指定演示文稿中的隐藏幻灯片。  |

**示例**

``` JavaScript
//本示例打印活动演示文稿的所有幻灯片（不论可见或隐藏）。
function test() {
    let active = Application.ActivePresentation
    active.PrintOptions.PrintHiddenSlides = msoTrue
    active.PrintOut()
}
```

### PrintOptions.PrintInBackground

决定是否在后台打印指定的演示文稿。可读/写 **MsoTriState** 类型。

**语法**

**express.PrintInBackground**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

|                                                                                  |
|----------------------------------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。                                    |
| **msoCTrue**                                                                     |
| **msoFalse**                                                                     |
| **msoTriStateMixed**                                                             |
| **msoTriStateToggle**                                                            |
| **msoTrue** 默认值。在后台打印指定的演示文稿，这意味着可以在打印的同时继续工作。 |

**示例**

``` JavaScript
//本示例在后台打印活动演示文稿。
function test() {
    let active = Application.ActivePresentation
    active.PrintOptions.PrintInBackground = msoTrue
    active.PrintOut()
}
```

### PrintOptions.Ranges

返回一个 **PrintRanges** 对象，该对象代表演示文稿中要打印的幻灯片范围。只读。

**语法**

**express.Ranges**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

如果不想打印整个演示文稿，必须使用 **Add** 方法为每个要打印的连续的幻灯片组创建一个 **PrintRange** 对象。例如，如果要打印指定演示文稿中的第一张、第三张到第五张以及第八张和第九张幻灯片，则必须创建三个 **PrintRange** 对象：一个代表第一张幻灯片；一个代表第三张到第五张幻灯片；一个代表第八张和第九张幻灯片。有关详细信息，请参阅此属性的示例。

RangeType 属性必须设置为 **ppPrintSlideRange** 以应用 **PrintRanges** 集合中的打印范围。

要从 **PrintRanges** 集合中清除现有的所有打印范围，请使用 **ClearAll** 方法。

指定 **PrintOut** 方法的 **To** 和 **From** 参数将设置 **PrintRanges** 对象的内容。

**示例**

``` JavaScript
/*本示例打印活动演示文稿中的第一张幻灯片、第三张到第五张幻灯片以及第八张和第九张幻灯片。*/
function test() {
  let active = Application.ActivePresentation
      let print = active.PrintOptions
          print.RangeType = ppPrintSlideRange
          let ranges = print.Ranges
              ranges.Add(1, 1)
              ranges.Add(3, 5)
              ranges.Add(8, 9)
      active.PrintOut()
}
```

### PrintOptions.RangeType

返回或设置演示文稿打印范围的类型。可读/写 **PpPrintRangeType** 类型。

**语法**

**express.RangeType**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

|                                                                 |
|-----------------------------------------------------------------|
| PpSlideShowRangeType 可以是下列 PpSlideShowRangeType 常量之一。 |
| **ppShowAll**                                                   |
| **ppShowNamedSlideShow**                                        |
| **ppShowSlideRange**                                            |

若要打印在 **PrintRanges** 集合中定义的幻灯片范围，必须先将 **RangeType** 属性设置为 **ppPrintSlideRange** 。将 **RangeType** 设置为除 **ppPrintSlideRange** 之外的值会使在 **PrintRanges** 集合中定义的范围失效。但是这不会以任何方式影响 **PrintRanges** 集合的内容。也就是说，如果定义了一些打印范围，将 **RangeType** 属性设置为除 **ppPrintSlideRange** 之外的值，然后将 **RangeType** 重新设置为 **ppPrintSlideRange** ，则之前定义的打印范围将保持不变。

指定 **PrintOut** 方法的 *To* 和 *From* 参数将设置该属性的值。

**示例**

``` JavaScript
/*本示例打印当前演示文稿的当前幻灯片。*/
function test() {
  let active = Application.ActivePresentation
      active.PrintOptions.RangeType = ppPrintCurrent
      active.PrintOut()
}
```

### PrintOptions.SlideShowName

返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（ **ActionSetting** 对象）；返回或设置自定义幻灯片放映的名称以便打印（ **PrintOptions** 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（ **PublishObject** 对象）。 **String** 类型，可读/写。

**语法**

**express.SlideShowName**

\*express是一个代表 PrintOptions 对象的变量。

**说明**

必须将 **RangeType** 属性设置为 **ppPrintNamedSlideShow** 才能打印自定义幻灯片放映。

**示例**

``` JavaScript
//本示例打印名为“tech talk”的现有自定义幻灯片放映。
function test() {
    let print = Application.ActivePresentation.PrintOptions
    print.RangeType = Application.Enum.ppPrintNamedSlideShow
    print.SlideShowName = "tech talk"
    Application.ActivePresentation.PrintOut()
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/PrintOptions](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# View 对象

## 简介

代表指定文档窗口中当前编辑的视图。

## 说明

View 对象可以代表任何文档窗口视图：普通视图、幻灯片视图、大纲视图、幻灯片浏览视图、备注页视图、幻灯片母版视图、讲义母版视图或备注母版视图。 View 对象的某些属性和方法仅在特定视图中使用。如果试图使用不适合 View 对象的属性或方法，将产生错误。

## 属性

| 序号 | 名称                               | 说明                                                                                             |
|:----:|:-----------------------------------|:-------------------------------------------------------------------------------------------------|
|  1   | [Application](#View.Application)   | 返回一个 Application 对象，该对象表示指定对象的创建者。                                          |
|  2   | [PrintOptions](#View.PrintOptions) | 返回一个 PrintOptions 对象，该对象代表与指定演示文稿共同保存的打印选项。只读。                   |
|  3   | [Slide](#View.Slide)               | 返回或设置一个 Slide 对象，该对象代表在指定文档窗口视图中当前显示的幻灯片。可读/写。             |
|  4   | [Type](#View.Type)                 | 返回一个 PpViewType 常量，该常量代表视图的类型。只读。                                           |
|  5   | [Zoom](#View.Zoom)                 | 以百分比形式返回或设置指定视图的缩放设置。可以为 10% 到 400% 之间的值。 Integer 类型，可读/写 。 |

## 成员属性

### View.Application

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/
function test() {
    let pptPres = Application.ActivePresentation
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

### View.PrintOptions

返回一个 PrintOptions 对象，该对象代表与指定演示文稿共同保存的打印选项。只读。

**语法**

**express.PrintOptions**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*以下示例打印当前演示文稿中的隐藏幻灯片，并调整幻灯片的尺寸以适应纸张大小*/
function test(){
    let pres = Application.ActivePresentation
    let opt = pres.PrintOptions
    opt.PrintHiddenSlides = true
    opt.FitToPage = true
    pres.PrintOut()
}
```

### View.Slide

返回或设置一个 Slide 对象，该对象代表在指定文档窗口视图中当前显示的幻灯片。可读/写。

**语法**

**express.Slide**

\*express是一个代表 View 对象的变量。

**说明**

如果当前显示的幻灯片来源于一个嵌入演示文稿，则可以使用 **Slide** 对象（从 **Slide** 属性返回）的 **Parent** 属性返回包含该幻灯片的嵌入演示文稿。（ **SlideShowWindow** 对象或 **DocumentWindow** 对象的 **Presentation** 属性返回创建该窗口的演示文稿，而不是该嵌入演示文稿。）

**示例**

``` JavaScript
/*本示例显示当前窗口的第二个幻灯片的名称*/
Application.ActivePresentation.Slides.Item(2).Name
```

### View.Type

返回一个 PpViewType 常量，该常量代表视图的类型。只读。

**语法**

**express.Type**

\*express是一个代表 View 对象的变量。

**说明**

PpViewType 可以是下列 PpViewType 常量之一。 ppViewHandoutMaster ppViewMasterThumbnails ppViewNormal ppViewNotesMaster ppViewNotesPage ppViewOutline ppViewPrintPreview ppViewSlide ppViewSlideMaster ppViewSlideSorter ppViewThumbnails ppViewTitleMaster

### View.Zoom

以百分比形式返回或设置指定视图的缩放设置。可以为 10% 到 400% 之间的值。 **Integer** 类型，可读/写 。

**语法**

**express.Zoom**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*以下示例将第一个文档窗口的视图缩小为 30%*/
Application.ActivePresentation.Windows.Item(1).View.Zoom = 30
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/View](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CommandBar 对象

## 简介

代表容器应用程序中的一个命令栏。 CommandBar 对象是 CommandBars 集合的成员。

## 说明

使用 CommandBars ( *index* ) 可返回一个 CommandBar 对象，其中 *index* 是命令栏的名称或索引号。下面的示例遍历命令栏集合以查找名为“Forms”的命令栏。如果找到，则本示例将显示该命令栏并保护其停靠状态。在本示例中，变量 cb 代表一个 CommandBar 对象。

``` JavaScript
function test()
{  
    let foundFlag = false  let cmdBars = Application.ActiveDocument.CommandBars   
    for(let i=1;i <= cmdBars.Count;i++) 
    {      
        let cb = cmdBars.Item(i)      
        if(cb.Name == "Forms") 
        {          
            cb.Protection = Application.Enum.msoBarNoChangeDock
            cb.Visible = true           
            foundFlag = true        
        }  
    }  
    if(!foundFlag) 
    {
        alert("The collection does not contain a Forms command bar.")  
    }
} 
```

可以使用名称或索引号在容器应用程序的可用菜单栏和工具栏列表中指定菜单栏或工具栏，但必须使用名称指定菜单、快捷菜单或子菜单（所有这些内容均由 CommandBar 对象代表）。本示例在 “Tools” 菜单底部添加一个新的菜单项。单击该新菜单项将运行名为“qtrReport”的过程。

``` JavaScript
function test()
{  
    let newItem = Application.ActiveDocument.CommandBars.Item("Tools").Controls.Add(Application.Enum.msoControlButton)  
    newItem.BeginGroup = true   
    newItem.Caption = "Make Report"  　　　　 newItem.FaceID = 0  
    newItem.OnAction = "qtrReport"
}
```

如果两个或两个以上的自定义菜单或子菜单同名， CommandBars(index) 将返回第一个具有该名称的菜单或子菜单。为了确保返回正确的菜单或子菜单，可找到显示该菜单的弹出式控件，然后对该弹出式控件应用 CommandBar 属性以返回代表该菜单的命令栏。假定在名为“Custom Tools”的工具栏中，第三个控件是一个弹出式控件，则本示例可将 “Save” 命令添加到该菜单的底部。

``` JavaScript
function test()
{
    let viewMenu = Application.ActiveDocument.CommandBars.Item("Custom Tools").Controls.Item(3)
    viewMenu.Controls.Add(null,3)    //ID of Save command is 3
}
```

## 方法

| 序号 | 名称                                   | 说明                                                       |
|:----:|:---------------------------------------|:-----------------------------------------------------------|
|  1   | [Delete](#CommandBar.Delete)           | 从集合中删除 CommandBar 对象。                             |
|  2   | [FindControl](#CommandBar.FindControl) | 获取一个符合指定条件的 CommandBarControl 对象。            |
|  3   | [Reset](#CommandBar.Reset)             | 将内置命令栏重置为其默认配置。                             |
|  4   | [ShowPopup](#CommandBar.ShowPopup)     | 将指定的命令栏作为快捷菜单，在指定坐标或当前光标位置显示。 |

## 属性

| 序号 | 名称                                     | 说明                                                                                                                                                                   |
|:----:|:-----------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AdaptiveMenu](#CommandBar.AdaptiveMenu) | 获取一个 Boolean 类型的值，该值指定命令栏是否应包含自适应菜单。可读写。                                                                                                |
|  2   | [Application](#CommandBar.Application)   | 获取一个 Application 对象，代表 CommandBar 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。                                      |
|  3   | [BuiltIn](#CommandBar.BuiltIn)           | 如果指定的命令栏是容器应用程序的内置命令栏，则获取 True 。如果是自定义命令栏，则返回 False 。只读。                                                                    |
|  4   | [Context](#CommandBar.Context)           | 获取或设置一个可确定命令栏保存位置的字符串。该字符串由应用程序定义和解释。可读写。                                                                                     |
|  5   | [Controls](#CommandBar.Controls)         | 获取一个代表命令栏上的所有控件的 CommandBarControls 对象。只读。                                                                                                       |
|  6   | [Creator](#CommandBar.Creator)           | 获取一个 32 位整数，指示创建 CommandBar 对象时所使用的应用程序。只读。                                                                                                 |
|  7   | [Enabled](#CommandBar.Enabled)           | 获取或设置用于指定是否启用了指定 CommandBar 的 Boolean 值。可读写。                                                                                                    |
|  8   | [Height](#CommandBar.Height)             | 获取或设置 CommandBar 的高度。可读写。                                                                                                                                 |
|  9   | [Index](#CommandBar.Index)               | 获取一个 Long 类型的值，该值代表集合中 CommandBar 对象的索引号。只读。                                                                                                 |
|  10  | [Left](#CommandBar.Left)                 | 设置或获取从对象左边缘算起 CommandBar 相对于屏幕的水平距离（以像素为单位）。可读写。                                                                                   |
|  11  | [Name](#CommandBar.Name)                 | 获取内置的 CommandBar 对象的名称。只读。                                                                                                                               |
|  12  | [NameLocal](#CommandBar.NameLocal)       | 获取以容器应用程序的语言版本显示的内置命令栏名称，或者返回或设置自定义命令栏的名称。可读写。                                                                           |
|  13  | [Parent](#CommandBar.Parent)             | 获取 CommandBar 对象的 Parent 对象。只读。                                                                                                                             |
|  14  | [Position](#CommandBar.Position)         | 获取或设置一个代表命令栏位置的 MsoBarPosition 常量。可读写。                                                                                                           |
|  15  | [RowIndex](#CommandBar.RowIndex)         | 获取或设置一个命令栏相对于同一停靠区域中其他命令栏的停靠顺序。该属性值可以是大于零的整数，也可以是以下 MsoBarRow 常量之一： msoBarRowFirst 或 msoBarRowLast 。可读写。 |
|  16  | [Top](#CommandBar.Top)                   | 设置或获取指定的命令栏顶边到屏幕顶边的距离。对于固定命令栏，此属性返回或设置从命令栏到停靠区域顶部的距离。可读写。                                                     |
|  17  | [Type](#CommandBar.Type)                 | 获取命令栏的类型。只读。                                                                                                                                               |
|  18  | [Visible](#CommandBar.Visible)           | 获取或设置命令栏的 Visible 属性。如果命令栏可见，则为 true 。可读写。                                                                                                  |
|  19  | [Width](#CommandBar.Width)               | 获取或设置指定命令栏的宽度（以像素为单位）。可读写。                                                                                                                   |

## 成员方法

### CommandBar.Delete

从集合中删除 **CommandBar** 对象。

**语法**

**express.Delete ()**

\*express是一个代表 CommandBar 对象的变量。

**说明**

对于 **Scripts** 集合，使用 **Delete** 方法将删除指定的 WPS 文档、ET 工作表或 WPP 幻灯片中的所有脚本。在宿主应用程序中，脚本定位标记以形状表示。因此，与每个 **msoScriptAnchor** 类型的脚本定位标记相关联的 **Shape** 对象将从 ET 和 WPP 的 **Shapes** 集合以及 WPS 的 **InlineShapes** 和 **Shapes** 集合中删除。

**示例**

``` JavaScript
//本示例删除所有不可见的自定义命令栏。
function test()
{  
    let foundFlag = false 
    let delBars = 0
    let cmdBars = Application.ActiveDocument.CommandBars;
    for (let bar = 1; bar <= cmdBars.Count; bar++) {
        if (cmdBars.Item(bar).BuiltIn == false 
            && cmdBars.Item(bar).Visible == false) {
                
            cmdBars.Item(bar).Delete()
            foundFlag = true 
            delBars = delBars + 1
        }
    }
    if (!foundFlag) {
     alert("No command bars have been deleted.")
    }
    else {
     alert(delBars + " custom bar(s) deleted.")
    }
} 
```

### CommandBar.FindControl

获取一个符合指定条件的 **CommandBarControl** 对象。

**语法**

**express.FindControl ( *Type* , *Id* , *Tag* , *Visible* , *Recursive* )**

\*express是一个代表 CommandBar 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                       |
|-------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------|
| *Type*      | 可选      | Variant  | 控件的类型。                                                                                                                               |
| *Id*        | 可选      | Variant  | 控件的标识符。                                                                                                                             |
| *Tag*       | 可选      | Variant  | 控件的标记值。                                                                                                                             |
| *Visible*   | 可选      | Variant  | 如果为 true，则只能在搜索中包含可见的命令栏控件。默认值为 false。可见的命令栏包括所有可见的工具栏和执行 FindControl 方法时打开的所有菜单。 |
| *Recursive* | 可选      | Variant  | 如果该值为 true，那么将在命令栏及其所有弹出式子工具栏中查找。此参数仅应用于 CommandBar 对象。默认值为 false。                              |

**说明**

如果 **CommandBars** 集合中有两个或者更多的控件符合搜索条件，那么 FindControl 返回找到的第一个控件。如果没有控件符合搜索条件，那么 **FindControl** 返回 Nothing。

### CommandBar.Reset

将内置命令栏重置为其默认配置。

**语法**

**express.Reset ()**

\*express是一个代表 CommandBar 对象的变量。

**说明**

重置一个内置控件将恢复该控件原来对应的动作，并将各属性恢复成初始状态。重置一个内置命令栏将删除其中的自定义控件并恢复其内置控件。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例用 User 值来根据用户级别调整命令栏。如果 User 为“Level 1”，将显示命令栏“Custom”。如果 User 为任何其他值，那么将重置为默认状态并且命令栏“Custom”被禁用。
function test()
{
    let myBar = Application.CommandBars.Item("Custom")
    if (user == "Level 1" )
    {
        myBar.Visible = true
    }
    else 
    {
        Application.CommandBars.Item("Visual Basic").Reset()
        myBar.Enabled = false 
    }
}
```

### CommandBar.ShowPopup

将指定的命令栏作为快捷菜单，在指定坐标或当前光标位置显示。

**语法**

**express.ShowPopup ( *x* , *y* )**

\*express是一个代表 CommandBar 对象的变量。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                                                                       |
|------|-----------|----------|----------------------------------------------------------------------------|
| *x*  | 可选      | Variant  | 快捷菜单所在位置的 x 坐标。如果省略此参数，那么将取当前光标位置的 x 坐标。 |
| *y*  | 可选      | Variant  | 快捷菜单所在位置的 y 坐标。如果省略此参数，那么将取当前光标位置的 y 坐标。 |

**说明**

当菜单左对齐时， **ShowPopup** 方法显示的快捷菜单的左上角位于 (x, y + 1)；当菜单右对齐时，快捷菜单的右上角位于 (x + 1, y + 1)。您可以使用 Windows 函数 **GetSystemMetrics(SM_MENUDROPALIGNMENT)** 检查下拉菜单对齐的系统布置。

当 (x, y) 坐标的屏幕位置导致全部或部分弹出式菜单超出可视屏幕的边界，弹出式菜单将发生平移，使其显示在可视区域中。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

| 注释                                                                      |
|---------------------------------------------------------------------------|
| 如果命令栏的 **Position** 属性没有设置为 **msoBarPopup** ，则此方法失败。 |

**示例**

``` JavaScript
//本示例可实现的功能为：创建一个包含两个控件的快捷菜单。ShowPopup 方法用于显示该快捷菜单。
function test()
{
    let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarPopup, null, false)
    myBar.Controls.Add(Application.Enum.msoControlButton, 3)
    myBar.Controls.Add(Application.Enum.msoControlComboBox)
    myBar.ShowPopup()
}
```

## 成员属性

### CommandBar.AdaptiveMenu

获取一个 **Boolean** 类型的值，该值指定命令栏是否应包含自适应菜单。可读写。

**语法**

**express.AdaptiveMenu**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBar.Application

获取一个 **Application** 对象，代表 **CommandBar** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBar.BuiltIn

如果指定的命令栏是容器应用程序的内置命令栏，则获取 **True** 。如果是自定义命令栏，则返回 **False** 。只读。

**语法**

**express.BuiltIn**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例删除所有未在屏幕上显示的自定义命令栏。
function test()
{
    let foundFlag = false 
    let deletedBars = 0
    for(let i=1;i <= Application.CommandBars.Count;i++) 
    {
        let bar = Application.CommandBars.Item(i)
        if(bar.BuiltIn == false && bar.Visible == false) 
        {
            bar.Delete()
            foundFlag = true 
            deletedBars++
        }
    }
    if(!foundFlag) 
    {
        alert("No command bars have been deleted.")
    }
    else 
    {
        alert(deletedBars + " custom command bar(s) deleted.")
    }
}
```

### CommandBar.Context

获取或设置一个可确定命令栏保存位置的字符串。该字符串由应用程序定义和解释。可读写。

**语法**

**express.Context**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例显示一个消息框，该消息框包含命令栏“Custom”的上下文相关字符串。本示例运行于 WPS 和其他支持 Context 属性的应用程序中。
function test()
{
    let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)
    myBar.Controls.Add(Application.Enum.msoControlButton, 2)
    myBar.Visible = true 
    alert(myBar.Context)
}
```

### CommandBar.Controls

获取一个代表命令栏上的所有控件的 **CommandBarControls** 对象。只读。

**语法**

**express.Controls**

\*express是一个代表 CommandBar 对象的变量。

**示例**

``` JavaScript
/*本示例在命令栏“Custom”中添加一个组合框控件，然后在该列表中添充两项。此外，本示例还设置项的行数，组合框的宽度，并将该组合框的默认值设置为空。*/
function test()
{  
    let myControl = Application.ActiveDocument.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlComboBox, null, null, 1)
    myControl.AddItem("First Item", 1)
    myControl.AddItem("Second Item", 2)
    myControl.DropDownLines = 3
    myControl.DropDownWidth = 75
    myControl.ListHeaderCount = 0
} 
```

### CommandBar.Creator

获取一个 32 位整数，指示创建 **CommandBar** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBar.Enabled

获取或设置用于指定是否启用了指定 **CommandBar** 的 **Boolean** 值。可读写。

**语法**

**express.Enabled**

\*express是一个代表 CommandBar 对象的变量。

**说明**

对于命令栏，如果将该属性设置为 **t** **rue** ，则该命令栏的名称将出现在可用命令栏列表中。

对于内置控件，如果将 **Enabled** 属性设置为 **true** ，则将由应用程序确定其状态。但如果将该属性设置为 **false** ，则强行禁用该控件。

### CommandBar.Height

获取或设置 **CommandBar** 的高度。可读写。

**语法**

**express.Height**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在名为 Custom 的命令栏上添加一个自定义控件。本示例设置该自定义控件的高度为命令栏高度的两倍，并设置该控件宽度为 50 像素。请注意命令栏如何根据控件的大小自我调整。
function test()
{
    let myBar = Application.CommandBars.Item("Custom")
    let barHeight = myBar.Height
    let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("Save").Id, null, null, true)
    myControl.Height = barHeight * 2
    myControl.Width = 50
    myBar.Visible = true
}
```

### CommandBar.Index

获取一个 **Long** 类型的值，该值代表集合中 **CommandBar** 对象的索引号。只读。

**语法**

**express.Index**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBar.Left

设置或获取从对象左边缘算起 **CommandBar** 相对于屏幕的水平距离（以像素为单位）。可读写。

**语法**

**express.Left**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例将名为 Custom 的命令栏从其固定位置沿窗口顶部移动窗口的左边。
function test()
{
    let myBar = Application.CommandBars.Item("Custom")
    myBar.Position = 1
    myBar.RowIndex = 2
    myBar.Left = 0
}
```

### CommandBar.Name

获取内置的 **CommandBar** 对象的名称。只读。

**语法**

**express.Name**

\*express是一个代表 CommandBar 对象的变量。

**说明**

内置命令栏的本地名称显示在标题栏中（在命令栏未固定时）以及可用命令栏列表中（无论该列表显示在容器应用程序的什么位置）。对于内置命令栏， **Name** 属性返回该命令栏的美式英语名称。使用 **NameLocal** 属性可返回其本地化的名称。如果更改了一个自定义命令栏的 **LocalName** 属性值，则该命令栏的 **Name** 值也将发生更改，反之亦然。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例可实现的功能为：在命令栏集合中查找命令栏“Custom”，如果找到，那么本示例显示该命令栏。
function test()
{
    let foundFlag =  false
    for(let i=1;i <= Application.CommandBars.Count;i++)
    {
        let bar = Application.CommandBars.Item(i)
        if(bar.Name == "Custom")
        {
            foundFlag = true
            bar.Visible = true
        }
    }
    if(!foundFlag)
    {
        alert("'Custom' bar isn't in collection.")
    }
    else
    {
        alert("'Custom' bar is now visible.")
    }
}
```

### CommandBar.NameLocal

获取以容器应用程序的语言版本显示的内置命令栏名称，或者返回或设置自定义命令栏的名称。可读写。

**语法**

**express.NameLocal**

\*express是一个代表 CommandBar 对象的变量。

**说明**

无论有效命令栏列表显示在容器应用程序的哪个位置，内置命令栏的本地名都将显示在标题栏中（当该命令栏未固定时），以及有效命令栏列表中。

如果更改了一个自定义命令栏的 **LocalName** 属性值，则该命令栏的 **Name** 值也将发生更改，反之亦然。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

| 注释                                         |
|----------------------------------------------|
| 如果要对内置命令栏设置该属性，就会出现错误。 |

**示例**

``` JavaScript
//本示例可实现的功能为：显示容器应用程序中第一个命令栏的名称和本地化名称。 
function test()
{
    let Com = Application.CommandBars.Item(1)
    alert("The name of the command bar is " + Com.Name)
    alert("The localized name of the command bar is " + Com.NameLocal)
}
```

### CommandBar.Parent

获取 **CommandBar** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBar.Position

获取或设置一个代表命令栏位置的 **MsoBarPosition** 常量。可读写。

**语法**

**express.Position**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例对命令栏集合中的元素逐个进行如下操作：将自定义命令栏停靠在应用程序窗口的底部，将内置命令栏停靠在窗口的顶部。  
function test()
{
    for(let i=1;i <= Application.CommandBars.Count;i++) 
    {
        let bar = Application.CommandBars.Item(i)
        if(bar.Visible == true) 
        {
            if(bar.BuiltIn) {
                bar.Position = Application.Enum.msoBarTop
            }
            else {
                bar.Position = Application.Enum.msoBarBottom
            }
        }
    }
}
```

### CommandBar.RowIndex

获取或设置一个命令栏相对于同一停靠区域中其他命令栏的停靠顺序。该属性值可以是大于零的整数，也可以是以下 **MsoBarRow** 常量之一： **msoBarRowFirst** 或 **msoBarRowLast** 。可读写。

**语法**

**express.RowIndex**

\*express是一个代表 CommandBar 对象的变量。

**说明**

多个命令栏可享有同一行索引号，具有较低数值的命令栏先固定。如果存在两个或两个以上命令栏具有同一行索引号的情况，那么最后指定的命令栏在这一组中最先显示。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例将名为“Custom”的命令栏的位置调整到默认值以左 110 个像素的位置，并通过将其行索引号改为 msoBarRowFirst 使该命令栏最先固定。
function test()
{
    let myBar = Application.CommandBars.Item("Custom")
    myBar.RowIndex = Application.Enum.msoBarRowFirst
    myBar.Left = 140
}
```

### CommandBar.Top

设置或获取指定的命令栏顶边到屏幕顶边的距离。对于固定命令栏，此属性返回或设置从命令栏到停靠区域顶部的距离。可读写。

**语法**

**express.Top**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例将名为“Custom”的浮动命令栏的左上角定位为距离屏幕左边 140 像素、距离屏幕顶部 100 像素。
function test()
{
    let myBar = Application.CommandBars.Item("Custom")
    myBar.Position = Application.Enum.msoBarFloating
    myBar.Left = 140
    myBar.Top = 100
}
```

### CommandBar.Type

获取命令栏的类型。只读。

**语法**

**express.Type**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例查找名为“Custom”的命令栏上的第一个控件，并使用 Type 属性确定该控件是否为按钮。如果该控件为按钮，本示例将复制“Copy”（位于“Standard”工具栏上）按钮的图标并将其粘贴到该控件上。
function test()
{
    let oldCtrl = Application.CommandBars.Item("Custom").Controls.Item(1)
    if(oldCtrl.Type == Application.Enum.msoControlButton) 
    {
        let newCtrl = Application.CommandBars.FindControl(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("Copy").ID)
        newCtrl.CopyFace()
        oldCtrl.PasteFace()
    }
}
```

### CommandBar.Visible

获取或设置命令栏的 **Visible** 属性。如果命令栏可见，则为 **true** 。可读写。

**语法**

**express.Visible**

\*express是一个代表 CommandBar 对象的变量。

**说明**

默认情况下，新创建的自定义命令栏的 **Visible** 属性为 **false** 。

必须先将命令栏的 **Enabled** 属性设置为 **true** ，才能将其 **Visible** 属性设置为 **true** 。

**示例**

``` JavaScript
/*
本示例将遍历命令栏集合，以查找“Forms”命令栏。如果找到“Forms”命令栏，则显示该命令栏并保护其停靠状态。
*/

function test()
{  
    let foundFlag = false 
    for(let i=1;i <= Application.CommandBars.Count;i++) {
    let cmdbar = Application.CommandBars.Item(i)
    if(cmdbar.Name == "Forms") {
        cmdbar.Protection = msoBarNoChangeDock
        cmdbar.Visible = true 
        foundFlag = true 
    }
    }
    if(!foundFlag) {
        alert("'Forms'command bar is not in the collection.")
    }
} 
```

### CommandBar.Width

获取或设置指定命令栏的宽度（以像素为单位）。可读写。

**语法**

**express.Width**

\*express是一个代表 CommandBar 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在名为 Custom 的命令栏上添加一个自定义控件。本示例设置该自定义控件的高度为命令栏高度的两倍，并设置宽度为 50 像素。请注意命令栏如何根据控件的大小自我调整。
function test()
{
    let myBar = Application.CommandBars.Item("Custom")
    let barHeight = myBar.Height
    let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("Save").Id, null, null, true)
    myControl.Height = barHeight * 2
    myControl.Width = 50
    myBar.Visible = true
}
```

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CommandBar](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CommandBarButton 对象

## 简介

代表命令栏中的一个按钮控件。

## 说明

使用 Controls(index) 可返回一个 CommandBarButton 对象，其中 *index* 是控件的索引号。请注意，该控件的 Type 属性必须是 msoControlButton 。假定在名为“Custom”的命令栏上，第二个控件是一个按钮，则下面的示例将更改该按钮的样式。

``` JavaScript
function test()
{
    let cmdBars = Application.ActiveDocument.CommandBars
    let c = cmdBars.Item("Custom").Controls.Item(2)
    if ( c.Type == Application.Enum.msoControlButton ) {
        if ( c.Style == Application.Enum.msoButtonIcon ) {
            c.Style = Application.Enum.msoButtonIconAndCaption
        }
        else {
            c.Style = Application.Enum.msoButtonIcon
        }
    }
}
```

## 方法

| 序号 | 名称                                   | 说明                                                                                       |
|:----:|:---------------------------------------|:-------------------------------------------------------------------------------------------|
|  1   | [Copy](#CommandBarButton.Copy)         | 将一个命令栏按钮控件复制到已有的命令栏中。                                                 |
|  2   | [CopyFace](#CommandBarButton.CopyFace) | 将命令栏按钮控件的图标复制到“剪贴板”中。                                                   |
|  3   | [Delete](#CommandBarButton.Delete)     | 将 CommandBarButton 对象从其集合中删除。                                                   |
|  4   | [Execute](#CommandBarButton.Execute)   | 运行分配给指定 CommandBarButton 控件的过程或内置命令。                                     |
|  5   | [Move](#CommandBarButton.Move)         | 将指定的 CommandBarButton 控件移动到已有的命令栏。                                         |
|  6   | [Reset](#CommandBarButton.Reset)       | 将内置 CommandBarButton 控件重置为其初始功能和图符。                                       |
|  7   | [SetFocus](#CommandBarButton.SetFocus) | 将键盘的焦点移到指定的 CommandBarButton 控件。如果该按钮被禁用或不可见，那么此方法将失败。 |

## 属性

| 序号 | 名称                                                     | 说明                                                                                                                                                                                                   |
|:----:|:---------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CommandBarButton.Application)             | 获取一个 Application 对象，代表 CommandBarButton 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。                                                                |
|  2   | [BeginGroup](#CommandBarButton.BeginGroup)               | 如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 True。可读写。                                                                                                                                |
|  3   | [BuiltIn](#CommandBarButton.BuiltIn)                     | 如果指定的命令栏控件是容器应用程序的控件，则为 True 。如果是自定义控件或是已设置了 OnAction 属性的内置控件，则返回 False 。只读。                                                                      |
|  4   | [BuiltInFace](#CommandBarButton.BuiltInFace)             | 如果命令栏按钮控件的图标是其原始内置图标，则为 True 。可读写。                                                                                                                                         |
|  5   | [Caption](#CommandBarButton.Caption)                     | 获取或设置命令栏控件的标题文字。可读写。                                                                                                                                                               |
|  6   | [Creator](#CommandBarButton.Creator)                     | 获取一个 32 位整数，指示创建 CommandBarButton 对象时所使用的应用程序。只读。                                                                                                                           |
|  7   | [DescriptionText](#CommandBarButton.DescriptionText)     | 获取或设置命令栏按钮控件的说明。可读写。                                                                                                                                                               |
|  8   | [Enabled](#CommandBarButton.Enabled)                     | 如果启用了指定的 CommandBar 或 CommandBarControl ，则此属性为 true 。可读写。                                                                                                                          |
|  9   | [FaceId](#CommandBarButton.FaceId)                       | 获取或设置某个 CommandBarButton 控件的图符的 ID 号。可读写。                                                                                                                                           |
|  10  | [Height](#CommandBarButton.Height)                       | 获取或设置命令栏控件的高度。可读写。                                                                                                                                                                   |
|  11  | [HelpContextId](#CommandBarButton.HelpContextId)         | 获取或设置附加于 CommandBarButton 控件的“帮助”主题的帮助上下文 ID 号。可读写。                                                                                                                         |
|  12  | [HelpFile](#CommandBarButton.HelpFile)                   | 获取或设置附加于 CommandBarButton 控件的“帮助”主题的文件名。可读写                                                                                                                                     |
|  13  | [HyperlinkType](#CommandBarButton.HyperlinkType)         | 设置或获取一个 MsoCommandBarButtonHyperlinkType 常量，代表与指定的命令栏按钮相关联的超链接的类型。可读写。                                                                                             |
|  14  | [Id](#CommandBarButton.Id)                               | 获取内置的 CommandBarButton 控件的 ID。只读。                                                                                                                                                          |
|  15  | [Index](#CommandBarButton.Index)                         | 获取一个 Long 类型的值，该值代表集合中 CommandBarButton 对象的索引号。只读。                                                                                                                           |
|  16  | [IsPriorityDropped](#CommandBarButton.IsPriorityDropped) | 如果当前已根据使用统计信息和布局空间将 CommandBarButton 控件从菜单或工具栏中删除，则获取 True 。（请注意，这与通过 Visible 属性设置的控件可见性不同）。只读。                                          |
|  17  | [Left](#CommandBarButton.Left)                           | 设置或获取指定的 CommandBarButton 控件相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。                                                                                   |
|  18  | [Mask](#CommandBarButton.Mask)                           | 获取或设置代表 CommandBarButton 对象的屏蔽图像的 IPictureDisp 对象。屏蔽图像可决定按钮图像的透明部分。可读写。                                                                                         |
|  19  | [OLEUsage](#CommandBarButton.OLEUsage)                   | 获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 CommandBarButton 控件。可读写。                                                                      |
|  20  | [OnAction](#CommandBarButton.OnAction)                   | 获取或设置一个 JS函数的名称，该过程在用户单击或更改 CommandBarButton 控件的值时运行。可读写。                                                                                                          |
|  21  | [Parameter](#CommandBarButton.Parameter)                 | 获取或设置一个应用程序可用于执行 CommandBarButton 控件中命令的字符串。可读写。                                                                                                                         |
|  22  | [Parent](#CommandBarButton.Parent)                       | 获取 CommandBarButton 对象的 Parent 对象。只读。                                                                                                                                                       |
|  23  | [Picture](#CommandBarButton.Picture)                     | 获取或设置一个代表 CommandBarButton 对象的图像的 IPictureDisp 对象。可读写。                                                                                                                           |
|  24  | [Priority](#CommandBarButton.Priority)                   | 获取或设置 CommandBarButton 控件的优先级。对于固定命令栏，当一行中无法容纳其中的所有控件时，控件的优先级将决定该控件是否会从命令栏中删除。无法容纳在一行中的控件将在命令栏上从右向左进行删减。可读写。 |
|  25  | [ShortcutText](#CommandBarButton.ShortcutText)           | 获取或设置当控钮出现在菜单、子菜单或快捷菜单上时 CommandBarButton 控件旁边显示的快捷键文字。可读写。                                                                                                   |
|  26  | [State](#CommandBarButton.State)                         | 获取或设置 CommandBarButton 控件的外观。可读写。                                                                                                                                                       |
|  27  | [Style](#CommandBarButton.Style)                         | 获取或设置 CommandBarButton 控件的显示方式。可读写。                                                                                                                                                   |
|  28  | [Tag](#CommandBarButton.Tag)                             | 获取或设置有关 CommandBarButton 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。                                                                                                 |
|  29  | [TooltipText](#CommandBarButton.TooltipText)             | 获取或设置 CommandBarButton 的 “屏幕提示” 中显示的文本。可读写。                                                                                                                                       |
|  30  | [Top](#CommandBarButton.Top)                             | 获取指定的 CommandBarButton 控件顶边到屏幕顶边的距离（以像素为单位）。只读。                                                                                                                           |
|  31  | [Type](#CommandBarButton.Type)                           | 获取 CommandBarButton 控件的类型。只读。                                                                                                                                                               |
|  32  | [Visible](#CommandBarButton.Visible)                     | 获取或设置 CommandBarButton 控件的 Visible 属性。如果 CommandBarButton 可见，则此属性为 true 。可读写。                                                                                                |
|  33  | [Width](#CommandBarButton.Width)                         | 获取或设置指定的 CommandBarButton 控件的宽度（以像素为单位）。可读写。                                                                                                                                 |

## 成员方法

### CommandBarButton.Copy

将一个命令栏按钮控件复制到已有的命令栏中。

**语法**

**express.Copy ( *Bar* , *Before* )**

\*express是一个代表 CommandBarButton 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                                                   |
|----------|-----------|----------|------------------------------------------------------------------------------------------------------------------------|
| *Bar*    | 可选      | Variant  | 一个代表目标命令栏的 CommandBar 对象。如果省略此参数，控件将被复制到它原来所在的命令栏中。                             |
| *Before* | 可选      | Variant  | 一个指示新控件在命令栏上的位置的数值。新控件将插入到位于此位置的控件之前。如果省略此参数，控件将被复制到命令栏的末尾。 |

### CommandBarButton.CopyFace

将命令栏按钮控件的图标复制到“剪贴板”中。

**语法**

**express.CopyFace ()**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

可使用 **PasteFace** 方法将“剪贴板”上的内容粘贴到一个按钮图标上。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

### CommandBarButton.Delete

将 **CommandBarButton** 对象从其集合中删除。

**语法**

**express.Delete ( *Temporary* )**

\*express是一个代表 CommandBarButton 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                          |
|-------------|-----------|----------|-------------------------------------------------------------------------------|
| *Temporary* | 可选      | Variant  | 如果为 true，则从当前会话中删除此控件。应用程序在下次会话中将再次显示该控件。 |

### CommandBarButton.Execute

运行分配给指定 **CommandBarButton** 控件的过程或内置命令。

**语法**

**express.Execute ()**

\*express是一个代表 CommandBarButton 对象的变量。

**示例**

``` JavaScript
/*
本 ET 示例创建一个命令栏，然后向其中添加内置命令栏按钮控件。该按钮将执行 ET 标准命令格的第一个函数函数。
```

### CommandBarButton.Move

将指定的 **CommandBarButton** 控件移动到已有的命令栏。

**语法**

**express.Move ( *Bar* , *Before* )**

\*express是一个代表 CommandBarButton 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                          |
|----------|-----------|----------|-----------------------------------------------------------------------------------------------|
| *Bar*    | 可选      | Variant  | 一个代表控件的目标命令栏的 Command 对象。如果忽略该参数，则控件将移动到当前所在命令栏的末端。 |
| *Before* | 可选      | Variant  | 表示控件位置的数字。控件将插到该位置的控件之前。如果忽略该参数，控件插入到同一命令栏。        |

**示例**

``` JavaScript
/*
本示例将名为 Custom 的命令栏上的第一个组合框控件移动到该命令栏上的第七个控件之前。本示例将 Tag 设置为“Selection box”并赋予控件较低的优先级，以便在一行容纳不下所有控件时将其隐藏。
*/
function test()
{
    let allControls = Application.ActiveDocument.CommandBars.Item("Custom").Controls
    for(let i=1; i<=allControls.Count; i++) {
    if(allControls.Item(i).Type == Application.Enum.msoControlComboBox) {
        allControls.Item(i).Move(null,7)
        allControls.Item(i).Tag = "Selection box"
        allControls.Item(i).Priority = 5
    }
}
}
```

### CommandBarButton.Reset

将内置 **CommandBarButton** 控件重置为其初始功能和图符。

**语法**

**express.Reset ()**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

重置一个内置控件将恢复该控件原来对应的动作，并将各属性恢复成初始状态。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例自定义一个命令栏按钮。首先，将按钮属性重置为其默认状态。然后设置各种按钮属性。
function test()
{
    let cbButton = Application.CommandBars.Item("Custom").Controls.Item(2)
    cbButton.Reset()
    cbButton.BuiltInFace = true 
    cbButton.Caption = "Compute Total"
    cbButton.DescriptionText = "This button computes the total of all purchases."
    cbButton.Enabled = true 
    cbButton.TooltipText = "Click to compute total amount for all items in your cart."
}
```

### CommandBarButton.SetFocus

将键盘的焦点移到指定的 **CommandBarButton** 控件。如果该按钮被禁用或不可见，那么此方法将失败。

**语法**

**express.SetFocus ()**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

带有焦点的控件与其他控件的差别是十分细微的。在使用此方法后，会看到该控件处于三维突出显示的状态。按方向键将使焦点在工具栏中各控件间切换，就好像是按键盘控制键到达该控件一样。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例创建一个名为“Custom”的命令栏，并在其中添加一个 ComboBox 控件和一个 Button 控件，然后使用 SetFocus 方法为 ComboBox 控件设置焦点。
function test()
{
    let focusBar = Application.CommandBars.Add("Custom")
    Application.CommandBars.Item("Custom").Visible = true
    Application.CommandBars.Item("Custom").Position = Application.Enum.msoBarTop
                                        
    let testComboBox = Application.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlComboBox, 1)
                                        
    let testButton = Application.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlButton)
    testButton.FaceId = 17
    // Set the focus to the combo box.
    testComboBox.SetFocus()
}
```

## 成员属性

### CommandBarButton.Application

获取一个 **Application** 对象，代表 **CommandBarButton** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.BeginGroup

如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 True。可读写。

**语法**

**express.BeginGroup**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.BuiltIn

如果指定的命令栏控件是容器应用程序的控件，则为 **True** 。如果是自定义控件或是已设置了 **OnAction** 属性的内置控件，则返回 **False** 。只读。

**语法**

**express.BuiltIn**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.BuiltInFace

如果命令栏按钮控件的图标是其原始内置图标，则为 **True** 。可读写。

**语法**

**express.BuiltInFace**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

此属性只能设置为 **True** ，其功能是将图标重新设置为内置的图标。可读写 **Boolean** 类型。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例可实现的功能为：确定命令栏“Custom”中第一个控件的图标是否为其内置的按钮图标。如果是，那么本示例将该按钮图标复制到“剪贴板”上。
function test()
{
    let myControl = Application.CommandBars.Item("Custom").Controls.Item(1)
    if ( myControl.BuiltInFace == true)
    {
        myControl.CopyFace()
    }
}
```

### CommandBarButton.Caption

获取或设置命令栏控件的标题文字。可读写。

**语法**

**express.Caption**

\*express是一个代表 CommandBarButton 对象的变量。

**示例**

``` JavaScript
/*本示例在自定义命令栏中添加一个带拼写检查按钮图标的命令栏控件，然后将其标题设置为“Spelling checker”。*/

function test()
{
    let myBar = Application.ActiveDocument.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)
    myBar.Visible = true 
    let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, 2)
    myControl.DescriptionText = "Starts the spelling checker"
    myControl.Caption = "Spelling checker"
}
```

### CommandBarButton.Creator

获取一个 32 位整数，指示创建 **CommandBarButton** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.DescriptionText

获取或设置命令栏按钮控件的说明。可读写。

**语法**

**express.DescriptionText**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

该说明不显示给用户，但有助于其他开发者将控件行为归档。

### CommandBarButton.Enabled

如果启用了指定的 **CommandBar** 或 **CommandBarControl** ，则此属性为 true 。可读写。

**语法**

**express.Enabled**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

对于命令栏，如果将该属性设置为 **True** ，则该命令栏的名称将出现在可用命令栏列表中。

对于内置控件，如果将 **Enabled** 属性设置为 **True** ，则将由应用程序确定其状态。但如果将该属性设置为 **False** ，则强行禁用该控件。

### CommandBarButton.FaceId

获取或设置某个 **CommandBarButton** 控件的图符的 ID 号。可读写。

**语法**

**express.FaceId**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

**FaceId** 属性确定一个命令栏按钮的外观而非其功能。 **CommandBarControl** 对象的 **Id** 属性确定按钮的功能。

对于具有自定义图标的命令栏按钮，其 **FaceId** 属性值为 0（零）。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例向自定义命令栏中添加一个命令栏按钮。虽然此按钮的图标与内置“制作图表”按钮的图标相同，但单击此按钮相当于单击“文件”菜单上的“打开”命令，这是因为此按钮的 ID 号为 23。
function test()
{
    let newBar = Application.CommandBars.Add("Custom2", Application.Enum.msoBarTop, null, true)
    newBar.Visible = true 
    let con = newBar.Controls.Add(Application.Enum.msoControlButton, 23)
    con.FaceId = 17
}
```

### CommandBarButton.Height

获取或设置命令栏控件的高度。可读写。

**语法**

**express.Height**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在名为 Custom 的命令栏上添加一个自定义控件。本示例设置该自定义控件的高度为命令栏高度的两倍，并设置该控件宽度为 50 像素。请注意命令栏如何根据控件的大小自我调整。
function test()
{
    let myBar = Application.CommandBars.Item("Custom")
    let barHeight = myBar.Height + 1
    let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("Save").Id, null, null, true)
    myControl.Height = barHeight * 2
    myControl.Width = 50
    myBar.Visible = true
}
```

### CommandBarButton.HelpContextId

获取或设置附加于 **CommandBarButton** 控件的“帮助”主题的帮助上下文 ID 号。可读写。

**语法**

**express.HelpContextId**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

若要使用此属性，必须同时设置 HelpFile 属性。帮助主题响应 Shift+F1 按键操作。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.HelpFile

获取或设置附加于 **CommandBarButton** 控件的“帮助”主题的文件名。可读写

**语法**

**express.HelpFile**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

若要使用该属性，必须同时设置 HelpContextID 属性。帮助主题响应 Shift+F1 用户按键操作。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.HyperlinkType

设置或获取一个 **MsoCommandBarButtonHyperlinkType** 常量，代表与指定的命令栏按钮相关联的超链接的类型。可读写。

**语法**

**express.HyperlinkType**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例检查名为“Custom”的命令栏上的指定命令栏按钮的 HyperlinkType 属性。如果 HyperlinkType 设置为 msoCommandBarButtonHyperlinkNone，则本示例会将该属性设置为 msoCommandBarButtonHyperlinkOpen，并将 URL 设置为 www.wps.cn。
function test()
{
    let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)
    let myButton = myBar.Controls.Add(Application.Enum.msoControlButton)
    myButton.FaceId = 277
    myButton.HyperlinkType = Application.Enum.msoCommandBarButtonHyperlinkNone
    if ( myButton.HyperlinkType > Application.Enum.msoCommandBarButtonHyperlinkOpen )     {
        myButton.HyperlinkType = Application.Enum.msoCommandBarButtonHyperlinkOpen
        myButton.TooltipText = "www.wps.cn"
    }
}
```

### CommandBarButton.Id

获取内置的 **CommandBarButton** 控件的 ID。只读。

**语法**

**express.Id**

\*express是一个代表 CommandBarButton 对象的变量。

**示例**

``` JavaScript
/*
在本示例中，如果名为“Custom2”的命令栏上第一个控件的按钮的 ID 值小于 25，则更改其按钮图符。
*/
function test()
{
    let ctrl = Application.ActiveDocument.CommandBars.Item("Custom").Controls.Item(1)
    if ( ctrl.Id < 25 ) {
        ctrl.FaceId = 17
        ctrl.Tag = "Changed control"
    }
}

/*下面的示例将名为“Standard”的工具栏上每个控件的标题更改为该控件当前的 Id 属性值。*/

function test()
{
    let ctl = Application.ActiveDocument.CommandBars.Item("Standard").Controls
    for (let i = 1; i <= ctl.Count; i++) {
        ctl.Item(i).Caption = String(ctl.Id)
    }
}
```

### CommandBarButton.Index

获取一个 **Long** 类型的值，该值代表集合中 **CommandBarButton** 对象的索引号。只读。

**语法**

**express.Index**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

第一个命令栏控件的位置为 1。 **CommandBarControls** 集合中的分隔符不计算在内。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在名为“Custom2”的命令栏中搜索 ID 值为 23 的控件。如果能找到这样的控件，并且其索引号大于 5，则将该控件放置在命令栏上的第一个位置上。
function test()
{
    let myBar = Application.CommandBars.Item("Custom2")
    let ctrl1 = myBar.FindControl(null,  23) 
    if( ctrl1.Index >  5 )
    {
        ctrl1.Move(null, 1)
    }
}
```

### CommandBarButton.IsPriorityDropped

如果当前已根据使用统计信息和布局空间将 **CommandBarButton** 控件从菜单或工具栏中删除，则获取 **True** 。（请注意，这与通过 Visible 属性设置的控件可见性不同）。只读。

**语法**

**express.IsPriorityDropped**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

如果 **IsPriorityDropped** 为 **True** ，则 **Visible** 属性设置为 **True** 的控件将不能在个性化菜单或工具栏中立即显示出来。

为了确定何时将某一特定菜单项的 **IsPriorityDropped** 属性设置为 **True** ，WPS Office会记录下该菜单项的使用次数的总和，并记录下用户使用了同一菜单中的其他菜单项而没有使用该特定菜单项的应用程序会话数。当该会话数达到一定的阈值之后，便会减少使用次数总和。当总次数减为零时，便将 **IsPriorityDropped** 设置为 **True** 。编程人员不能设置会话值和阈值，也不能设置 **IsPriorityDropped** 属性。但他们可以使用 **AdaptiveMenus** 属性来禁用应用程序中特定菜单的自适应菜单。

为了确定何时将某一特定工具栏控件的 **IsPriorityDropped** 属性设置为 **True** ，Office 维护了一个列表，记录下了该工具栏中各个控件的最后使用顺序。工具栏会在空间允许的前提下尽可能多地显示控件，控件按使用时间由近向远排列，最近使用过的控件显示在最前面。 **Priority** 设置为 1 的控件将始终显示，而且为了显示这样的控件，在必要时工具栏还将换行。编程人员可使用 **Priority** 属性来确保始终显示特定的工具栏控件，或重新定位工具栏以便使其有足够的空间显示所有的控件。

使用下表可以预测在菜单项的 **IsPriorityDropped** 属性被设置为 **True** 之前，它仍将出现在个性化菜单中的会话数。

| 命令栏控件的使用次数 | 应用程序的会话数 |
|----------------------|------------------|
| 0, 1                 | 3                |
| 2                    | 6                |
| 3                    | 9                |
| 4, 5                 | 12               |
| 6- 8                 | 17               |
| 9-13                 | 23               |
| 14-24                | 29               |
| 25 及 25 以上        | 31               |

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.Left

设置或获取指定的 **CommandBarButton** 控件相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。

**语法**

**express.Left**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.Mask

获取或设置代表 **CommandBarButton** 对象的屏蔽图像的 **IPictureDisp** 对象。屏蔽图像可决定按钮图像的透明部分。可读写。

**语法**

**express.Mask**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

在创建要作为屏蔽图像使用的图像时，所有要透明的区域应该为白色，所有要显示的区域应该为黑色。

通常在设置了 **CommandBarButton** 对象的图片后设置屏蔽。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

### CommandBarButton.OLEUsage

获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 **CommandBarButton** 控件。可读写。

**语法**

**express.OLEUsage**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

本属性允许用户在某 Office 应用程序与另一个 Office 应用程序合并时，指定加载应用程序的各个命令栏控件在该 Office 应用程序中的表示方法。如果客户端和服务器均提供命令栏，那么命令栏控件将逐个嵌入客户端。系统将从服务器中略去标记为只用于客户端（或既不用于客户端也不用于服务器）的自定义控件，而从客户端中略去标记为只用于服务器（或既不用于客户端也不用于服务器）的控件。剩下的控件再进行合并。

如果要合并的应用程序之一不是 Office 应用程序，则使用常规 OLE 菜单合并，该过程将由 **OLEMenuGroup** 属性控制。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例向名为“Tools”的命令栏中添加一个新按钮并设置其 OLEUsage 属性。
function test()
{
    let myControl = Application.CommandBars.Item("Tools").Controls.Add(Application.Enum.msoControlButton, null, null, null, true)
    myControl.OLEUsage = Application.Enum.msoControlOLEUsageNeither
}
```

### CommandBarButton.OnAction

获取或设置一个 JS函数的名称，该过程在用户单击或更改 **CommandBarButton** 控件的值时运行。可读写。

**语法**

**express.OnAction**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

容器应用程序将确定该值是否是一个合法的宏名。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在名为“Custom”的命令栏中添加一个命令栏控件。COM 加载项“FinanceAddIn”将在每次单击该控件时运行。
function test()
{
    let myBar = Application.CommandBars.Item("Custom")
    let myControl = myBar.Controls.Add(Application.Enum.msoControlButton)
    myControl.FaceId = 2
    myControl.OnAction = "!<FinanceAddIn>"
    myBar.Visible = true
}
```

### CommandBarButton.Parameter

获取或设置一个应用程序可用于执行 **CommandBarButton** 控件中命令的字符串。可读写。

**语法**

**express.Parameter**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

如果为内置控件设置一个指定的参数，那么应用程序将修改其默认的操作，前提是该应用程序能分析和使用这个新值。如果为自定义控件设置参数，那么该参数可用来为 Visual Basic 程序传送信息，或保存关于该控件的信息（类似于另一个 Tag 属性值）。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例将一个新参数赋予控件并将焦点置于新按钮上。
function test()
{
    let myControl = Application.CommandBars.Item("Custom").Controls.Item(4)
    myControl.Copy(null, 1)
    myControl.Parameter = "2"
    myControl.SetFocus()
}
```

### CommandBarButton.Parent

获取 **CommandBarButton** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CommandBarButton 对象的变量。

### CommandBarButton.Picture

获取或设置一个代表 **CommandBarButton** 对象的图像的 **IPictureDisp** 对象。可读写。

**语法**

**express.Picture**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

在更改按钮图像时，还希望使用 **Mask** 属性设置屏蔽图像。屏蔽图像决定按钮图像透明的部分。通常在设置了 **CommandBarButton** 对象的图片后设置屏蔽。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

### CommandBarButton.Priority

获取或设置 CommandBarButton 控件的优先级。对于固定命令栏，当一行中无法容纳其中的所有控件时，控件的优先级将决定该控件是否会从命令栏中删除。无法容纳在一行中的控件将在命令栏上从右向左进行删减。可读写。

**语法**

**express.Priority**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

有效的优先级数为从 0 至 7 之间的数字，默认值为 3。优先级为 1 表示不能从工具栏中删除该控件。而优先级为其他值的控件将被忽略。

菜单项的命令栏控件不能使用 **Priority** 属性。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.ShortcutText

获取或设置当控钮出现在菜单、子菜单或快捷菜单上时 **CommandBarButton** 控件旁边显示的快捷键文字。可读写。

**语法**

**express.ShortcutText**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

只能为包含 **OnAction** 宏的命令栏按钮设置此属性。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在消息框中显示 ET 工作表菜单栏上“打开”命令（“文件”菜单）的快捷键文字。
alert (CommandBars.Item("Worksheet Menu Bar").Controls.Item("File").Controls.Item("New...").ShortcutText)
```

### CommandBarButton.State

获取或设置 CommandBarButton 控件的外观。可读写。

**语法**

**express.State**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

内置命令栏按钮的 **State** 属性是只读的。 **Type** 属性的值可作为 **MsoButtonState** 枚举中的值之一。

| 注释                                                   |
|--------------------------------------------------------|
| WPS Office Assistant 在 WPS Office 2015 版中已被淘汰。 |

**示例**

``` JavaScript
//本示例创建一个名为“Custom”的命令栏并添加两个按钮，然后将左边的按钮设置为 msoButtonUp，并将右边的按钮设置为 msoButtonDown。
function test()
{
    // Add new command bar.
    let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)
    // Add 2 buttons to new command bar.
    myBar.Controls.Add(Application.Enum.msoControlButton)
    myBar.Controls.Add(Application.Enum.msoControlButton)
    myBar.Visible = true 
    
    // Paste Bold button face and set State of first button.
    let myControl1 = myBar.Controls.Item(1)
    let imgSource = Application.CommandBars.FindControl(Application.Enum.msoControlButton, 113)
    myControl1.State = Application.Enum.msoButtonUp
    
    // Paste italic button face and set State of second button.
    let myControl2 = myBar.Controls.Item(2)
    imgSource = Application.CommandBars.FindControl(Application.Enum.msoControlButton, 114)
    myControl2.State = Application.Enum.msoButtonDown
}
```

### CommandBarButton.Style

获取或设置 **CommandBarButton** 控件的显示方式。可读写。

**语法**

**express.Style**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

**Style** 属性的值可作为 **MsoButtonStyle** 枚举中的值之一。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.Tag

获取或设置有关 **CommandBarButton** 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。

**语法**

**express.Tag**

\*express是一个代表 CommandBarButton 对象的变量。

**示例**

``` JavaScript
/*
本示例可实现的功能为：将自定义工具栏上的按钮的标记设置为“Spelling Button”，并在消息框中显示此标记。
*/
function test() {
    Application.ActiveDocument.CommandBars.Item("Custom").Controls.Item(1).Tag = "Spelling Button"
    alert(Application.ActiveDocument.CommandBars.Item("Custom").Controls.Item(1).Tag)
}

```

### CommandBarButton.TooltipText

获取或设置 **CommandBarButton** 的 **“屏幕提示”** 中显示的文本。可读写。

**语法**

**express.TooltipText**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

默认情况下， **Caption** 属性的值将用作 **“屏幕提示”** 。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.Top

获取指定的 **CommandBarButton** 控件顶边到屏幕顶边的距离（以像素为单位）。只读。

**语法**

**express.Top**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.Type

获取 **CommandBarButton** 控件的类型。只读。

**语法**

**express.Type**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarButton.Visible

获取或设置 **CommandBarButton** 控件的 **Visible** 属性。如果 **CommandBarButton** 可见，则此属性为 **true** 。可读写。

**语法**

**express.Visible**

\*express是一个代表 CommandBarButton 对象的变量。

### CommandBarButton.Width

获取或设置指定的 **CommandBarButton** 控件的宽度（以像素为单位）。可读写。

**语法**

**express.Width**

\*express是一个代表 CommandBarButton 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CommandBarButton](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FileDialogSelectedItems 对象

## 简介

String 值的集合，这些值对应于通过 FileDialog 对象显示的文件对话框中用户所选文件或文件夹路径。

## 说明

示例

``` JavaScript
function test() {
    //Declare a variable as a FileDialog object.

    //Create a FileDialog object as a File Picker dialog box.
    let fd = Application.FileDialog(Application.Enum.msoFileDialogFilePicker)

    //Declare a variable to contain the path
    //of each selected item. Even though the path is aString,
    //the variable must be a Variant because For Each...Next
    //routines only work with Variants and Objects.

    //Use a With...End With block to reference the FileDialog object.
            
         //Allow the selection of multiple file. 
        fd.AllowMultiSelect = true 

        //Use the Show method to display the File Picker dialog box and return the user's action.
        //The user pressed the button.
        if(fd.Show() == -1) {

            //Step through each string in the FileDialogSelectedItems collection
            for(let i = 1;i <= fd.SelectedItems.Count;i++) {
                let vrtSelectedItem = fd.SelectedItems.Item(i)
                //vrtSelectedItem is aString that contains the path of each selected item.
                //You can use any file I/O functions that you want to work with this path.
                //This example displays the path in a message box.
                alert("Selected item's path: " + vrtSelectedItem)

            }
        }
        //The user pressed Cancel.
        else { }
        

    //Set the object variable to Nothing.
    fd = null

}
```

使用 FileDialog 对象的 SelectedItems 属性可返回一个 FileDialogSelectedItems 集合。下面的示例显示一个 “File Picker” 对话框，并在消息框中显示每个选定的文件。

## 方法

| 序号 | 名称                                  | 说明                                                                                                             |
|:----:|:--------------------------------------|:-----------------------------------------------------------------------------------------------------------------|
|  1   | [Item](#FileDialogSelectedItems.Item) | 获取一个 String 类型的值，对应于用户在使用 FileDialog 对象的 Show 方法显示的文件对话框中所选择的文件之一的路径。 |

## 属性

| 序号 | 名称                                                | 说明                                                                                                                                           |
|:----:|:----------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FileDialogSelectedItems.Application) | 获取一个 Application 对象，代表 FileDialogSelectedItems 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Count](#FileDialogSelectedItems.Count)             | 获取一个 Long 类型的值，指示 FileDialogSelectedItem s 集合中的项数。只读。                                                                     |
|  3   | [Creator](#FileDialogSelectedItems.Creator)         | 获取一个 32 位整数，指示创建 FileDialogSelectedItems 对象时所使用的应用程序。只读。                                                            |
|  4   | [Parent](#FileDialogSelectedItems.Parent)           | 获取 FileDialogSelectedItems 对象的 Parent 对象。只读。                                                                                        |

## 成员方法

### FileDialogSelectedItems.Item

获取一个 **String** 类型的值，对应于用户在使用 **FileDialog** 对象的 **Show** 方法显示的文件对话框中所选择的文件之一的路径。

**语法**

**express.Item ()**

\*express是一个代表 FileDialogSelectedItems 对象的变量。

## 成员属性

### FileDialogSelectedItems.Application

获取一个 **Application** 对象，代表 **FileDialogSelectedItems** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 FileDialogSelectedItems 对象的变量。

### FileDialogSelectedItems.Count

获取一个 **Long** 类型的值，指示 **FileDialogSelectedItem** s 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 FileDialogSelectedItems 对象的变量。

**示例**

下面的示例显示一个“File Picker”对话框，并在消息框中显示每个选定的文件。

``` JavaScript
function test() {

    //Declare a variable as a FileDialog object.

    //Create a FileDialog object as a File Picker dialog box.
    let fd = Application.FileDialog(Application.Enum.msoFileDialogFilePicker)

    //Declare a variable to contain the path
    //of each selected item. Even though the path is aString,
    //the variable must be a Variant because For Each...Next
    //routines only work with Variants and Objects.

    //Use a With...End With block to reference the FileDialog object.
            
        //Allow the selection of multiple file.
        fd.AllowMultiSelect = true 

        //Use the Show method to display the File Picker dialog box and return the user's action.
        //The user pressed the button.
        if(fd.Show() == -1) {

            //Step through each string in the FileDialogSelectedItems collection
            for(let cnt = 1;cnt <= fd.SelectedItems.Count;cnt++) {
                let vrtSelectedItem = fd.SelectedItems
                //vrtSelectedItem is aString that contains the path of each selected item.
                //You can use any file I/O functions that you want to work with this path.
                //This example displays the path in a message box.
                alert("Selected item's path: " + vrtSelectedItem.Item(cnt))

            }
        }
        //The user pressed Cancel.
        else { }


    //Set the object variable to Nothing.
    fd = null

}
```

### FileDialogSelectedItems.Creator

获取一个 32 位整数，指示创建 **FileDialogSelectedItems** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 FileDialogSelectedItems 对象的变量。

### FileDialogSelectedItems.Parent

获取 **FileDialogSelectedItems** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FileDialogSelectedItems 对象的变量。

**类型**

Object

> WPS 开发文档： [WPS 基础接口/通用 API 参考/FileDialogSelectedItems](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PictureEffects 对象

## 简介

代表 PictureEffects 对象的集合。

## 说明

``` JavaScript
//下面的代码将设置 Wpp 幻灯片中某个形状的几个图片效果填充属性。
function test()
{
    // Setup a slide with one picture shape.
    let pes = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects
    // Insert a 150% Saturation effect.
    pes.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5
    // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.
    let brightnessContrast = pes.Insert(Application.Enum.msoEffectBrightnessContrast)
    brightnessContrast.EffectParameters.Item(1).Value = -0.5
    brightnessContrast.EffectParameters.Item(2).Value = 0.25
}
```

## 方法

| 序号 | 名称                             | 说明                              |
|:----:|:---------------------------------|:----------------------------------|
|  1   | [Delete](#PictureEffects.Delete) | 从集合中删除 PictureEffect 对象。 |
|  2   | [Insert](#PictureEffects.Insert) | 将图片效果插入到复合效果链中。    |

## 属性

| 序号 | 名称                                       | 说明                                                                           |
|:----:|:-------------------------------------------|:-------------------------------------------------------------------------------|
|  1   | [Application](#PictureEffects.Application) | 获取一个 Application 对象，该对象代表 PictureEffects 对象的容器应用程序。只读  |
|  2   | [Count](#PictureEffects.Count)             | 检索包含在 PictureEffects 集合中的 PictureEffect 对象数的计数。只读            |
|  3   | [Creator](#PictureEffects.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 PictureEffects 对象的应用程序。只读 |
|  4   | [Item](#PictureEffects.Item)               | 检索位于指定索引处的 PictureEffect 对象。只读                                  |

## 成员方法

### PictureEffects.Delete

从集合中删除 **PictureEffect** 对象。

**语法**

**express.Delete ( *Index* )**

\*express是一个代表 PictureEffects 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                  |
|---------|-----------|----------|---------------------------------------|
| *Index* | 可选      | Integer  | 要删除的 PictureEffect 对象的索引号。 |

### PictureEffects.Insert

将图片效果插入到复合效果链中。

**语法**

**express.Insert ( *EffectType* , *Position* )**

\*express是一个代表 PictureEffects 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型             | 说明                             |
|--------------|-----------|----------------------|----------------------------------|
| *EffectType* | 必选      | MsoPictureEffectType | 一个指定图片效果类型的枚举。     |
| *Position*   | 可选      | Integer              | 效果在图片效果的复合链中的位置。 |

**说明**

图片效果被处理为由各个项构成的链，以便创建一个最终复合图像，链中的项是按照顺序应用的。效果链将允许向链中添加效果、对效果重新排序或从链中删除效果。

**示例**

``` JavaScript
//下面的代码将设置 Wpp 幻灯片中某个形状的几个图片效果填充属性。
function test()
{
    // Setup a slide with one picture shape.
    let pes = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects
    // Insert a 150% Saturation effect.
    pes.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5
    // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.
    let brightnessContrast = pes.Insert(Application.Enum.msoEffectBrightnessContrast)
    brightnessContrast.EffectParameters.Item(1).Value = -0.5
    brightnessContrast.EffectParameters.Item(2).Value = 0.25
}
```

## 成员属性

### PictureEffects.Application

获取一个 **Application** 对象，该对象代表 **PictureEffects** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 PictureEffects 对象的变量。

### PictureEffects.Count

检索包含在 **PictureEffects** 集合中的 **PictureEffect** 对象数的计数。只读

**语法**

**express.Count**

\*express是一个代表 PictureEffects 对象的变量。

### PictureEffects.Creator

获取一个 32 位整数，该整数指示在其中创建了 **PictureEffects** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 PictureEffects 对象的变量。

### PictureEffects.Item

检索位于指定索引处的 **PictureEffect** 对象。只读

**语法**

**express.Item**

\*express是一个代表 PictureEffects 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/PictureEffects](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# WorkflowTemplate 对象

## 简介

代表可用于当前文档的工作流之一。

## 说明

一个对应于 “启动新工作流” 对话框中显示的选项之一的 WorkflowTemplate 对象。在网页上，工作流模板以选项列表的形式显示。

``` JavaScript
/*以下示例显示当前文档中每个工作流模板的名称，然后显示特定模板的工作流特定配置用户界面。应该注意到，调用 GetWorkflowTemplates 方法涉及到服务器的往返行程。*/
function test(){
    const objWorkflowTemplates = Application.ActiveWorkbook.GetWorkflowTemplates();
    for(let i = 1; i <= objWorkflowTemplates.Count; i++)
        Debug.Print(objWorkflowTemplates.Item(i).Name);
        
    var objWorkflowTemplate = objWorkflowTemplates.Item(1);
    objWorkflowTemplate.Show();
}
```

## 方法

| 序号 | 名称                           | 说明                                                     |
|:----:|:-------------------------------|:---------------------------------------------------------|
|  1   | [Show](#WorkflowTemplate.Show) | 显示指定 WorkflowTemplate 对象的工作流特定配置用户界面。 |

## 属性

| 序号 | 名称                                                         | 说明                                                                         |
|:----:|:-------------------------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#WorkflowTemplate.Application)                 | 获取一个代表 WorkflowTemplate 对象的容器应用程序的 Application 对象。只读。  |
|  2   | [Creator](#WorkflowTemplate.Creator)                         | 获取一个 32 位整数，指示创建 WorkflowTemplate 对象时所使用的应用程序。只读。 |
|  3   | [Description](#WorkflowTemplate.Description)                 | 获取工作流模板的说明。只读。                                                 |
|  4   | [DocumentLibraryName](#WorkflowTemplate.DocumentLibraryName) | 获取与工作流模板关联的文档库的名称。只读。                                   |
|  5   | [DocumentLibraryURL](#WorkflowTemplate.DocumentLibraryURL)   | 获取存储工作流模板的文档库的 URL 地址。只读。                                |
|  6   | [Id](#WorkflowTemplate.Id)                                   | 获取用于创建工作流实例的模板的 ID。只读。                                    |
|  7   | [Name](#WorkflowTemplate.Name)                               | 获取 WorkflowTemplate 对象的名称。只读。                                     |

## 成员方法

### WorkflowTemplate.Show

显示指定 **WorkflowTemplate** 对象的工作流特定配置用户界面。

**语法**

**express.Show ()**

\*express是一个代表 WorkflowTemplate 对象的变量。

**返回值**

Integer

**示例**

``` JavaScript
/*以下示例显示当前文档中每个工作流模板的名称，然后显示特定模板的工作流特定配置用户界面。*/
function test(){
    const objWorkflowTemplates = Application.ActiveWorkbook.GetWorkflowTemplates();
    for(let i = 1; i <= objWorkflowTemplates.Count; i++)
        Debug.Print(objWorkflowTemplates.Item(i).Name);

    var objWorkflowTemplate = objWorkflowTemplates.Item(1);
    objWorkflowTemplate.Show();
}
```

## 成员属性

### WorkflowTemplate.Application

获取一个代表 **WorkflowTemplate** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 WorkflowTemplate 对象的变量。

### WorkflowTemplate.Creator

获取一个 32 位整数，指示创建 **WorkflowTemplate** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 WorkflowTemplate 对象的变量。

### WorkflowTemplate.Description

获取工作流模板的说明。只读。

**语法**

**express.Description**

\*express是一个代表 WorkflowTemplate 对象的变量。

### WorkflowTemplate.DocumentLibraryName

获取与工作流模板关联的文档库的名称。只读。

**语法**

**express.DocumentLibraryName**

\*express是一个代表 WorkflowTemplate 对象的变量。

### WorkflowTemplate.DocumentLibraryURL

获取存储工作流模板的文档库的 URL 地址。只读。

**语法**

**express.DocumentLibraryURL**

\*express是一个代表 WorkflowTemplate 对象的变量。

**类型**

String

### WorkflowTemplate.Id

获取用于创建工作流实例的模板的 ID。只读。

**语法**

**express.Id**

\*express是一个代表 WorkflowTemplate 对象的变量。

### WorkflowTemplate.Name

获取 **WorkflowTemplate** 对象的名称。只读。

**语法**

**express.Name**

\*express是一个代表 WorkflowTemplate 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/WorkflowTemplate](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
