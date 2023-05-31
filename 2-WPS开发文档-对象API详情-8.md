# FileSystem 对象

## 简介

FileSystem 对象包括对文件和文件夹的一些基本和常见的操作接口，后续会根据实际需求持续增加。

## 说明

以下代码示例用于判断在临时目录下是否有a.txt这个文件

``` JavaScript
let tempPath = Application.Env.GetTempPath()
if (Application.FileSystem.Exists(tempPath + 'a.txt'))
    alert("a.txt文件存在")
```

## 方法

| 序号 | 名称                                                   | 说明                                           |
|:----:|:-------------------------------------------------------|:-----------------------------------------------|
|  1   | [AppendFile](#FileSystem.AppendFile)                   | 往文件末尾添加数据                             |
|  2   | [copyFileSync](#FileSystem.copyFileSync)               | 生成文件副本                                   |
|  3   | [Exists](#FileSystem.Exists)                           | 判断一个文件和文件夹是否存在。                 |
|  4   | [existsSync](#FileSystem.existsSync)                   | 判断目录是否存在                               |
|  5   | [Mkdir](#FileSystem.Mkdir)                             | 根据给定的path创建一个文件夹。                 |
|  6   | [mkdirSync](#FileSystem.mkdirSync)                     | 创建目录                                       |
|  7   | [mkdtempSync](#FileSystem.mkdtempSync)                 | 创建临时目录                                   |
|  8   | [readAsBinaryString](#FileSystem.readAsBinaryString)   | 读取指定路径的文件，返回二进制字符串数据。     |
|  9   | [readdirSync](#FileSystem.readdirSync)                 | 获取目录下的子目录对象数组（包含目录相关属性） |
|  10  | [ReadFile](#FileSystem.ReadFile)                       | 获取文件的内容                                 |
|  11  | [readFileString](#FileSystem.readFileString)           | 读取指定路径的文件，返回字符串。               |
|  12  | [Remove](#FileSystem.Remove)                           | 删除指定的path所代表的文件或文件夹。           |
|  13  | [rmdirSync](#FileSystem.rmdirSync)                     | 删除目录                                       |
|  14  | [tmpdir](#FileSystem.tmpdir)                           | 获取系统的临时文件目录                         |
|  15  | [unlinkSync](#FileSystem.unlinkSync)                   | 删除文件                                       |
|  16  | [writeAsBinaryString](#FileSystem.writeAsBinaryString) | 写二进制字符串对象到指定文件。                 |
|  17  | [WriteFile](#FileSystem.WriteFile)                     | 创建文件                                       |
|  18  | [writeFileString](#FileSystem.writeFileString)         | 写字符串到指定路径的文件。                     |

## 属性

| 序号 | 名称                           | 说明 |
|:----:|:-------------------------------|:-----|
|  1   | [Reserve](#FileSystem.Reserve) |      |

## 成员方法

### FileSystem.AppendFile

往文件末尾添加数据

**语法**

**express.AppendFile ( *FilePath* , *Data* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明             |
|------------|-----------|----------|------------------|
| *FilePath* | 必选      | String   | 要写入的文件名   |
| *Data*     | 必选      | String   | 要写入文件的内容 |

### FileSystem.copyFileSync

生成文件副本

**语法**

**express.copyFileSync ( *SrcFilePath* , *DestFilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                     |
|----------------|-----------|----------|--------------------------|
| *SrcFilePath*  | 必选      | String   | 源文件的路径包括文件名称 |
| *DestFilePath* | 必选      | String   | 新创建的文件副本文件名称 |

### FileSystem.Exists

判断一个文件和文件夹是否存在。

**语法**

**express.Exists ( *path* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                             |
|--------|-----------|----------|--------------------------------------------------|
| *path* | 必选      | String   | 文件或文件夹的路径，这个路径一定要是一个全路径。 |

### FileSystem.existsSync

判断目录是否存在

**语法**

**express.existsSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明               |
|------------|-----------|----------|--------------------|
| *FilePath* | 必选      | String   | 要判断的文件的名称 |

### FileSystem.Mkdir

根据给定的path创建一个文件夹。

**语法**

**express.Mkdir ( *path* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                               |
|--------|-----------|----------|----------------------------------------------------|
| *path* | 必选      | String   | 要创建的文件夹的路径，这个路径一定要是一个全路径。 |

### FileSystem.mkdirSync

创建目录

**语法**

**express.mkdirSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明           |
|------------|-----------|----------|----------------|
| *FilePath* | 必选      | String   | 创建目录的路径 |

### FileSystem.mkdtempSync

创建临时目录

**语法**

**express.mkdtempSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明           |
|------------|-----------|----------|----------------|
| *FilePath* | 必选      | String   | 临时目录的路径 |

### FileSystem.readAsBinaryString

读取指定路径的文件，返回二进制字符串数据。

**语法**

**express.readAsBinaryString ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明     |
|------------|-----------|----------|----------|
| *FilePath* | 必选      | String   | 文件路径 |

### FileSystem.readdirSync

获取目录下的子目录对象数组（包含目录相关属性）

**语法**

**express.readdirSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明               |
|------------|-----------|----------|--------------------|
| *FilePath* | 必选      | String   | 要获取的目录的路径 |

### FileSystem.ReadFile

获取文件的内容

**语法**

**express.ReadFile ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明         |
|------------|-----------|----------|--------------|
| *FilePath* | 必选      | String   | 读取文件内容 |

### FileSystem.readFileString

读取指定路径的文件，返回字符串。

**语法**

**express.readFileString ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明     |
|------------|-----------|----------|----------|
| *FilePath* | 必选      | String   | 文件路径 |

### FileSystem.Remove

删除指定的path所代表的文件或文件夹。

**语法**

**express.Remove ( *path* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                     |
|--------|-----------|----------|----------------------------------------------------------|
| *path* | 必选      | String   | 要删除的文件或文件夹的路径，这个路径一定要是一个全路径。 |

### FileSystem.rmdirSync

删除目录

**语法**

**express.rmdirSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                 |
|------------|-----------|----------|----------------------|
| *FilePath* | 必选      | String   | 要删除的文件目录路径 |

### FileSystem.tmpdir

获取系统的临时文件目录

**语法**

**express.tmpdir ()**

\*express是一个代表 FileSystem 对象的变量。

### FileSystem.unlinkSync

删除文件

**语法**

**express.unlinkSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明             |
|------------|-----------|----------|------------------|
| *FilePath* | 必选      | String   | 要删除的文件路径 |

### FileSystem.writeAsBinaryString

写二进制字符串对象到指定文件。

**语法**

**express.writeAsBinaryString ( *FilePath* , *binaryString* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明             |
|----------------|-----------|----------|------------------|
| *FilePath*     | 必选      | String   | 文件路径         |
| *binaryString* | 必选      | String   | 二进制字符串数据 |

### FileSystem.WriteFile

创建文件

**语法**

**express.WriteFile ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明     |
|------------|-----------|----------|----------|
| *FilePath* | 必选      | String   | 文件路径 |

### FileSystem.writeFileString

写字符串到指定路径的文件。

**语法**

**express.writeFileString ( *FilePath* , *data* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明       |
|------------|-----------|----------|------------|
| *FilePath* | 必选      | String   | 文件路径   |
| *data*     | 必选      | String   | 字符串数据 |

## 成员属性

### FileSystem.Reserve

**语法**

**express.Reserve**

\*express是一个代表 FileSystem 对象的变量。

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/文件系统/FileSystem](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Application 对象

## 简介

代表整个 ET 应用程序，它是整个应用程序api对象树的根对象。

## 说明

整个et应用程序的api对象结构是一个树状结构，而 Application 对象是树状结构的根对象，同时， Application 对象也提供了对于应用程序相关的各种访问接口，例如：

- 应用程序的设置选项、环境、版本号等相关信息
- 一些常见的属性，如 ActiveCell 、 ActiveWorkbook 、 ActiveSheet 等

## 示例

使用 Application 属性可返回 Application 对象。下例对 Application 对象应用 Windows 属性。

``` JavaScript
function test() {
    let app = Application.Application;
    alert(app.Windows.Count);
}
```

下例在 ET 中打开指定的工作簿。

``` JavaScript
function test(){
    let docPath = "d:/abc.xlsx";
    Application.Workbooks.Open(docPath);
} 
```

## 方法

| 序号 | 名称                              | 说明                                                                                         |
|:----:|:----------------------------------|:---------------------------------------------------------------------------------------------|
|  1   | [Evaluate](#Application.Evaluate) | 将一个 ET 名称转换为一个对象或者一个值。                                                     |
|  2   | [Goto](#Application.Goto)         | 选定任意工作簿中的任意区域，并且如果该工作簿未处于活动状态，就激活该工作簿。                 |
|  3   | [InputBox](#Application.InputBox) | 显示一个接收用户输入的对话框。返回此对话框中输入的信息。                                     |
|  4   | [Quit](#Application.Quit)         | 退出ET                                                                                       |
|  5   | [Run](#Application.Run)           | 运行一个宏或者调用一个函数。该方法可用于运行宏编辑器编写的宏，或者运行 DLL 或 XLL 中的函数。 |

## 属性

| 序号 | 名称                                                | 说明                                                                                                                                                                        |
|:----:|:----------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ActiveCell](#Application.ActiveCell)               | 返回一个 Range 对象，该对象代表活动窗口（顶部窗口）或指定窗口中的活动单元格。如果窗口中没有显示工作表，此属性无效。只读。                                                   |
|  2   | [ActiveSheet](#Application.ActiveSheet)             | 返回一个对象，它代表活动工作簿中或指定的窗口或工作簿中的活动工作表（最上面的工作表）。如果没有活动的工作表，则返回 null 。                                                  |
|  3   | [ActiveWindow](#Application.ActiveWindow)           | 返回一个 Window 对象，该对象代表活动窗口（顶部窗口）。只读。如果没有打开的窗口，则返回 null 。                                                                              |
|  4   | [ActiveWorkbook](#Application.ActiveWorkbook)       | 返回一个 Workbook 对象，该对象代表活动窗口（顶部窗口）中的工作簿。只读。如果没有打开的窗口，或者“信息”窗口或“剪贴板”窗口为活动窗口，则返回 null 。                          |
|  5   | [Application](#Application.Application)             | 该属性返回一个代表 ET 应用程序的 Application 对象。 只读。                                                                                                                  |
|  6   | [Build](#Application.Build)                         | 返回 ET 内部版本号。 Number 类型，只读。                                                                                                                                    |
|  7   | [Caption](#Application.Caption)                     | 返回或设置一个 String 值，它代表出现在 ET 主窗口标题栏中显示的名称。                                                                                                        |
|  8   | [Cells](#Application.Cells)                         | 返回一个 Range 对象，该对象代表活动工作表中的所有单元格。如果活动文档不是工作表，则此属性无效。                                                                             |
|  9   | [Columns](#Application.Columns)                     | 返回一个 Range 对象，它代表活动工作表中的所有列。如果活动文档不是工作表，则 Columns 属性无效。                                                                              |
|  10  | [Dialogs](#Application.Dialogs)                     | 返回一个 Dialogs 集合，该集合代表所有内置对话框。只读。                                                                                                                     |
|  11  | [Height](#Application.Height)                       | 返回一个 Double 值，它代表主应用程序窗口的高度（以磅为单位）。                                                                                                              |
|  12  | [Hwnd](#Application.Hwnd)                           | 返回一个 Number 类型的值，该值表示 ET 窗口的最高级别的窗口句柄。只读。                                                                                                      |
|  13  | [Left](#Application.Left)                           | 返回或设置一个 Double 值，它代表从屏幕左边缘到 ET 主窗口左边缘的距离（以磅为单位）。                                                                                        |
|  14  | [Name](#Application.Name)                           | 返回一个代表对象名称的 String 值。                                                                                                                                          |
|  15  | [Names](#Application.Names)                         | 返回一个 Names 集合，该集合代表活动工作簿中的所有名称。只读 Names 对象。                                                                                                    |
|  16  | [NewWorkbook](#Application.NewWorkbook)             | 返回一个 NewFile 对象。                                                                                                                                                     |
|  17  | [Path](#Application.Path)                           | 返回一个 String 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。                                                                                            |
|  18  | [Rows](#Application.Rows)                           | 返回一个 Range 对象，它代表活动工作表中的所有行。如果活动文档不是工作表，则 Rows 属性失效。 Range 对象，只读。                                                              |
|  19  | [Selection](#Application.Selection)                 | 为 Application 对象返回在活动窗口中选定的对象。                                                                                                                             |
|  20  | [Top](#Application.Top)                             | 返回或设置一个 Double 值，它代表从屏幕上边缘到 ET 主窗口上边缘的距离（以磅为单位）。                                                                                        |
|  21  | [Value](#Application.Value)                         | 返回一个代表应用程序名称的 String 值。                                                                                                                                      |
|  22  | [Version](#Application.Version)                     | 返回一个 String 值，它代表 ET 版本号。                                                                                                                                      |
|  23  | [Width](#Application.Width)                         | 返回或设置一 个 Double 值，它代表应用程序窗口的宽度（以磅为单位）。                                                                                                         |
|  24  | [WindowState](#Application.WindowState)             | 返回或设置窗口的状态。 XlWindowState 类型，可读写。                                                                                                                         |
|  25  | [Windows](#Application.Windows)                     | 返回一个 Windows 集合，它代表所有工作簿中的所有窗口。 Windows 对象，只读。                                                                                                  |
|  26  | [Workbooks](#Application.Workbooks)                 | 返回一个 Workbooks 集合，该集合表示所有打开的工作簿。只读。                                                                                                                 |
|  27  | [WorksheetFunction](#Application.WorksheetFunction) | 返回 WorksheetFunction 对象。只读。                                                                                                                                         |
|  28  | [Worksheets](#Application.Worksheets)               | 对于 Application 对象，返回一个 Sheets 集合，它代表活动工作簿中的所有工作表。对于 Workbook 对象，返回一个 Sheets 集合，它代表指定工作簿中的所有工作表。 Sheets 对象，只读。 |

## 成员方法

### Application.Evaluate

将一个 ET 名称转换为一个对象或者一个值。

**语法**

**express.Evaluate ( *Name* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                               |
|--------|-----------|----------|------------------------------------|
| *Name* | 必选      | Variant  | 公式或使用 ET 命名约定的对象名称。 |

**说明**

该方法可使用下列 ET 名称类型：

- A1 格式引用。可以通过 A1 格式表示法引用单个单元格。所有引用均视为绝对引用。
- 区域。在引用中可以使用区域、交集和联合运算符（分别为冒号、空格和逗号）。
- 定义的名称。可用宏语言指定任何名称。
- 外部引用。可以使用 ! 运算符引用另一工作簿中的单元格或已定义的名称，例如，Evaluate("\[BOOK1.XLS\]Sheet1!A1")。
- 图表对象。可以指定任何图表对象名称（如“Legend”、“Plot Area”或“Series 1”），以访问该对象的属性和方法。例如，Charts("Chart1").Evaluate("Legend").Font.Name 返回图例中所用字体的名称。

**示例**

``` JavaScript
/* 用字符串参数调用 Evaluate 方法 */
function test() {
    Application.Evaluate("A1").Value2 = 25
    let trigVariable = Application.Evaluate("SIN(45)")
    let firstCellInSheet = Application.Workbooks.Item("工作簿1.xlsx").Sheets.Item(1).Evaluate("A1")
    alert(trigVariable)
    alert(firstCellInSheet)
}

/* 此示例将工作表 Sheet1 上 A1 单元格的字体设置为加粗。 */
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    let boldCell = "A1"
    Application.Evaluate(boldCell).Font.Bold = true
}
```

### Application.Goto

选定任意工作簿中的任意区域，并且如果该工作簿未处于活动状态，就激活该工作簿。

**语法**

**express.Goto ( *Reference* , *Scroll* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                      |
|-------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------|
| *Reference* | 可选      | Variant  | 目标。可以是 Range 对象、包含 R1C1-样式批注的单元格引用的字符串。如果省略该参数，目标将为最近一次用 Goto 方法选定的区域。 |
| *Scroll*    | 可选      | Boolean  | 如果为 True，则滚动窗口直至区域的左上角出现在窗口的左上角中。如果为 False，则不滚动窗口。默认值为 False。                 |

**说明**

该方法与 Select 方法的区别：

1.  如果指定的区域不在位于最前面屏幕的工作表中，ET 将在选定该区域之前切换至该工作表。（如果对不在屏幕的最前面的工作表中的区域使用 Select 方法，则选定该区域时并不激活该工作表）
2.  该方法具有让用户滚动目标窗口的 Scroll 参数
3.  当使用 Goto 方法时，前一次选定区域（Goto 方法运行前）被增加到以前选定区域的数组中（有关详细信息，请参阅 PreviousSelections 属性）。可以使用该功能快速跳过选定区域，选定区域最多为四个
4.  Select 方法具有 Replace 参数，而 Goto 方法没有该参数

**示例**

``` JavaScript
/* 本示例选定工作表 Sheet1 中的单元格 A154，并滚动工作表以显示该区域。 */
function test() {
    let rg = Application.ActiveWorkbook.Worksheets.Item("Sheet1").Range("A154");
    Application.Goto(rg, true)
}
```

### Application.InputBox

显示一个接收用户输入的对话框。返回此对话框中输入的信息。

**语法**

**express.InputBox ( *Prompt* , *Title* , *Default* , *Left* , *Top* , *HelpFile* , *HelpContextID* , *Type* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                        |
|-----------------|-----------|----------|-------------------------------------------------------------------------------------------------------------|
| *Prompt*        | 必选      | String   | 要在对话框中显示的消息。可为字符串、数字、日期、或布尔值（在显示之前，ET 自动将其值强制转换为 String）。    |
| *Title*         | 可选      | Variant  | 输入框的标题。如果省略该参数，默认标题将为“Input”。                                                         |
| *Default*       | 可选      | Variant  | 指定一个初始值，该值在对话框最初显示时出现在文本框中。如果省略该参数，文本框将为空。该值可以是 Range 对象。 |
| *Left*          | 可选      | Variant  | 指定对话框相对于屏幕左上角的 X 坐标（以磅为单位）。                                                         |
| *Top*           | 可选      | Variant  | 指定对话框相对于屏幕左上角的 Y 坐标（以磅为单位）。                                                         |
| *HelpFile*      | 可选      | Variant  | 此输入框使用的帮助文件名。如果存在 HelpFile 和 HelpContextID 参数，对话框中将出现一个帮助按钮。             |
| *HelpContextID* | 可选      | Variant  | HelpFile 中帮助主题的上下文 ID 号。                                                                         |
| *Type*          | 可选      | Variant  | 指定返回的数据类型。如果省略该参数，对话框将返回文本。                                                      |

**说明**

下表列出了可以在 Type 参数中传递的值。可以为下列值之一或其中几个值的和。例如，对于一个可接受文本和数字的输入框，将 *Type* 设置为 1 + 2。

| 值  | 含义                                |
|-----|-------------------------------------|
| 0   | 公式                                |
| 1   | 数字                                |
| 2   | 文本（字符串）                      |
| 4   | 逻辑值（ **True** 或 **False** ）   |
| 8   | 单元格引用，作为一个 **Range** 对象 |
| 16  | 错误值，如 \#N/A                    |
| 64  | 数值数组                            |

使用 **InputBox** 可以显示一个简单的对话框，以便可以输入要在宏中使用的信息。此对话框有一个“确定” 按钮和一个“取消” 按钮。如果选择了“确定” 按钮，则 **InputBox** 将返回对话框中输入的值。如果单击“取消” 按钮，则 **InputBox** 返回 **False** 。

如果 *Type* 为 0， **InputBox** 将以文本格式返回公式。例如，“=2\*PI()/360”。如果公式中有引用，将以 A1-样式引用返回（使用 **ConvertFormula** 转换引用样式）。

如果 *Type* 为 8， **InputBox** 将返回一个 **Range** 对象。您必须用 **Set** 语句将结果指定给一个 **Range** 对象，如下例所示。

``` JavaScript
function test() {
    let myRange = Application.InputBox("Sample", undefined, undefined, undefined, undefined, undefined, undefined,8)
}
```

如果不使用 **Set** 语句，此变量将被设置为这个区域的值，而不是 **Range** 这个对象本身。

如果使用 **InputBox** 方法要求用户输入公式，则必须使用 **FormulaLocal** 属性来将此公式指定给一个 **Range** 对象。输入的公式使用用户语言。

**InputBox** 方法与 **InputBox** 函数的区别在于：它可以对用户的输入进行选择性验证，也可用于 ET 对象、误差值、和公式的输入。注意， ` Application.InputBox ` 调用的是 **InputBox** 方法，不带对象识别符的 ` InputBox ` 调用的是 **InputBox** 函数。

**示例**

``` JavaScript
/* 本示例提示用户输入数字。 */
function test() {
    let myNum = Application.InputBox("Enter a number")

}

/* 本示例提示用户在 Sheet1 中选取一个单元格。示例使用 Type 参数证实返回值是有效的单元格引用 (一个 Range 对象)。 */
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    let myCell = Application.InputBox("Select a cell", undefined, undefined, undefined, undefined, undefined, undefined, 8)
}
```

### Application.Quit

退出ET

**语法**

**express.Quit ()**

\*express是一个代表 Application 对象的变量。

**说明**

使用此方法时，如果未保存的工作簿处于打开状态，则 ET 将显示一个对话框，询问是否要保存所作更改。要防止发生这种情况，请在使用 **Quit** 方法前保存所有工作簿或将 **DisplayAlerts** 属性设置为 **False** 。如果该属性为 **False** ，则在 ET 退出时，即使有未保存的工作簿，也不会显示对话框，而且不保存就退出。

如果将工作簿的 **Saved** 属性设置为 **True** ，而没有将工作簿保存到磁盘上，则 ET 在退出时不会提示保存该工作簿。

**示例**

``` JavaScript
/* 本示例保存所有打开的工作簿，然后退出 ET。 */
function test() {
    let w = Application.Workbooks
    for (let i = 1; i <= w.Count; i++) {
        w.Item(i).Save()
    }
    Application.Quit()
}
```

### Application.Run

运行一个宏或者调用一个函数。该方法可用于运行宏编辑器编写的宏，或者运行 DLL 或 XLL 中的函数。

**语法**

**express.Run ( *Macro* , *Arg1-Arg30* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                                                                                       |
|--------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Macro*      | 可选      | Variant  | 要运行的宏。它可以是具有宏名称的字符串、表示函数所在位置的 Range 对象，或者是一个已注册的 DLL (XLL) 函数的注册号。如果使用字符串，将在当前工作表的上下文中对该字符串求值。 |
| *Arg1-Arg30* | 可选      | Variant  | 应传递给函数的参数。                                                                                                                                                       |

**说明**

此方法不可使用命名参数，参数必须通过位置进行传递。

**Run** 方法返回被调用的宏返回的任何值。如果将对象作为参数传递给宏，该对象将转换为相应的值（通过对该对象应用 **Value** 属性）。这意味着不能用 **Run** 方法将对象传递给宏。

## 成员属性

### Application.ActiveCell

返回一个 Range 对象，该对象代表活动窗口（顶部窗口）或指定窗口中的活动单元格。如果窗口中没有显示工作表，此属性无效。只读。

**语法**

**express.ActiveCell**

\*express是一个代表 Application 对象的变量。

**说明**

如果不指定对象识别符，此属性返回活动窗口中的活动单元格。

请仔细区分活动单元格和选定区域。活动单元格为选定区域内部的一个单元格。而选定区域可以包含多个单元格，但只有一个单元格为活动单元格。

下列表达式都是返回活动单元格，并且都是等效的。

``` JavaScript
Application.ActiveCell
Application.ActiveWindow.ActiveCell
```

**示例**

``` JavaScript
/* 此示例在消息框中显示活动单元格的值。由于如果活动表不是工作表则 ActiveCell 属性无效，所以此示例使用 ActiveCell 属性之前先激活 Sheet1。 */
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    alert(Application.ActiveCell.Value2)
}

/* 此示例更改活动单元格的字体格式设置。 */
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    let font = Application.ActiveCell.Font
    font.Bold = true
    font.Italic = true
}
```

### Application.ActiveSheet

返回一个对象，它代表活动工作簿中或指定的窗口或工作簿中的活动工作表（最上面的工作表）。如果没有活动的工作表，则返回 **null** 。

**语法**

**express.ActiveSheet**

\*express是一个代表 Application 对象的变量。

**说明**

如果不指定对象识别符，则此属性返回活动工作簿中的活动工作表。

如果某个工作簿出现在若干个窗口中，那么该工作簿的 **ActiveSheet** 属性在不同窗口中可能不同。

**示例**

``` JavaScript
/* 此示例显示活动工作表的名称。 */
function test() {
    alert("The name of the active sheet is " + Application.ActiveSheet.Name)
}
```

### Application.ActiveWindow

返回一个 **Window** 对象，该对象代表活动窗口（顶部窗口）。只读。如果没有打开的窗口，则返回 **null** 。

**语法**

**express.ActiveWindow**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 本示例显示活动窗口的名称（Caption 属性）。 */
function test() {
    alert("The name of the active window is " + Application.ActiveWindow.Caption)
}
```

### Application.ActiveWorkbook

返回一个 **Workbook** 对象，该对象代表活动窗口（顶部窗口）中的工作簿。只读。如果没有打开的窗口，或者“信息”窗口或“剪贴板”窗口为活动窗口，则返回 **null** 。

**语法**

**express.ActiveWorkbook**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 本示例显示活动工作簿的名称。 */
function test() {
    alert("The name of the active workbook is " + Application.ActiveWorkbook.Name)
}
```

### Application.Application

该属性返回一个代表 ET 应用程序的 **Application** 对象。 只读。

**语法**

**express.Application**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息。 */
function test() {
    let myObject = Application
    alert(myObject.Application.Value)
}
```

### Application.Build

返回 ET 内部版本号。 **Number** 类型，只读。

**语法**

**express.Build**

\*express是一个代表 Application 对象的变量。

**说明**

通常检测 Version 属性较为安全，除非必须要获知内部版本号。

**示例**

``` JavaScript
/* 本示例检测 Build 属性。 */
function test() {
    if(Application.Build > 2500) {
        alert("build-dependent code here")
    }
}
```

### Application.Caption

返回或设置一个 **String** 值，它代表出现在 ET 主窗口标题栏中显示的名称。

**语法**

**express.Caption**

\*express是一个代表 Application 对象的变量。

**说明**

如果未设置该名称，或将其设置为 **Empty** , 则此属性返回“ET”。

**示例**

``` JavaScript
/* 此示例将出现在 ET 主窗口标题栏中的名称设置为自定义的名称。 */
function test() {
    Application.Caption = "KingSoft"
}
```

### Application.Cells

返回一个 **Range** 对象，该对象代表活动工作表中的所有单元格。如果活动文档不是工作表，则此属性无效。

**语法**

**express.Cells**

\*express是一个代表 Application 对象的变量。

**说明**

因为 **Item** 属性是 Range 对象的 <span class="glossary" "appendpopup(this,'defdefaultproperty_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none;">默认属性</span> ，所以可以在 **Cells** 关键字后面紧接着指定行和列索引。有关详细信息，请参阅 **Item** 属性和本主题的示例。

在不使用对象识别符的情况下，使用此属性将返回一个 Range 对象，它代表活动工作表中所有的单元格。

### Application.Columns

返回一个 Range 对象，它代表活动工作表中的所有列。如果活动文档不是工作表，则 Columns 属性无效。

**语法**

**express.Columns**

\*express是一个代表 Application 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 **ActiveSheet.Columns** ` `` ` 。

此属性在应用于一个是多重选定区域的 Range 对象时，会只从该区域的第一个子区域中返回列。例如，如果 Range 对象有两个子区域 A1:B2 和 C3:D4，那么， ` Selection.Columns.Count ` 的返回值是 2，而不是 4。若要对一个可能包含多重选定区域的区域使用此属性，请测试 ` Areas.Count ` 以确定此区域内是否包含多个子区域。如果包含，请对此区域内的每个子区域进行循环。

` ` ` `

### Application.Dialogs

返回一个 **Dialogs** 集合，该集合代表所有内置对话框。只读。

**语法**

**express.Dialogs**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 本示例显示“文件”菜单的“打开”对话框。 */
function test() {
    Application.Dialogs.Item(xlDialogOpen).Show()
}
```

### Application.Height

返回一个 **Double** 值，它代表主应用程序窗口的高度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 Application 对象的变量。

**说明**

如果窗口处于最小化，则此属性为只读，并引用图标的高度。如果窗口处于最大化，则无法设置此属性。使用 **WindowState** 属性可确定窗口的状态。

### Application.Hwnd

返回一个 **Number** 类型的值，该值表示 ET 窗口的最高级别的窗口句柄。只读。

**语法**

**express.Hwnd**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 在本示例中，ET 向用户通知 ET 窗口的最高级别的窗口句柄。 */
function test() {
    MsgBox("The top-level window handle is: " + Application.Hwnd)
}
```

### Application.Left

返回或设置一个 **Double** 值，它代表从屏幕左边缘到 ET 主窗口左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 Application 对象的变量。

### Application.Name

返回一个代表对象名称的 **String** 值。

**语法**

**express.Name**

\*express是一个代表 Application 对象的变量。

### Application.Names

返回一个 **Names** 集合，该集合代表活动工作簿中的所有名称。只读 **Names** 对象。

**语法**

**express.Names**

\*express是一个代表 Application 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 ` ActiveWorkbook.Names ` 。

### Application.NewWorkbook

返回一个 **NewFile** 对象。

**语法**

**express.NewWorkbook**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 在本示例中，ET 将 wkbOne 变量设置为 NewFile 对象。 */
function test() {
    // Create a reference to an instance of the NewFile object.
    let wkbOne = Application.NewWorkbook
}
```

### Application.Path

返回一个 **String** 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。

**语法**

**express.Path**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 此示例显示 ET 的完整路径。 */
function test() {
    alert("The path is " + Application.Path)
}
```

### Application.Rows

返回一个 **Range** 对象，它代表活动工作表中的所有行。如果活动文档不是工作表，则 **Rows** 属性失效。 **Range** 对象，只读。

**语法**

**express.Rows**

\*express是一个代表 Application 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 ` ActiveSheet.Rows ` 。

该属性在应用于是多个选定区域的 **Range** 对象时，只从该区域中第一个子区域内返回行。例如，如果 **Range** 对象有两个子区域：A1:B2 和 C3:D4，则 ` Selection.Rows.Count ` 返回 2 而不是 4。若要在一个可能包含多个选定区域的区域中使用此属性，请测试 ` Areas.Count ` 以确定该区域是否包含多个选择区域。如果是，请对此区域内的每个子区域进行循环，如第 3 个示例所示。

**示例**

``` JavaScript
/* 此示例删除 Sheet1 的第三行。 */
function test() {
    Application.Worksheets.Item("Sheet1").Rows.Item(3).Delete()
}

/* 此示例检查第一张工作表上当前区域中的行，如果某行的第一个单元格值与前一行的第一个单元格的值相等，则删除此行。 */
function test() {
    let last;
    let cnt = Application.Selection.Rows.Count;
    let delCnt = 0;
    for (i = 1; i <= cnt - delCnt; i++) {
        let rw = Application.Selection.Rows.Item(i);
        cur = rw.Cells.Item(1, 1).Value();
        if (cur === last) {
            rw.Delete();
            i--;
            delCnt ++;
        }
        last = cur;
    }
}


/* 此示例显示 Sheet1 选定区域中的行数。如果选择了多个子区域，此示例将对每一个子区域进行循环。 */
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    let areaCount = Application.Selection.Areas.Count
    if (areaCount <= 1) {
        alert(`The selection contains${Application.Selection.Rows.Count}rows.`)
    } else {
        let i = 1
        for (item in Application.Selection.Areas) {
            alert(`Area${i}of the selection contains${item.Rows.Count}rows.`)
            i += 1
        }
    }
}
```

### Application.Selection

为 **Application** 对象返回在活动窗口中选定的对象。

**语法**

**express.Selection**

\*express是一个代表 Application 对象的变量。

**说明**

返回的对象类型取决于当前所选内容（例如，如果选择了单元格，此属性将返回 Range 对象）。如果未选择任何内容， **Selection** 属性将返回 **Nothing** 。

在不使用对象识别符的情况下，使用此属性等效于使用 ` Application.Selection ` 。

**示例**

``` JavaScript
/* 本示例清空 Sheet1 的选定对象（假定选定对象为单元格区域）。 */
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.Selection.Clear()
}

/* 本示例显示选定对象的 wpsjs 加载项 对象类型。 */
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    alert("The selection object type is " + typeof(Application.Selection))
}
```

### Application.Top

返回或设置一个 **Double** 值，它代表从屏幕上边缘到 ET 主窗口上边缘的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 Application 对象的变量。

**说明**

如果应用程序窗口已最小化，则此属性将控制窗口图标的位置（在屏幕的任何位置）。

### Application.Value

返回一个代表应用程序名称的 **String** 值。

**语法**

**express.Value**

\*express是一个代表 Application 对象的变量。

**说明**

此属性始终返回 "Microsoft Excel"。

### Application.Version

返回一个 **String** 值，它代表 ET 版本号。

**语法**

**express.Version**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 此示例显示包含 ET 版本号和操作系统名称的消息框。 */
function test() {
    alert(`Welcome to ET version ${Application.Version} running on ${Application.OperatingSystem}!`)
}
```

### Application.Width

返回或设置一 **个 Double** 值，它代表应用程序窗口的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 Application 对象的变量。

**说明**

如果窗口处于最小化状态，则 **Width** 为只读，并返回窗口图标的宽度。

**示例**

``` JavaScript
/*  此示例将活动窗口的尺寸扩大为可用的最大尺寸（假定该窗口尚未最大化）。 */
function test() {
    let window = Application.ActiveWindow
    window.WindowState = Application.Enum.xlNarrow
    window.Top = 1
    window.Left = 1
    window.Height = Application.UsableHeight
    window.Width = Application.UsableWidth
}
```

### Application.WindowState

返回或设置窗口的状态。 XlWindowState 类型，可读写。

**语法**

**express.WindowState**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 本示例将 ET 应用程序窗口最大化。 */
function test() {
    Application.WindowState = Application.Enum.xlMaximized
}

/*  此示例将活动窗口的尺寸扩大为可用的最大尺寸（假定该窗口尚未最大化）。 */
function test() {
    let window = Application.ActiveWindow
    window.WindowState = Application.Enum.xlNarrow
    window.Top = 1
    window.Left = 1
    window.Height = Application.UsableHeight
    window.Width = Application.UsableWidth
}
```

### Application.Windows

返回一个 **Windows** 集合，它代表所有工作簿中的所有窗口。 **Windows** 对象，只读。

**语法**

**express.Windows**

\*express是一个代表 Application 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 ` Application.Windows ` 。

此属性返回的集合中既包括可见窗口，也包括隐藏窗口。

**示例**

``` JavaScript
/* 此示例关闭 ET 中第一个打开或隐藏的窗口。 */
function test() {
    Application.Windows.Item(1).Close()
}

/* 此示例将活动工作簿的窗口一命名为“Consolidated Balance Sheet”。此名称将被用作对 Windows 集合的索引。 */
function test() {
    Application.ActiveWorkbook.Windows.Item(1).Caption = "Consolidated Balance Sheet"
    Application.ActiveWorkbook.Windows.Item("Consolidated Balance Sheet").ActiveSheet.Calculate()
}
```

### Application.Workbooks

返回一个 **Workbooks** 集合，该集合表示所有打开的工作簿。只读。

**语法**

**express.Workbooks**

\*express是一个代表 Application 对象的变量。

**说明**

在不使用对象识别符的情况下，使用该属性相当于使用 。

``` JavaScript
Application.Workbooks
```

由 **Workbooks** 属性返回的集合并不包含打开的加载宏（一种特殊的隐藏工作簿）。但如果知道文件名，则可返回单个打开的加载宏。例如，

``` JavaScript
Application.Workbooks.Item("Oscar.xla")
```

将把名为“Oscar.xla”的打开的加载宏作为 **Workbook** 对象返回。

**示例**

``` JavaScript
/* 本示例激活 Book1.xls 工作簿。 */
function test() {
    Application.Workbooks.Item("Book1").Activate()
}

/* 本示例打开 Large.xlsx 工作簿。 */
function test() {
    Application.Workbooks.Open("C:/Users/wps/Desktop/large.xlsx")
}

/* 本示例关闭除正在运行本示例的工作簿以外的所有其他工作簿，并保存其更改。 */
function test() {
    const Count = Application.Workbooks.Count
    for(let i = 1; i <= Count; i++){
        if(Application.Workbooks.Item(i).Name != Application.ThisWorkbook.Name){
            Application.Workbooks.Item(i).Close(true)
        }
    }  
}
```

### Application.WorksheetFunction

返回 **WorksheetFunction** 对象。只读。

**语法**

**express.WorksheetFunction**

\*express是一个代表 Application 对象的变量。

**示例**

``` JavaScript
/* 本示例对单元格区域 A1:A10 应用 Min 工作表函数，并显示结果。 */
function test() {
    let myRange = Application.Worksheets.Item("Sheet1").Range("A1:C10")
    let answer = Application.WorksheetFunction.Min(myRange)
    alert(answer)
}
```

### Application.Worksheets

对于 **Application** 对象，返回一个 Sheets 集合，它代表活动工作簿中的所有工作表。对于 **Workbook** 对象，返回一个 Sheets 集合，它代表指定工作簿中的所有工作表。 Sheets 对象，只读。

**语法**

**express.Worksheets**

\*express是一个代表 Application 对象的变量。

**说明**

在不使用对象识别符的情况下，使用此属性将返回活动工作簿中所有的工作表。

此属性不返回宏表；使用 **Excel4MacroSheets** 属性或 **Excel4IntlMacroSheets** 属性可返回这些表。

**示例**

``` JavaScript
/* 此示例显示活动工作簿中 sheet1 上单元格 A1 中的值。。 */
function test() {
    alert(Application.Worksheets.Item("Sheet1").Range('A1').Value2)
}

/* 此示例显示活动工作簿中每个工作表的名称。 */
function test() {
    for (let i = 1; i <= Application.Worksheets.Count; i++) {
        alert(Application.Worksheets.Item(i).Name)
    }
}

/* 此示例向活动工作簿添加新工作表，并设置该工作表的名称。 */
function test() {
    let newSheet = Application.Worksheets.Add()
    newSheet.Name = "current Budget"
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ColorFormat 对象

## 简介

代表单色对象的颜色、带有渐变或图案填充的对象的前景或背景色，或者指针的颜色。

## 说明

可以将颜色设为显式的红-绿-蓝值（使用 RGB 属性），或设为配色方案中的一种颜色（使用 SchemeColor 属性）。

使用下表中列出的属性之一可返回 ColorFormat 对象。

| 使用此属性     | 对象         | 返回一个 ColorFormat 对象，该对象代表          |
|----------------|--------------|------------------------------------------------|
| BackColor      | FillFormat   | 背景填充色（用于阴影或图案填充格式）           |
| ForeColor      | FillFormat   | 前景填充色（对于纯色填充格式，即代表填充颜色） |
| BackColor      | LineFormat   | 线条背景色（用于图案线条）                     |
| ForeColor      | LineFormat   | 线条前景色（对于纯色线条，即代表线条颜色）     |
| ForeColor      | ShadowFormat | 阴影颜色                                       |
| ExtrusionColor | ThreeDFormat | 有延伸的对象的侧边颜色                         |

使用 RGB 属性可将颜色设置为显示的红-绿-蓝值。

``` JavaScript
/*下例向 myDocument 中添加一个矩形，然后设置矩形填充的前景色、背景色和渐变。*/
function test(){
let myDocument = Worksheets.Item(1)
let myFill=  myDocument.Shapes.AddShape(msoShapeRectangle, 
        90, 90, 90, 50).Fill
    myFill.ForeColor.RGB = (128, 0, 0)
    myFill.BackColor.RGB = (170, 170, 170)
    myFill.TwoColorGradient(msoGradientHorizontal, 1)
}
```

## 属性

| 序号 | 名称                                              | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ColorFormat.Application)           | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Brightness](#ColorFormat.Brightness)             | 返回或设置指定对象的发光度。可读/写。                                                                                                                                                                                           |
|  3   | [Creator](#ColorFormat.Creator)                   | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [ObjectThemeColor](#ColorFormat.ObjectThemeColor) | 返回或设置一种映射到主题配色方案的颜色。可读/写 MsoThemeColorIndex 。                                                                                                                                                           |
|  5   | [Parent](#ColorFormat.Parent)                     | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  6   | [RGB](#ColorFormat.RGB)                           | 返回或设置一个 Long 值，它代表指定颜色的红绿蓝 (RGB) 值。                                                                                                                                                                       |
|  7   | [SchemeColor](#ColorFormat.SchemeColor)           | 返回或设置一个 Integer 值，它代表 Color 对象的颜色，作为当前配色方案中的索引。                                                                                                                                                  |
|  8   | [TintAndShade](#ColorFormat.TintAndShade)         | 返回或设置一个 Single ，使颜色变深或变浅。                                                                                                                                                                                      |
|  9   | [Type](#ColorFormat.Type)                         | 返回一个代表颜色格式类型的 MsoColorType 值。                                                                                                                                                                                    |

## 成员属性

### ColorFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ColorFormat 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
let myObject = ActiveWorkbook
if(myObject.Application.Value == "ET") {
    alert("This is an ET Application object.")
}
else {
    alert("This is not an ET Application object.")
}
}
```

### ColorFormat.Brightness

返回或设置指定对象的发光度。可读/写。

**语法**

**express.Brightness**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

该属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数字。

**示例**

``` JavaScript
/*下面的代码示例设置活动工作表中第一个形状的填充色的亮度。*/
Application.ActiveSheet.Shapes.Item(1).Fill.ForeColor.Brightness = 0.5
```

### ColorFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ColorFormat 对象的变量。

### ColorFormat.ObjectThemeColor

返回或设置一种映射到主题配色方案的颜色。可读/写 **MsoThemeColorIndex** 。

**语法**

**express.ObjectThemeColor**

\*express是一个代表 ColorFormat 对象的变量。

### ColorFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ColorFormat 对象的变量。

### ColorFormat.RGB

返回或设置一个 **Long** 值，它代表指定颜色的红绿蓝 (RGB) 值。

**语法**

**express.RGB**

\*express是一个代表 ColorFormat 对象的变量。

### ColorFormat.SchemeColor

返回或设置一个 **Integer** 值，它代表 **Color** 对象的颜色，作为当前配色方案中的索引。

**语法**

**express.SchemeColor**

\*express是一个代表 ColorFormat 对象的变量。

### ColorFormat.TintAndShade

返回或设置一个 **Single** ，使颜色变深或变浅。

**语法**

**express.TintAndShade**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

### ColorFormat.Type

返回一个代表颜色格式类型的 **MsoColorType** 值。

**语法**

**express.Type**

\*express是一个代表 ColorFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ColorFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Dialogs 对象

## 简介

ET 中所有 Dialog 对象的集合。

## 说明

每个 Dialog 对象代表一个内置对话框。不能在集合中新建或添加内置对话框。 Dialog 对象只能在 Show 方法中用来显示相应的对话框。

ET Visual Basic 对象库包含许多内置对话框的内置常量。每个常量的前缀均为“xlDialog”，随后即为对话框的名称。例如， “应用名称” 对话框常量为 xlDialogApplyNames ， “查找文件” 对话框常量为 xlDialogFindFile 。这些常量是 XlBuiltinD ialog 枚举类型的成员。

使用 Dialogs 属性可返回 Dialogs 集合。下例显示可用的 ET 内置对话框的个数。

``` JavaScript
alert(Application.Dialogs.Count)
```

使用 Dialogs ( *index* )（其中 *index* 是用于标识对话框的内置常量）可返回单个 Dialog 对象。下例运行内置的 “打开文件” 对话框。

``` JavaScript
let dlgAnswer = Application.Dialogs.Item(xlDialogOpen).Show()
```

## 方法

| 序号 | 名称                  | 说明                   |
|:----:|:----------------------|:-----------------------|
|  1   | [Item](#Dialogs.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Dialogs.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Dialogs.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Dialogs.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Dialogs.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Dialogs.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Dialogs 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型        | 说明                          |
|---------|-----------|-----------------|-------------------------------|
| *Index* | 必选      | XlBuiltInDialog | Variant。对象的名称或索引号。 |

**示例**

``` JavaScript
/* 此示例显示“打开”对话框，并选定“只读”选项*/
Application.Dialogs.Item(xlDialogOpen).Show(null, null, true)
```

## 成员属性

### Dialogs.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Dialogs 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
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

### Dialogs.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Dialogs 对象的变量。

### Dialogs.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Dialogs 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Dialogs.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Dialogs 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Dialogs](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DialogSheetView 对象

## 简介

代表工作簿中的当前 对话框 工作表视图。

## 说明

若要访问此对象，活动工作簿中必须有已开发的对话框工作表。如果没有对话框工作表，该对象的视图属性将返回空字符串值。

``` JavaScript
/*以下示例将为活动工作簿启用对话框工作表视图。*/
Application.Worksheets.Item("Sheet1").DialogSheetView.Visible = true
```

## 属性

| 序号 | 名称                                        | 说明                                                                                                                                                               |
|:----:|:--------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#DialogSheetView.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#DialogSheetView.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Parent](#DialogSheetView.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  4   | [Sheet](#DialogSheetView.Sheet)             | 返回指定的 DialogSheetView 对象的工作表名称。只读。                                                                                                                |

## 成员属性

### DialogSheetView.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 DialogSheetView 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### DialogSheetView.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 DialogSheetView 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### DialogSheetView.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 DialogSheetView 对象的变量。

### DialogSheetView.Sheet

返回指定的 **DialogSheetView** 对象的工作表名称。只读。

**语法**

**express.Sheet**

\*express是一个代表 DialogSheetView 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/DialogSheetView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# GroupShapes 对象

## 简介

代表一组形状中的单个形状。

## 说明

每个形状都由一个 Shape 对象代表。将 Item 方法与此对象一起使用，您可以在不取消分组的情况下处理组合的各个形状。

使用 GroupItems 属性可返回 GroupShapes 集合。使用 GroupItems ( *index* )（其中 *index* 是分组的形状中单个形状的编号）可从 GroupShapes 集合中返回单个形状。

``` JavaScript
/*下例向 myDocument 添加三个三角形，将它们分成一组，设置整个组的颜色，然后只更改第二个三角形的颜色。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let shape1 = myDocument.Shapes
    shape1.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"
    shape1.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"
    shape1.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"
    let shape2 = shape1.Range(["shpOne", "shpTwo", "shpThree"]).Group()
        shape2.Fill.PresetTextured(msoTextureBlueTissuePaper)
        shape2.GroupItems.Item(2).Fill.PresetTextured(msoTextureGreenMarble)
}
```

## 方法

| 序号 | 名称                      | 说明                   |
|:----:|:--------------------------|:-----------------------|
|  1   | [Item](#GroupShapes.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#GroupShapes.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#GroupShapes.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#GroupShapes.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#GroupShapes.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  5   | [Range](#GroupShapes.Range)             | 返回一个 ShapeRange 对象，它代表 Shapes 集合中形状的子集。                                                                                                                                                                      |

## 成员方法

### GroupShapes.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 GroupShapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 Shape 对象。

**示例**

``` JavaScript
/*本示例对某个形状区域中第二个形状的 OnAction 属性进行设置。如果变量 sr 不代表 ShapeRange 对象，则本示例无效。*/
function test(){
let sr = Application.Worksheets.Item(1).Shapes
sr.Item(2).OnAction = "ShapeAction"
}
```

## 成员属性

### GroupShapes.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 GroupShapes 对象的变量。

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

### GroupShapes.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 GroupShapes 对象的变量。

### GroupShapes.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 GroupShapes 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### GroupShapes.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 GroupShapes 对象的变量。

### GroupShapes.Range

返回一个 **ShapeRange** 对象，它代表 **Shapes** 集合中形状的子集。

**语法**

**express.Range**

\*express是一个代表 GroupShapes 对象的变量。

**说明**

虽然使用 **Range** 属性可返回任意数量的形状，但如果要返回集合的单个成员，用 **Item** 方法更加简单。例如， ` Shapes(1) ` 比 ` Shapes.Range(1) ` 简单。

**示例**

``` JavaScript
/*此示例设置 myDocument 中第一个形状和第三个形状的填充图案。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
myDocument.Shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick)
}
```

``` JavaScript
/*若要为 Index 指定一个整数或字符串数组，可以使用 Array 函数。例如，以下指令返回用名称指定的两个形状。*/
function test(){
let arShapes = ["Oval 4", "Rectangle 5"]
let objRange = Application.ActiveSheet.Shapes.Range(arShapes)
}
```

``` JavaScript
/*在 ET 中，不能用此属性返回包含工作表上的所有 Shape 对象的 ShapeRange 对象。如果要达到该目的，可用下列代码： */
function test(){
Application.Worksheets.Item(1).Shapes.SelectAll     // select all shapes
let sr = Selection.ShapeRange           // create ShapeRange
}
```

``` JavaScript
/*此示例设置 myDocument 中形状“Oval 4”和“Rectangle 5”的填充图案。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let arShapes = ["Oval 4", "Rectangle 5"]
let objRange = myDocument.Shapes.Range(arShapes)
objRange.Fill.Patterned(msoPatternHorizontalBrick)
}
```

``` JavaScript
/*此示例设置 myDocument 中第一个形状的填充图案。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let myRange = myDocument.Shapes.Range(1)
myRange.Fill.Patterned(msoPatternHorizontalBrick)
}
```

``` JavaScript
/*此示例创建一个数组，其中包含 myDocument 中所有的自选图形，并用该数组定义一个形状区域，然后在该区域中水平分布所有的形状。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let shapes2 = myDocument.Shapes
    numShapes = shapes2.Count
    if(numShapes > 1) {
        let numAutoShapes = 0
        let autoShpArray = []
        for(let i = 1; i <= numShapes; i++) {
            if(shapes2.Item(i).Type == msoAutoShape) {
                autoShpArray[numAutoShapes] = shapes2.Item(i).Name
                numAutoShapes++
            }
        }
        if(numAutoShapes > 1) {
            let asRange = shapes2.Range(autoShpArray)
            asRange.Distribute(msoDistributeHorizontally, false)
        }
    }
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/GroupShapes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# LegendEntry 对象

## 简介

代表图表图例中的图例项。

## 说明

LegendEntry 对象是 LegendEntries 集合的成员。 LegendEntries 集合包含图例中所有的 LegendEntry 对象。

每个图例项都有两部分：一部分是该项的文本，它是与该图例项相关联的数据系列或趋势线的名称；另一部分是项标志，它在图表中以直观的方式将图例项以及与之相关联的数据系列或趋势线链接起来。项标志及其相关数据系列或趋势线的格式设置属性包含在 LegendKey 对象中。

不能修改图例项的文字。 LegendEntry 对象支持字体格式，且可被删除。图例项不支持图案格式。图例项的位置和尺寸是固定的。

没有返回相应于某图例项的数据系列或趋势线的直接方法。

在删除了图例项之后，唯一的恢复方法是：通过将图表的 HasLegend 属性设置为 False ，然后再设回 True ，从而删除并重新创建包含这些图例项的图例。

使用 LegendEntries ( *index* )（其中 *index* 是图例项索引号）可返回一个 LegendEntry 对象。不能按名称返回图例项。

图例项的编号代表图例项在图例中的位置。 ` LegendEntries(1) ` 位于图例的顶部； ` LegendEntries(LegendEntries.Count) ` 位于图例的底部。

``` JavaScript
/*下例更改工作表“Sheet1”上嵌入的第一个图表中图例顶部的图例项（这通常是第一个数据系列的图例）的文字字体。*/
Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.Legend.LegendEntries(1).Font.Italic = true
```

## 方法

| 序号 | 名称                          | 说明       |
|:----:|:------------------------------|:-----------|
|  1   | [Delete](#LegendEntry.Delete) | 删除对象。 |
|  2   | [Select](#LegendEntry.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#LegendEntry.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#LegendEntry.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Font](#LegendEntry.Font)               | 返回一个 Font 对象，它代表指定对象的字体。                                                                                                                                                                                      |
|  4   | [Format](#LegendEntry.Format)           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  5   | [Height](#LegendEntry.Height)           | 返回一个 Double 值，它代表对象的高度（以磅为单位）。                                                                                                                                                                            |
|  6   | [Index](#LegendEntry.Index)             | 返回 Long 值，它代表对象在其同类对象所组成的集合内的索引号。                                                                                                                                                                    |
|  7   | [Left](#LegendEntry.Left)               | 返回一个 Double 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。                                                                                                                                                      |
|  8   | [LegendKey](#LegendEntry.LegendKey)     | 返回一个 LegendKey 对象，该对象表示与图例项相关的图例标示。                                                                                                                                                                     |
|  9   | [Parent](#LegendEntry.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  10  | [Top](#LegendEntry.Top)                 | 返回一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                            |
|  11  | [Width](#LegendEntry.Width)             | 返回一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                            |

## 成员方法

### LegendEntry.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 LegendEntry 对象的变量。

**返回值**

Variant

### LegendEntry.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 LegendEntry 对象的变量。

**返回值**

Variant

## 成员属性

### LegendEntry.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 LegendEntry 对象的变量。

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

### LegendEntry.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 LegendEntry 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### LegendEntry.Font

返回一个 **Font** 对象，它代表指定对象的字体。

**语法**

**express.Font**

\*express是一个代表 LegendEntry 对象的变量。

### LegendEntry.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 LegendEntry 对象的变量。

### LegendEntry.Height

返回一个 **Double** 值，它代表对象的高度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 LegendEntry 对象的变量。

### LegendEntry.Index

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 LegendEntry 对象的变量。

### LegendEntry.Left

返回一个 **Double** 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 LegendEntry 对象的变量。

### LegendEntry.LegendKey

返回一个 **LegendKey** 对象，该对象表示与图例项相关的图例标示。

**语法**

**express.LegendKey**

\*express是一个代表 LegendEntry 对象的变量。

**示例**

``` JavaScript
/*本示例将 Chart1 上图例项一的图例标示设置为三角形。本示例应在二维折线图上运行。*/
Application.Charts.Item("Chart1").Legend.LegendEntries(1).LegendKey.MarkerStyle = xlMarkerStyleTriangle
```

### LegendEntry.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 LegendEntry 对象的变量。

### LegendEntry.Top

返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 LegendEntry 对象的变量。

### LegendEntry.Width

返回一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 LegendEntry 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/LegendEntry](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
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
# OLEObjects 对象

## 简介

指定工作表中所有 OLEObject 对象的集合。

## 说明

每一个 OLEObject 对象都代表一个 ActiveX 控件或一个链接或嵌入的 OLE 对象。

工作表上的 ActiveX 控件有两个名称：一个是包含该控件的形状的名称，查看工作表时可在 “名称” 框中看到它；另一个是控件的代码名称，可以在 “属性” 窗口中 “(名称)” 右边的单元格中看到它。在您首次向工作表添加控件时，形状名称和代码名称是一致的。但是，如果您更改这两个名称中的任意一个，另一个不会随之自动更改。

使用 OLEObjects 方法可返回 OLEObjects 集合。下例隐藏工作表一上的所有 OLE 对象。

``` JavaScript
Application.Worksheets.Item(1).OLEObjects().Visible = false
```

使用 Add 方法可创建一个新 OLE 对象并将它添加到 OLEObjects 集合中草药。下例创建一个新 OLE 对象（代表位图文件 Arcade.bmp）并将它添加到工作表一。

``` JavaScript
Application.Worksheets.Item(1).OLEObjects().Add ("arcade.gif")
```

下例新建一个 ActiveX 控件（一个列表框），并将其添加到第一张工作表中。

``` JavaScript
Application.Worksheets.Item(1).OLEObjects().Add ("Forms.ListBox.1")
```

在控件事件过程名称中使用的是控件的代码名称，但是，当您从工作表的 Shapes 或 OLEObjects 集合中返回控件时，必须使用形状名称而不是代码名称，以便按名称引用控件。例如，假定您给工作表添加一个复选框，而默认的形状名称和代码名称都是 CheckBox1。如果通过在 “属性” 窗口的 “(名称)” 旁边键入“chkFinished”而更改了控件代码名称，则在事件过程名称中必须使用 chkFinished，但是您仍然需要使用 CheckBox1 从 Shapes 或 OLEObject 集合中返回控件，如下例所示。

``` JavaScript
function chkFinished_Click(){
    Application.ActiveSheet.OLEObjects("CheckBox1").Object.Value = 1
}
```

## 方法

| 序号 | 名称                                     | 说明                                                                         |
|:----:|:-----------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Add](#OLEObjects.Add)                   | 向工作表中添加新的 OLE 对象。返回 一个 OLEObject 对象，它代表新的 OLE 对象。 |
|  2   | [BringToFront](#OLEObjects.BringToFront) | 将对象放到 z-次序前面。返回 Variant值                                        |
|  3   | [Copy](#OLEObjects.Copy)                 | 将对象复制到剪贴板。返回 Variant值                                           |
|  4   | [CopyPicture](#OLEObjects.CopyPicture)   | 将所选对象作为图片复制到剪贴板。返回 Variant值 。                            |
|  5   | [Cut](#OLEObjects.Cut)                   | 将对象剪切到剪贴板。返回Variant值                                            |
|  6   | [Delete](#OLEObjects.Delete)             | 删除对象。返回Variant值                                                      |
|  7   | [Duplicate](#OLEObjects.Duplicate)       | 复制对象，并返回对新复制对象的引用。返回 Object 值                           |
|  8   | [Item](#OLEObjects.Item)                 | 从集合中返回一个对象， 它代表集合包含的某个对象。                            |
|  9   | [Select](#OLEObjects.Select)             | 选择对象。返回Variant值                                                      |
|  10  | [SendToBack](#OLEObjects.SendToBack)     | 将对象放到 z-次序的后面。返回Variant值                                       |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#OLEObjects.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [AutoLoad](#OLEObjects.AutoLoad)       | 如果打开包含指定 OLE 对象的工作簿时自动载入该 OLE 对象，则为 True 。 Boolean 类型，可读写。                                                                                                                                     |
|  3   | [Border](#OLEObjects.Border)           | 返回一个 Border 对象，它代表对象的边框。                                                                                                                                                                                        |
|  4   | [Count](#OLEObjects.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  5   | [Creator](#OLEObjects.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  6   | [Enabled](#OLEObjects.Enabled)         | 如果启用对象，则为 True 。 Boolean 类型，可读写。                                                                                                                                                                               |
|  7   | [Height](#OLEObjects.Height)           | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |
|  8   | [Interior](#OLEObjects.Interior)       | 返回一个 Interior 对象，它代表指定对象的内部。                                                                                                                                                                                  |
|  9   | [Left](#OLEObjects.Left)               | 返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                       |
|  10  | [Locked](#OLEObjects.Locked)           | 返回或设置一个 Boolean 值，它指明对象是否已被锁定。                                                                                                                                                                             |
|  11  | [Parent](#OLEObjects.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  12  | [Placement](#OLEObjects.Placement)     | 返回或设置一个包含 XlPlacement 常量的 Variant 值，它代表对象附加到它所在的单元格的方式。                                                                                                                                        |
|  13  | [PrintObject](#OLEObjects.PrintObject) | 如果打印文档时也打印指定对象，则为 True 。 Boolean 类型，可读写。                                                                                                                                                               |
|  14  | [Shadow](#OLEObjects.Shadow)           | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                               |
|  15  | [ShapeRange](#OLEObjects.ShapeRange)   | 返回一个 ShapeRange 对象，它代表指定的一个或多个对象。只读。                                                                                                                                                                    |
|  16  | [SourceName](#OLEObjects.SourceName)   | 返回或设置一个 String 值，它代表指定对象的链接源名称。                                                                                                                                                                          |
|  17  | [Top](#OLEObjects.Top)                 | 返回或设置一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                      |
|  18  | [Visible](#OLEObjects.Visible)         | 返回或设置一个 Boolean 值，它确定对象是否可见。可读写。                                                                                                                                                                         |
|  19  | [Width](#OLEObjects.Width)             | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |
|  20  | [ZOrder](#OLEObjects.ZOrder)           | 返回指定对象的 z-次序位置。 Long 型，只读。                                                                                                                                                                                     |

## 成员方法

### OLEObjects.Add

向工作表中添加新的 OLE 对象。返回 一个 **OLEObject** 对象，它代表新的 OLE 对象。

**语法**

**express.Add ( *ClassType* , *Height* , *Link* , *DisplayAsIcon* , *IconFileName* , *IconIndex* , *IconLabel* , *Left* , *Width* )**

\*express是一个代表 OLEObjects 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                |
|-----------------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ClassType*     | 可选      | Variant  | （必须指定 ClassType 或 FileName）。一个字符串，包含要创建的对象的程序标识符。如果指定了 ClassType 参数，则忽略 FileName 和 Link。                                                                                  |
| *Height*        | 可选      | Variant  | （必须指定 ClassType 或 FileName）。一个字符串，指定用于创建 OLE 对象的文件。                                                                                                                                       |
| *Link*          | 可选      | Variant  | 如果为 True，则让基于 FileName 的新 OLE 对象链接到该文件。如果该对象未链接到文件，则该对象被创建为文件副本。默认值是 False。                                                                                        |
| *DisplayAsIcon* | 可选      | Variant  | 如果为 True，则以图标或正常图片方式显示新的 OLE 对象。如果该参数设置为 True，则可以使用 IconFileName 和 IconIndex 来指定图标。                                                                                      |
| *IconFileName*  | 可选      | Variant  | 一个字符串，指定要显示的图标所在的文件。仅当 DisplayAsIcon 为 True 时，才使用该参数。如果不指定该参数，或文件中不包含图标，则使用 OLE 类的默认图标。                                                                |
| *IconIndex*     | 可选      | Variant  | 图标文件中包含的图标数目。仅当 DisplayAsIcon 参数为 True 并且 IconFileName 参数引用包含图标的有效文件时，才使用该参数。如果由 IconFileName 参数指定的文件中不存在具有指定索引号的图标，则使用该文件中的第一个图标。 |
| *IconLabel*     | 可选      | Variant  | 一个字符串，指定在图标下方显示一个标签。仅当 DisplayAsIcon 为 True 时，才使用该参数。如果省略该参数，或者该参数为空字符串 ("")，则不显示任何标题。                                                                  |
| *Left*          | 可选      | Variant  | 以磅为单位给出新对象的初始坐标，该坐标是相对于工作表上单元格 A1 的左上角或图表的左上角的坐标。                                                                                                                      |
| *Width*         | 可选      | Variant  | 以磅为单位给出新对象的初始大小。                                                                                                                                                                                    |

**示例**

``` JavaScript
funtion test(){
  /*在工作表 Sheet1 中新建一个 WPS OLE 对象。*/
  Application.ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects().Add("WPS.Document")
  /*为第一张工作表添加一个命令按钮。*/
  Application.Worksheets.Item(1).OLEObjects().Add("Forms.CommandButton.1", false, false, 40, 40, 150, 10)
  
}
```

### OLEObjects.BringToFront

将对象放到 z-次序前面。返回 Variant值

**语法**

**express.BringToFront ()**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Copy

将对象复制到剪贴板。返回 Variant值

**语法**

**express.Copy ()**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.CopyPicture

将所选对象作为图片复制到剪贴板。返回 **Variant值** 。

**语法**

**express.CopyPicture ( *Appearance* , *Format* )**

\*express是一个代表 OLEObjects 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型            | 说明                                    |
|--------------|-----------|---------------------|-----------------------------------------|
| *Appearance* | 可选      | XlPictureAppearance | 指定图片的复制方式。默认值为 xlScreen。 |
| *Format*     | 可选      | XlCopyPictureFormat | 图片的格式。默认值为 xlPicture。        |

### OLEObjects.Cut

将对象剪切到剪贴板。返回Variant值

**语法**

**express.Cut ()**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Delete

删除对象。返回Variant值

**语法**

**express.Delete ()**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Duplicate

复制对象，并返回对新复制对象的引用。返回 Object 值

**语法**

**express.Duplicate ()**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Item

从集合中返回一个对象， 它代表集合包含的某个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 OLEObjects 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 可选      | Variant  | 对象的名称或索引号。 |

**说明**

对象的文本名称就是 **Name** 和 **Value** 属性的值。

**示例**

``` JavaScript
/*此示例将工作表 Sheet1 中的第一个 OLE 对象删除。*/
Application.Worksheets.Item("sheet1").OLEObjects().Item(1).Delete()
```

### OLEObjects.Select

选择对象。返回Variant值

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 OLEObjects 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                            |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | 如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

### OLEObjects.SendToBack

将对象放到 z-次序的后面。返回Variant值

**语法**

**express.SendToBack ()**

\*express是一个代表 OLEObjects 对象的变量。

## 成员属性

### OLEObjects.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 OLEObjects 对象的变量。

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

### OLEObjects.AutoLoad

如果打开包含指定 OLE 对象的工作簿时自动载入该 OLE 对象，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoLoad**

\*express是一个代表 OLEObjects 对象的变量。

**说明**

ActiveX 忽略此属性。打开一个工作簿时总会载入 ActiveX 控件。

对于大多数 OLE 对象类型，此属性不能设为 **True** 。对于新 OLE 对象，默认情况下其 **AutoLoad** 属性设为 **False** ；当 ET 载入工作簿时，设为“False”可节省时间和内存。自动载入 OLE 对象的好处在于，对于代表易变动的数据的对象，可立即重建到数据源的链接，而且，如果需要，可对这些对象进行重新映射。

**示例**

``` JavaScript
/*此示例对活动工作表中第一个 OLE 对象的 AutoLoad 属性进行设置。*/
Application.ActiveSheet.OLEObjects(1).AutoLoad = false
```

### OLEObjects.Border

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 OLEObjects 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### OLEObjects.Enabled

如果启用对象，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Enabled**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Height

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Interior

返回一个 **Interior** 对象，它代表指定对象的内部。

**语法**

**express.Interior**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Left

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Locked

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

\*express是一个代表 OLEObjects 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True** ；如果在工作表处于受保护状态时仍能修改对象，则返回 **False** 。

### OLEObjects.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Placement

返回或设置一个包含 **XlPlacement** 常量的 **Variant** 值，它代表对象附加到它所在的单元格的方式。

**语法**

**express.Placement**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.PrintObject

如果打印文档时也打印指定对象，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.PrintObject**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.ShapeRange

返回一个 **ShapeRange** 对象，它代表指定的一个或多个对象。只读。

**语法**

**express.ShapeRange**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.SourceName

返回或设置一个 **String** 值，它代表指定对象的链接源名称。

**语法**

**express.SourceName**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Top

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Visible

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.Width

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 OLEObjects 对象的变量。

### OLEObjects.ZOrder

返回指定对象的 z-次序位置。 **Long** 型，只读。

**语法**

**express.ZOrder**

\*express是一个代表 OLEObjects 对象的变量。

**说明**

在任何对象集合中，z-次序尾端的对象为 *collection* (1)，z-次序前端的对象为 *collection* ( *collection* . **Count** )。例如，如果活动工作表中有嵌入图表，z-次序尾端的图表为 ` ActiveSheet.ChartObjects(1) ` ，z-次序前端的图表为 ` ActiveSheet.ChartObjects(ActiveSheet.ChartObjects.Count) ` 。

**示例**

``` JavaScript
/*此示例显示 Sheet1 上嵌入的第一张图表的 z-次序位置。*/
alert("The chart's z-order position is " + Application.Worksheets.Item("Sheet1").ChartObjects(1).ZOrder
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/OLEObjects](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Page 对象

## 简介

代表工作表的页面。使用 PageSetup 对象及相关方法和属性可通过编程方式定义工作簿的页面布局。

## 说明

使用 Item 方法可访问工作簿的特定页面。下面的示例访问活动工作簿的第一页。

``` JavaScript
let objPage = Application.ActiveWindow.Panes.Item(1).Pages.Item(1)
```

## 属性

| 序号 | 名称                               | 说明                                 |
|:----:|:-----------------------------------|:-------------------------------------|
|  1   | [CenterFooter](#Page.CenterFooter) | 指定要在页脚中居中对齐的图片或文本。 |
|  2   | [CenterHeader](#Page.CenterHeader) | 指定要在页眉中居中对齐的图片或文本。 |
|  3   | [LeftFooter](#Page.LeftFooter)     | 指定要在页脚中左对齐的图片或文本。   |
|  4   | [LeftHeader](#Page.LeftHeader)     | 指定要在页眉中左对齐的图片或文本。   |
|  5   | [RightFooter](#Page.RightFooter)   | 指定要在页脚中右对齐的图片或文本。   |
|  6   | [RightHeader](#Page.RightHeader)   | 指定要在页眉中右对齐的图片或文本。   |

## 成员属性

### Page.CenterFooter

指定要在页脚中居中对齐的图片或文本。

**语法**

**express.CenterFooter**

\*express是一个代表 Page 对象的变量。

### Page.CenterHeader

指定要在页眉中居中对齐的图片或文本。

**语法**

**express.CenterHeader**

\*express是一个代表 Page 对象的变量。

### Page.LeftFooter

指定要在页脚中左对齐的图片或文本。

**语法**

**express.LeftFooter**

\*express是一个代表 Page 对象的变量。

### Page.LeftHeader

指定要在页眉中左对齐的图片或文本。

**语法**

**express.LeftHeader**

\*express是一个代表 Page 对象的变量。

### Page.RightFooter

指定要在页脚中右对齐的图片或文本。

**语法**

**express.RightFooter**

\*express是一个代表 Page 对象的变量。

### Page.RightHeader

指定要在页眉中右对齐的图片或文本。

**语法**

**express.RightHeader**

\*express是一个代表 Page 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Page](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotLine 对象

## 简介

table { word-break:break-all; }

PivotLine 对象是 ET 数据透视表中的行或列的线条。

## 说明

table { word-break:break-all; }

PivotLine 只包含可见项，因此 PivotLine 集合中不存在折叠的项目子项以及隐藏级别中的项目。

PivotLine 在所有位置始终具有一个 PivotItem。这意味着与普通 PivotLine 相比，代表数据透视表中分类汇总的 PivotLine 包含较少的 PivotItem。

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
<td style="text-align: left;" data-valian="middle"><a href="#PivotLine.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 <span>Application</span> 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLine.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLine.LineType">LineType</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个表示 PivotLine 类型的 <span>XlPivotLineType</span> 常量。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLine.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定 PivotLine 对象的父对象。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLine.PivotLineCells">PivotLineCells</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回 PivotLine 中 PivotCell 对象的集合。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLine.Position">Position</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置 PivotLine 对象的位置。只读。</p></td>
</tr>
</tbody>
</table>

## 成员属性

### PivotLine.Application

table { word-break:break-all; }

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotLine 对象的变量。

**说明**

table { word-break:break-all; }

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### PivotLine.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotLine 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotLine.LineType

table { word-break:break-all; }

返回一个表示 PivotLine 类型的 **XlPivotLineType** 常量。只读。

**语法**

**express.LineType**

\*express是一个代表 PivotLine 对象的变量。

### PivotLine.Parent

table { word-break:break-all; }

返回指定 **PivotLine** 对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotLine 对象的变量。

### PivotLine.PivotLineCells

table { word-break:break-all; }

返回 PivotLine 中 **PivotCell** 对象的集合。只读。

**语法**

**express.PivotLineCells**

\*express是一个代表 PivotLine 对象的变量。

### PivotLine.Position

table { word-break:break-all; }

返回或设置 **PivotLine** 对象的位置。只读。

**语法**

**express.Position**

\*express是一个代表 PivotLine 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotLine](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PlotArea 对象

## 简介

代表图表的绘图区。

## 说明

该区域为绘制图表数据的区域。二维图表中的绘图区包含数据标志、网格线、数据标签、趋势线和可选的置于图表区内的图表项。三维图表的绘图区中除包含上述各项外，还在图表中包含背景墙、基底、坐标轴、坐标轴标题和刻度线标签。绘图区被图表区所包围。二维图表的图表区包含坐标轴、图表标题、坐标轴标题和图例。三维图表的图表区包含图表标题和图例。有关设置图表区格式的详细信息，请参阅 ChartArea 对象。

## 方法

| 序号 | 名称                                   | 说明                 |
|:----:|:---------------------------------------|:---------------------|
|  1   | [ClearFormats](#PlotArea.ClearFormats) | 清除对象的格式设置。 |
|  2   | [Select](#PlotArea.Select)             | 选择对象。           |

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
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Format">Format</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 ChartFormat 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Height">Height</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.InsideHeight">InsideHeight</a></td>
<td style="text-align: left;" data-valian="middle"><p>以磅为单位返回绘图区内部高度。可读写 Double 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.InsideLeft">InsideLeft</a></td>
<td style="text-align: left;" data-valian="middle"><p>以 <span>磅</span> 为单位返回从图表边界到绘图区内部左边界的距离。可读写 Double 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.InsideTop">InsideTop</a></td>
<td style="text-align: left;" data-valian="middle"><p>以 <span>磅</span> 为单位返回从图表边界到绘图区内部上边界的距离。可读写 Double 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.InsideWidth">InsideWidth</a></td>
<td style="text-align: left;" data-valian="middle"><p>以 <span>磅</span> 为单位返回绘图区内部宽度。可读写 Double 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Left">Left</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Name">Name</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 String 值，它代表对象的名称。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Position">Position</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置图表上绘制区域的位置。可读/写 <span>XlChartElementPosition</span> 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Top">Top</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#PlotArea.Width">Width</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PlotArea.ClearFormats

清除对象的格式设置。

**语法**

**express.ClearFormats ()**

\*express是一个代表 PlotArea 对象的变量。

**返回值**

Variant

### PlotArea.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 PlotArea 对象的变量。

**返回值**

Varint

## 成员属性

### PlotArea.Application

如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PlotArea 对象的变量。

**示例**

``` JavaScript
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

### PlotArea.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PlotArea 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。Creator 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PlotArea.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Height

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.InsideHeight

以磅为单位返回绘图区内部高度。可读写 **Double** 类型。

**语法**

**express.InsideHeight**

\*express是一个代表 PlotArea 对象的变量。

**说明**

这种度量方式的图形区域不包含坐标轴标签。图形区域的 Height 属性使用包含坐标轴标签的封闭矩形。

**示例**

``` JavaScript
//本示例在 Chart1 中的绘图区内绘制带点线的矩形。
function test()
{
    let pa = Application.Charts.Item("chart1").PlotArea
    pa.Shapes.AddShape(msoShapeRectangle,pa.InsideLeft,pa.InsideTop,pa.InsideWidth,pa.InsideHeight)
    pa.Fill.Transparency = 1
    pa.Line.DashStyle = msoLineDashDot
}
```

### PlotArea.InsideLeft

以 磅 为单位返回从图表边界到绘图区内部左边界的距离。可读写 **Double** 类型。

**语法**

**express.InsideLeft**

\*express是一个代表 PlotArea 对象的变量。

**说明**

这种度量方式的图形区域不包含坐标轴标签。图形区域的 **Left** 属性使用包含坐标轴标签的封闭矩形。

**示例**

``` JavaScript
//本示例在 Chart1 中的绘图区内绘制带点线的矩形。
function test()
{
    let pa = Application.Charts.Item("chart1").PlotArea
    pa.Shapes.AddShape(msoShapeRectangle,pa.InsideLeft,pa.InsideTop,pa.InsideWidth,pa.InsideHeight)
    pa.Fill.Transparency = 1
    pa.Line.DashStyle = msoLineDashDot
}
```

### PlotArea.InsideTop

以 磅 为单位返回从图表边界到绘图区内部上边界的距离。可读写 **Double** 类型。

**语法**

**express.InsideTop**

\*express是一个代表 PlotArea 对象的变量。

**说明**

这种度量方式的图形区域不包含坐标轴标签。图形区域的 **Top** 属性使用包含坐标轴标签的封闭矩形。

**示例**

``` JavaScript
//本示例在 Chart1 中的绘图区内绘制带点线的矩形。
function test()
{
    let pa = Application.Charts.Item("chart1").PlotArea
    pa.Shapes.AddShape(msoShapeRectangle,pa.InsideLeft,pa.InsideTop,pa.InsideWidth,pa.InsideHeight)
    pa.Fill.Transparency = 1
    pa.Line.DashStyle = msoLineDashDot
}
```

### PlotArea.InsideWidth

以 磅 为单位返回绘图区内部宽度。可读写 **Double** 类型。

**语法**

**express.InsideWidth**

\*express是一个代表 PlotArea 对象的变量。

**说明**

这种度量方式的图形区域不包含坐标轴标签。图形区域的 **Width** 属性使用包含坐标轴标签的封闭矩形。

**示例**

``` JavaScript
//本示例在 Chart1 中的绘图区内绘制带点线的矩形。
function test()
{
    let pa = Application.Charts.Item("chart1").PlotArea
    pa.Shapes.AddShape(msoShapeRectangle,pa.InsideLeft,pa.InsideTop,pa.InsideWidth,pa.InsideHeight)
    pa.Fill.Transparency = 1
    pa.Line.DashStyle = msoLineDashDot
}
```

### PlotArea.Left

table { word-break:break-all; }

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Position

返回或设置图表上绘制区域的位置。可读/写 **XlChartElementPosition** 类型。

**语法**

**express.Position**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Top

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 PlotArea 对象的变量。

### PlotArea.Width

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 PlotArea 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PlotArea](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Protection 对象

## 简介

代表工作表可使用的各种保护选项类型。

## 说明

使用 Worksheet 对象的 Protection 属性可返回一个 Protection 对象。

返回一个 Protection 对象后，就可用该对象的下列属性来设置或返回保护选项。

- AllowDeletingColumns
- AllowDeletingRows
- AllowFiltering
- AllowFormattingCells
- AllowFormattingColumns
- AllowFormattingRows
- AllowInsertingColumns
- AllowInsertingHyperlinks
- AllowInsertingRows
- AllowSorting
- AllowUsingPivotTables

下例通过在最上面的行中放三个成员并保留该工作表说明了如何使用 Protection 对象的 AllowInserting.Columns 属性。

然后，此示例检查插入列的保护设置是否是 False ，如果必要，则将其设置为 True 。最后，通知用户插入一个列。

``` JavaScript
function SetProtection(){
    Range("A1").Formula = "1"
    Range("B1").Formula = "3"
    Range("C1").Formula = "4" Application.ActiveSheet.Protect()
    
    // Check the protection setting of the worksheet and act accordingly.
    if(Application.ActiveSheet.Protection.AllowInsertingColumns == false){ Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, true)
       alert("Insert a column between 1 and 3")
    }
    else{
       alert("Insert a column between 1 and 3")
    }
}
```

## 属性

| 序号 | 名称                                                             | 说明                                                                                  |
|:----:|:-----------------------------------------------------------------|:--------------------------------------------------------------------------------------|
|  1   | [AllowDeletingColumns](#Protection.AllowDeletingColumns)         | 如果允许删除受保护工作表上的列，则返回 True 。 Boolean 类型，只读。                   |
|  2   | [AllowDeletingRows](#Protection.AllowDeletingRows)               | 如果允许删除受保护工作表上的行，则返回 True 。 Boolean 类型，只读。                   |
|  3   | [AllowEditRanges](#Protection.AllowEditRanges)                   | 返回 AllowEditRanges 对象。                                                           |
|  4   | [AllowFiltering](#Protection.AllowFiltering)                     | 如果允许用户使用工作表受保护之前设置的“自动筛选”，则返回 True 。 Boolean 类型，只读。 |
|  5   | [AllowFormattingCells](#Protection.AllowFormattingCells)         | 如果允许对受保护的工作表上的单元格进行格式设置，则返回 True 。 Boolean 类型，只读。   |
|  6   | [AllowFormattingColumns](#Protection.AllowFormattingColumns)     | 如果允许对受保护的工作表上的列进行格式设置，则返回 True 。 Boolean 类型，只读。       |
|  7   | [AllowFormattingRows](#Protection.AllowFormattingRows)           | 如果允许对受保护的工作表上的行进行格式设置，则返回 True 。 Boolean 类型，只读。       |
|  8   | [AllowInsertingColumns](#Protection.AllowInsertingColumns)       | 如果允许在受保护的工作表上插入列，则返回 True 。 Boolean 类型，只读。                 |
|  9   | [AllowInsertingHyperlinks](#Protection.AllowInsertingHyperlinks) | 如果允许在受保护的工作表上插入超链接，则返回 True 。 Boolean 类型，只读。             |
|  10  | [AllowInsertingRows](#Protection.AllowInsertingRows)             | 如果允许用户在受保护的工作表上插入行，则返回 True 。 Boolean 类型，只读。             |
|  11  | [AllowSorting](#Protection.AllowSorting)                         | 如果允许在受保护的工作表上使用排序选项，则返回 True 。 Boolean 类型，只读。           |
|  12  | [AllowUsingPivotTables](#Protection.AllowUsingPivotTables)       | 如果允许用户在受保护的工作表上处理数据透视表，则返回 True 。 Boolean 类型，只读。     |

## 成员属性

### Protection.AllowDeletingColumns

如果允许删除受保护工作表上的列，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowDeletingColumns**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowDeletingColumns** 属性。

对于受保护的工作表，必须取消对包含要删除的单元格的列的锁定。

**示例**

``` JavaScript
function ProtectionOptions(){
    /*本示例取消对受保护的工作表上的列 A 的锁定，然后允许用户删除列 A 并通知用户。*/
    Application.ActiveSheet.Unprotect()
    
    // Unlock column A.
    Columns.Item("A:A").Locked = false
    
    // Allow column A to be deleted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowDeletingColumns == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("Column A can be deleted on this protected worksheet.")
}
```

### Protection.AllowDeletingRows

如果允许删除受保护工作表上的行，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowDeletingRows**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowDeletingRows** 属性。

对于受保护的工作表，必须取消对包含要删除的单元格的行的锁定。

**示例**

``` JavaScript
function ProtectionOptions(){
    /*本示例取消对受保护的工作表上的第 1 行的锁定，然后允许用户删除第 1 行并通知用户。*/
    Application.ActiveSheet.Unprotect()
    
    // Unlock row 1.
    Rows.Item("1:1").Locked = false
    
    // Allow row 1 to be deleted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowDeletingRows == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("Row 1 can be deleted on this protected worksheet.")
}
```

### Protection.AllowEditRanges

返回 **AllowEditRanges** 对象。

**语法**

**express.AllowEditRanges**

\*express是一个代表 Protection 对象的变量。

**示例**

``` JavaScript
function UseAllowEditRanges(){
    /*在本示例中，ET 允许用户编辑活动工作表上的区域 A1:A4，并将指定区域的标题和地址通知用户。*/
    let wksOne = Application.ActiveSheet

    // Unprotect worksheet.
    wksOne.Unprotect()

    // Establish a range that can allow edits on the protected worksheet.
    wksOne.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), "123")

    //Notify the user the title and address of the range.
    let WKSone = wksOne.Protection.AllowEditRanges.Item(1)

    MsgBox("Title of range: " + WKSone.Title)
    MsgBox("Address of range: " + WKSone.Range.Address)
}
```

### Protection.AllowFiltering

如果允许用户使用工作表受保护之前设置的“自动筛选”，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowFiltering**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowFiltering** 属性。

**AllowFiltering** 属性允许用户更改已有的“自动筛选”上的筛选条件。用户不能创建或删除受保护的工作表上的“自动筛选”。

若要筛选受保护的工作表上的单元格，则必须先取消对这些单元格的锁定。

**示例**

``` JavaScript
function ProtectionOptions(){
    /*本示例允许用户筛选受保护的工作表上的第 1 行并通知用户。*/
    Application.ActiveSheet.Unprotect()
        
    // Unlock row 1.
    Rows.Item("1:1").Locked = false
        
    // Allow row 1 to be filtered on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowFiltering == false){
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("Row 1 can be filtered on this protected worksheet.")
}
```

### Protection.AllowFormattingCells

如果允许对受保护的工作表上的单元格进行格式设置，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowFormattingCells**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowFormattingCells** 属性。

使用该属性可禁用“保护”选项卡，从而允许用户更改所有格式，但不能取消对区域的锁定或隐藏。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户对受保护的工作表上的单元格进行格式设置，并通知用户。*/
    Application.ActiveSheet.Unprotect()
    
    // Allow cells to be formatted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowFormattingCells == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, true)
    }
    alert("Cells can be formatted on this protected worksheet.")
}
```

### Protection.AllowFormattingColumns

如果允许对受保护的工作表上的列进行格式设置，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowFormattingColumns**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowFormattingColumns** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户对受保护的工作表上的列进行格式设置，并通知用户。*/
    Application.ActiveSheet.Unprotect()

    // Allow columns to be formatted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowFormattingColumns == false ){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, true)
    }
    alert("Columns can be formatted on this protected worksheet.")
}
```

### Protection.AllowFormattingRows

如果允许对受保护的工作表上的行进行格式设置，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowFormattingRows**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowFormattingRows** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户对受保护的工作表上的行进行格式设置，并通知用户。*/
    Application.ActiveSheet.Unprotect()

    // Allow rows to be formatted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowFormattingRows == false){
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, true)
    }
    alert("Rows can be formatted on this protected worksheet.")
}
```

### Protection.AllowInsertingColumns

如果允许在受保护的工作表上插入列，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowInsertingColumns**

\*express是一个代表 Protection 对象的变量。

**说明**

默认情况下，插入的列继承了其左边的列的格式设置，这意味着该列可能包含锁定的单元格。也就是说，用户可能不能删除插入的列。

可以使用 **Protect** 方法参数设置 **AllowInsertingColumns** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户在受保护的工作表上插入列，并通知用户。*/
    Application.ActiveSheet.Unprotect()
        
    // Allow columns to be inserted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowInsertingColumns == false){
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, true)
    }
    alert("Columns can be inserted on this protected worksheet.")
}
```

### Protection.AllowInsertingHyperlinks

如果允许在受保护的工作表上插入超链接，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowInsertingHyperlinks**

\*express是一个代表 Protection 对象的变量。

**说明**

在受保护的工作表上，只能将超链接插入未锁定或未保护的单元格中。

可以使用 **Protect** 方法参数设置 **AllowInsertingHyperlinks** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户在受保护的工作表上的单元格 A1 中插入超链接，并通知用户。*/
    Application.ActiveSheet.Unprotect()
        
    // Unlock cell A1.
    Range("A1").Locked = false
        
    // Allow hyperlinks to be inserted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowInsertingHyperlinks == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("Hyperlinks can be inserted on this protected worksheet.")
}
```

### Protection.AllowInsertingRows

如果允许用户在受保护的工作表上插入行，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowInsertingRows**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowInsertingRows** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户在受保护的工作表上插入行，并通知用户。*/
    Application.ActiveSheet.Unprotect()

    // Allow rows to be inserted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowInsertingRows == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, true)
    }
    alert("Rows can be inserted on this protected worksheet.")
}
```

### Protection.AllowSorting

如果允许在受保护的工作表上使用排序选项，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowSorting**

\*express是一个代表 Protection 对象的变量。

**说明**

在受保护的工作表中，只能对未锁定或未保护的单元格进行排序。

可以使用 Protect 方法参数设置 **AllowSorting** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户对受保护的工作表上未锁定或未保护的单元格进行排序，并通知用户。*/
    Application.ActiveSheet.Unprotect()

    // Unlock cells A1 through B5.
    Range("A1:B5").Locked = false

    // Allow sorting to be performed on the protected worksheet.
    if(Application.ActiveSheet.Protection.AllowSorting == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("For cells A1 through B5, sorting can be performed on the protected worksheet.")
}
```

### Protection.AllowUsingPivotTables

如果允许用户在受保护的工作表上处理数据透视表，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowUsingPivotTables**

\*express是一个代表 Protection 对象的变量。

**说明**

**AllowUsingPivotTables** 属性应用于 <span class="glossary" "appendpopup(this,'idh_xldefnonolap_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none;">非 OLAP 源数据</span> 。

可以使用 **Protect** 方法参数设置 **AllowUsingPivotTables** 属性。

**示例**

``` JavaScript
function test() {
    /*本示例允许用户访问数据透视表并通知用户。本示例假定非 OLAP 数据透视表位于活动的工作表上。*/
    Application.ActiveSheet.Unprotect()

    // Allow pivot tables to be manipulated on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowUsingPivotTables == false) {
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("Pivot tables can be manipulated on the protected worksheet.")
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Protection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Range 对象

## 简介

代表某一单元格、某一行、某一列、某一选定区域（该区域可包含一个或若干连续单元格区域），或者某一三维区域。

## 说明

示例部分中说明了以下用于返回 Range 对象的属性和方法：

- Range 属性
- Cells 属性
- Range 和 Cells
- Offset 属性
- Union 方法

## 示例

使用 Range ( *arg* )（其中 *arg* 为区域名称）可返回一个代表单个单元格或单元格区域的 Range 对象。下例将单元格 A1 中的值赋给单元格 A5。

``` JavaScript
Application.Worksheets.Item("Sheet1").Range("A5").Value2 = Application.Worksheets.Item("Sheet1").Range("A1").Value2 
```

下例通过为区域 A1:H8 中的每个单元格设置公式，用随机数字填充该区域。如果在不带对象识别符（句点左边的对象）的情况下使用 Range 属性，该属性会返回活动表上的一个区域。如果活动表不是工作表，则该方法失败。在使用没有显式对象识别符的 Range 属性之前，请先使用 Activate 方法激活一个工作表。

``` JavaScript
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    //Range is on the active sheet
    Application.Range("A1:H8").Formula = "=Rand()" 
}
```

下例清除区域名为“ *Criteria* ”的区域中的内容。 注释： 如果用文本参数指定区域地址，必须以 A1 样式记号指定该地址（不能用 R1C1 样式记号）。

``` JavaScript
Application.Worksheets.Item(1).Range("Criteria").ClearContents()
```

使用 Cells ( *row* , *column* )（其中 *row* 是行号， *column* 是列标）可返回一个单元格。下例将单元格 A1 赋值为 24。

``` JavaScript
Application.Worksheets.Item(1).Cells.Item(1, 1).Value2 = 24
```

下例设置单元格 A2 的公式。

``` JavaScript
Application.ActiveSheet.Cells.Item(2, 1).Formula = "=Sum(B1:B5)"
```

虽然也可用 Range("A1") 返回单元格 A1，但有时用 Cells 属性更为方便，因为对行或列使用变量。下例在 Sheet1 上创建行号和列标。注意，当工作表激活以后，使用 Cells 属性时不必明确声明工作表（它将返回活动工作表上的单元格）。 注释：虽然可用 Visual Basic 字符串函数转换 A1 样式引用，但使用 Cells(1, 1) 记号更为简便（而且也是更好的编程习惯）。

``` JavaScript
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    for(let i = 1; i <= 5; i++)
    {
        Application.Cells.Item(1, i + 1).Value2 = 1990 + i 
    } 

    for(let j = 1; j <= 4; j++)
    {
        Application.Cells.Item(j + 1, 1).Value2 = "Q" + j
    }
}
```

使用 *expression* . Cells ( *row* , *column* )（其中 *expression* 是返回 Range 对象的表达式， *row* 和 *column* 是相对于该区域左上角的偏移量）可返回区域中的一部分。下例设置单元格 C5 中的公式。

``` JavaScript
Application.Worksheets.Item(1).Range("C5:C10").Cells.Item(1, 1).Formula = "=Rand()"
```

使用 Range ( *cell1, cell2* )（其中 *cell1* 和 *cell2* 是指定起始和终止单元格的 Range 对象）可返回一个 Range 对象。下例设置单元格区域 A1:J10 的边框线条的样式。 注释： 请注意每个 Cells 属性之前的句点。如果前导的 With 语句应用于 Cells 属性，那么这些句点就是必需的。本示例中，句点指示单元格处于工作表一上。如果没有句点， Cells 属性将返回活动工作表上的单元格。

``` JavaScript
Application.Worksheets.Item(1).Range(Worksheets.Item(1).Cells.Item(1, 1),Application.Worksheets.Item(1).Cells.Item(10, 10)).Borders.LineStyle = xlThick
```

使用 Offset ( *row, column* )（其中 *row* 和 *column* 为行偏移量和列偏移量）可返回相对于另一区域在指定偏移量处的区域。下例选定位于当前选定区域左上角单元格的向下三行且向右一列处的单元格。由于必须选定位于活动工作表上的单元格，因此必须先激活工作表。

``` JavaScript
function test()
{
    Application.Worksheets.Item("Sheet1").Activate() 
    //Can't select unless the sheet is active 
    Application.Selection.Offset(3, 1).Range("A1").Select()
}
```

使用 Union ( *range1, range2* , ...) 可返回多块区域，即该区域由两个或多个连续的单元格区域所组成。下例创建由单元格区域 A1:B2 和 C3:D4 组合定义的对象，然后选定该定义区域。

``` JavaScript
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    let r1 = Range("A1:B2") let r2 = Range("C3:D4")
    let myMultiAreaRange = Union(r1, r2)
    Application.myMultiAreaRange.Select()
}
```

如果您处理包含多个区域的选定内容， Areas 属性是很有用的。它将多区域选定内容拆分为单个的 Range 对象，然后将对象返回为一个集合。您可以对返回的集合使用 Count 属性，以查找包含多个区域的选定内容，如下例所示。

``` JavaScript
function test()
{ 
    let NumberOfSelectedAreas = Application.Selection.Areas.Count
    if(NumberOfSelectedAreas > 1)
    {
        MsgBox("You cannot carry out this command " + "on multi-area selections")
    }
}
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
<td style="text-align: left;" data-valian="middle"><a href="#Range.Activate">Activate</a></td>
<td style="text-align: left;" data-valian="middle"><p>激活单个单元格，该单元格必须处于当前选定区域内。要选择单元格区域，请使用 Select 方法。返回值为Variant类型。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Address">Address</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 String 值，它代表宏语言的区域引用。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.ApplyNames">ApplyNames</a></td>
<td style="text-align: left;" data-valian="middle"><p><em>表达式</em> .ApplyNames( <em>Names</em> , <em>IgnoreRelativeAbsolute</em> , <em>UseRowColumnNames</em> , <em>OmitColumn</em> , <em>OmitRow</em> , <em>Order</em> , <em>AppendLast</em> )</p>
<p><em>表达式</em> 一个代表 Range 对象的变量。</p>
<p>返回值为Variant类型。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.AutoFit">AutoFit</a></td>
<td style="text-align: left;" data-valian="middle"><p>更改区域中的列宽或行高以达到最佳匹配。</p>
<p>返回值为Variant类型。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Calculate">Calculate</a></td>
<td style="text-align: left;" data-valian="middle"><p>计算所有打开的工作簿、工作簿中的某张特定工作表或工作表指定区域中的单元格，如下表所示。</p>
<p>返回值类型为Variant。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Clear">Clear</a></td>
<td style="text-align: left;" data-valian="middle"><p>清除整个对象。</p>
<p>返回值为Variant类型。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.ClearComments">ClearComments</a></td>
<td style="text-align: left;" data-valian="middle"><p>清除指定区域的所有单元格批注。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.ClearContents">ClearContents</a></td>
<td style="text-align: left;" data-valian="middle"><p>清除区域中的公式。</p>
<p>返回值类型为Variant。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.ClearFormats">ClearFormats</a></td>
<td style="text-align: left;" data-valian="middle"><p>清除对象的格式设置。</p>
<p>返回值类型为Variant。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Copy">Copy</a></td>
<td style="text-align: left;" data-valian="middle"><p>将单元格区域复制到指定的区域或剪贴板中。</p>
<p>返回值类型为Variant。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.CreateNames">CreateNames</a></td>
<td style="text-align: left;" data-valian="middle"><p>在指定区域中依据工作表中的文本标签创建名称。</p>
<p>返回值类型为Variant。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.CurrentRegion">CurrentRegion</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Range 对象，该对象表示当前区域。当前区域是以空行与空列的组合为边界的区域。只读。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Cut">Cut</a></td>
<td style="text-align: left;" data-valian="middle"><p>将对象剪切到剪贴板，或者将其粘贴到指定的目的地。</p>
<p>返回值类型为Variant。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Delete">Delete</a></td>
<td style="text-align: left;" data-valian="middle"><p>删除对象。</p>
<p>返回值类型为Variant。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Find">Find</a></td>
<td style="text-align: left;" data-valian="middle"><p>在区域中查找特定信息。</p>
<p>返回值为 一个 Range 对象，它代表第一个在其中找到该信息的单元格。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">16</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.FindNext">FindNext</a></td>
<td style="text-align: left;" data-valian="middle"><p>继续由 Find 方法开始的搜索。查找匹配相同条件的下一个单元格，并返回表示该单元格的 Range 对象。该操作不影响选定内容和活动单元格。</p>
<p>返回值类型为Range。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">17</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.FindPrevious">FindPrevious</a></td>
<td style="text-align: left;" data-valian="middle"><p>继续由 Find 方法开始的搜索。查找匹配相同条件的上一个单元格，并返回代表该单元格的 Range 对象。该操作不影响选定内容和活动单元格。</p>
<p>返回值类型为Range。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">18</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Insert">Insert</a></td>
<td style="text-align: left;" data-valian="middle"><p>在工作表或宏表中插入一个单元格或单元格区域，其他单元格相应移位以腾出空间。</p>
<p>返回值类型为Variant。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">19</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Range 对象，它代表对指定区域某一偏移量处的区域。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">20</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Justify">Justify</a></td>
<td style="text-align: left;" data-valian="middle"><p>调整区域内的文字，使之均衡地填充该区域。</p>
<p>返回值类型为Variant。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">21</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Merge">Merge</a></td>
<td style="text-align: left;" data-valian="middle"><p>由指定的 Range 对象创建合并单元格。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">22</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Offset">Offset</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Range 对象，它代表位于指定单元格区域的一定的偏移量位置上的区域。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">23</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.PrintPreview">PrintPreview</a></td>
<td style="text-align: left;" data-valian="middle"><p>按对象打印后的外观效果显示对象的预览。</p>
<p>返回值类型为Variant。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">24</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Range">Range</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Range 对象，它代表一个单元格或单元格区域。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">25</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Resize">Resize</a></td>
<td style="text-align: left;" data-valian="middle"><p>调整指定区域的大小。返回 Range 对象，该对象代表调整后的区域。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">26</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Run">Run</a></td>
<td style="text-align: left;" data-valian="middle"><p>在该处运行 ET 宏。区域必须位于宏表上。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">27</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Select">Select</a></td>
<td style="text-align: left;" data-valian="middle"><p>选择对象。</p>
<p>返回值类型为Variant。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">28</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.UnMerge">UnMerge</a></td>
<td style="text-align: left;" data-valian="middle"><p>将合并区域分解为独立的单元格。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">29</td>
<td style="text-align: left;" data-valian="middle"><a href="#Range.Value">Value</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Variant 型，它代表指定单元格的值。</p></td>
</tr>
</tbody>
</table>

## 属性

| 序号 | 名称                                              | 说明                                                                                                                                                                                                                                                                                                                                                    |
|:----:|:--------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddIndent](#Range.AddIndent)                     | 返回或设置一个 Variant 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。                                                                                                                                                                                                                                                   |
|  2   | [Application](#Range.Application)                 | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 Application 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 Application 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。                                                                                                                             |
|  3   | [Areas](#Range.Areas)                             | 返回一个 Areas 集合，该集合表示多重区域选择中的所有区域。只读。                                                                                                                                                                                                                                                                                         |
|  4   | [Cells](#Range.Cells)                             | 返回一个 Range 对象，该对象代表指定区域中的单元格。                                                                                                                                                                                                                                                                                                     |
|  5   | [Column](#Range.Column)                           | 返回指定区域中第一块中的第一列的列号。Long 类型，只读。                                                                                                                                                                                                                                                                                                 |
|  6   | [ColumnWidth](#Range.ColumnWidth)                 | 返回或设置指定区域中所有列的列宽。 Variant 类型，可读写。                                                                                                                                                                                                                                                                                               |
|  7   | [Columns](#Range.Columns)                         | 返回一个 Range 对象，该对象代表指定区域中的列。                                                                                                                                                                                                                                                                                                         |
|  8   | [Count](#Range.Count)                             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                                                                                                                                              |
|  9   | [Dependents](#Range.Dependents)                   | 返回一个 Range 对象，该对象表示包含一个单元格所有从属单元格的区域。如果有多个从属单元格，这可能有多个选择（多个 Range 对象）。 Range 类型，只读。                                                                                                                                                                                                       |
|  10  | [End](#Range.End)                                 | 返回一个 Range 对象，该对象代表包含源区域的区域尾端的单元格。等同于按键 End+ 向上键、End+ 向下键、End+ 向左键或 End+ 向右键。 Range 对象，只读。                                                                                                                                                                                                        |
|  11  | [EntireColumn](#Range.EntireColumn)               | 返回一个 Range 对象，该对象表示包含指定区域的整列（或多列）。只读。                                                                                                                                                                                                                                                                                     |
|  12  | [EntireRow](#Range.EntireRow)                     | 返回一个 Range 对象，该对象表示包含指定区域的整行（或多行）。只读。                                                                                                                                                                                                                                                                                     |
|  13  | [Font](#Range.Font)                               | 返回一个 Font 对象，它代表指定对象的字体。                                                                                                                                                                                                                                                                                                              |
|  14  | [Formula](#Range.Formula)                         | 返回或设置一个 Variant 值，它代表 A1 样式表示法和宏语言中的对象的公式。                                                                                                                                                                                                                                                                                 |
|  15  | [Height](#Range.Height)                           | 返回或设置一个 Variant 值，该值代表区域的高度（以磅为单位）。                                                                                                                                                                                                                                                                                           |
|  16  | [Hidden](#Range.Hidden)                           | 返回或设置一个 Variant 值，它指明是否隐藏行或列。                                                                                                                                                                                                                                                                                                       |
|  17  | [HorizontalAlignment](#Range.HorizontalAlignment) | 返回或设置一个 Variant 值，它代表指定对象的水平对齐方式。                                                                                                                                                                                                                                                                                               |
|  18  | [ID](#Range.ID)                                   | 返回或设置一个 String 值，它代表将页面另存为网页时指定单元格的识别标志。                                                                                                                                                                                                                                                                                |
|  19  | [Left](#Range.Left)                               | 返回一个 Variant 值，它代表从列 A 的左边缘到区域的左边缘的距离（以磅为单位）。                                                                                                                                                                                                                                                                          |
|  20  | [MergeArea](#Range.MergeArea)                     | 返回一个 Range 对象，该对象代表包含指定单元格的合并区域。如果指定的单元格不在合并区域内，则该属性返回指定的单元格。只读。 Variant 类型。                                                                                                                                                                                                                |
|  21  | [MergeCells](#Range.MergeCells)                   | 如果区域包含合并单元格，则为 True 。 Variant 型，可读写。                                                                                                                                                                                                                                                                                               |
|  22  | [Name](#Range.Name)                               | 返回或设置一个 Variant 值，它代表对象的名称。                                                                                                                                                                                                                                                                                                           |
|  23  | [NumberFormat](#Range.NumberFormat)               | 返回或设置一个 Variant 值，它代表对象的格式代码。                                                                                                                                                                                                                                                                                                       |
|  24  | [Parent](#Range.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                                                                                                                                            |
|  25  | [Previous](#Range.Previous)                       | 返回一个代表下一个单元格的 Range 对象。                                                                                                                                                                                                                                                                                                                 |
|  26  | [Row](#Range.Row)                                 | 返回区域中第一个子区域的第一行的行号。 Long 类型，只读。                                                                                                                                                                                                                                                                                                |
|  27  | [RowHeight](#Range.RowHeight)                     | 以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置指定区域中所有行的行高。如果指定区域中的各行的行高不等，则返回 null 。 Variant 类型，可读写。 |
|  28  | [Rows](#Range.Rows)                               | 返回一个 Range 对象，它代表指定单元格区域中的行。 Range 对象，只读。                                                                                                                                                                                                                                                                                    |
|  29  | [Text](#Range.Text)                               | 返回或设置指定对象中的文本。 String 型，只读。                                                                                                                                                                                                                                                                                                          |
|  30  | [Top](#Range.Top)                                 | 返回或设置一个 Variant 值，它代表行 1 上边缘到区域上边缘的距离（以磅为单位）。                                                                                                                                                                                                                                                                          |
|  31  | [Value2](#Range.Value2)                           | 返回或设置单元格值。 Variant 类型，可读写。                                                                                                                                                                                                                                                                                                             |
|  32  | [Width](#Range.Width)                             | 返回一个 Variant 值，它代表区域的宽度（以单位表示）。                                                                                                                                                                                                                                                                                                   |
|  33  | [Worksheet](#Range.Worksheet)                     | 返回一个 Worksheet 对象，该对象表示包含指定区域的工作表。只读。                                                                                                                                                                                                                                                                                         |

## 成员方法

### Range.Activate

激活单个单元格，该单元格必须处于当前选定区域内。要选择单元格区域，请使用 **Select** 方法。返回值为Variant类型。

**语法**

**express.Activate ()**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例选定工作表 Sheet1 上的单元格区域 A1:C3，并激活单元格 B2。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.Range("A1:C3").Select()
    Application.Range("B2").Activate()
}
```

### Range.Address

返回一个 **String** 值，它代表宏语言的区域引用。

**语法**

**express.Address ( *RowAbsolute* , *ColumnAbsolute* , *ReferenceStyle* , *External* , *RelativeTo* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型         | 说明                                                                                                                                      |
|------------------|-----------|------------------|-------------------------------------------------------------------------------------------------------------------------------------------|
| *RowAbsolute*    | 可选      | Variant          | 如果为 True，则以绝对引用返回引用的行部分。默认值为 True。                                                                                |
| *ColumnAbsolute* | 可选      | Variant          | 如果为 True，则以绝对引用返回引用的列部分。默认值为 True。                                                                                |
| *ReferenceStyle* | 可选      | XlReferenceStyle | 引用样式。默认值为 xlA1。                                                                                                                 |
| *External*       | 可选      | Variant          | 如果为 True，则返回外部引用。如果为 False，则返回本地引用。默认值为 False。                                                               |
| *RelativeTo*     | 可选      | Variant          | 如果 RowAbsolute 和 ColumnAbsolute 为 False，并且 ReferenceStyle 为 xlR1C1，则必须包括相对引用的起始点。此参数是定义起始点的 Range 对象。 |

**说明**

如果引用包含多个单元格， *RowAbsolute* 和 *ColumnAbsolute* 将应用于所有的行和列。

**示例**

``` JavaScript
/*下例对工作表 Sheet1 中的同一单元格地址使用了四种不同的表达方式。示例中的注释为将要显示于消息框中的地址。*/
function test()
{
    let mc = Application.Worksheets.Item("Sheet1").Cells.Item(1, 1)
    MsgBox(mc.Address())                                    // $A$1
    MsgBox(mc.Address(false))                               // $A1
    MsgBox(mc.Address(undefined,undefined,xlR1C1))          // R1C1
    MsgBox(mc.Address(false,false,xlR1C1,undefined,Worksheets.Item(1).Cells.Item(3, 3)))      // R[-2]C[-2]
}
```

### Range.ApplyNames

***表达式* .ApplyNames( *Names* , *IgnoreRelativeAbsolute* , *UseRowColumnNames* , *OmitColumn* , *OmitRow* , *Order* , *AppendLast* )**

*表达式* 一个代表 **Range** 对象的变量。

返回值为Variant类型。

**语法**

**express.ApplyNames ( *Names* , *IgnoreRelativeAbsolute* , *UseRowColumnNames* , *OmitColumn* , *OmitRow* , *Order* , *AppendLast* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型          | 说明                                                                                                                                                                                                                                                |
|--------------------------|-----------|-------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Names*                  | 可选      | Variant           | 要应用的名称数组。如果省略该参数，则工作表中所有的名称都将应用到该区域上。                                                                                                                                                                          |
| *IgnoreRelativeAbsolute* | 可选      | Variant           | 要应用的名称数组。如果省略该参数，则工作表中所有的名称都将应用到该区域上。 如果为 True，则用名称取代引用，而不管名称或引用的类型如何。如果为 False，则只用绝对名称替换绝对引用，用相对名称替换相对引用，并只用混合名称替换混合引用。默认值为 True。 |
| *UseRowColumnNames*      | 可选      | Variant           | 如果为 True，则当无法找到指定区域的名称时，使用该区域所在行或列区域的名称。如果为 False，则忽略 OmitColumn 和 OmitRow 参数。默认值为 True。                                                                                                         |
| *OmitColumn*             | 可选      | Variant           | 如果为 True，则用行方向的名称替换整个引用。仅当引用的单元格与公式位于同一列中，且处于行方向命名的区域中时，才能省略列方向名称。默认值为 True。                                                                                                      |
| *OmitRow*                | 可选      | Variant           | 如果为 True，则用列方向的名称替换整个引用。仅当引用的单元格与公式位于同一行中，且处于列方向命名的区域中时，才能省略行方向名称。默认值为 True。                                                                                                      |
| *Order*                  | 可选      | XlApplyNamesOrder | 确定用行方向区域名称和列方向区域名称替换单元格引用时，首先列出哪个区域名称。                                                                                                                                                                        |
| *AppendLast*             | 可选      | Variant           | 如果为 True，则替换 Names 中的名称定义以及所定义的姓氏的定义。如果为 False，则只替换 Names 中的名称定义。默认值为 False。                                                                                                                           |

**说明**

可用 **Array** 函数为 *Names* 参数创建名称列表。

如果要对整个工作表应用名称，可用 Cells.ApplyNames。

不能“取消应用”名称；要删除名称，请使用 **Delete** 方法。

**示例**

``` JavaScript
/*本示例对整个工作表应用名称。*/
Application.Cells.ApplyNames(["Sales","Profits"])
```

### Range.AutoFit

更改区域中的列宽或行高以达到最佳匹配。

返回值为Variant类型。

**语法**

**express.AutoFit ()**

\*express是一个代表 Range 对象的变量。

**说明**

**Range** 对象必须是行或行区域，或者列或列区域。否则，该方法将产生错误。

一个列宽单位等于“常规”样式中一个字符的宽度。

**示例**

``` JavaScript
/*本示例调整工作表 Sheet1 中从 A 到 I 的列，以获得最适当的列宽。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Columns.Item("A:I").AutoFit()
}

/*本示例调整工作表 Sheet1 上从 A 到 E 的列，以获得最适当的列宽，但该调整仅依据单元格区域 A1:E1 中的内容进行。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("A1:E1").Columns.AutoFit()
}
```

### Range.Calculate

计算所有打开的工作簿、工作簿中的某张特定工作表或工作表指定区域中的单元格，如下表所示。

返回值类型为Variant。

**语法**

**express.Calculate ()**

\*express是一个代表 Range 对象的变量。

**说明**

| 要计算           | 依照此示例                                   |
|------------------|----------------------------------------------|
| 所有打开的工作簿 | Application.Calculate()（或只是Calculate()） |
| 指定工作表       | Worksheets.Item(1).Calculate()               |
| 指定区域         | Worksheets.Item(1).Rows.Item(2).Calculate()  |

**示例**

``` JavaScript
/*此示例计算 Sheet1 上已用区域中 A 列、B 列和 C 列的公式。*/
function test()
{
    Application.Worksheets.Item("Sheet1").UsedRange.Columns.Item("A:C").Calculate()
}
```

### Range.Clear

清除整个对象。

返回值为Variant类型。

**语法**

**express.Clear ()**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*此示例清除工作表 Sheet1 中 A1:G37 单元格区域的公式和格式设置。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("A1:G37").Clear()
}
```

### Range.ClearComments

清除指定区域的所有单元格批注。

**语法**

**express.ClearComments ()**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例清除 E5 单元格的所有批注。*/
Application.Worksheets.Item(1).Range("E5").ClearComments()
```

### Range.ClearContents

清除区域中的公式。

返回值类型为Variant。

**语法**

**express.ClearContents ()**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*此示例清除工作表 Sheet1 上 A1:G37 单元格区域的公式，但保留其格式设置。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("A1:G37").ClearContents()
}
```

### Range.ClearFormats

清除对象的格式设置。

返回值类型为Variant。

**语法**

**express.ClearFormats ()**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*此示例清除工作表 Sheet1 上 A1:G37 单元格区域的所有格式设置。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("A1:G37").ClearFormats()
}

/*此示例清除工作表 Sheet1 上第一个嵌入式图表的格式设置。*/
function test()
{
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.ClearFormats()
}
```

### Range.Copy

将单元格区域复制到指定的区域或剪贴板中。

返回值类型为Variant。

**语法**

**express.Copy ( *Destination* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                              |
|---------------|-----------|----------|-------------------------------------------------------------------|
| *Destination* | 可选      | Variant  | 指定区域要复制到的新域。如果省略此参数，ET 会将区域复制到剪贴板。 |

**示例**

``` JavaScript
/*此示例将工作表 Sheet1 上单元格区域 A1:D4 中的公式复制到工作表 Sheet2 上的单元格区域 E5:H8 中。*/
Application.Worksheets.Item("Sheet1").Range("A1:D4").Copy(Worksheets.Item("Sheet2").Range("E5"))
```

### Range.CreateNames

在指定区域中依据工作表中的文本标签创建名称。

返回值类型为Variant。

**语法**

**express.CreateNames ( *Top* , *Left* , *Bottom* , *Right* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                            |
|----------|-----------|----------|-----------------------------------------------------------------|
| *Top*    | 可选      | Variant  | 如果该值为 True，则使用顶部行中的标签创建名称。默认值为 False。 |
| *Left*   | 可选      | Variant  | 如果该值为 True，则使用左侧列中的标签创建名称。默认值为 False。 |
| *Bottom* | 可选      | Variant  | 如果该值为 True，则使用底部行中的标签创建名称。默认值为 False。 |
| *Right*  | 可选      | Variant  | 如果该值为 True，则使用右侧列中的标签创建名称。默认值为 False。 |

**说明**

如果未指定 *Top* 、 *Left* 、 *Bottom* 或 *Right* 之一，ET 将根据指定区域的形状猜测文本标志。

**示例**

``` JavaScript
/*本示例用单元格区域 A1:A3 中的文字创建区域 B1:B3 的名称。注意，指定区域时必须包括那些含有名称文字的单元格，即便只是为单元格区域 B1:B3 创建名称。*/
function test()
{
    let rangeToName = Application.Worksheets.Item("Sheet1").Range("A1:B3")
    rangeToName.CreateNames(null, true)
}
```

### Range.CurrentRegion

返回一个 **Range** 对象，该对象表示当前区域。当前区域是以空行与空列的组合为边界的区域。只读。

**语法**

**express.CurrentRegion ()**

\*express是一个代表 Range 对象的变量。

**说明**

该属性对于许多自动展开选择以包括整个当前区域的操作很有用，如 **AutoFormat** 方法。

该属性不能用于被保护的工作表。

**示例**

``` JavaScript
/*本示例选定工作表 Sheet1 上的当前区域。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.CurrentRegion.Select()
}

/*本示例假定在 Sheet1 中有一个包含标题行的表。本示例选定该表，但不选定标题行。运行本示例之前，活动单元格必须处于该表中。*/
function test()
{
    let tbl = Application.ActiveCell.CurrentRegion
    tbl.Offset(1, 0).Resize(tbl.Rows.Count-1,tbl.Columns.Count).Select()
}
```

### Range.Cut

将对象剪切到剪贴板，或者将其粘贴到指定的目的地。

返回值类型为Variant。

**语法**

**express.Cut ( *Destination* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                   |
|---------------|-----------|----------|------------------------------------------------------------------------|
| *Destination* | 可选      | Variant  | 应在其中粘贴对象的目标区域。如果省略此参数，区域对象会被剪切到剪贴板。 |

**说明**

剪切的区域必须由相邻的单元格组成。

### Range.Delete

删除对象。

返回值类型为Variant。

**语法**

**express.Delete ( *Shift* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                               |
|---------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Shift* | 可选      | Variant  | 仅用于 Range 对象。指定如何调整单元格以替换删除的单元格。可以为以下 XlDeleteShiftDirection 常量之一：xlShiftToLeft 或 xlShiftUp。如果省略此参数，ET 将根据区域的形状确定调整方式。 |

### Range.Find

在区域中查找特定信息。

返回值为 一个 **Range** 对象，它代表第一个在其中找到该信息的单元格。

**语法**

**express.Find ( *What* , *After* , *LookIn* , *LookAt* , *SearchOrder* , *SearchDirection* , *MatchCase* , *MatchByte* , *SearchFormat* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型          | 说明                                                                                                                                                                                                                                                                     |
|-------------------|-----------|-------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *What*            | 可选      | Variant           | 要搜索的数据。可为字符串或任意 ET 数据类型。                                                                                                                                                                                                                             |
| *After*           | 可选      | Variant           | 表示搜索过程将从其之后开始进行的单元格。此单元格对应于从用户界面搜索时的活动单元格的位置。请注意：After 必须是区域中的单个单元格。要记住搜索是从该单元格之后开始的；直到此方法绕回到此单元格时，才对其进行搜索。如果不指定该参数，搜索将从区域的左上角的单元格之后开始。 |
| *LookIn*          | 可选      | Variant           | 信息类型。                                                                                                                                                                                                                                                               |
| *LookAt*          | 可选      | Variant           | 可为以下 XlLookAt 常量之一：xlWhole 或 xlPart。                                                                                                                                                                                                                          |
| *SearchOrder*     | 可选      | Variant           | 可为以下 XlSearchOrder 常量之一：xlByRows 或 xlByColumns。                                                                                                                                                                                                               |
| *SearchDirection* | 可选      | XlSearchDirection | 搜索的方向。                                                                                                                                                                                                                                                             |
| *MatchCase*       | 可选      | Variant           | 如果为 True，则搜索区分大小写。默认值为 False。                                                                                                                                                                                                                          |
| *MatchByte*       | 可选      | Variant           | 只在已经选择或安装了双字节语言支持时适用。如果为 True，则双字节字符只与双字节字符匹配。如果为 False，则双字节字符可与其对等的单字节字符匹配。                                                                                                                            |
| *SearchFormat*    | 可选      | Variant           | 搜索的格式。                                                                                                                                                                                                                                                             |

**说明**

如果未发现匹配项，则返回 **Nothing** 。 **Find** 方法不影响选定区域或当前活动的单元格。

每次使用此方法后，参数 *LookIn* 、 *LookAt* 、 *SearchOrder* 和 *MatchByte* 的设置都将被保存。如果下次调用此方法时不指定这些参数的值，就使用保存的值。设置这些参数将更改 **“查找”** 对话框中的设置，如果省略这些参数，更改 **“查找”** 对话框中的设置将更改使用的保存值。要避免出现这一问题，每次使用此方法时请明确设置这些参数。

使用 **FindNext** 和 **FindPrevious** 方法可重复搜索。

当搜索到指定查找区域的末尾时，此方法将绕回到区域的开始位置继续搜索。发生绕回后，要停止搜索，可保存第一个找到的单元格地址，然后测试后面找到的每个单元格地址是否与其相同。

若要对单元格进行模式更为复杂的搜索，请结合使用 For Each...Next 语句和 **Like** 运算符。例如，下列代码在单元格区域 A1:C5 中搜索字体名称以“Cour”开始的单元格。当 ET 找到匹配单元格以后，就将其字体改为 Times New Roman。

For Each c In \[A1:C5\] If c.Font.Name Like "Cour\*" Then c.Font.Name = "Times New Roman" End If Next

**示例**

``` JavaScript
/*本示例在第一个工作表的单元格区域 A1:A500 中查找包含值 2 的所有单元格，并将这些单元格的值更改为 5。*/
function test()
{
    let add = Application.Worksheets.Item(1).Range("A1:A500") 
    let c = add.Find(2,undefined,xlValues)
    if(c != null) {
        let firstAddress = c.Address
        do {
            c.Value2 = 5
            c = add.FindNext(c)
        }while(c != null && c.Address != firstAddress)
    }
}
```

### Range.FindNext

继续由 **Find** 方法开始的搜索。查找匹配相同条件的下一个单元格，并返回表示该单元格的 **Range** 对象。该操作不影响选定内容和活动单元格。

返回值类型为Range。

**语法**

**express.FindNext ( *After* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                   |
|---------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *After* | 可选      | Variant  | 指定一个单元格，查找将从该单元格之后开始。此单元格对应于从用户界面搜索时的活动单元格位置。注意，After 必须是查找区域中的单个单元格。注意，搜索是从该单元格之后开始的；直到本方法环绕到此单元格时，才检测其内容。如果未指定本参数，查找将从区域的左上角单元格之后开始。 |

**说明**

当查找到指定查找区域的末尾时，本方法将环绕至区域的开始继续搜索。发生环绕后，为停止查找，可保存第一次找到的单元格地址，然后测试下一个查找到的单元格地址是否与其相同。

**示例**

``` JavaScript
/*本示例在单元格区域 A1:A500 中查找值为 2 的单元格，并将这些单元格的值变为 5。*/
function test()
{
    let add = Application.Worksheets.Item(1).Range("A1:A500") 
    let c = add.Find(2,undefined,xlValues)
    if(c != null) {
        let firstAddress = c.Address
        do {
            c.Value2 = 5
            c = add.FindNext(c)
        }while(c != null && c.Address != firstAddress)
    }
}
```

### Range.FindPrevious

继续由 **Find** 方法开始的搜索。查找匹配相同条件的上一个单元格，并返回代表该单元格的 **Range** 对象。该操作不影响选定内容和活动单元格。

返回值类型为Range。

**语法**

**express.FindPrevious ( *After* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                   |
|---------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *After* | 可选      | Variant  | 指定一个单元格，查找将从该单元格之前开始。此单元格对应于从用户界面搜索时的活动单元格的位置。请注意：After 必须是区域中的单个单元格。注意，搜索是从该单元格之前开始的；直到本方法环绕到此单元格时，才检测其内容。如果未指定本参数，查找将从区域的左上角单元格之前开始。 |

**说明**

当查找到指定查找区域的起始位置时，本方法将环绕至区域的末尾继续搜索。发生绕回后，要停止搜索，可保存第一个找到的单元格地址，然后测试后面找到的每个单元格地址是否与其相同。

**示例**

``` JavaScript
/*本示例演示 FindPrevious 方法如何与 Find 方法和 FindNext 方法一起使用。运行本示例之前，请确保工作表 Sheet1 的 B 列中至少出现过两次“Phoenix”单词。*/
function test()
{
    let fc = Application.Worksheets.Item("Sheet1").Columns.Item("B").Find("Phoenix")
        MsgBox("The first occurrence is in cell " + fc.Address)
    let Mc = Application.Worksheets.Item("Sheet1").Columns.Item("B").FindNext(fc)
        MsgBox("The next occurrence is in cell " + Mc.Address)
    let gc = Application.Worksheets.Item("Sheet1").Columns.Item("B").FindPrevious(fc)
        MsgBox("The previous occurrence is in cell " + gc.Address)
}
```

### Range.Insert

在工作表或宏表中插入一个单元格或单元格区域，其他单元格相应移位以腾出空间。

返回值类型为Variant。

**语法**

**express.Insert ( *Shift* , *CopyOrigin* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                                                             |
|--------------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------|
| *Shift*      | 可选      | Variant  | 指定单元格的调整方式。可为以下 XlInsertShiftDirection 常量之一：xlShiftToRight 或 xlShiftDown。如果省略此参数，ET 将根据区域的形状确定调整方式。 |
| *CopyOrigin* | 可选      | Variant  | 复制的起点。                                                                                                                                     |

### Range.Item

返回一个 **Range** 对象，它代表对指定区域某一偏移量处的区域。

**语法**

**express.Item ( *RowIndex* , *ColumnIndex* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                                                                                     |
|---------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------|
| *RowIndex*    | 必选      | Variant  | 要访问的单元格的索引号，顺序为从左到右，然后往下。 Range.Item(1) 返回区域左上角单元格； Range.Item(2) 返回紧靠左上角单元格右侧的单元格。 |
| *ColumnIndex* | 可选      | Variant  | 指明要访问的单元格所在列的列号的数字或字符串，1 或 “A”表示区域中的第一列。                                                               |

**说明**

语法 1 使用行号和列号或列标作为索引参数。关于此语法的详细信息，请参阅 **Range** 对象。 **RowIndex** 和 **ColumnIndex** 参数为相对偏移量。换句话说，如果 **RowIndex** 指定为 1，则返回区域内第一行中的单元格，而非工作表的第一行。例如，如果选定区域为单元格 C3，则 Selection.Cells(2, 2) 返回单元格 D4（使用 **Item** 属性可在原始区域之外进行索引）。

**示例**

``` JavaScript
/*此示例基于单元格 A1 的内容填写 Sheet1 的单元格区域 A2。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("A1").Item(2).FillDown()    
}
```

### Range.Justify

调整区域内的文字，使之均衡地填充该区域。

返回值类型为Variant。

**语法**

**express.Justify ()**

\*express是一个代表 Range 对象的变量。

**说明**

如果区域不够大，ET 将显示一个消息，通知您文本将延伸到区域的下方。如果单击“确定” 按钮，被调整的文本将替换超出选定区域的单元格的内容。为防止该消息出现，请将 **DisplayAlerts** 属性值设置为 **False** 。设置该属性后，文本将始终替换区域下方单元格的内容。

**示例**

``` JavaScript
/*本示例调整工作表 Sheet1 上单元格 A1 中的文字。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("A1").Justify()
}
```

### Range.Merge

由指定的 **Range** 对象创建合并单元格。

**语法**

**express.Merge ( *Across* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                  |
|----------|-----------|----------|---------------------------------------------------------------------------------------|
| *Across* | 可选      | Variant  | 如果为 True，则将指定区域中每一行的单元格合并为一个单独的合并单元格。默认值是 False。 |

**说明**

合并区域的值在该区域左上角的单元格中指定。

### Range.Offset

返回 **Range** 对象，它代表位于指定单元格区域的一定的偏移量位置上的区域。

**语法**

**express.Offset ( *RowOffset* , *ColumnOffset* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                                     |
|----------------|-----------|----------|------------------------------------------------------------------------------------------|
| *RowOffset*    | 可选      | Variant  | 区域偏移的行数（正数、负数或 0（零））。正数表示向下偏移，负数表示向上偏移。默认值是 0。 |
| *ColumnOffset* | 可选      | Variant  | 区域偏移的列数（正数、负数或 0（零））。正数表示向右偏移，负数表示向左偏移。默认值是 0。 |

**示例**

``` JavaScript
/*此示例激活 Sheet1 上活动单元格向右偏移三列、向下偏移三行处的单元格。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.Offset(3,3).Activate()
}

/*此示例假设 Sheet1 中包含一个具有标题行的表格。此示例先选定该表格，但并不选择行首。运行此示例之前，活动单元格必须位于表格中。*/
function test()
{
    let tbl = Application.ActiveCell.CurrentRegion
    tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1,tbl.Columns.Count).Select()
}
```

### Range.PrintPreview

按对象打印后的外观效果显示对象的预览。

返回值类型为Variant。

**语法**

**express.PrintPreview ( *EnableChanges* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                          |
|-----------------|-----------|----------|-------------------------------------------------------------------------------|
| *EnableChanges* | 可选      | Variant  | 传递 Boolean 值，以指定用户是否可更改边距和打印预览中可用的其他页面设置选项。 |

**示例**

``` JavaScript
/*此示例在打印预览中显示工作表 Sheet1。*/
Application.Worksheets.Item("Sheet1").PrintPreview()
```

### Range.Range

返回一个 **Range** 对象，它代表一个单元格或单元格区域。

**语法**

**express.Range ( *Cell1* , *Cell2* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                       |
|---------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Cell1* | 必选      | Variant  | 区域名称。必须为采用宏语言的 A1 样式引用。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可包括货币符号，但它们会被忽略掉。您可以在区域中任一部分使用局部定义名称。如果使用名称，则假定该名称使用的是宏语言。 |
| *Cell2* | 可选      | Variant  | 区域左上角和右下角的单元格。可以是一个包含单个单元格、整列或整行的 Range 对象，或者也可以是一个用宏语言为单个单元格命名的字符串。                                                                                                          |

**说明**

如果在没有对象识别符时使用，则该属性是 ActiveSheet.Range 的快捷方式（它返回活动表的一个区域，如果活动表不是一张工作表，则该属性无效）。

当应用于 **Range** 对象时，该属性与 **Range** 对象相关。例如，如果选中单元格 C3，那么 Selection.Range("B1") 返回单元格 D3，因为它同 **Selection** 属性返回的 **Range** 对象相关。此外，代码 ActiveSheet.Range("B1") 总是返回单元格 B1。

**示例**

``` JavaScript
/*此示例将 Sheet1 上 A1 单元格的值设置为 3.14159。*/
Application.Worksheets.Item("Sheet1").Range("A1").Value2 = 3.14159

/*此示例在 Sheet1 的 A1 单元格中创建一个公式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Formula = "=10*RAND()"

/*此示例在 Sheet1 上的单元格区域 A1:D10 中进行循环。如果某个单元格的值小于 0.001，则此代码将用 0（零）来取代该值。*/
function test()
{
    let c = Application.Worksheets.Item("Sheet1").Range("A1:D10")
    for(let i = 1; i <= c.Rows.Count;i++) {
        for(let j = 1 ;j <= c.Columns.Count;j++) {
            if(c.Item(i,j).Value2 < 0.001) {
                c.Item(i,j).Value2 = 0
            }
        }
    }
}

/*此示例在名为“TestRange”的区域上进行循环，并显示该区域中空白单元格的个数。*/
function test()
{    
    let numBlanks = 0
    let c = Application.Range("TestRange")
    for(let i = 1 ; i <= c.Rows.Count; i++) {
        for(let j = 1; j <= c.Columns.Count; j++) {
            if(c.Item(i,j).Value2 == null) {
                numBlanks = numBlanks + 1
            }
        }
    }
    MsgBox("There are " + numBlanks + " empty cells in this range")
}

/*此示例将 Sheet1 中单元格区域 A1:C5 上的字体样式设置为斜体。此示例使用 Range 属性的语法 2。*/
Application.Worksheets.Item("Sheet1").Range(Cells.Item(1, 1), Cells.Item(5, 3)).Font.Italic = true
```

### Range.Resize

调整指定区域的大小。返回 **Range** 对象，该对象代表调整后的区域。

**语法**

**express.Resize ( *RowSize* , *ColumnSize* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                       |
|--------------|-----------|----------|------------------------------------------------------------|
| *RowSize*    | 可选      | Variant  | 新区域中的行数。如果省略该参数，则该区域中的行数保持不变。 |
| *ColumnSize* | 可选      | Variant  | 新区域中的列数。如果省略该参数。则该区域中的列数保持不变。 |

**示例**

``` JavaScript
/*本示例调整 Sheet1 中选定区域的大小，使之增加一行和一列。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    numRows = Application.Selection.Rows.Count
    numColumns = Application.Selection.Columns.Count
    Application.Selection.Resize(numRows + 1, numColumns + 1).Select()
}

/*本示例假定在 Sheet1 中有一个包含标题行的表。本示例选定该表，但不选定标题行。运行本示例之前，活动单元格必须处于该表中。*/
function test()
{  
    let tbl = Application.ActiveCell.CurrentRegion
    tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1,tbl.Columns.Count).Select()
}
```

### Range.Run

在该处运行 ET 宏。区域必须位于宏表上。

**语法**

**express.Run ( *Arg1* , *...* , *Arg30* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明               |
|---------|-----------|----------|--------------------|
| *Arg1*  | 可选      | Variant  | 传递给函数的参数。 |
| *...*   | 可选      | Variant  | 传递给函数的参数。 |
| *Arg30* | 可选      | Variant  | 传递给函数的参数。 |

**说明**

此方法不可使用命名参数，参数必须通过位置进行传递。

**Run** 方法返回被调用的宏返回的任何值。如果将对象作为参数传递给宏，该对象将转换为相应的值（通过对该对象应用 **Value** 属性）。这意味着不能用 **Run** 方法将对象传递给宏。

### Range.Select

选择对象。

返回值类型为Variant。

**语法**

**express.Select ()**

\*express是一个代表 Range 对象的变量。

**说明**

要选择单元格或单元格区域，请使用 **Select** 方法。要使单个单元格成为活动单元格，请使用 **Activate** 方法。

### Range.UnMerge

将合并区域分解为独立的单元格。

**语法**

**express.UnMerge ()**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将包含单元格 A3 的合并区域分解。*/
function test()
{
    if(Application.Range("a3").MergeCells){
        Application.Range("a3").MergeArea.UnMerge()
    }
    else {
        MsgBox( "not merged")
    }
}
```

### Range.Value

返回一个 **Variant** 型，它代表指定单元格的值。

**语法**

**express.Value ( *RangeValueDataType* )**

\*express是一个代表 Range 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                               |
|----------------------|-----------|----------|----------------------------------------------------|
| *RangeValueDataType* | 可选      | Variant  | 区域值数据类型。可以为 xlRangeValueDataType 常量。 |

**说明**

用 XML 电子表格文件中的内容设置单元格区域时，只使用工作簿中第一张工作表中的值。无法设置或得到以 XML 电子表格格式表示的不相连的单元格区域。

JSAPI里的Value是个方法，可以进行取值，不能进行赋值。要给单元格赋值，请用Range.Value2属性。

**示例**

``` JavaScript
/*本示例获取 Sheet1 上 A1 单元格的值，赋给变量"rngValue"。*/
let rngValue = Application.Worksheets.Item("Sheet1").Range("A1").Value()
```

## 成员属性

### Range.AddIndent

返回或设置一个 **Variant** 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。

**语法**

**express.AddIndent**

\*express是一个代表 Range 对象的变量。

**说明**

如果将此属性的值设为 **True** ，那么在单元格中文本的对齐方式设为水平或垂直等距分布时，将自动缩进文本。

要将文本对齐方式设为等距分布，可在 **Orientation** 属性的值为 **xlVertical** 时，将 **VerticalAlignment** 属性设为 **xlVAlignDistributed** ；在 **Orientation** 属性的值为 **xlHorizontal** 时，将 **HorizontalAlignment** 属性设为 **xlHAlignDistributed** 。

**示例**

``` JavaScript
/*此示例将 Sheet1 的单元格 A1 中文本的水平对齐方式设置为等距分布，然后缩进该文本。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("A1").HorizontalAlignment = xlHAlignDistributed
    Application.Worksheets.Item("Sheet1").Range("A1").AddIndent = true
}
```

### Range.Application

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    let myObject = ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        MsgBox("This is an ET Application object.")
    }
    else {
        MsgBox("This is not an ET Application object.")
    }
}
```

### Range.Areas

返回一个 **Areas** 集合，该集合表示多重区域选择中的所有区域。只读。

**语法**

**express.Areas**

\*express是一个代表 Range 对象的变量。

**说明**

对于单一选择区域， **Areas** 属性返回只包含一个对象的集合，即原始 **Range** 对象本身。对于多重选择区域， **Areas** 属性返回一个集合，该集合包含与每个选定区域相对应的对象。

**示例**

``` JavaScript
/*本示例在用户选定多个区域并试图执行某一命令时显示提示信息。该示例必须在工作表上运行。*/
function test()
{
    if(Application.Selection.Areas.Count > 1) {
        MsgBox("Cannot do this to a multi-area selection.")
    }
}
```

### Range.Cells

返回一个 Range 对象，该对象代表指定区域中的单元格。

**语法**

**express.Cells**

\*express是一个代表 Range 对象的变量。

**说明**

因为 **Item** 属性是 **Range** 对象的 <span class="glossary" "appendpopup(this,'defdefaultproperty_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none;">默认属性</span> ，所以可以在 **Cells** 关键字后面紧接着指定行和列索引。有关详细信息，请参阅 **Item** 属性和本主题的示例。

在不使用对象识别符的情况下，使用此属性将返回一个 **Range** 对象，它代表活动工作表中所有的单元格。

**示例**

``` JavaScript
/*本示例将 Sheet1 上单元格区域 A1: C5 的字体样式设置为斜体。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.Worksheets.Item("Sheet1").Range(Cells.Item(1, 1), Cells.Item(5, 3)).Font.Italic = true
}

/*本示例搜索列“myRange”中的数据。如果发现某单元格的值与上面的一个单元格的值相等，则本示例将显示这个包含重复数据的单元格的地址。*/
function test()
{
    let r = Range("myRange")
    for(let n = 1;n<=r.Rows.Count;n++) {
        if(r.Cells.Item(n, 1).Value2 == r.Cells.Item(n + 1, 1).Value2) {
            MsgBox("Duplicate data in " + r.Cells.Item(n + 1, 1).Address())
        }
    }
}
```

### Range.Column

返回指定区域中第一块中的第一列的列号。Long 类型，只读。

**语法**

**express.Column**

\*express是一个代表 Range 对象的变量。

**说明**

A 列返回 1，B 列返回 2，依此类推。

若要返回区域中最后一列的列号，请使用下列语句。

myRange.Columns(myRange.Columns.Count).Column

**示例**

``` JavaScript
/*本示例将工作表 Sheet1 上每隔一列的列宽设置为 4 磅。*/
function test()
{
    let col = Application.Worksheets.Item("Sheet1").Columns
    for (let i = 1; i <= col.Count; i++) {
        if ((col.Item(i).Column) % 2 == 0) {
            col.Item(i).ColumnWidth = 4
        }
    }
}
```

### Range.ColumnWidth

返回或设置指定区域中所有列的列宽。 **Variant** 类型，可读写。

**语法**

**express.ColumnWidth**

\*express是一个代表 Range 对象的变量。

**说明**

一个列宽单位等于“常规”样式中一个字符的宽度。对于比例字体，则使用字符 0（零）的宽度。

使用 **Width** 属性可返回以磅为单位的列宽。

如果区域中所有列的列宽都相等， **ColumnWidth** 属性返回该宽度值。如果区域中的列宽不等，该属性返回 **null** 。

**示例**

``` JavaScript
/*本示例使工作表 Sheet1 上 A 列的列宽加倍。*/
function test()
{
    let itm = Application.Worksheets.Item("Sheet1").Columns.Item("A")
    itm.ColumnWidth = itm.ColumnWidth * 2
}
```

### Range.Columns

返回一个 **Range** 对象，该对象代表指定区域中的列。

**语法**

**express.Columns**

\*express是一个代表 Range 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 ActiveSheet.Columns。

此属性在应用于一个是多重选定区域的 **Range** 对象时，会只从该区域的第一个子区域中返回列。例如，如果 **Range** 对象有两个子区域 A1:B2 和 C3:D4，那么，Selection.Columns.Count 的返回值是 2，而不是 4。若要对一个可能包含多重选定区域的区域使用此属性，请测试 Areas.Count 以确定此区域内是否包含多个子区域。如果包含，请对此区域内的每个子区域进行循环。

**示例**

``` JavaScript
/*此示例将名为“myRange”区域第一列中每一单元格的值置为 0（零〕。*/
Application.Range("myRange").Columns.Item(1).Value2 = 0

/*此示例显示 Sheet1 上选定区域中的列数。如果选择了多个子区域，此示例将对每一个子区域进行循环。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    let areaCount = Application.Selection.Areas.Count
    MsgBox(areaCount)
    if (areaCount <= 1) {
        MsgBox("The selection contains " + Application.Selection.Columns.Count + " columns.")
    }
    else {
        for (let i = 1; i <= areaCount; i++) {
            MsgBox("Area " + i + " of the selection contains " + Application.Selection.Areas.Item(i).Columns.Count + " columns.")
        }
    }
}
```

### Range.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*此示例显示 Sheet1 上选定区域中的列数。此示例还将检测选定区域中是否包含多重选定区域，如果包含，则对多重选定区域中每一子区域进行循环。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    let iAreaCount = Application.Selection.Areas.Count
    if(iAreaCount <= 1) {
        MsgBox("The selection contains " + Application.Selection.Columns.Count + " columns.")
    }
    else {
        for(let i = 1;i<=iAreaCount;i++) {
            MsgBox("Area " + i + " of the selection contains " + Application.Selection.Areas.Item(i).Columns.Count + " columns.")
        }
    }
}
```

### Range.Dependents

返回一个 **Range** 对象，该对象表示包含一个单元格所有从属单元格的区域。如果有多个从属单元格，这可能有多个选择（多个 **Range** 对象）。 **Range** 类型，只读。

**语法**

**express.Dependents**

\*express是一个代表 Range 对象的变量。

**说明**

**Dependents** 属性仅在活动工作表中有效，不能追踪远程引用。

**示例**

``` JavaScript
/*本示例选定工作表 Sheet1 中单元格 A1 的从属单元格。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.Range("A1").Dependents.Select()
}
```

### Range.End

返回一个 **Range** 对象，该对象代表包含源区域的区域尾端的单元格。等同于按键 End+ 向上键、End+ 向下键、End+ 向左键或 End+ 向右键。 **Range** 对象，只读。

**语法**

**express.End**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例选定包含单元格 B4 的区域中 B 列顶端的单元格。*/
Application.Range("B4").End(xlUp).Select()

/*本示例选定包含单元格 B4 的区域中第 4 行尾端的单元格。*/
Application.Range("B4").End(xlToRight).Select()

/*本示例将选定区域从单元格 B4 延伸至第四行最后一个包含数据的单元格。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.Range("B4", Range("B4").End(xlToRight)).Select()
}
```

### Range.EntireColumn

返回一个 **Range** 对象，该对象表示包含指定区域的整列（或多列）。只读。

**语法**

**express.EntireColumn**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例对包含活动单元格的列中的第一个单元格赋值。本示例必须在工作表上运行。*/
Application.ActiveCell.EntireColumn.Cells.Item(1, 1).Value2 = 5
```

### Range.EntireRow

返回一个 **Range** 对象，该对象表示包含指定区域的整行（或多行）。只读。

**语法**

**express.EntireRow**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例对包含活动单元格的行中的第一个单元格赋值。本示例必须在工作表上运行。*/
Application.ActiveCell.EntireRow.Cells.Item(1, 1).Value2 = 5
```

### Range.Font

返回一个 **Font** 对象，它代表指定对象的字体。

**语法**

**express.Font**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*此示例判断单元格 A1 的字体名称是否为 Arial，并通知用户。*/
function CheckFont(){
    Application.Range("A1").Select()
                                
    // Determine if the font name for selected cell is Arial.
    if(Application.Range("A1").Font.Name == "Arial") {
        MsgBox("The font name for this cell is 'Arial'")
    }
    else {
        MsgBox("The font name for this cell is not 'Arial'")
    }
}
```

### Range.Formula

返回或设置一个 **Variant** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

**语法**

**express.Formula**

\*express是一个代表 Range 对象的变量。

**说明**

此属性对于 <span class="glossary" "appendpopup(this,'xldefolap_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none;">OLAP</span> 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式， **Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

**示例**

``` JavaScript
/*此示例设置 Sheet1 中 A1 单元格的公式。*/
Application.Worksheets.Item("Sheet1").Range("A1").Formula = "=$A$4+$A$10"
```

### Range.Height

返回或设置一个 **Variant** 值，该值代表区域的高度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 Range 对象的变量。

### Range.Hidden

返回或设置一个 **Variant** 值，它指明是否隐藏行或列。

**语法**

**express.Hidden**

\*express是一个代表 Range 对象的变量。

**说明**

将此属性设置为 **True** 以隐藏行或列。指定的区域必须占据整个行或整个列。

请不要将此属性与 **FormulaHidden** 属性混淆。

**示例**

``` JavaScript
/*此示例隐藏 Sheet1 的 C 列。*/
Application.Worksheets.Item("Sheet1").Columns.Item("C").Hidden = true
```

### Range.HorizontalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

\*express是一个代表 Range 对象的变量。

**说明**

此属性的 值可设为以下常量之一：

|                   |
|-------------------|
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlLeft**        |
| **xlRight**       |

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### Range.ID

返回或设置一个 **String** 值，它代表将页面另存为网页时指定单元格的识别标志。

**语法**

**express.ID**

\*express是一个代表 Range 对象的变量。

**说明**

可将 ID 标签用作其他 HTML 文档或相同网页上的超链接引用。

**示例**

``` JavaScript
/*此示例将活动工作表上 A1 单元格的 ID 设置为“target”。*/
Application.ActiveSheet.Range("A1").ID = "target"

/*之后，将文档另存为网页，并将下面的 HTML 行添加到网页中。 然后，当用户在 Web 浏览器中查看该页并用鼠标单击该超链接时，浏览器将显示该单元格。*/
<A HREF="#target">Quarterly earnings</A>
```

### Range.Left

返回一个 **Variant** 值，它代表从列 A 的左边缘到区域的左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 Range 对象的变量。

**说明**

如果该区域不连续，则使用第一块区域。如果该区域有若干列宽，则使用区域中最左边的列。

### Range.MergeArea

返回一个 **Range** 对象，该对象代表包含指定单元格的合并区域。如果指定的单元格不在合并区域内，则该属性返回指定的单元格。只读。 **Variant** 类型。

**语法**

**express.MergeArea**

\*express是一个代表 Range 对象的变量。

**说明**

**MergeArea** 属性只应用于单个单元格区域。

**示例**

``` JavaScript
/*本示例为包含单元格 A3 的合并区域赋值。*/
function test()
{
    let ma = Application.Range("a3").MergeArea
    if(ma.Address == "$A$3") {
        MsgBox("not merged")
    }
    else {
        ma.Cells.Item(1, 1).Value2 = "42"
    }
}
```

### Range.MergeCells

如果区域包含合并单元格，则为 **True** 。 **Variant** 型，可读写。

**语法**

**express.MergeCells**

\*express是一个代表 Range 对象的变量。

**说明**

选定包含合并单元格的区域时，所选定的区域可能与所期望选定的区域不同。使用 **Address** 属性检验选定区域的地址。

**示例**

``` JavaScript
/*本示例为包含单元格 A3 的合并区域赋值。*/
function test()
{
    let ma = Application.Range("a3").MergeArea
    if(Application.Range("a3").MergeCells) {
        MsgBox(123)
        ma.Cells.Item(1, 1).Value2 = "42"
    }
}
```

### Range.Name

返回或设置一个 **Variant** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Range 对象的变量。

### Range.NumberFormat

返回或设置一个 **Variant** 值，它代表对象的格式代码。

**语法**

**express.NumberFormat**

\*express是一个代表 Range 对象的变量。

**说明**

如果指定区域中的所有单元格包含不同的数字格式，则此属性返回 **Null** 。

格式代码与 **“设置单元格格式”** 对话框中的 **“格式代码”** 选项是同一个字符串。 **Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

**示例**

``` JavaScript
/*以下这些示例分别对 Sheet1 中的 A17 单元格、第一行和 C 列的数字格式进行设置。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("A17").NumberFormat = "General"
    Application.Worksheets.Item("Sheet1").Rows.Item(1).NumberFormat = "hh:mm:ss"
    Application.Worksheets.Item("Sheet1").Columns.Item("C").NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
}
```

### Range.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Range 对象的变量。

### Range.Previous

返回一个代表下一个单元格的 **Range** 对象。

**语法**

**express.Previous**

\*express是一个代表 Range 对象的变量。

**说明**

如果指定对象为区域，则此属性的作用是仿效 Shift+Tab；但此属性只是返回上一单元格，并不选定它。

在保护工作表上，该属性返回上一个未锁定的单元格。在未保护的工作表上，该属性通常返回指定单元格左侧相邻的单元格。

**示例**

``` JavaScript
/*此示例选定 sheet1 中上一个未锁定单元格。如果 sheet1 未保护，选定的单元格将是紧靠活动单元格左边的单元格。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.Previous.Select()
}
```

### Range.Row

返回区域中第一个子区域的第一行的行号。 **Long** 类型，只读。

**语法**

**express.Row**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例将工作表 sheet1 中每隔一行的行高设置为 4 磅。*/
function test()
{
    let rw = Application.Worksheets.Item("Sheet1").Rows
    for(let i = 1; i <= rw.Count; i++) {
        if((rw.Item(i).Row) % 2 == 0) {
           rw.Item(i).RowHeight = 4
        }
    }
}
```

### Range.RowHeight

以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位返回或设置指定区域中所有行的行高。如果指定区域中的各行的行高不等，则返回 **null** 。 **Variant** 类型，可读写。

**语法**

**express.RowHeight**

\*express是一个代表 Range 对象的变量。

**说明**

可使用 **Height** 属性返回整个单元格区域的高度。

以下是 **RowHeight** 和 **Height** 的不同之处：

- **Height** 为只读。
- 如果要返回几行的 **RowHeight** 属性，可得到每一行的行高（如果所有的行等高），或得到 **null** （如果它们不等高）。如果要返回几行的 **Height** 属性，将得到所有行高的总和。

**示例**

``` JavaScript
/*本示例使工作表 sheet1 上第一行的行高加倍。*/
function test()
{
    let ro= Application.Worksheets.Item("Sheet1").Rows.Item(1)
    ro.RowHeight =ro.RowHeight * 2
}
```

### Range.Rows

返回一个 **Range** 对象，它代表指定单元格区域中的行。 **Range** 对象，只读。

**语法**

**express.Rows**

\*express是一个代表 Range 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 ActiveSheet.Rows。

该属性在应用于是多个选定区域的 **Range** 对象时，只从该区域中第一个子区域内返回行。例如，如果 **Range** 对象有两个子区域：A1:B2 和 C3:D4，则 Selection.Rows.Count 返回 2 而不是 4。若要在一个可能包含多个选定区域的区域中使用此属性，请测试 Areas.Count 以确定该区域是否包含多个选择区域。如果是，请对此区域内的每个子区域进行循环，如第 3 个示例所示。

**示例**

``` JavaScript
/*此示例删除 Sheet1 的第三行。*/
Application.Worksheets.Item("Sheet1").Rows.Item(3).Delete()

/*此示例检查第一张工作表上当前区域中的行，如果某行的第一个单元格值与前一行的第一个单元格的值相等，则删除此行。*/
function test()
{
    let rw = Application.Worksheets.Item(1).Cells.Item(1, 1).CurrentRegion.Rows
    for(let i = 2; i <= rw.Count; i++) {
        let ths= rw.Item(i).Cells.Item(1, 1).Value2
        let last =rw.Item(i-1).Cells.Item(1, 1).Value2
        if(ths == last) {
            rw.Item(i).Delete()
            last = ths
        }
    }
}

/*此示例显示 Sheet1 选定区域中的行数。如果选择了多个子区域，此示例将对每一个子区域进行循环。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    let areaCount = Application.Selection.Areas.Count
    if(areaCount <= 1) {
        MsgBox("The selection contains " + Selection.Rows.Count & " rows.")
    }
    else {
      let  i = 1
      let a = Application.Selection.Areas
      for(let b=1;b<=a.Count;b++) {
          MsgBox("Area " + i + " of the selection contains " + a.Item(b).Rows.Count + " rows.")
          i = i + 1
      }  
    }
}
```

### Range.Text

返回或设置指定对象中的文本。 **String** 型，只读。

**语法**

**express.Text**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*此示例演示包含格式数字的单元格的 Text 和 Value 属性的区别。*/
function test()
{
    let c = Application.Worksheets.Item("Sheet1").Range("B14")
    c.Value2 = 1198.3  
    c.NumberFormat = "$#,##0_);($#,##0)"
    MsgBox(c.Value2)
    MsgBox(c.Text)
}
```

### Range.Top

返回或设置一个 **Variant** 值，它代表行 1 上边缘到区域上边缘的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 Range 对象的变量。

**说明**

如果该区域不连续，则使用第一块区域。如果该区域有若干行，则使用区域中的顶端行（行号最小的行）。

### Range.Value2

返回或设置单元格值。 **Variant** 类型，可读写。

**语法**

**express.Value2**

\*express是一个代表 Range 对象的变量。

**说明**

该属性与 **Value** 属性的唯一区别是： **Value2** 属性不使用 **Currency** 和 **Date** 数据类型。可以通过使用 **Double** 数据类型，以浮点数形式返回这些数据类型格式的数值。

**示例**

``` JavaScript
/*本示例使用 Value2 属性对两个单元格的值进行相加。*/
Application.Range("a1").Value2 = Range("b1").Value2 + Range("c1").Value2
```

### Range.Width

返回一个 **Variant** 值，它代表区域的宽度（以单位表示）。

**语法**

**express.Width**

\*express是一个代表 Range 对象的变量。

### Range.Worksheet

返回一个 **Worksheet** 对象，该对象表示包含指定区域的工作表。只读。

**语法**

**express.Worksheet**

\*express是一个代表 Range 对象的变量。

**示例**

``` JavaScript
/*本示例显示包含活动单元格的工作表的名称。本示例必须在工作表上运行。*/
MsgBox(Application.ActiveCell.Worksheet.Name)

/*本示例显示包含名为“testRange”的区域的工作表名称。*/
MsgBox(Application.Range("testRange").Worksheet.Name)
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Range](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Shapes 对象

## 简介

指定的工作表上的所有 Shape 对象的集合。

## 说明

每个 Shape 对象都代表绘图层中的一个对象，如自选图形、任意多边形、OLE 对象或图片。

注释： 如果您想处理文档中的一部分形状（例如，只针对文档中的自选图形或只对选定的形状进行一些操作），则必须构造一个 ShapeRange 集合，其中包含您要处理的形状。

## 示例

使用 Shapes 属性可返回 Shapes 集合。下例选定 myDocument 上的所有形状。

注释： 如果您要同时对工作表上的所有形状进行操作（例如删除或设置属性），请选定所有形状，然后对选定区域使用 ShapeRange 属性，以创建一个 ShapeRange 对象，该对象包含工作表上的所有形状，然后对 ShapeRange 对象应用相应的属性或方法。

``` JavaScript
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.SelectAll()
}
```

使用 ` Shapes ` ( *index* )（其中 *index* 是形状的名称或索引号）可返回一个 Shape 对象。下例设置 ` myDocument ` 上形状一的预设阴影的填充。

``` JavaScript
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.Item(1).Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientBrass)
}
```

使用 ` Shapes.Range ` ( *index* )（其中 *index* 是形状的名称或索引号，或是它们的一个数组）可返回一个 ShapeRange 集合，该集合代表 Shapes 集合的一个子集。下例设置 *myDocument* 上形状一和三的填充图案。

``` JavaScript
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick)
}
```

工作表上的 ActiveX 控件有两个名称：一个是包含该控件的形状的名称，查看工作表时可在 “名称” 框中看到它；另一个是控件的代码名称，可以在 “属性” 窗口中 “(名称)” 右边的单元格中看到它。在您首次向工作表添加控件时，形状名称和代码名称是一致的。但是，如果您更改这两个名称中的任意一个，另一个不会随之自动更改。

在控件事件过程名称中使用的是控件的代码名称，但是，当您从工作表的 Shapes 或 OLEObjects 集合中返回控件时，必须使用形状名称而不是代码名称，以便按名称引用控件。例如，假定您给工作表添加一个复选框，而默认的形状名称和代码名称都是 CheckBox1。如果通过在“属性” 窗口的“(名称)” 旁边键入“chkFinished”而更改了控件代码名称，则在事件过程名称中必须使用 chkFinished，但是您仍然需要使用 CheckBox1 从 Shapes 或 OLEObject 集合中返回控件，如下例所示。

``` JavaScript
function test() {
    Application.ActiveSheet.OLEObjects("CheckBox1").Object.Value = 1
}
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
<td style="text-align: left;" data-valian="middle"><a href="#Shapes.AddOLEObject">AddOLEObject</a></td>
<td style="text-align: left;" data-valian="middle"><p>创建 OLE 对象。返回一个代表新 OLE 对象的 Shape 对象。</p>
<p>返回值类型为Shape。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#Shapes.AddPicture">AddPicture</a></td>
<td style="text-align: left;" data-valian="middle"><p>现有文件创建图片。返回一个代表新图片的 Shape 对象。</p>
<p>返回值类型为 Shape。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#Shapes.AddShape">AddShape</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Shape 对象，该对象表示工作表中的新自选形状。</p>
<p>返回值类型为Shape。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#Shapes.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>从集合中返回一个对象。</p>
<p>返回值： 包含在集合中的一个 Shape 对象。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#Shapes.Range">Range</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 ShapeRange 对象，它代表 Shapes 集合中形状的子集。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#Shapes.SelectAll">SelectAll</a></td>
<td style="text-align: left;" data-valian="middle"><p>选择指定的 Shap es 集合中的所有形状。</p></td>
</tr>
</tbody>
</table>

## 属性

| 序号 | 名称                               | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Shapes.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Shapes.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |

## 成员方法

### Shapes.AddOLEObject

创建 OLE 对象。返回一个代表新 OLE 对象的 **Shape** 对象。

返回值类型为Shape。

**语法**

**express.AddOLEObject ( *ClassType* , *Filename* , *Link* , *DisplayAsIcon* , *IconFileName* , *IconIndex* , *IconLabel* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                       |
|-----------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *ClassType*     | 可选      | Variant  | （必须指定 ClassType 或 FileName）。该字符串包含要创建的对象的程序标识符。如果指定了 ClassType 参数，则忽略 FileName 和 Link。                                                                                                                                                                             |
| *Filename*      | 可选      | Variant  | 创建对象所用的文件名称。如果未指定路径，则使用当前的工作文件夹。必须为对象指定 ClassType 或 FileName 参数，但不能同时指定。                                                                                                                                                                                |
| *Link*          | 可选      | Variant  | 如果为 True，则将 OLE 对象链接到创建该对象所使用的文件；如果为 False，则 OLE 对象成为该文件的独立副本。如果为 ClassType 指定值，则该参数必须为 False。默认值为 False。                                                                                                                                     |
| *DisplayAsIcon* | 可选      | Variant  | 如果为 True，则将 OLE 对象显示为图标。默认值为 False。                                                                                                                                                                                                                                                     |
| *IconFileName*  | 可选      | Variant  | 包含将要显示的图标的文件。                                                                                                                                                                                                                                                                                 |
| *IconIndex*     | 可选      | Variant  | IconFileName 内的图标索引。指定文件中图标的顺序与图标在“更改图标”对话框中出现的顺序相对应（选中“显示为图标”复选框后，可从“对象”对话框访问该对话框）。文件中的第一个图标的索引号为 0（零）。如果 IconFileName 中不存在给定索引号的图标，则使用索引号为 1 的图标（即文件中的第二个图标）。默认值为 0（零）。 |
| *IconLabel*     | 可选      | Variant  | 显示在图标下面的标签（题注）。                                                                                                                                                                                                                                                                             |
| *Left*          | 可选      | Variant  | 新对象的左上角相对于文档左上角的位置（以磅为单位）。默认值为 0（零）。                                                                                                                                                                                                                                     |
| *Top*           | 可选      | Variant  | 新对象的左上角相对于文档左上角的位置（以磅为单位）。默认值为 0（零）。                                                                                                                                                                                                                                     |
| *Width*         | 可选      | Variant  | 以磅为单位给出 OLE 对象的初始尺寸。                                                                                                                                                                                                                                                                        |
| *Height*        | 可选      | Variant  | 以磅为单位给出 OLE 对象的初始尺寸。                                                                                                                                                                                                                                                                        |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加链接的 WPS 文档。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.AddOLEObject(null,"c:\\my documents\\testing.doc",true,null,null,null,null,100,100,200,300)
}

/*本示例向 myDocument 中添加新的命令按钮。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.AddOLEObject("Forms.CommandButton.1",null,null,null,null,null,null,100,100,100,200)
}
```

### Shapes.AddPicture

现有文件创建图片。返回一个代表新图片的 **Shape** 对象。

返回值类型为 Shape。

**语法**

**express.AddPicture ( *Filename* , *LinkToFile* , *SaveWithDocument* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型    | 说明                                             |
|--------------------|-----------|-------------|--------------------------------------------------|
| *Filename*         | 必选      | String      | 要在其中创建 OLE 对象的文件。                    |
| *LinkToFile*       | 必选      | MsoTriState | 要链接至的文件。                                 |
| *SaveWithDocument* | 必选      | MsoTriState | 将图片与文档一起保存。                           |
| *Left*             | 必选      | Single      | 图片左上角相对于文档左上角的位置（以磅为单位）。 |
| *Top*              | 必选      | Single      | 图片左上角相对于文档顶部的位置（以磅为单位）。   |
| *Width*            | 必选      | Single      | 图片的宽度（以磅为单位）。                       |
| *Height*           | 必选      | Single      | 图片的高度（以磅为单位）。                       |

**说明**

|                                                   |
|---------------------------------------------------|
| **MsoTriState** 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                      |
| **msoFalse** 使图片成为其源文件的独立副本。       |
| **msoTriStateMixed**                              |
| **msoTriStateToggle**                             |
| **msoTrue** 建立图片与其源文件之间的链接。        |

|                                                                                                                   |
|-------------------------------------------------------------------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。                                                             |
| **msoCTrue**                                                                                                      |
| **msoFalse** 在文档中只存储链接信息。                                                                             |
| **msoTriStateMixed**                                                                                              |
| **msoTriStateToggle**                                                                                             |
| **msoTrue** 将链接图片与该图片插入的文档一起保存。如果 LinkToFile 为 **msoFalse** ，则该参数必须为 **msoTrue** 。 |

**示例**

``` JavaScript
/*本示例向 myDocument 中添加由文件“Music.bmp”创建的图片。插入的图片链接到创建该图片的文件，并与 myDocument 一起保存。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.AddPicture("c:\\microsoft office\\clipart\\music.bmp", true, true, 100, 100, 70, 70)
}
```

### Shapes.AddShape

返回一个 **Shape** 对象，该对象表示工作表中的新自选形状。

返回值类型为Shape。

**语法**

**express.AddShape ( *Type* , *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型         | 说明                                                       |
|----------|-----------|------------------|------------------------------------------------------------|
| *Type*   | 必选      | MsoAutoShapeType | 指定要创建的自选形状的类型。                               |
| *Left*   | 必选      | Single           | 自选形状边框的左上角相对于文档左上角的位置（以磅为单位）。 |
| *Top*    | 必选      | Single           | 自选形状边框的左上角相对于文档左上角的位置（以磅为单位）。 |
| *Width*  | 必选      | Single           | 自选形状边框的宽度（以磅为单位）。                         |
| *Height* | 必选      | Single           | 自选形状边框的高度（以磅为单位）。                         |

**说明**

要更改已添加的自选形状的类型，请设置 **AutoShapeType** 属性。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加矩形。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200)
}
```

### Shapes.Item

从集合中返回一个对象。

返回值： 包含在集合中的一个 **Shape** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**说明**

对象的文本名称就是 **Name** 属性的值。

**示例**

``` JavaScript
/*本示例对 Shapes 集合中第二个形状的 OnAction 属性进行设置。如果变量 ss 代表的不是 Shapes 对象，则本示例无效。*/
function test()
{
    let ss = Application.Worksheets.Item(1).Shapes
    ss.Item(2).OnAction = "ShapeAction"
}
```

### Shapes.Range

返回一个 **ShapeRange** 对象，它代表 **Shapes** 集合中形状的子集。

**语法**

**express.Range ( *Index* )**

\*express是一个代表 Shapes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                             |
|---------|-----------|----------|------------------------------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | t 包含在该区域中的各单个形状。可以是指定形状索引号的整数、指定形状名称的字符串，也可以是包含整数或字符串的数组。 |

**说明**

虽然使用 **Range** 属性可返回任意数量的形状，但如果要返回集合中单个成员时，用 Item 方法更加简单。例如， ` Shapes(1) ` 比 ` Shapes.Range(1) ` 简单。

若要为 **Index** 指定一个整数或字符串数组，可以使用 **Array** 函数。例如，以下指令返回用名称指定的两个形状。

` Dim arShapes() As Variant Dim objRange As Object arShapes = Array("Oval 4", "Rectangle 5") Set objRange = ActiveSheet.Shapes.Range(arShapes) `

在 ET 中，不能用此属性返回包含工作表上的所有 Shape 对象的 ShapeRange 对象。如果要达到该目的，可用下列代码：

` Worksheets(1).Shapes.SelectAll ' select all shapes set sr = Selection.ShapeRange ' create ShapeRange `

### Shapes.SelectAll

选择指定的 **Shap es** 集合中的所有形状。

**语法**

**express.SelectAll ()**

\*express是一个代表 Shapes 对象的变量。

**示例**

``` JavaScript
/*本示例选择 myDocument 中的所有形状，并创建包含所有这些形状的 ShapeRange 集合。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.SelectAll()

    let sr = Application.Selection.ShapeRange
}
```

## 成员属性

### Shapes.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Shapes 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET"){
        alert("This is an ET Application object.")
    }
    else{
        alert("This is not an ET Application object.")
    }
}
```

### Shapes.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Shapes 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Shapes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# UniqueValues 对象

## 简介

UniqueValues 对象使用 DupeUnique 属性返回或设置一个枚举，该枚举确定规则是查找区域中的重复值还是唯一值。

## 方法

| 序号 | 名称                                                       | 说明                                                                       |
|:----:|:-----------------------------------------------------------|:---------------------------------------------------------------------------|
|  1   | [Delete](#UniqueValues.Delete)                             | 删除指定的条件格式规则对象。                                               |
|  2   | [ModifyAppliesToRange](#UniqueValues.ModifyAppliesToRange) | 设置此格式规则所应用于的单元格区域。                                       |
|  3   | [SetLastPriority](#UniqueValues.SetLastPriority)           | 为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则。 |

## 属性

| 序号 | 名称                                       | 说明                                                                                                                                                               |
|:----:|:-------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#UniqueValues.Application)   | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [AppliesTo](#UniqueValues.AppliesTo)       | 返回一个 Range 对象，指定格式规则将应用于的单元格区域。                                                                                                            |
|  3   | [Borders](#UniqueValues.Borders)           | 返回一个 Borders 集合，该集合在条件格式规则的计算结果为 True 时指定单元格边框的格式。只读。                                                                        |
|  4   | [Creator](#UniqueValues.Creator)           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  5   | [DupeUnique](#UniqueValues.DupeUnique)     | 返回或设置 XlDupeUnique 枚举的常量之一，指定条件格式规则是查找唯一的值还是查找重复的值。                                                                           |
|  6   | [Font](#UniqueValues.Font)                 | 返回一个 Font 对象，该对象在条件格式规则的计算结果为 True 时指定字体格式。只读。                                                                                   |
|  7   | [Interior](#UniqueValues.Interior)         | 返回一个 Interior 对象，该对象在条件格式规则的计算结果为 True 时指定单元格的内部属性。只读。                                                                       |
|  8   | [NumberFormat](#UniqueValues.NumberFormat) | 在条件格式规则的计算结果为 True 时返回或设置应用于单元格的数字格式。 Variant 型，可读写。                                                                          |
|  9   | [Parent](#UniqueValues.Parent)             | 返回指定对象的父对象。只读。                                                                                                                                       |
|  10  | [Priority](#UniqueValues.Priority)         | 返回或设置条件格式规则的优先级值。当工作表中存在多个条件格式规则时，优先级确定求值的顺序。                                                                         |
|  11  | [PTCondition](#UniqueValues.PTCondition)   | 返回一个 Boolean 值，指明是否将条件格式应用于数据透视表图表。只读。                                                                                                |
|  12  | [ScopeType](#UniqueValues.ScopeType)       | 返回或设置 XlPivotConditionScope 枚举的常量之一，该常量确定条件格式在应用于数据透视表图表时的范围。                                                                |
|  13  | [StopIfTrue](#UniqueValues.StopIfTrue)     | 返回或设置一个 Boolean 值，该值确定在当前规则的计算结果为 True 时是否应计算单元格上的其他格式规则。                                                                |
|  14  | [Type](#UniqueValues.Type)                 | 返回 XlFormatConditionType 枚举的常量之一，该常量指定条件格式的类型。只读。                                                                                        |

## 成员方法

### UniqueValues.Delete

删除指定的条件格式规则对象。

**语法**

**express.Delete ()**

\*express是一个代表 UniqueValues 对象的变量。

### UniqueValues.ModifyAppliesToRange

设置此格式规则所应用于的单元格区域。

**语法**

**express.ModifyAppliesToRange ( *Range* )**

\*express是一个代表 UniqueValues 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                       |
|---------|-----------|----------|----------------------------|
| *Range* | 必选      | Range    | 此格式规则将应用于的区域。 |

**说明**

该区域必须采用 A1 引用样式，并全部包含在 **FormatConditions** 集合的父工作表内。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可以使用货币符号，但会被忽略。

您也可以在区域的任意部分中使用局部定义名称，但该名称必须使用宏语言。

### UniqueValues.SetLastPriority

为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则。

**语法**

**express.SetLastPriority ()**

\*express是一个代表 UniqueValues 对象的变量。

**说明**

优先级的实际值将等于工作表上条件格式规则的总数。当工作表中有多个条件格式规则时，此方法将导致优先级值大于此规则的规则的优先级增加 1。

**注释：**

|                                          |
|------------------------------------------|
| 将在工作表级别应用条件格式规则的优先级。 |

## 成员属性

### UniqueValues.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 UniqueValues 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### UniqueValues.AppliesTo

返回一个 **Range** 对象，指定格式规则将应用于的单元格区域。

**语法**

**express.AppliesTo**

\*express是一个代表 UniqueValues 对象的变量。

### UniqueValues.Borders

返回一个 **Borders** 集合，该集合在条件格式规则的计算结果为 **True** 时指定单元格边框的格式。只读。

**语法**

**express.Borders**

\*express是一个代表 UniqueValues 对象的变量。

**说明**

对于条件格式对象，您只能为单元格的上边框、下边框和侧边框设置属性。

### UniqueValues.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 UniqueValues 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### UniqueValues.DupeUnique

返回或设置 **XlDupeUnique** 枚举的常量之一，指定条件格式规则是查找唯一的值还是查找重复的值。

**语法**

**express.DupeUnique**

\*express是一个代表 UniqueValues 对象的变量。

### UniqueValues.Font

返回一个 **Font** 对象，该对象在条件格式规则的计算结果为 **True** 时指定字体格式。只读。

**语法**

**express.Font**

\*express是一个代表 UniqueValues 对象的变量。

**说明**

条件格式对象并不支持 **Font** 对象的所有属性。您可以设置字体样式、下划线、颜色和删除线属性。

### UniqueValues.Interior

返回一个 **Interior** 对象，该对象在条件格式规则的计算结果为 **True** 时指定单元格的内部属性。只读。

**语法**

**express.Interior**

\*express是一个代表 UniqueValues 对象的变量。

### UniqueValues.NumberFormat

在条件格式规则的计算结果为 **True** 时返回或设置应用于单元格的数字格式。 **Variant** 型，可读写。

**语法**

**express.NumberFormat**

\*express是一个代表 UniqueValues 对象的变量。

**说明**

数字格式是使用 **“单元格格式”** 对话框的 **“数字”** 选项卡上显示的相同格式代码指定的。您可以使用内置的数字格式，例如 ` "General" ` 或者 创建自定义数字格式 。

### UniqueValues.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 UniqueValues 对象的变量。

### UniqueValues.Priority

返回或设置条件格式规则的优先级值。当工作表中存在多个条件格式规则时，优先级确定求值的顺序。

**语法**

**express.Priority**

\*express是一个代表 UniqueValues 对象的变量。

**说明**

设置优先级时，值必须为介于 1 和工作表上条件格式规则总数之间的正整数。对于工作表上的所有规则，优先级必须为唯一值，因此，如果更改指定条件格式规则的优先级，将可能会导致工作表上其他规则的优先级值移位。

### UniqueValues.PTCondition

返回一个 **Boolean** 值，指明是否将条件格式应用于数据透视表图表。只读。

**语法**

**express.PTCondition**

\*express是一个代表 UniqueValues 对象的变量。

### UniqueValues.ScopeType

返回或设置 **XlPivotConditionScope** 枚举的常量之一，该常量确定条件格式在应用于数据透视表图表时的范围。

**语法**

**express.ScopeType**

\*express是一个代表 UniqueValues 对象的变量。

**说明**

默认值为 **xlSelectionScope** ，它使用 **AppliesTo** 属性设置范围。

### UniqueValues.StopIfTrue

返回或设置一个 **Boolean** 值，该值确定在当前规则的计算结果为 **True** 时是否应计算单元格上的其他格式规则。

**语法**

**express.StopIfTrue**

\*express是一个代表 UniqueValues 对象的变量。

**说明**

为了支持向后兼容性，此属性的默认值为 **True** ，而这与在用户界面中创建的格式规则的行为相反。

### UniqueValues.Type

返回 **XlFormatConditionType** 枚举的常量之一，该常量指定条件格式的类型。只读。

**语法**

**express.Type**

\*express是一个代表 UniqueValues 对象的变量。

**说明**

此属性将始终返回 **Long** 值“8”，等价于 **xlUniqueValues** 。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/UniqueValues](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# UpBars 对象

## 简介

涨柱线将图表组中第一个系列的数据点与最后一个系列中相应的有较大值的数据点连接起来（从第一个系列向上生长）。只有至少包含两个系列的二维折线图才能有涨柱线。此对象不是集合。没有代表单个涨柱线的对象；或者打开图表组中所有数据点的涨柱线，或者将其全部关闭。

如果 HasUpDownBars 属性是 False ， UpBars 对象的绝大部分属性将被禁用。

## 说明

使用 UpBars 属性可返回 UpBars 对象。下例打开 Sheet5 上嵌入式图表一中图表组一的涨跌柱线。然后，本示例将涨柱线的颜色设置为蓝色，而将跌柱线设置为红色。

``` JavaScript
function TEST()
{
   let rng = Worksheets.Item("sheet5").ChartObjects(1).Chart.ChartGroups(1)
   rng.HasUpDownBars = true
   rng.UpBars.Interior.Color.RGB = (0, 0, 255)
   rng.DownBars.Interior.Color.RGB = (255, 0, 0)
}
```

## 方法

| 序号 | 名称                     | 说明       |
|:----:|:-------------------------|:-----------|
|  1   | [Delete](#UpBars.Delete) | 删除对象。 |
|  2   | [Select](#UpBars.Select) | 选择对象。 |

## 属性

| 序号 | 名称                               | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#UpBars.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#UpBars.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Format](#UpBars.Format)           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  4   | [Name](#UpBars.Name)               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  5   | [Parent](#UpBars.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### UpBars.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 UpBars 对象的变量。

**返回值**

Variant

### UpBars.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 UpBars 对象的变量。

**返回值**

Variant

## 成员属性

### UpBars.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 UpBars 对象的变量。

**示例**

``` JavaScript
//本示例显示一条有关创建 myObject 的应用程序的消息。
function test()
{
    let myObject = Application.ActiveWorkbook
    if(myObject.Value == "ET") 
    {
        alert("This is an ET Application object.")
    }
    else 
    {
        alert("This is not an ET Application object.")
    }
}
```

### UpBars.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 UpBars 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### UpBars.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 UpBars 对象的变量。

### UpBars.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 UpBars 对象的变量。

### UpBars.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 UpBars 对象的变量。

**说明**

> WPS 开发文档： [WPS 基础接口/表格 API 参考/UpBars](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Comments 对象

## 简介

Comment 对象的集合，这些对象代表选定内容、区域或文档中的批注。

## 方法

| 序号 | 名称                   | 说明                                                  |
|:----:|:-----------------------|:------------------------------------------------------|
|  1   | [Add](#Comments.Add)   | 返回一个 Comment 对象，该对象代表添加到区域中的备注。 |
|  2   | [Item](#Comments.Item) | 返回集合中的单个 Comment 对象。                       |

## 属性

| 序号 | 名称                                 | 说明                                                                           |
|:----:|:-------------------------------------|:-------------------------------------------------------------------------------|
|  1   | [Application](#Comments.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                 |
|  2   | [Count](#Comments.Count)             | 返回一个 Number 类型的值，该值代表 Comments 集合中的项目数。只读。             |
|  3   | [Creator](#Comments.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。   |
|  4   | [Parent](#Comments.Parent)           | 返回一个 Object 类型值，该值代表指定 Comments 对象的父对象。                   |
|  5   | [ShowBy](#Comments.ShowBy)           | 返回或设置审阅者的名字，该审阅者的批注显示在批注窗格中。 String 类型，可读写。 |

## 成员方法

### Comments.Add

返回一个 **Comment** 对象，该对象代表添加到区域中的备注。

**语法**

**express.Add ( *Range* , *Text* )**

\*express是一个代表 Comments 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明               |
|---------|-----------|----------|--------------------|
| *Range* | 必选      | Range    | 要添加备注的区域。 |
| *Text*  | 可选      | Variant  | 备注文本。         |

**示例**

``` JavaScript
/* 在选区开始位置，插入一条备注 */
function test(){
  Application.Selection.Collapse(wdCollapseEnd);
  Application.ActiveDocument.Comments.Add(Application.Selection.Range,"review this");
}

/* 在活动文档的第三段添加一条备注 */
function test(){
  let myRange = Application.ActiveDocument.Paragraphs.Item(3).Range
  Application.ActiveDocument.Comments.Add(myRange, "original third paragraph")
}
```

### Comments.Item

返回集合中的单个 **Comment** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Comments 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                           |
|---------|-----------|----------|----------------------------------------------------------------|
| *Index* | 必选      | Number   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

## 成员属性

### Comments.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Comments 对象的变量。

### Comments.Count

返回一个 **Number** 类型的值，该值代表 **Comments** 集合中的项目数。只读。

**语法**

**express.Count**

\*express是一个代表 Comments 对象的变量。

### Comments.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Comments 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Comments.Parent

返回一个 **Object** 类型值，该值代表指定 **Comments** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Comments 对象的变量。

### Comments.ShowBy

返回或设置审阅者的名字，该审阅者的批注显示在批注窗格中。 **String** 类型，可读写。

**语法**

**express.ShowBy**

\*express是一个代表 Comments 对象的变量。

**说明**

可选择显示一位审阅者或者所有审阅者所做的批注。若要查看所有审阅者的批注，可将本属性设置为“所有审阅者”。

**示例**

``` JavaScript
/*下面的示例显示“Don Funk”所做的批注。*/
function test() {
    if (Application.ActiveDocument.Comments.Count >= 1) {
        Application.ActiveDocument.ActiveWindow.View.SplitSpecial = wdPaneComments
        Application.ActiveDocument.Comments.ShowBy = "Don Funk"
    }
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Comments](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ContentControl 对象

## 简介

一个内容控件。内容控件是固定在文档中并可能带有标记的区域，这些区域作为特定类型的内容的容器。单个内容控件可能包含日期、列表或带格式文本段落等内容。 ContentControl 对象是 ContentControls 集合的成员。

## 说明

使用 ContentControls 集合的 Add 方法可创建内容控件。使用 Add 方法的 *Type* 参数可以指定要创建的内容控件的型。以下示例创建了一个新的下拉列表内容控件，并在该列表中添加了若干个项。

``` JavaScript
function test() {
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
    //List entries
    objCC.DropdownListEntries.Add("Cat")
    objCC.DropdownListEntries.Add("Dog")
    objCC.DropdownListEntries.Add("Horse")
    objCC.DropdownListEntries.Add("Monkey")
    objCC.DropdownListEntries.Add("Snake")
    objCC.DropdownListEntries.Add("Other")
}
```

使用 Type 属性可以更改内容控件的类型。例如，您或许希望将日期控件更改为文本控件。但是，您可能无法更改所有内容控件的类型；一些内容控件可能不允许更改其类型。此外，根据内容控件的内容，您也许不能更改控件的类型。例如，如果作为更改目标的内容控件不允许使用现有内容控件中内容的类型，那么，类型更改尝试将受到禁止并将生成运行时错误。

以下示例插入一个日期内容控件并且设置其值，然后将该控件更改为文本内容控件。

``` JavaScript
function test() {
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDate)
    objCC.Range.Text = ("January 1, 2007")
    objCC.Type = wdContentControlText
}
```

使用 SetPlaceholderText 方法可以将占位符文本从默认字符串更改为更适合该控件的内容。使用 Title 属性可以指定该控件的标题文本。在将光标放置到该控件的内部或者将鼠标指针放置到控件上方时，将在控件上方显示此标题。

根据您具有的内容控件的类型，您可能无法使用 ContentControl 对象的某些属性和方法。

并非所有内容控件属性都适用于所有不同类型的内容控件。下表列出了哪些属性适用于哪些类型的内容控件。

| 属性/方法                 | 适用于                                                                               |
|---------------------------|--------------------------------------------------------------------------------------|
| BuildingBlockCategory属性 | BuildingBlock 库内容控件 (wdContentControlBuildingBlockGallery)                      |
| BuildingBlockType属性     | BuildingBlock 库内容控件 (wdContentControlBuildingBlockGallery)                      |
| DateDisplayFormat属性     | 日期内容控件 (wdContentControlDate)                                                  |
| DateDisplayLocale属性     | 日期内容控件 (wdContentControlDate)                                                  |
| DateStorageFormat属性     | 日期内容控件 (wdContentControlDate)                                                  |
| DropdownListEntries属性   | 组合框和下拉列表内容控件（wdContentControlComboBox 和 wdContentControlDropdownList） |
| MultiLine属性             | 纯文本内容控件 (wdContentControlText)                                                |
| Ungroup方法               | 组内容控件 (wdContentControlGroup)                                                   |

## 方法

| 序号 | 名称                                                     | 说明                                                                             |
|:----:|:---------------------------------------------------------|:---------------------------------------------------------------------------------|
|  1   | [Copy](#ContentControl.Copy)                             | 将活动文档中的内容控件复制到剪贴板。                                             |
|  2   | [Cut](#ContentControl.Cut)                               | 从活动文档中删除内容控件并将该内容控件移动到剪贴板中。                           |
|  3   | [Delete](#ContentControl.Delete)                         | 删除指定的内容控件以及内容控件中的内容。                                         |
|  4   | [SetCheckedSymbol](#ContentControl.SetCheckedSymbol)     | 设置用于代表复选框内容控件的选中状态的符号。                                     |
|  5   | [SetPlaceholderText](#ContentControl.SetPlaceholderText) | 设置在用户输入自己的文本之前显示在内容控件中的占位符文本。                       |
|  6   | [SetUncheckedSymbol](#ContentControl.SetUncheckedSymbol) | 设置用于代表复选框内容控件的未选中状态的符号。                                   |
|  7   | [Ungroup](#ContentControl.Ungroup)                       | 从文档中删除一个组内容控件，以便它的子内容控件不再嵌套，并且可以自由地进行编辑。 |

## 属性

| 序号 | 名称                                                             | 说明                                                                                                                                    |
|:----:|:-----------------------------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ContentControl.Application)                       | 返回一个 Application 对象，该对象代表 WPS 应用程序。                                                                                    |
|  2   | [BuildingBlockCategory](#ContentControl.BuildingBlockCategory)   | 返回或设置一个 String ，代表生成块内容控件的类别。可读写。                                                                              |
|  3   | [BuildingBlockType](#ContentControl.BuildingBlockType)           | 返回或设置一个 WdBuildingBlockTypes 常量，该常量代表生成块内容控件的生成块的类型。可读写。                                              |
|  4   | [Checked](#ContentControl.Checked)                               | 返回或设置 Boolean 类型的值，代表复选框的当前状态（选中/未选中）。可读/写。                                                             |
|  5   | [Creator](#ContentControl.Creator)                               | 返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 Long 类型。                                                            |
|  6   | [DateCalendarType](#ContentControl.DateCalendarType)             | 返回或设置一个 WdCalendarType 常量，该常量表示日历内容控件的日历类型。可读写。                                                          |
|  7   | [DateDisplayFormat](#ContentControl.DateDisplayFormat)           | 返回或设置 String 值，该值表示日期的显示格式。可读/写。                                                                                 |
|  8   | [DateDisplayLocale](#ContentControl.DateDisplayLocale)           | 返回一个 WdLanguageID ，它代表日期内容控件中显示的日期的语言格式。可读写。                                                              |
|  9   | [DateStorageFormat](#ContentControl.DateStorageFormat)           | 返回或设置 WdContentControlDateStorageFormat 值，该值表示在将日期内容控件绑定到活动文档的 XML 数据存储时的日期存储和检索格式。可读/写。 |
|  10  | [DefaultTextStyle](#ContentControl.DefaultTextStyle)             | 返回或设置一个 Variant ，它代表用于设置文本内容控件中文本格式的字符样式的名称。可读写。                                                 |
|  11  | [DropdownListEntries](#ContentControl.DropdownListEntries)       | 返回 ContentControlListEntries 集合，该集合表示下拉列表内容控件或组合框内容控件中的项。只读。                                           |
|  12  | [ID](#ContentControl.ID)                                         | 返回一个代表内容控件标识的 String 类型的值。只读。                                                                                      |
|  13  | [LockContentControl](#ContentControl.LockContentControl)         | 返回或设置 Boolean 值，该值表示用户能否从活动文档中删除内容控件。可读/写。                                                              |
|  14  | [LockContents](#ContentControl.LockContents)                     | 返回或设置 Boolean 值，该值表示用户能否编辑内容控件的内容。可读/写。                                                                    |
|  15  | [MultiLine](#ContentControl.MultiLine)                           | 返回 或设置 Boolean 值，该值表示文本内容控件是否允许多行文本。可读/写。                                                                 |
|  16  | [Parent](#ContentControl.Parent)                                 | 返回一个 Object 类型的值，该值代表指定 ContentControl 对象的父对象。                                                                    |
|  17  | [ParentContentControl](#ContentControl.ParentContentControl)     | 返回一个 ContentControl 类型的值，该值代表嵌套在格式文本控件或组控件内的内容控件的父内容控件。只读。                                    |
|  18  | [PlaceholderText](#ContentControl.PlaceholderText)               | 返回一个 BuildingBlock 对象，该对象代表内容控件的占位符文本。只读。                                                                     |
|  19  | [Range](#ContentControl.Range)                                   | 返回 Range ，它表示活动文档中的内容控件的内容。只读。                                                                                   |
|  20  | [ShowingPlaceholderText](#ContentControl.ShowingPlaceholderText) | 返回 Boolean 值，该值指示是否显示内容控件的占位符文本。只读。                                                                           |
|  21  | [Tag](#ContentControl.Tag)                                       | t返回或设置一个 String 类型的值，该值代表用于识别内容控件的值。可读写。                                                                 |
|  22  | [Temporary](#ContentControl.Temporary)                           | 返回或设置 Boolean 值，该值表示当用户编辑内容控件的内容时是否从活动文档中删除该控件。可读/写。                                          |
|  23  | [Title](#ContentControl.Title)                                   | 返回或设置 String 值，该值表示内容控件的标题。可读/写。                                                                                 |
|  24  | [Type](#ContentControl.Type)                                     | 返回或设置 WdContentControlType ，它表示内容控件的类型。可读/写。                                                                       |
|  25  | [XMLMapping](#ContentControl.XMLMapping)                         | 返回 XMLMapping 对象，该对象表示内容控件到文档的数据存储中的 XML 数据的映射。只读。                                                     |

## 成员方法

### ContentControl.Copy

将活动文档中的内容控件复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 ContentControl 对象的变量。

**说明**

使用 **Copy** 方法时，原始内容控件仍保留在活动文档中，但控件的副本（包括所有文本和属性设置）将移动到剪贴板中。然后，可以将内容控件粘贴到活动文档的其他部分。可以使用 **Selection** 对象的 **Paste** 方法或 **Range** 对象的 **Paste** 方法插入复制的内容控件，也可以使用 WPS 中的 **“粘贴”** 功能。

### ContentControl.Cut

从活动文档中删除内容控件并将该内容控件移动到剪贴板中。

**语法**

**express.Cut ()**

\*express是一个代表 ContentControl 对象的变量。

**说明**

可以使用 **Selection** 对象的 **Paste** 方法或 **Range** 对象的 **Paste** 方法插入剪切内容控件，也可以使用 WPS 中的 **“粘贴”** 功能。

### ContentControl.Delete

删除指定的内容控件以及内容控件中的内容。

**语法**

**express.Delete ( *DeleteContents* )**

\*express是一个代表 ContentControl 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                                                                                |
|------------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------|
| *DeleteContents* | 可选      | Boolean  | 指定是否删除内容控件的内容。如果值为 True，将同时删除内容控件及其内容。如果值为 False，将删除内容控件，但将其内容保留在活动文档中。默认值为 False。 |

**示例**

``` JavaScript
/* 删除当前活动文档第一个内容控件 */
Application.ActiveDocument.ContentControls.Item(1).Delete()
```

### ContentControl.SetCheckedSymbol

设置用于代表复选框内容控件的选中状态的符号。

**语法**

**express.SetCheckedSymbol ( *CharacterNumber* , *Font* )**

\*express是一个代表 ContentControl 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                |
|-------------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *CharacterNumber* | 必选      | Long     | 指定符号的 Unicode 字符编号。该值始终是 31（字体开始处的控件符号的数目）加上此符号在符号表中的位置（从左向右计算）相应的数字。例如，若要指定一个位于 Symbol 字体符号表中第 37 号位置的 delta 字符，请将 CharacterNumber 设置为 68。 |
| *Font*            | 可选      | String   | 包含符号的字体的名称。                                                                                                                                                                                                              |

**示例**

``` JavaScript
/*下面的代码示例将指定内容控件的选中符号设置为“MS Gothic”字体“Ballot Box with X”符号。 */
function test() {
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlCheckBox)
    objCC.SetCheckedSymbol(2612, "MS Gothic")
}
```

### ContentControl.SetPlaceholderText

设置在用户输入自己的文本之前显示在内容控件中的占位符文本。

**语法**

**express.SetPlaceholderText ( *BuildingBlock* , *Range* , *Text* )**

\*express是一个代表 ContentControl 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型      | 说明                                                  |
|-----------------|-----------|---------------|-------------------------------------------------------|
| *BuildingBlock* | 可选      | BuildingBlock | 指定 BuildingBlock 对象，该对象包含占位符文本的内容。 |
| *Range*         | 可选      | Range         | 指定 Range 对象，该对象包含占位符文本的内容。         |
| *Text*          | 可选      | String        | 指定占位符文本的内容。                                |

**说明**

在指定占位符文本时，只使用一个参数。如果使用多个参数，则 WPS 将使用在第一个参数中指定的文本。如果省略所有参数，则占位符文本将为空白。

**示例**

``` JavaScript
/*以下示例在活动文档中插入新的下拉列表内容控件，设置标题和占位符文本，然后在列表中插入几个新项。*/
function test() {
    let objMap
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
    objCC.Title = "My Favorite Animal"
    objCC.SetPlaceholderText, "Select your favorite animal "

    /*List entries*/
    objCC.DropdownListEntries.Add("Cat")
    objCC.DropdownListEntries.Add("Dog")
    objCC.DropdownListEntries.Add("Horse")
    objCC.DropdownListEntries.Add("Monkey")
    objCC.DropdownListEntries.Add("Snake")
    objCC.DropdownListEntries.Add("Other")
}
```

### ContentControl.SetUncheckedSymbol

设置用于代表复选框内容控件的未选中状态的符号。

**语法**

**express.SetUncheckedSymbol ( *CharacterNumber* , *Font* )**

\*express是一个代表 ContentControl 对象的变量。

**参数**

| 名称              | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                |
|-------------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *CharacterNumber* | 必选      | Long     | 指定符号的 Unicode 字符编号。该值始终是 31（字体开始处的控件符号的数目）加上此符号在符号表中的位置（从左向右计算）相应的数字。例如，若要指定一个位于 Symbol 字体符号表中第 37 号位置的 delta 字符，请将 CharacterNumber 设置为 68。 |
| *Font*            | 可选      | String   | 包含符号的字体的名称。                                                                                                                                                                                                              |

**示例**

``` JavaScript
/*下面的代码示例将指定内容控件的未选中符号设置为“MS Gothic”字体“Ballot Box”符号。*/
function test() {
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlCheckBox)
    objCC.SetUncheckedSymbol(2610, "MS Gothic")
}
```

### ContentControl.Ungroup

从文档中删除一个组内容控件，以便它的子内容控件不再嵌套，并且可以自由地进行编辑。

**语法**

**express.Ungroup ()**

\*express是一个代表 ContentControl 对象的变量。

**说明**

如果您在非 **wdContentControlGroup** 类型的控件上运行该方法，则该方法将失败。

## 成员属性

### ContentControl.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 ContentControl 对象的变量。

### ContentControl.BuildingBlockCategory

返回或设置一个 **String** ，代表生成块内容控件的类别。可读写。

**语法**

**express.BuildingBlockCategory**

\*express是一个代表 ContentControl 对象的变量。

**说明**

此属性仅适用于生成块内容控件，并且相当于 **“内容控件属性”** 对话框中的 **“类别”** 选项。您可以将此属性设置为任何字符串；但如果您将它设置为不存在相应类别的字符串，则 **“类别”** 选项的值被设置为“(所有类别)”。

**示例**

``` JavaScript
/*以下示例创建了一个新的生成块内容控件，并且指定了生成块的类型以及库。*/
function test() {
    let objBB = Application.Selection.ContentControls.Add(wdContentControlBuildingBlockGallery)
    objBB.BuildingBlockType = wdTypeEquations
    objBB.BuildingBlockCategory = "General"
}
```

### ContentControl.BuildingBlockType

返回或设置一个 **WdBuildingBlockTypes** 常量，该常量代表生成块内容控件的生成块的类型。可读写。

**语法**

**express.BuildingBlockType**

\*express是一个代表 ContentControl 对象的变量。

**说明**

此属性仅适用于生成块内容控件，并且相当于 **“内容控件属性”** 对话框中的 **“库”** 选项。您只能为下列生成块类型设置此属性：

- 自定义 1 到自定义 5
- 自动图文集
- 文档部件
- 自定义自动图文集
- 自定义文档部件
- 公式

**示例**

``` JavaScript
/*以下示例创建了一个新的生成块内容控件，并且指定了生成块的类型以及库。*/
function test() {
    let objBB = Application.Selection.ContentControls.Add(wdContentControlBuildingBlockGallery)
    objBB.BuildingBlockType = wdTypeEquations
    objBB.BuildingBlockCategory = "General"
}
```

### ContentControl.Checked

返回或设置 **Boolean** 类型的值，代表复选框的当前状态（选中/未选中）。可读/写。

**语法**

**express.Checked**

\*express是一个代表 ContentControl 对象的变量。

**说明**

使用 **Checked** 属性获取/设置复选框内容控件的当前状态。如果该控件不是复选框，尝试访问该属性将失败，并导致运行时错误“该属性仅可用于复选框内容控件”。

**示例**

``` JavaScript
/*下面的代码示例设置指定复选框内容控件的 Checked 属性。*/
function test() {
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlCheckBox)
    objCC.Title = "Send Reminder"
    objCC.Checked = true
}
```

### ContentControl.Creator

返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ContentControl 对象的变量。

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

### ContentControl.DateCalendarType

返回或设置一个 **WdCalendarType** 常量，该常量表示日历内容控件的日历类型。可读写。

**语法**

**express.DateCalendarType**

\*express是一个代表 ContentControl 对象的变量。

### ContentControl.DateDisplayFormat

返回或设置 **String** 值，该值表示日期的显示格式。可读/写。

**语法**

**express.DateDisplayFormat**

\*express是一个代表 ContentControl 对象的变量。

**说明**

默认格式是在用户的系统上在 WPS 中指定的格式设置，通常取决于 Microsoft Windows 中的位置设置。例如，英语（美国）的默认日期格式为“mm/dd/yyyy”。使用 **DateDisplayFormat** 属性可以指定另一种日期格式。

**示例**

``` JavaScript
function test() {
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDate)
    objCC.Title = "Review Period End Date"
    objCC.DateDisplayFormat = "MMMM d, yyyy"
    objCC.DateStorageFormat = wdContentControlDateStorageDate
    objCC.Range.Text = "January 1, 2007"
}
```

### ContentControl.DateDisplayLocale

返回一个 **WdLanguageID** ，它代表日期内容控件中显示的日期的语言格式。可读写。

**语法**

**express.DateDisplayLocale**

\*express是一个代表 ContentControl 对象的变量。

**说明**

默认情况下， **DateDisplayLocale** 属性最初设置为 WPS 中指定的编辑语言（前提是它与 Microsoft Windows 中指定的位置不同）。

### ContentControl.DateStorageFormat

返回或设置 **WdContentControlDateStorageFormat** 值，该值表示在将日期内容控件绑定到活动文档的 XML 数据存储时的日期存储和检索格式。可读/写。

**语法**

**express.DateStorageFormat**

\*express是一个代表 ContentControl 对象的变量。

**说明**

**DateStorageFormat** 属性用于以日期格式、日期/时间格式或文本格式存储日期。

**示例**

``` JavaScript
/*以下示例将日期内容控件添加到活动文档中，并指定日期、日期显示格式和日期存储格式。*/
function test() {
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDate)
    objCC.Title = "Review Period End Date"
    objCC.DateDisplayFormat = "MMMM d, yyyy"
    objCC.DateStorageFormat = wdContentControlDateStorageDate
    objCC.Range.Text = "January 1, 2007"
}
```

### ContentControl.DefaultTextStyle

返回或设置一个 **Variant** ，它代表用于设置文本内容控件中文本格式的字符样式的名称。可读写。

**语法**

**express.DefaultTextStyle**

\*express是一个代表 ContentControl 对象的变量。

### ContentControl.DropdownListEntries

返回 **ContentControlListEntries** 集合，该集合表示下拉列表内容控件或组合框内容控件中的项。只读。

**语法**

**express.DropdownListEntries**

\*express是一个代表 ContentControl 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动文档中插入新的下拉列表内容控件，设置标题和占位符文本，然后向列表中添加几个新项。*/
function test() {
    let objMap
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
    objCC.Title = "My Favorite Animal"
    objCC.SetPlaceholderText, "Select your favorite animal "

    /*List entries*/
    objCC.DropdownListEntries.Add("Cat")
    objCC.DropdownListEntries.Add("Dog")
    objCC.DropdownListEntries.Add("Horse")
    objCC.DropdownListEntries.Add("Monkey")
    objCC.DropdownListEntries.Add("Snake")
    objCC.DropdownListEntries.Add("Other")
}
```

### ContentControl.ID

返回一个代表内容控件标识的 **String** 类型的值。只读。

**语法**

**express.ID**

\*express是一个代表 ContentControl 对象的变量。

**说明**

**ID** 属性是一个内部号码，您不能更改它，但可以使用它在代码中标识内容控件。此号码对于每个内容控件都是唯一的，并且不会更改。

**示例**

``` JavaScript
Application.ActiveDocument.ContentControls.Item(1).ID
```

### ContentControl.LockContentControl

返回或设置 **Boolean** 值，该值表示用户能否从活动文档中删除内容控件。可读/写。

**语法**

**express.LockContentControl**

\*express是一个代表 ContentControl 对象的变量。

**说明**

此属性的默认值为 **False** 。此属性对应于 **“内容控制属性”** 对话框中的 **“不能删除内容控件”** 复选框。 如果 Temporary 属性设置为 **True** ，则您不能设置此属性。

**示例**

``` JavaScript
/* 插入日期内容控件，设置内容后，并且禁止编辑以及禁止删除 */
function test() {
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDate)
    objCC.Range.Text = "January 1, 2007"
    objCC.LockContents = true
    objCC.LockContentControl = true
}
```

### ContentControl.LockContents

返回或设置 **Boolean** 值，该值表示用户能否编辑内容控件的内容。可读/写。

**语法**

**express.LockContents**

\*express是一个代表 ContentControl 对象的变量。

**说明**

此属性的默认值为 **False** 。此属性对应于 **“内容控制属性”** 对话框中的 **“不能编辑内容”** 复选框。

**示例**

``` JavaScript
/* 设置当前文档第一个内容控件可以编辑 */
Application.ActiveDocument.ContentControls.Item(1).LockContents = false
```

### ContentControl.MultiLine

返回 或设置 **Boolean** 值，该值表示文本内容控件是否允许多行文本。可读/写。

**语法**

**express.MultiLine**

\*express是一个代表 ContentControl 对象的变量。

**说明**

此属性只应用于文本内容控件和格式文本内容控件。

### ContentControl.Parent

返回一个 **Object** 类型的值，该值代表指定 **ContentControl** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 ContentControl 对象的变量。

### ContentControl.ParentContentControl

返回一个 **ContentControl** 类型的值，该值代表嵌套在格式文本控件或组控件内的内容控件的父内容控件。只读。

**语法**

**express.ParentContentControl**

\*express是一个代表 ContentControl 对象的变量。

**说明**

如果内容控件未嵌套，则此属性返回错误。

### ContentControl.PlaceholderText

返回一个 **BuildingBlock** 对象，该对象代表内容控件的占位符文本。只读。

**语法**

**express.PlaceholderText**

\*express是一个代表 ContentControl 对象的变量。

**说明**

可以使用 **SetPlaceholderText** 方法设置内容控件的占位符文本。

### ContentControl.Range

返回 **Range** ，它表示活动文档中的内容控件的内容。只读。

**语法**

**express.Range**

\*express是一个代表 ContentControl 对象的变量。

**说明**

使用 **Range** 属性可以访问文本、文本格式和其他文本属性。有关详细信息，请参阅 处理 Range 对象。

**示例**

``` JavaScript
/* 插入日期内容控件，设置内容 */
function test() {
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDate)
    objCC.Range.Text = "January 1, 2007"
}
```

### ContentControl.ShowingPlaceholderText

返回 **Boolean** 值，该值指示是否显示内容控件的占位符文本。只读。

**语法**

**express.ShowingPlaceholderText**

\*express是一个代表 ContentControl 对象的变量。

**示例**

``` JavaScript
/* 添加一个文本内容控件，如果可以设置占位文本，那就设置为"hello" */
function test() {
    let txt = Application.ActiveDocument.ContentControls.Add(wdContentControlText);
    if (txt.ShowingPlaceholderText) {
        txt.SetPlaceholderText(null, null, "hello");
    }
}
```

### ContentControl.Tag

t返回或设置一个 **String** 类型的值，该值代表用于识别内容控件的值。可读写。

**语法**

**express.Tag**

\*express是一个代表 ContentControl 对象的变量。

**说明**

**Tag** 属性与 **Title** 属性的不同之处在于，用户在编辑文档时从不会显示标记。当文档处于打开状态时，开发人员可以利用该属性来存储编程操作的值。

### ContentControl.Temporary

返回或设置 **Boolean** 值，该值表示当用户编辑内容控件的内容时是否从活动文档中删除该控件。可读/写。

**语法**

**express.Temporary**

\*express是一个代表 ContentControl 对象的变量。

**说明**

默认值为 **False** 。此属性对应于 **“内容控件属性”** 对话框中的 **“内容被编辑后删除内容控件”** 复选框。

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
<td>如果 <strong><span>LockContentControl</span></strong> 属性设置为 <strong>True</strong> ，则您不能设置此属性。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

### ContentControl.Title

返回或设置 **String** 值，该值表示内容控件的标题。可读/写。

**语法**

**express.Title**

\*express是一个代表 ContentControl 对象的变量。

**示例**

``` JavaScript
/* 插入文本内容控件并设置标题 */
function test() {
    let txt = Application.ActiveDocument.ContentControls.Add(wdContentControlText);
    txt.Title = "txt title"
}
```

### ContentControl.Type

返回或设置 **WdContentControlType** ，它表示内容控件的类型。可读/写。

**语法**

**express.Type**

\*express是一个代表 ContentControl 对象的变量。

**说明**

可以使用 **Type** 属性将内容控件的类型从一种类型更改为另一种类型。但是，能否更改控件的类型将取决于原始类型和更改时内容控件中的内容。所有内容控件都可以更改为格式文本或生成块库类型的内容控件，因为这些类型允许任意内容。对于其他类型，如果内容对于要更改为的目标类型来说是有效的，则允许更改类型。否则，将拒绝更改，并导致运行时错误。

**示例**

``` JavaScript
/* 以下示例检查并确定指定的内容控件是否是下拉列表框或组合框。如果它是这两种类型之一，则上移列表的最后一项，使它成为列表的第一项。 */
function test() {
    let objCC = Application.ActiveDocument.ContentControls.Item(1);
    if (objCC.Type == wdContentControlComboBox || objCC.Type == wdContentControlDropdownList) {
        let objCL = objCC.DropdownListEntries.Item(objCC.DropdownListEntries.Count)
        for (let i = 1; i <= objCC.DropdownListEntries.Count - 1; i++) {
            objCL.MoveUp()
        }
    }
}
```

### ContentControl.XMLMapping

返回 **XMLMapping** 对象，该对象表示内容控件到文档的数据存储中的 XML 数据的映射。只读。

**语法**

**express.XMLMapping**

\*express是一个代表 ContentControl 对象的变量。

**示例**

``` JavaScript
/*以下示例设置内置的“作者”文档属性，并向活动文档添加新内容控件，然后设置该控件到该内置文档属性的值的映射。*/
function test() {
    Application.ActiveDocument.BuiltInDocumentProperties.Item("Author").Value = "David Jaffe"
    let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlText, Application.ActiveDocument.Paragraphs.Item(1).Range)
    let objMap = objCC.XMLMapping
    blnMap = objMap.SetMapping("/ns1:coreProperties[1]/ns0:creator[1]")
    if (blnMap == false) {
        alert("Unable to map the content control.")
    }
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/ContentControl](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ContentControls 对象

## 简介

ContentControl 对象的集合。内容控件是文档中绑定的、有可能添加标签的区域，它们充当特定类型的内容的容器。单个内容控件可能包含诸如日期、列表或格式化文本段落等内容。

## 方法

| 序号 | 名称                          | 说明                                                                                 |
|:----:|:------------------------------|:-------------------------------------------------------------------------------------|
|  1   | [Add](#ContentControls.Add)   | 将指定类型的新内容控件添加到活动文档中，并返回表示新内容控件的 ContentControl 对象。 |
|  2   | [Item](#ContentControls.Item) | 返回 ContentControl 对象，该对象表示文档中的内容控件集合内的指定内容控件。           |

## 属性

| 序号 | 名称                                        | 说明                                                                         |
|:----:|:--------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#ContentControls.Application) | 返回一个 Application 对象，该对象代表 WPS 应用程序。                         |
|  2   | [Count](#ContentControls.Count)             | 返回 ContentControls 集合中的项目数。只读 Number 类型。                      |
|  3   | [Creator](#ContentControls.Creator)         | 返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 Long 类型。 |
|  4   | [Parent](#ContentControls.Parent)           | 返回一个 Object 类型的值，该值代表指定 ContentControls 对象的父对象。        |

## 成员方法

### ContentControls.Add

将指定类型的新内容控件添加到活动文档中，并返回表示新内容控件的 **ContentControl** 对象。

**语法**

**express.Add ( *Type* , *Range* )**

\*express是一个代表 ContentControls 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型             | 说明                                                                                                        |
|---------|-----------|----------------------|-------------------------------------------------------------------------------------------------------------|
| *Type*  | 可选      | WdContentControlType | 指定要插入活动文档中的内容控件的类型。如果省略它，WPS 将插入 RTF 内容控件。                                 |
| *Range* | 可选      | Variant              | 指定内容控件在活动文档中的放置位置。如果省略它， WPS 会将内容控件放在插入点所在位置，或者替换当前所选内容。 |

**说明**

只能在 RTF 内容控件、生成块库内容控件和组内容控件中嵌套内容控件。如果插入点或当前所选项在不同类型的内容控件的内部，则此方法将引发错误。在这种情况下，可以移动插入点，也可以使用 *Range* 参数来指定文档内的位置。

**示例**

``` JavaScript
/* 添加一个下拉列表控件 */
Application.ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
```

### ContentControls.Item

返回 **ContentControl** 对象，该对象表示文档中的内容控件集合内的指定内容控件。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 ContentControls 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                             |
|---------|-----------|----------|----------------------------------|
| *Index* | 可选      | Variant  | 指定要返回的内容控件的序号位置。 |

## 成员属性

### ContentControls.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 ContentControls 对象的变量。

### ContentControls.Count

返回 **ContentControls** 集合中的项目数。只读 **Number** 类型。

**语法**

**express.Count**

\*express是一个代表 ContentControls 对象的变量。

### ContentControls.Creator

返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ContentControls 对象的变量。

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

### ContentControls.Parent

返回一个 **Object** 类型的值，该值代表指定 **ContentControls** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 ContentControls 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/ContentControls](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Field 对象

## 简介

代表一个域。 Field 对象是 Fields 集合的一个成员。 Fields 集合代表所选内容、范围或文档中的域。

## 说明

使用 Fields ( *Index* ) 可返回一个 Field 对象，其中 *Index* 是索引号。索引号代表域在所选内容、范围或文档中的位置。以下示例显示活动文档中第一个域的域代码和结果。

## 方法

| 序号 | 名称                                | 说明                                           |
|:----:|:------------------------------------|:-----------------------------------------------|
|  1   | [Copy](#Field.Copy)                 | 将指定域复制到剪贴板。                         |
|  2   | [Cut](#Field.Cut)                   | 将指定域从文档中移到剪贴板上。                 |
|  3   | [Delete](#Field.Delete)             | 删除指定的域。                                 |
|  4   | [DoClick](#Field.DoClick)           | 单击指定域。                                   |
|  5   | [Select](#Field.Select)             | 选择指定的域。                                 |
|  6   | [Unlink](#Field.Unlink)             | 将指定域替换为其最新结果。                     |
|  7   | [Update](#Field.Update)             | 更新域的结果。如果该域更新成功，则返回 True 。 |
|  8   | [UpdateSource](#Field.UpdateSource) | 将对 INCLUDETEXT 域所做的更改保存回源文档。    |

## 属性

| 序号 | 名称                              | 说明                                                                                                            |
|:----:|:----------------------------------|:----------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Field.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                                                  |
|  2   | [Code](#Field.Code)               | 返回一个代表域代码的 Range 对象。可读写。                                                                       |
|  3   | [Creator](#Field.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                    |
|  4   | [Data](#Field.Data)               | 返回或设置 ADDIN 域中的数据。可读写 String 类型。                                                               |
|  5   | [Index](#Field.Index)             | 返回一个 Number 类型的值，该值代表项目在集合中的位置。只读。                                                    |
|  6   | [InlineShape](#Field.InlineShape) | 返回一个 InlineShape 对象，该对象可以代表图片、OLE 对象以及作为 INCLUDEPICTURE 或 EMBED 域结果的 ActiveX 控件。 |
|  7   | [Kind](#Field.Kind)               | 返回 Field 对象的链接类型。只读 WdFieldKind 类型。                                                              |
|  8   | [LinkFormat](#Field.LinkFormat)   | 返回一个 LinkFormat 对象，该对象代表指定域的链接选项。只读。                                                    |
|  9   | [Locked](#Field.Locked)           | 如果指定域已被锁定，则该属性值为 True 。可读写 Boolean 类型。                                                   |
|  10  | [Next](#Field.Next)               | 返回集合中的下一个对象。只读。                                                                                  |
|  11  | [OLEFormat](#Field.OLEFormat)     | 返回一个 OLEFormat 对象，该对象代表指定域的 OLE 特征（链接除外）。只读。                                        |
|  12  | [Parent](#Field.Parent)           | 返回一个 Object 类型值，该值代表指定 Field 对象的父对象。                                                       |
|  13  | [Previous](#Field.Previous)       | 返回集合中的上一个对象。只读。                                                                                  |
|  14  | [Result](#Field.Result)           | 返回一个 Range 对象，该对象代表域的结果。可读写。                                                               |
|  15  | [ShowCodes](#Field.ShowCodes)     | 如果显示指定域的域代码而不是域结果，则该属性值为 True 。 Boolean 类型，可读写。                                 |
|  16  | [Type](#Field.Type)               | 返回域的类型。只读 WdFieldType 类型。                                                                           |

## 成员方法

### Field.Copy

将指定域复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 Field 对象的变量。

### Field.Cut

将指定域从文档中移到剪贴板上。

**语法**

**express.Cut ()**

\*express是一个代表 Field 对象的变量。

**示例**

``` JavaScript
/* 本示例删除活动文档的第一个域，并将该域粘贴到插入点 */
function test() {
    if(Application.ActiveDocument.Fields.Count>=1){
        Application.ActiveDocument.Fields.Item(1).Cut()
        Application.Selection.Collapse(wdCollapseEnd)
        Application.Selection.Paste()
    }
}
```

### Field.Delete

删除指定的域。

**语法**

**express.Delete ()**

\*express是一个代表 Field 对象的变量。

### Field.DoClick

单击指定域。

**语法**

**express.DoClick ()**

\*express是一个代表 Field 对象的变量。

**说明**

如果是 GOTOBUTTON 域，该方法将插入点移至指定位置或选定指定的书签。如果是 MACROBUTTON 域，该方法将运行指定的宏。如果是 HYPERLINK 域，该方法将跳转至目标位置。

**示例**

``` JavaScript
/*如果所选内容中的第一个域为 GOTOBUTTON 域，本示例可实现：单击此域（将插入点移至指定位置，或选择指定的书签）。*/
function test()
{
let fldTemp = Selection.Fields.Item(1)
if(fldTemp.Type == wdFieldGoToButton){
    fldTemp.DoClick()
}
}
```

### Field.Select

选择指定的域。

**语法**

**express.Select ()**

\*express是一个代表 Field 对象的变量。

**说明**

使用此方法后，可使用 **Selection** 对象处理所选项目。有关详细信息，请参阅 处理 Selection 对象。

**示例**

``` JavaScript
/* 以下示例更新并选定活动文档中的第一个域 */
function test() {
    if(Application.ActiveDocument.Fields.Count >= 1 ){
        Application.ActiveDocument.Fields.Item(1).Update()
        Application.ActiveDocument.Fields.Item(1).Select()
    }
}
```

### Field.Unlink

将指定域替换为其最新结果。

**语法**

**express.Unlink ()**

\*express是一个代表 Field 对象的变量。

**说明**

如果断开一个域的链接，则当前结果会转换为文字或图形，并且无法再自动更新。注意，某些域（例如，XE（索引项）域和 SEQ（序列）域）不能断开链接。

**示例**

``` JavaScript
/*以下示例断开“Sales.doc”中的第一个域的链接。*/
Documents.Item("Sales.doc").Fields.Item(1).Unlink()
```

### Field.Update

更新域的结果。如果该域更新成功，则返回 **True** 。

**语法**

**express.Update ()**

\*express是一个代表 Field 对象的变量。

**返回值**

Boolean

**示例**

``` JavaScript
/*以下示例更新活动文档中的第一个域。返回 0（零）值表示更新域时没有出错。*/
function test()
{
if (ActiveDocument.Fields.Item(1).Update() == 0){
    MsgBox( "Update Successful")
}
else{
    MsgBox("Field "+ActiveDocument.Fields.Item(1).Update()+" has an error")
}
}
```

### Field.UpdateSource

将对 INCLUDETEXT 域所做的更改保存回源文档。

**语法**

**express.UpdateSource ()**

\*express是一个代表 Field 对象的变量。

**说明**

源文档必须是 WPS 文档格式。

**示例**

``` JavaScript
/*以下示例更新活动文档中的 INCLUDETEXT 域。*/
function test()
{
for (let i=1; i<= ActiveDocument.Fields.Count; i++) {
    if (ActiveDocument.Fields.Item(i).Type == wdFieldIncludeText) { 
        ActiveDocument.Fields.Item(i).UpdateSource()    
    }
}
}
```

## 成员属性

### Field.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Field 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Field.Code

返回一个代表域代码的 **Range** 对象。可读写。

**语法**

**express.Code**

\*express是一个代表 Field 对象的变量。

**说明**

域代码是包含在域字符 ( **{ }** ) 中的所有内容，包括前导和尾部空格。无需更改域结果的显示方式即可访问域代码。

**示例**

``` JavaScript
/* 获取当前文档的域代码 */
Application.ActiveDocument.Fields.Item(1).Code
```

### Field.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Field 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Field.Data

返回或设置 ADDIN 域中的数据。可读写 **String** 类型。

**语法**

**express.Data**

\*express是一个代表 Field 对象的变量。

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

### Field.Index

返回一个 **Number** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Field 对象的变量。

**示例**

``` JavaScript
/* 返回选定域再Fields集合中的位置 */
Application.Selection.Fields.Item(1).Index
```

### Field.InlineShape

返回一个 **InlineShape** 对象，该对象可以代表图片、OLE 对象以及作为 INCLUDEPICTURE 或 EMBED 域结果的 ActiveX 控件。

**语法**

**express.InlineShape**

\*express是一个代表 Field 对象的变量。

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

### Field.Kind

返回 **Field** 对象的链接类型。只读 **WdFieldKind** 类型。

**语法**

**express.Kind**

\*express是一个代表 Field 对象的变量。

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

### Field.LinkFormat

返回一个 **LinkFormat** 对象，该对象代表指定域的链接选项。只读。

**语法**

**express.LinkFormat**

\*express是一个代表 Field 对象的变量。

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

### Field.Locked

如果指定域已被锁定，则该属性值为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.Locked**

\*express是一个代表 Field 对象的变量。

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

### Field.Next

返回集合中的下一个对象。只读。

**语法**

**express.Next**

\*express是一个代表 Field 对象的变量。

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

### Field.OLEFormat

返回一个 **OLEFormat** 对象，该对象代表指定域的 OLE 特征（链接除外）。只读。

**语法**

**express.OLEFormat**

\*express是一个代表 Field 对象的变量。

### Field.Parent

返回一个 **Object** 类型值，该值代表指定 **Field** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Field 对象的变量。

### Field.Previous

返回集合中的上一个对象。只读。

**语法**

**express.Previous**

\*express是一个代表 Field 对象的变量。

**示例**

``` JavaScript
/*以下示例显示活动文档中倒数第二个域的域代码。*/
function test()
{
let aField = ActiveDocument.Fields.Item(ActiveDocument.Fields.Count).Previous
MsgBox( "Field code = " + aField.Code)
}
```

### Field.Result

返回一个 **Range** 对象，该对象代表域的结果。可读写。

**语法**

**express.Result**

\*express是一个代表 Field 对象的变量。

**说明**

无需改变域代码的显示方式即可获得一个域结果。使用 **Text** 属性可返回 **Range** 对象中的文本。

**示例**

``` JavaScript
/* 将选区第一个域结果加粗 */
function test() {
    if(Application.Selection.Fields.Count >= 1) { 
        let myRange = Application.Selection.Fields.Item(1).Result
        myRange.Bold = true
    }
}
```

### Field.ShowCodes

如果显示指定域的域代码而不是域结果，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowCodes**

\*express是一个代表 Field 对象的变量。

**示例**

``` JavaScript
/* 本示例更新当前文档第一个域，并显示其域代码 */
function test() {
    if( Application.ActiveDocument.Fields.Count >= 1) {
        Application.ActiveDocument.Fields.Item(1).Update()
        Application.ActiveDocument.Fields.Item(1).ShowCodes = true
    }
}
```

### Field.Type

返回域的类型。只读 WdFieldType 类型。

**语法**

**express.Type**

\*express是一个代表 Field 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Field](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ListGalleries 对象

## 简介

{\_description}

## 方法

| 序号 | 名称                        | 说明                                |
|:----:|:----------------------------|:------------------------------------|
|  1   | [Item](#ListGalleries.Item) | 返回集合中的单个 ListGallery 对象。 |

## 成员方法

### ListGalleries.Item

返回集合中的单个 **ListGallery** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 ListGalleries 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型          | 说明                 |
|---------|-----------|-------------------|----------------------|
| *Index* | 必选      | WDLISTGALLERYTYPE | 要返回的单个对象索引 |

> WPS 开发文档： [WPS 基础接口/文字 API 参考/ListGalleries](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Paragraph 对象

## 简介

代表所选内容、范围或文档中的一个段落。 Paragraph 对象是 Paragraphs 集合的一个成员。 Paragraphs 集合包含所选内容、范围或文档中的所有段落。

## 说明

使用 Paragraphs ( *Index* ) 可返回一个 Paragraph 对象，其中 *Index* 为索引号。以下示例将活动文档中的第一段右对齐。

``` JavaScript
Application.ActiveDocument.Paragraphs.Item(1).Alignment = wdAlignParagraphRight
```

使用 Add 、 InsertParagraph 、 InsertParagraphAfter 或 InsertParagraphBefore 方法可在文档中添加一个新的空白段落。以下示例在所选内容的第一段前添加一个段落标记。

``` JavaScript
function test() {  let S = Application.Selection.Paragraphs.Item(1).Range    Application.Selection.Paragraphs.Add(S)
} 
```

以下示例同样在所选内容的第一段前添加一个段落标记。

``` JavaScript
Application.Selection.Paragraphs.Item(1).Range.InsertParagraphBefore()
```

## 方法

| 序号 | 名称                                                            | 说明                                                                                     |
|:----:|:----------------------------------------------------------------|:-----------------------------------------------------------------------------------------|
|  1   | [CloseUp](#Paragraph.CloseUp)                                   | 清除指定段落前的任何间距。                                                               |
|  2   | [Indent](#Paragraph.Indent)                                     | 为一个或多个段落增加一个级别的缩进。                                                     |
|  3   | [IndentCharWidth](#Paragraph.IndentCharWidth)                   | 将段落缩进指定的字符数。                                                                 |
|  4   | [IndentFirstLineCharWidth](#Paragraph.IndentFirstLineCharWidth) | 将一个或多个段落的首行缩进指定的字符数。                                                 |
|  5   | [JoinList](#Paragraph.JoinList)                                 | 将列表段落与距指定段落前后最近的列表进行连接。                                           |
|  6   | [ListAdvanceTo](#Paragraph.ListAdvanceTo)                       | 为列表中的段落设置列表级别。                                                             |
|  7   | [ListNumberOriginal](#Paragraph.ListNumberOriginal)             | 返回一个 Integer 类型的值，该值代表段落的原始列表级别。只读。                            |
|  8   | [Next](#Paragraph.Next)                                         | 返回一个 Paragraph 对象，该对象代表下一段。                                              |
|  9   | [OpenOrCloseUp](#Paragraph.OpenOrCloseUp)                       | 切换段前间距。                                                                           |
|  10  | [OpenUp](#Paragraph.OpenUp)                                     | 为指定段落设置 12 磅的段前间距。                                                         |
|  11  | [Outdent](#Paragraph.Outdent)                                   | 为一个或多个段落删除一个级别的缩进。                                                     |
|  12  | [OutlineDemote](#Paragraph.OutlineDemote)                       | 对指定段落应用下一个标题级别样式（从“标题 1”到“标题 8”）。                               |
|  13  | [OutlineDemoteToBody](#Paragraph.OutlineDemoteToBody)           | 通过应用“正文”样式将指定段落降级为正文文本。                                             |
|  14  | [OutlinePromote](#Paragraph.OutlinePromote)                     | 对指定段落应用上一个标题级别样式（从“标题 1”到“标题 8”）。                               |
|  15  | [Previous](#Paragraph.Previous)                                 | 将上一段作为一个 Paragraph 对象返回。                                                    |
|  16  | [Reset](#Paragraph.Reset)                                       | 删除手动段落格式（不使用样式应用的格式）。                                               |
|  17  | [ResetAdvanceTo](#Paragraph.ResetAdvanceTo)                     | 将使用自定义列表级别的段落重置为原始级别设置。                                           |
|  18  | [SelectNumber](#Paragraph.SelectNumber)                         | 选择列表中的编号或项目符号。                                                             |
|  19  | [SeparateList](#Paragraph.SeparateList)                         | 将列表分隔为两个单独的列表。对于编号的列表，新列表将从起始编号（通常为 1）开始重新编号。 |
|  20  | [Space1](#Paragraph.Space1)                                     | 为指定段落设置单倍行距。                                                                 |
|  21  | [Space15](#Paragraph.Space15)                                   | 为指定段落设置 1.5 倍行距。                                                              |
|  22  | [Space2](#Paragraph.Space2)                                     | 为指定段落设置 2 倍行距。                                                                |
|  23  | [TabHangingIndent](#Paragraph.TabHangingIndent)                 | 将悬挂缩进量设置为指定的制表位数。                                                       |
|  24  | [TabIndent](#Paragraph.TabIndent)                               | 将指定段落的左缩进量设置为指定的制表位数。                                               |

## 属性

| 序号 | 名称                                                                          | 说明                                                                                                                                                                                            |
|:----:|:------------------------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddSpaceBetweenFarEastAndAlpha](#Paragraph.AddSpaceBetweenFarEastAndAlpha)   | 如果 WPS 将自动在指定段落的日文和拉丁文文字之间添加空格，则该属性值为 True 。如果仅对于某些指定段落将该属性设置为 True ，则该属性会返回 wdUndefined 。 Long 类型，可读写。                      |
|  2   | [AddSpaceBetweenFarEastAndDigit](#Paragraph.AddSpaceBetweenFarEastAndDigit)   | 如果 WPS 将自动在指定段落的日文文字和数字之间添加空格，则该属性值为 True 。如果仅对于某些指定段落将该属性设置为 True ，则该属性会返回 wdUndefined 。 Long 类型，可读写。                        |
|  3   | [Alignment](#Paragraph.Alignment)                                             | 返回或设置一个 WdParagraphAlignment 常量，该常量代表指定段落的对齐方式，可读写。                                                                                                                |
|  4   | [Application](#Paragraph.Application)                                         | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                                                  |
|  5   | [AutoAdjustRightIndent](#Paragraph.AutoAdjustRightIndent)                     | 如果 WPS 会根据您指定的每行字符数自动调整指定段落的右缩进，则该属性值为 True 。如果只将某些指定段落的 AutoAdjustRightIndent 属性设置为 True ，则该属性会返回 wdUndefined 。 Long 类型，可读写。 |
|  6   | [BaseLineAlignment](#Paragraph.BaseLineAlignment)                             | 返回或设置一个 WdBaselineAlignment 常量，该常量代表行中字体的垂直位置，可读写。                                                                                                                 |
|  7   | [Borders](#Paragraph.Borders)                                                 | 返回一个 Borders 集合，该集合代表指定段落的所有边框。                                                                                                                                           |
|  8   | [CharacterUnitFirstLineIndent](#Paragraph.CharacterUnitFirstLineIndent)       | 返回或设置首行或悬挂缩进的值（以字符为单位）。用正值设置首行缩进，用负值设置悬挂缩进。 Single 类型，可读写。                                                                                    |
|  9   | [CharacterUnitLeftIndent](#Paragraph.CharacterUnitLeftIndent)                 | 该属性返回或设置指定段落的左缩进量（以字符为单位）。 Single 类型，可读写。                                                                                                                      |
|  10  | [CharacterUnitRightIndent](#Paragraph.CharacterUnitRightIndent)               | 该属性返回或设置指定段落的右缩进量（以字符为单位）。 Single 类型，可读写。                                                                                                                      |
|  11  | [Creator](#Paragraph.Creator)                                                 | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                                                    |
|  12  | [DisableLineHeightGrid](#Paragraph.DisableLineHeightGrid)                     | 如果该属性的值为 True ，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 DisableLineHeightGrid 属性设置为 True ，则返回 wdUndefined 。 Long 类型，可读写。    |
|  13  | [DropCap](#Paragraph.DropCap)                                                 | 返回一个 DropCap 对象，该对象代表指定段落中格式为首字下沉的大写字母。只读。                                                                                                                     |
|  14  | [FarEastLineBreakControl](#Paragraph.FarEastLineBreakControl)                 | 如果为 True ，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 FarEastLineBreakControl 属性设定为 True ，则返回 wdUndefined 。 Long 类型，可读写。                     |
|  15  | [FirstLineIndent](#Paragraph.FirstLineIndent)                                 | 返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。 Single 类型，可读写。                                                                    |
|  16  | [Format](#Paragraph.Format)                                                   | 返回或设置一个 ParagraphFormat 对象，该对象代表指定的一个或多个段落的格式。                                                                                                                     |
|  17  | [HalfWidthPunctuationOnTopOfLine](#Paragraph.HalfWidthPunctuationOnTopOfLine) | 如果为 True ，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 True ，则此属性将返回 wdUndefined 。 Long 类型，可读写。                                        |
|  18  | [HangingPunctuation](#Paragraph.HangingPunctuation)                           | 如果为 True ，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 True ，则返回 wdUndefined 。 Long 类型，可读写。                                                             |
|  19  | [Hyphenation](#Paragraph.Hyphenation)                                         | 如果指定的段落进行自动断字，则该属性值为 True 。如果指定的段落不进行自动断字，则该属性值为 False 。可读写 Long 类型。                                                                           |
|  20  | [ID](#Paragraph.ID)                                                           | 在当前文档另存为网页时，返回或设置指定对象的标识标签。可读写 String 类型。                                                                                                                      |
|  21  | [IsStyleSeparator](#Paragraph.IsStyleSeparator)                               | 如果段落包含特殊隐藏段落标记，而利用该标记 WPS 可以显示合并不同段落样式的段落，则该属性值为 True 。 Boolean 类型，只读。                                                                        |
|  22  | [KeepTogether](#Paragraph.KeepTogether)                                       | 在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 True 。可读写 Long 类型。                                                                                           |
|  23  | [KeepWithNext](#Paragraph.KeepWithNext)                                       | 在 WPS 对文档重新分页时，如果指定段落与它的下一段位于同一页上，则该属性值为 True 。可读写 Long 类型。                                                                                           |
|  24  | [LeftIndent](#Paragraph.LeftIndent)                                           | 返回或设置一个 Single 类型的值，该值代表指定段落的左缩进值（以磅为单位）。可读写。                                                                                                              |
|  25  | [LineSpacing](#Paragraph.LineSpacing)                                         | 返回或设置指定段落的行距（以磅为单位）。 Single 类型，可读写。                                                                                                                                  |
|  26  | [LineSpacingRule](#Paragraph.LineSpacingRule)                                 | 返回或设置指定段落的行距。可读写 WdLineSpacing 类型。                                                                                                                                           |
|  27  | [LineUnitAfter](#Paragraph.LineUnitAfter)                                     | 返回或设置指定段落的段后间距（以网格线为单位）。可读写 Single 类型。                                                                                                                            |
|  28  | [LineUnitBefore](#Paragraph.LineUnitBefore)                                   | 返回或设置指定段落的段前间距（以网格线为单位）。可读写 Single 类型。                                                                                                                            |
|  29  | [MirrorIndents](#Paragraph.MirrorIndents)                                     | 返回或设置一个 Long 类型的值，该值代表左缩进和右缩进的宽度是否相同。该属性值可以是 True 、 False 或 wdUndefined 。可读写。                                                                      |
|  30  | [NoLineNumber](#Paragraph.NoLineNumber)                                       | 如果取消指定段落的行号，则该属性值为 True 。可读写 Long 类型。                                                                                                                                  |
|  31  | [OutlineLevel](#Paragraph.OutlineLevel)                                       | 返回或设置指定段落的大纲级别。可读写 WdOutlineLevel 类型。                                                                                                                                      |
|  32  | [PageBreakBefore](#Paragraph.PageBreakBefore)                                 | 如果在指定段落前插入了分页符，则该属性值为 True 。可读写 Long 类型。                                                                                                                            |
|  33  | [Parent](#Paragraph.Parent)                                                   | 返回一个 Object 类型值，该值代表指定 Paragraph 对象的父对象。                                                                                                                                   |
|  34  | [Range](#Paragraph.Range)                                                     | 返回一个 Range 对象，该对象代表指定段落中包含的文档部分。                                                                                                                                       |
|  35  | [ReadingOrder](#Paragraph.ReadingOrder)                                       | 返回或设置指定段落的读取次序而不改变其对齐方式。可读写 WdReadingOrder 类型。                                                                                                                    |
|  36  | [RightIndent](#Paragraph.RightIndent)                                         | 返回或设置指定段落的右缩进（以磅为单位）。可读写 Single 类型。                                                                                                                                  |
|  37  | [Shading](#Paragraph.Shading)                                                 | 返回一个 Shading 对象，该对象引用指定段落的底纹格式。                                                                                                                                           |
|  38  | [SpaceAfter](#Paragraph.SpaceAfter)                                           | 返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 Single 类型。                                                                                                                       |
|  39  | [SpaceAfterAuto](#Paragraph.SpaceAfterAuto)                                   | 如果 WPS 自动设置指定段落的段后间距，则该属性为 True 。可读/写 Long 类型。                                                                                                                      |
|  40  | [SpaceBefore](#Paragraph.SpaceBefore)                                         | 返回或设置指定段落的段前间距（以磅为单位）。可读/写 Single 类型。                                                                                                                               |
|  41  | [SpaceBeforeAuto](#Paragraph.SpaceBeforeAuto)                                 | 如果 WPS 自动设置指定段落的段前间距，则该属性为 True 。可读/写 Long 类型。                                                                                                                      |
|  42  | [Style](#Paragraph.Style)                                                     | 返回或设置指定对象的样式。可读写 Variant 类型。                                                                                                                                                 |
|  43  | [TabStops](#Paragraph.TabStops)                                               | 返回或设置 TabStops 集合，该集合表示指定段落的所有自定义制表位。可读写。                                                                                                                        |
|  44  | [TextboxTightWrap](#Paragraph.TextboxTightWrap)                               | 返回或设置一个 WdTextboxTightWrap 常量，该常量代表文本环绕形状或文本框的紧密程度。可读写。                                                                                                      |
|  45  | [WidowControl](#Paragraph.WidowControl)                                       | 在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 True 。可读写 Long 类型。                                                                                   |
|  46  | [WordWrap](#Paragraph.WordWrap)                                               | 如果 WPS 在指定段落或文本框架的西文单词中间断字换行，则该属性值为 True 。可读写 Long 类型。                                                                                                     |

## 成员方法

### Paragraph.CloseUp

清除指定段落前的任何间距。

**语法**

**express.CloseUp ()**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*以下两个语句是等价的：*/
function test()
{
ActiveDocument.Paragraphs.Item(1).CloseUp()
ActiveDocument.Paragraphs.Item(1).SpaceBefore = 0
}
```

``` JavaScript
/*本示例清除选定内容中首段的段前间距。*/
Selection.Paragraphs.Item(1).CloseUp()
```

### Paragraph.Indent

为一个或多个段落增加一个级别的缩进。

**语法**

**express.Indent ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

此方法等效于单击 **“格式”** 工具栏上的 **“增加缩进量”** 按钮。

**示例**

``` JavaScript
/*以下示例为活动文档中的所有段落增加两个级别的缩进，然后为第一段删除一个级别的缩进。*/
function test()
{
ActiveDocument.Paragraphs.Indent()
ActiveDocument.Paragraphs.Indent()
ActiveDocument.Paragraphs.Item(1).Outdent()
}
```

### Paragraph.IndentCharWidth

将段落缩进指定的字符数。

**语法**

**express.IndentCharWidth ( *Count* )**

\*express是一个代表 Paragraph 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                     |
|---------|-----------|----------|--------------------------|
| *Count* | 必选      | Integer  | 指定段落要缩进的字符数。 |

**说明**

此方法等效于单击 **“格式”** 工具栏上的 **“增加缩进量”** 按钮。

**示例**

``` JavaScript
/*以下示例将活动文档的第一段缩进 10 个字符。*/
ActiveDocument.Paragraphs.Item(1).IndentCharWidth(10)
```

### Paragraph.IndentFirstLineCharWidth

将一个或多个段落的首行缩进指定的字符数。

**语法**

**express.IndentFirstLineCharWidth ( *Count* )**

\*express是一个代表 Paragraph 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                               |
|---------|-----------|----------|------------------------------------|
| *Count* | 必选      | Integer  | 每个指定段落的首行要缩进的字符数。 |

**示例**

``` JavaScript
/*以下示例将活动文档中第一段的首行缩进 10 个字符。*/
Application.ActiveDocument.Paragraphs.Item(1).IndentCharWidth(10)
```

### Paragraph.JoinList

将列表段落与距指定段落前后最近的列表进行连接。

**语法**

**express.JoinList ()**

\*express是一个代表 Paragraph 对象的变量。

### Paragraph.ListAdvanceTo

为列表中的段落设置列表级别。

**语法**

**express.ListAdvanceTo ( *Level1 to Level9* )**

\*express是一个代表 Paragraph 对象的变量。

**参数**

| 名称               | 必选/可选 | 数据类型 | 说明           |
|--------------------|-----------|----------|----------------|
| *Level1 to Level9* | 可选      | Integer  | 指定列表级别。 |

**说明**

WPS 通过在指定段落与上一段落之间以隐藏文字形式插入必要的干预段落，来调整用于对指定段落的每个列表级别进行编号的编号值。

### Paragraph.ListNumberOriginal

返回一个 **Integer** 类型的值，该值代表段落的原始列表级别。只读。

**语法**

**express.ListNumberOriginal ( *Level* )**

\*express是一个代表 Paragraph 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *Level* | 必选      | Integer  | 指定列表级别。 |

### Paragraph.Next

返回一个 **Paragraph** 对象，该对象代表下一段。

**语法**

**express.Next ( *Count* )**

\*express是一个代表 Paragraph 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                         |
|---------|-----------|----------|------------------------------|
| *Count* | 可选      | Variant  | 要前移的段落数。默认值为 1。 |

**返回值**

Paragraph

**示例**

``` JavaScript
/*以下示例在活动文档中的前九段前面分别插入一个数字和一个制表符。*/
function test()
{
for(let n = 0; n <= 8; n++) {
    let myRange = ActiveDocument.Paragraphs.Item(1).Next(n).Range
    myRange.Collapse(wdCollapseStart)
    myRange.InsertAfter(n + 1 + "\t")
}
}
```

### Paragraph.OpenOrCloseUp

切换段前间距。

**语法**

**express.OpenOrCloseUp ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

如果指定段落的段前间距等于 0（零），则此方法将其设置为 12 磅；如果指定段落的段前间距大于 0（零），则此方法将其设置为 0（零）。

**示例**

``` JavaScript
/*以下示例切换活动文档中第一段的格式，或增加 12 磅的段前间距，或不设置段前间距。*/
ActiveDocument.Paragraphs.Item(1).OpenOrCloseUp()
```

### Paragraph.OpenUp

为指定段落设置 12 磅的段前间距。

**语法**

**express.OpenUp ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

您还可以使用 **SpaceBefore** 属性设置段落间距。下列两个语句是等效的：

``` JavaScript
function test()
{
ActiveDocument.Paragraphs.Item(1).OpenUp()
ActiveDocument.Paragraphs.Item(1).SpaceBefore = 12
}
```

**示例**

``` JavaScript
/*以下示例更改活动文档的第二段的格式，为该段落设置 12 磅的段前间距。*/
ActiveDocument.Paragraphs.Item(2).OpenUp()
```

### Paragraph.Outdent

为一个或多个段落删除一个级别的缩进。

**语法**

**express.Outdent ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

此方法等效于单击 **“格式”** 工具栏上的 **“减少缩进量”** 按钮。

**示例**

``` JavaScript
/*以下示例为活动文档中的所有段落增加两个级别的缩进，然后为第一段删除一个级别的缩进。*/
function test()
{
ActiveDocument.Paragraphs.Indent()
ActiveDocument.Paragraphs.Indent()
ActiveDocument.Paragraphs.Item(1).Outdent()
}
```

### Paragraph.OutlineDemote

对指定段落应用下一个标题级别样式（从“标题 1”到“标题 8”）。

**语法**

**express.OutlineDemote ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

如果一个段落的样式为“标题 2”，则此方法将该段落的标题样式降级到“标题 3”。

**示例**

``` JavaScript
/*以下示例将所选内容中的第一段降级。*/
Selection.Paragraphs.Item(1).OutlineDemote()
```

``` JavaScript
/*以下示例将活动文档中的第三段降级。*/
ActiveDocument.Paragraphs.Item(3).OutlineDemote()
```

### Paragraph.OutlineDemoteToBody

通过应用“正文”样式将指定段落降级为正文文本。

**语法**

**express.OutlineDemoteToBody ()**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*以下示例将所选内容中的第一段降级为正文文本。*/
Selection.Paragraphs.Item(1).OutlineDemoteToBody()
```

``` JavaScript
/*以下示例将活动窗口切换至大纲视图，并将所选内容中的第一段降级为正文文本。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdOutlineView
Selection.Paragraphs.Item(1).OutlineDemoteToBody()
}
```

### Paragraph.OutlinePromote

对指定段落应用上一个标题级别样式（从“标题 1”到“标题 8”）。

**语法**

**express.OutlinePromote ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

如果一个段落的样式为“标题 2”，则此方法将该段落的样式升级到“标题 1”。

**示例**

``` JavaScript
/*以下示例将所选内容中的第一段升级。*/
Selection.Paragraphs.Item(1).OutlinePromote()
```

``` JavaScript
/*以下示例将活动窗口切换至大纲视图，并将活动文档中的第一段升级。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdOutlineView
ActiveDocument.Paragraphs.Item(1).OutlinePromote()
}
```

### Paragraph.Previous

将上一段作为一个 **Paragraph** 对象返回。

**语法**

**express.Previous ( *Count* )**

\*express是一个代表 Paragraph 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                         |
|---------|-----------|----------|------------------------------|
| *Count* | 可选      | Variant  | 要后移的段落数。默认值为 1。 |

**返回值**

Paragraph

**示例**

``` JavaScript
/*以下示例选择活动文档中所选内容的上一段。*/
Selection.Previous(wdParagraph, 1).Select()
```

### Paragraph.Reset

删除手动段落格式（不使用样式应用的格式）。

**语法**

**express.Reset ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

如果手动将段落设置为右对齐，而基本样式为其他对齐方式，则用 **Reset** 方法将手动对齐方式改为与基本样式相同。

**示例**

``` JavaScript
/*本示例删除活动文档第二段的手动设定的段落格式。*/
ActiveDocument.Paragraphs.Item(2).Reset()
```

### Paragraph.ResetAdvanceTo

将使用自定义列表级别的段落重置为原始级别设置。

**语法**

**express.ResetAdvanceTo ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

WPS 通过删除指定段落与上一段落之间所有隐藏的干预段落，来调整用于对指定段落的每个列表级别进行编号的编号值。

### Paragraph.SelectNumber

选择列表中的编号或项目符号。

**语法**

**express.SelectNumber ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

如果在不包括列表的段落、选定内容或区域中调用 **SelectNumber** 方法，则显示一条错误消息。

**示例**

``` JavaScript
/*本示例选定列表中的项目符号或编号（该列表包括活动文档中的选定段落），并将项目符号或编号的字号增加至 17 磅。本示例假定选定内容中的第一段被设为项目符号或编号列表。*/
function SelectParaNumber(){
    Selection.Paragraphs.Item(1).SelectNumber()
    Selection.Font.Size = 17
}
```

### Paragraph.SeparateList

将列表分隔为两个单独的列表。对于编号的列表，新列表将从起始编号（通常为 1）开始重新编号。

**语法**

**express.SeparateList ()**

\*express是一个代表 Paragraph 对象的变量。

### Paragraph.Space1

为指定段落设置单倍行距。

**语法**

**express.Space1 ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

精确间距取决于各段内最大字符的字号。

您还可以使用 **LineSpacingRule** 属性设置段落的行距。下列两个语句是等效的：

``` JavaScript
function test()
{
ActiveDocument.Paragraphs.Item(1).Space1()
ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpaceSingle
}
```

**示例**

``` JavaScript
/*以下示例将活动文档中第一段的行距更改为单倍行距。*/
ActiveDocument.Paragraphs.Item(1).Space1()
```

### Paragraph.Space15

为指定段落设置 1.5 倍行距。

**语法**

**express.Space15 ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

精确间距为各段内最大字符的字号加上 6 磅。

您还可以使用 **LineSpacingRule** 属性设置段落的行距。下列两个语句是等效的：

``` JavaScript
function test()
{
ActiveDocument.Paragraphs.Item(1).Space15()
ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpace1pt5
}
```

**示例**

``` JavaScript
/*以下示例将活动文档中第一段的行距更改为 1.5 倍行距。*/
ActiveDocument.Paragraphs.Item(1).Space15()
```

### Paragraph.Space2

为指定段落设置 2 倍行距。

**语法**

**express.Space2 ()**

\*express是一个代表 Paragraph 对象的变量。

**说明**

精确间距为各段内最大字符的字号加上 12 磅。

您还可以使用 **LineSpacingRule** 属性设置段落的行距。下列两个语句是等效的：

``` JavaScript
function test()
{
ActiveDocument.Paragraphs.Item(1).Space2()
ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpaceDouble
}
```

**示例**

``` JavaScript
/*以下示例将所选内容中第一段的行距更改为 2 倍行距。*/
Selection.Paragraphs.Item(1).Space2()
```

### Paragraph.TabHangingIndent

将悬挂缩进量设置为指定的制表位数。

**语法**

**express.TabHangingIndent ( *Count* )**

\*express是一个代表 Paragraph 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                     |
|---------|-----------|----------|--------------------------------------------------------------------------|
| *Count* | 必选      | Integer  | 如果为正，则表示要缩进的制表位数；如果为负，则表示要删除的缩进制表位数。 |

**说明**

如果 *Count* 的值为负数，您还可以使用此方法删除悬挂缩进的制表位。

**示例**

``` JavaScript
/*以下示例将活动文档中的第一段悬挂缩进到第二个制表位。*/
ActiveDocument.Paragraphs.Item(1).TabHangingIndent(2)
```

``` JavaScript
/*以下示例将活动文档中第一段的悬挂缩进量减少一个制表位。*/
ActiveDocument.Paragraphs.Item(1).TabHangingIndent(-1)
```

### Paragraph.TabIndent

将指定段落的左缩进量设置为指定的制表位数。

**语法**

**express.TabIndent ( *Count* )**

\*express是一个代表 Paragraph 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                     |
|---------|-----------|----------|--------------------------------------------------------------------------|
| *Count* | 必选      | Integer  | 如果为正，则表示要缩进的制表位数；如果为负，则表示要删除的缩进制表位数。 |

**说明**

如果 *Count* 的值为负数，您还可以使用此方法删除缩进。

**示例**

``` JavaScript
/*以下示例将活动文档中的第一段缩进到第二个制表位。*/
ActiveDocument.Paragraphs.Item(1).TabIndent(2)
```

``` JavaScript
/*以下示例将活动文档中第一段的缩进量减少一个制表位。*/
ActiveDocument.Paragraphs.Item(1).TabIndent(-1)
```

## 成员属性

### Paragraph.AddSpaceBetweenFarEastAndAlpha

如果 WPS 将自动在指定段落的日文和拉丁文文字之间添加空格，则该属性值为 **True** 。如果仅对于某些指定段落将该属性设置为 **True** ，则该属性会返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndAlpha**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设置 WPS 在活动文档第一段的日文文字和拉丁文字之间自动添加空格。*/
ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndAlpha = true
```

### Paragraph.AddSpaceBetweenFarEastAndDigit

如果 WPS 将自动在指定段落的日文文字和数字之间添加空格，则该属性值为 **True** 。如果仅对于某些指定段落将该属性设置为 **True** ，则该属性会返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndDigit**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例可实现的功能是：设定 WPS 自动在活动文档第一段的中文文字和数字之间添加空格。*/
ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndDigit = true
```

### Paragraph.Alignment

返回或设置一个 **WdParagraphAlignment** 常量，该常量代表指定段落的对齐方式，可读写。

**语法**

**express.Alignment**

\*express是一个代表 Paragraph 对象的变量。

**说明**

对于所选择或安装的不同语言支持（例如美国英语），有些 **WdParagraphAlignment** 常量可能不可用。

**示例**

``` JavaScript
/*本示例将活动文档中的首段设置为右对齐。*/
function test(){
    Application.ActiveDocument.Paragraphs.Item(1).Alignment = wdAlignParagraphRight
}
```

### Paragraph.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Paragraph 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Paragraph.AutoAdjustRightIndent

如果 WPS 会根据您指定的每行字符数自动调整指定段落的右缩进，则该属性值为 **True** 。如果只将某些指定段落的 **AutoAdjustRightIndent** 属性设置为 **True** ，则该属性会返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.AutoAdjustRightIndent**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设定 WPS 根据指定的每行字符数，自动调整所选段落的右缩进。*/
Selection.ParagraphFormat.AutoAdjustRightIndent = true
```

### Paragraph.BaseLineAlignment

返回或设置一个 **WdBaselineAlignment** 常量，该常量代表行中字体的垂直位置，可读写。

**语法**

**express.BaseLineAlignment**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 自动调整活动文档中的基线字体对齐方式。*/
ActiveDocument.BaseLineAlignment = wdBaselineAlignAuto
```

### Paragraph.Borders

返回一个 **Borders** 集合，该集合代表指定段落的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Paragraph 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Paragraph.CharacterUnitFirstLineIndent

返回或设置首行或悬挂缩进的值（以字符为单位）。用正值设置首行缩进，用负值设置悬挂缩进。 **Single** 类型，可读写。

**语法**

**express.CharacterUnitFirstLineIndent**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中第一段的首行缩进设为一个字符。*/
Application.ActiveDocument.Paragraphs.Item(1).CharacterUnitFirstLineIndent = 1

/*本示例将活动文档中第二段的悬挂缩进设为 1.5 个字符。*/
Application.ActiveDocument.Paragraphs.Item(2).CharacterUnitFirstLineIndent = -1.5
```

### Paragraph.CharacterUnitLeftIndent

该属性返回或设置指定段落的左缩进量（以字符为单位）。 **Single** 类型，可读写。

**语法**

**express.CharacterUnitLeftIndent**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中第一段的左缩进设为从左边距缩进一个字符。*/
ActiveDocument.Paragraphs.Item(1).CharacterUnitLeftIndent = 1
```

### Paragraph.CharacterUnitRightIndent

该属性返回或设置指定段落的右缩进量（以字符为单位）。 **Single** 类型，可读写。

**语法**

**express.CharacterUnitRightIndent**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中的所有段落的右缩进设为从右边距缩进一个字符。*/
ActiveDocument.Paragraphs.CharacterUnitRightIndent = 1
```

### Paragraph.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Paragraph 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Paragraph.DisableLineHeightGrid

如果该属性的值为 **True** ，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 **DisableLineHeightGrid** 属性设置为 **True** ，则返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.DisableLineHeightGrid**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设定 WPS 在已指定每页行数的情况下，将选定段落中的字符与行网格对齐。*/
Selection.ParagraphFormat.DisableLineHeightGrid = true
```

### Paragraph.DropCap

返回一个 **DropCap** 对象，该对象代表指定段落中格式为首字下沉的大写字母。只读。

**语法**

**express.DropCap**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例为活动文档首段设置首字下沉。*/
function test()
{
let dropcap =  ActiveDocument.Paragraphs.Item(1).DropCap
dropcap.FontName = "Arial"
dropcap.Position = wdDropNormal
dropcap.LinesToDrop = 3
dropcap.DistanceFromText = InchesToPoints(0.1)
}
```

### Paragraph.FarEastLineBreakControl

如果为 **True** ，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 **FarEastLineBreakControl** 属性设定为 **True** ，则返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.FarEastLineBreakControl**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 将东亚语言文字的换行规则应用于活动文档的第一段。*/
ActiveDocument.Paragraphs.Item(1).FarEastLineBreakControl = true
```

### Paragraph.FirstLineIndent

返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。 **Single** 类型，可读写。

**语法**

**express.FirstLineIndent**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例为活动文档的首段设置 1 英寸的首行缩进。*/
ActiveDocument.Paragraphs.Item(1).FirstLineIndent = InchesToPoints(1)
```

``` JavaScript
/*本示例为活动文档的第二段设置 0.5 英寸的悬挂缩进。InchesToPoints 方法用来将英寸转化为磅值。*/
ActiveDocument.Paragraphs.Item(2).FirstLineIndent = InchesToPoints(-0.5)
```

### Paragraph.Format

返回或设置一个 **ParagraphFormat** 对象，该对象代表指定的一个或多个段落的格式。

**语法**

**express.Format**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例返回活动文档中第一段的格式，然后将该格式应用于文档的选定内容。*/
function test()
{
let paraFormat = ActiveDocument.Paragraphs.Item(1).Format.Duplicate
Selection.Paragraphs.Format = paraFormat
}
```

### Paragraph.HalfWidthPunctuationOnTopOfLine

如果为 **True** ，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 **True** ，则此属性将返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.HalfWidthPunctuationOnTopOfLine**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 将活动文档第一段的行首标点符号改为半角字符。*/
ActiveDocument.Paragraphs.Item(1).HalfWidthPunctuationOnTopOfLine = true
```

### Paragraph.HangingPunctuation

如果为 **True** ，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 **True** ，则返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.HangingPunctuation**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：允许活动文档第一段的标点可以溢出边界。*/
ActiveDocument.Paragraphs.Item(1).HangingPunctuation = true
```

### Paragraph.Hyphenation

如果指定的段落进行自动断字，则该属性值为 **True** 。如果指定的段落不进行自动断字，则该属性值为 **False** 。可读写 **Long** 类型。

**语法**

**express.Hyphenation**

\*express是一个代表 Paragraph 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例关闭活动文档中第一段的自动断字功能。*/
ActiveDocument.Paragraphs.Item(1).Hyphenation = false
```

### Paragraph.ID

在当前文档另存为网页时，返回或设置指定对象的标识标签。可读写 **String** 类型。

**语法**

**express.ID**

\*express是一个代表 Paragraph 对象的变量。

### Paragraph.IsStyleSeparator

如果段落包含特殊隐藏段落标记，而利用该标记 WPS 可以显示合并不同段落样式的段落，则该属性值为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.IsStyleSeparator**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例为包含内置“Normal”样式的样式分隔符的所有段落设置格式。*/
function StyleSep(){
    let pghDoc = ActiveDocument.Paragraphs
    for(let i = 1; i <= pghDoc.Count; i++ ){
        if(pghDoc.Item(i).IsStyleSeparator == true){
            pghDoc.Item(i).Range.Select()
            Selection.Style = "Normal"
        }
    }
}
```

### Paragraph.KeepTogether

在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.KeepTogether**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*以下示例设置活动文档中第一段的格式，以使在 WPS 对文档重新分页时同一段中的各行都位于同一页上。*/
ActiveDocument.Paragraphs.Item(1).KeepTogether = true
```

### Paragraph.KeepWithNext

在 WPS 对文档重新分页时，如果指定段落与它的下一段位于同一页上，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.KeepWithNext**

\*express是一个代表 Paragraph 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例将活动文档中的第三段至第六段保持在同一页上。*/
function test()
{
for(let i = 3; i <= 5; i++){
    ActiveDocument.Paragraphs.Item(i).KeepWithNext = true
}
}
```

### Paragraph.LeftIndent

返回或设置一个 **Single** 类型的值，该值代表指定段落的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*以下示例为活动文档中的第一段设置 1 英寸的左缩进。InchesToPoints 方法用来将英寸转换为磅。*/
ActiveDocument.Paragraphs.Item(1).LeftIndent = InchesToPoints(1)
 
```

### Paragraph.LineSpacing

返回或设置指定段落的行距（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.LineSpacing**

\*express是一个代表 Paragraph 对象的变量。

**说明**

使用 **LinesToPoints** 方法可将行数转换为相应的值（以磅为单位）。例如， ` LinesToPoints(2) ` 返回值 24。

可以在 **LineSpacingRule** 属性已设置为下列值之后对 **LineSpacing** 属性进行设置：

- **wdLineSpaceAtLeast** 行距可以大于或等于指定的 **LineSpacing** 值，但绝不能小于该值。
- **wdLineSpaceExactly** 指定的 **LineSpacing** 行距值绝不能更改，即使段落中使用了较大的字体。
- **wdLineSpaceMultiple** 必须指定 **LineSpacing** 属性值，以磅为单位。

**示例**

``` JavaScript
/*以下示例将活动文档中第一段的行距设置为总是至少 12 磅。*/
function test()
{
let pgh = ActiveDocument.Paragraphs.Item(1)
pgh.LineSpacingRule = wdLineSpaceAtLeast
pgh.LineSpacing = 12
}
```

### Paragraph.LineSpacingRule

返回或设置指定段落的行距。可读写 **WdLineSpacing** 类型。

**语法**

**express.LineSpacingRule**

\*express是一个代表 Paragraph 对象的变量。

**说明**

使用 **wdLineSpaceSingle** 、 **wdLineSpace1pt5** 或 **wdLineSpaceDouble** 可将行距设置为这些值之一。要将行距设置为固定磅值或多倍行距，还必须设置 **LineSpacing** 属性。

**示例**

``` JavaScript
/*以下示例为活动文档的第一段设置 2 倍行距。*/
ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpaceDouble
```

``` JavaScript
/*以下示例返回所选内容的第一段所用的行距标准。*/
lrule = Selection.Paragraphs.Item(1).LineSpacingRule
```

### Paragraph.LineUnitAfter

返回或设置指定段落的段后间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitAfter**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中第一段的段后间距设置为 1 个网格线。*/
ActiveDocument.Paragraphs.Item(1).LineUnitAfter = 1
```

### Paragraph.LineUnitBefore

返回或设置指定段落的段前间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitBefore**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档第二段的段前间距设置为 1 个网格线。*/
ActiveDocument.Paragraphs.Item(2).LineUnitBefore = 1
```

### Paragraph.MirrorIndents

返回或设置一个 **Long** 类型的值，该值代表左缩进和右缩进的宽度是否相同。该属性值可以是 **True** 、 **False** 或 **wdUndefined** 。可读写。

**语法**

**express.MirrorIndents**

\*express是一个代表 Paragraph 对象的变量。

### Paragraph.NoLineNumber

如果取消指定段落的行号，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.NoLineNumber**

\*express是一个代表 Paragraph 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。使用 **PageSetup** 对象的 **LineNumbering** 属性可设置行号。

**示例**

``` JavaScript
/*以下示例为活动文档设置行号。起始编号设置为 1，并且在文档中的所有节都连续编号。然后取消第二段的行号。*/
function test()
{
let lnbg = ActiveDocument.PageSetup.LineNumbering
lnbg.Active = true
lnbg.StartingNumber = 1
lnbg.CountBy = 1
lnbg.RestartMode = wdRestartContinuous
ActiveDocument.Paragraphs.Item(2).NoLineNumber = true
}
```

### Paragraph.OutlineLevel

返回或设置指定段落的大纲级别。可读写 **WdOutlineLevel** 类型。

**语法**

**express.OutlineLevel**

\*express是一个代表 Paragraph 对象的变量。

**说明**

如果段落应用了一种标题样式（从“标题 1”到“标题 9”），则大纲级别与该标题样式相同，并且不能更改。大纲级别仅在大纲视图或文档结构图窗格中可见。

**示例**

``` JavaScript
/*以下示例返回活动文档中第一段的大纲级别。*/
temp = ActiveDocument.Paragraphs.Item(1).OutlineLevel
```

``` JavaScript
/*以下示例设置活动文档每一段的大纲级别。首先将正文样式应用于所有段落。Mod 操作符用来确定文档中的后继段落采用何种大纲级别（1、2 或 3），然后将视图切换至大纲视图。*/
function test()
{
let myParas = ActiveDocument.Paragraphs
myParas.Style = wdStyleNormal
for(let x = 1; x <= myParas.Count; x++){
    if(x % 3 == 1){
    myParas.Item(x).OutlineLevel = wdOutlineLevel1
    }
    else if(x % 3 == 2){
    myParas.Item(x).OutlineLevel = wdOutlineLevel2
    }
    else{
        myParas.Item(x).OutlineLevel = wdOutlineLevel3
    }
}
ActiveDocument.ActiveWindow.View.Type = wdOutlineView
}
```

### Paragraph.PageBreakBefore

如果在指定段落前插入了分页符，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.PageBreakBefore**

\*express是一个代表 Paragraph 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例在所选内容的第一段前插入分页符。*/
Selection.Paragraphs.Item(1).PageBreakBefore = true
```

### Paragraph.Parent

返回一个 **Object** 类型值，该值代表指定 **Paragraph** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Paragraph 对象的变量。

### Paragraph.Range

返回一个 **Range** 对象，该对象代表指定段落中包含的文档部分。

**语法**

**express.Range**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*以下示例为活动文档中的第一段应用“标题 1”样式。*/
ActiveDocument.Paragraphs.Item(1).Range.Style = wdStyleHeading1
```

### Paragraph.ReadingOrder

返回或设置指定段落的读取次序而不改变其对齐方式。可读写 **WdReadingOrder** 类型。

**语法**

**express.ReadingOrder**

\*express是一个代表 Paragraph 对象的变量。

**说明**

使用 **Selection** 对象的 **LtrPara** 、 **LtrRun** 、 **RtlPara** 和 **RtlRun** 方法可更改段落的对齐方式和读取次序。

**示例**

``` JavaScript
/*以下示例将第一段的读取次序设置为从右向左。*/
ActiveDocument.Paragraphs.Item(1).ReadingOrder = wdReadingOrderRtl
```

### Paragraph.RightIndent

返回或设置指定段落的右缩进（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.RightIndent**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中第一段的右缩进设置为距右边距 1 英寸。InchesToPoints 方法用来将英寸转换为磅。*/
ActiveDocument.Paragraphs.Item(1).RightIndent = InchesToPoints(1)
```

### Paragraph.Shading

返回一个 **Shading** 对象，该对象引用指定段落的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*以下示例对所选内容的第一段应用黄色底纹。*/
function test()
{
let shad = Selection.Paragraphs.Item(1).Shading
shad.Texture = wdTexture12Pt5Percent
shad.BackgroundPatternColorIndex = wdYellow
shad.ForegroundPatternColorIndex = wdBlack
}
```

### Paragraph.SpaceAfter

返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceAfter**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档第一段的段后间距设置为 12 磅。*/
ActiveDocument.Paragraphs.Item(1).SpaceAfter = 12
```

### Paragraph.SpaceAfterAuto

如果 WPS 自动设置指定段落的段后间距，则该属性为 **True** 。可读/写 **Long** 类型。

**语法**

**express.SpaceAfterAuto**

\*express是一个代表 Paragraph 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceAfterAuto** 属性设置为 **True** ，则返回 **wdUndefined** 。该属性可被设置为 **True** 或 **False** 。

如果将 **SpaceAfterAuto** 设置为 **True** ，则 **SpaceAfter** 属性将被忽略。

**示例**

``` JavaScript
/*本示例显示关于活动文档的第一段的 SpaceAfterAuto 设置的报告。*/
function test()
{
let saa = ActiveDocument.Paragraphs.Item(1).SpaceAfterAuto
switch(saa) {
    case 1:
        MsgBox ( "Spacing after paragraphs is handled automatically for all paragraphs.")
        break
    case 0:
        MsgBox ( "Spacing after paragraphs is handled manually for all paragraphs.")
        break
    case undefined:
        MsgBox ( "Spacing after paragraphs is handled automatically for some paragraphs, manually for some paragraphs.")
        break
}
}
```

### Paragraph.SpaceBefore

返回或设置指定段落的段前间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceBefore**

\*express是一个代表 Paragraph 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中第二段的段前间距设置为 12 磅。*/
ActiveDocument.Paragraphs.Item(2).SpaceBefore = 12
```

### Paragraph.SpaceBeforeAuto

如果 WPS 自动设置指定段落的段前间距，则该属性为 **True** 。可读/写 **Long** 类型。

**语法**

**express.SpaceBeforeAuto**

\*express是一个代表 Paragraph 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceBeforeAuto** 属性设置为 **True** ，则该属性返回 **wdUndefined** 。该属性可被设置为 **True** 或 **False** 。

如果将 **SpaceBeforeAuto** 设置为 **True** ，则 **SpaceBefore** 属性将被忽略。

**示例**

``` JavaScript
/*本示例显示活动文档中第一段的 SpaceBeforeAuto 设置的报告。*/
function test()
{
let saa = ActiveDocument.Paragraphs.Item(1).SpaceBeforeAuto
switch(saa) {
    case 1:
        MsgBox ( "Spacing before paragraphs is handled automatically for all paragraphs.")
        break
    case 0:
        MsgBox ( "Spacing before paragraphs is handled manually for all paragraphs.")
        break
    case undefined:
        MsgBox ( "Spacing before paragraphs is handled automatically for some paragraphs, manually for some paragraphs.")
        break
}
}
```

### Paragraph.Style

返回或设置指定对象的样式。可读写 **Variant** 类型。

**语法**

**express.Style**

\*express是一个代表 Paragraph 对象的变量。

**说明**

要设置该属性，请指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个代表样式的对象。

**示例**

``` JavaScript
/*以下示例显示活动文档中各段的样式。*/
function test()
{
let para = ActiveDocument.Paragraphs
for(let i = 1; i <= para.Count; i++) {
    MsgBox (para.Item(i).Style)
}
}
```

``` JavaScript
/*以下示例将活动文档中所有段落交替设置为“标题 3”和“正文”样式。*/
function test()
{
let pghc = ActiveDocument.Paragraphs
for(let i = 1; i<= pghc.Count; i++){
    if(i % 2 ==0){
        pghc.Item(i).Style = wdStyleNormal
    }
    else{
        pghc.Item(i).Style = wdStyleHeading3
    }
}
}
```

### Paragraph.TabStops

返回或设置 **TabStops** 集合，该集合表示指定段落的所有自定义制表位。可读写。

**语法**

**express.TabStops**

\*express是一个代表 Paragraph 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例将文档中各段的制表位设置为与第一段的制表位一致。*/
function test()
{
let para1Tabs = ActiveDocument.Paragraphs.Item(1).TabStops
ActiveDocument.Paragraphs.TabStops = para1Tabs
}
```

### Paragraph.TextboxTightWrap

返回或设置一个 **WdTextboxTightWrap** 常量，该常量代表文本环绕形状或文本框的紧密程度。可读写。

**语法**

**express.TextboxTightWrap**

\*express是一个代表 Paragraph 对象的变量。

### Paragraph.WidowControl

在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.WidowControl**

\*express是一个代表 Paragraph 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例设置活动文档中第一段的格式，以便段落中的首行或末行可以单独出现在页面顶端或底端。*/
ActiveDocument.Paragraphs.Item(1).WidowControl = false
```

### Paragraph.WordWrap

如果 WPS 在指定段落或文本框架的西文单词中间断字换行，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.WordWrap**

\*express是一个代表 Paragraph 对象的变量。

**说明**

如果仅将某些指定段落或文本框架的此属性设置为 **True** ，则此属性返回 **wdUndefined** 。根据您所选择或安装的语言支持（例如，美国英语），这种用法可能不可用。

**示例**

``` JavaScript
/*以下示例设置 WPS 在活动文档的第一段中的西文单词中间断字换行。*/
ActiveDocument.Paragraphs.Item(1).WordWrap = false
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Paragraph](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Revision 对象

## 简介

代表由修订标记所标记的修改。 Revision 对象是 Revisions 集合的一个成员。 Revisions 集合包含某范围或文档中的所有修订标记。

## 方法

| 序号 | 名称                       | 说明                                             |
|:----:|:---------------------------|:-------------------------------------------------|
|  1   | [Accept](#Revision.Accept) | 接受指定修订、删除修订标记并将更改合并到文档中。 |
|  2   | [Reject](#Revision.Reject) | 拒绝接受指定的修订。将删除修订标记，不改变原文   |

## 属性

| 序号 | 名称                                             | 说明                                                                                      |
|:----:|:-------------------------------------------------|:------------------------------------------------------------------------------------------|
|  1   | [Application](#Revision.Application)             | 返回一个代表 WPS 应用程序的 Application 对象。                                            |
|  2   | [Author](#Revision.Author)                       | 返回或设置完成指定修订的用户姓名                                                          |
|  3   | [Cells](#Revision.Cells)                         | 返回一个 Cells 集合，该集合代表已经用修订标记进行标记的表格单元格。只读。                 |
|  4   | [Creator](#Revision.Creator)                     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。              |
|  5   | [Date](#Revision.Date)                           | 修订的日期和时间                                                                          |
|  6   | [FormatDescription](#Revision.FormatDescription) | 返回一个 String 类型的值，该值代表对修订中所作格式更改的说明。只读。                      |
|  7   | [Index](#Revision.Index)                         | 返回一个 Long 类型的值，该值代表项目在集合中的位置。只读。                                |
|  8   | [MovedRange](#Revision.MovedRange)               | 返回一个 Range 对象，该对象代表在带修订的文档中从一个位置移至另一个位置的文本区域。只读。 |
|  9   | [Parent](#Revision.Parent)                       | 返回一个 Object 类型值，该值代表指定 Revision 对象的父对象。                              |
|  10  | [Range](#Revision.Range)                         | 返回一个 Range 对象，该对象代表修订标记内包含的文档部分。                                 |
|  11  | [Style](#Revision.Style)                         | 返回一个 Style 对象，该对象代表与一个修订标记相关联的样式。                               |
|  12  | [Type](#Revision.Type)                           | 返回修订的类型。只读 [WdRevisionType]() 类型。                                            |

## 成员方法

### Revision.Accept

接受指定修订、删除修订标记并将更改合并到文档中。

**语法**

**express.Accept ()**

\*express是一个代表 Revision 对象的变量。

**示例**

``` JavaScript
//接受当前文档的第一个修订
function test()
{
    let rev;
    if (Application.ActiveDocument.Revisions.Count >= 1)
        rev = Application.ActiveDocument.Revisions.Item(1)
    rev.Accept();
}
```

### Revision.Reject

拒绝接受指定的修订。将删除修订标记，不改变原文

**语法**

**express.Reject ()**

\*express是一个代表 Revision 对象的变量。

**示例**

``` JavaScript
//拒绝当前文档的第二个修订。
function test()
{
    let rev;
    if (Application.ActiveDocument.Revisions.Count >= 2)
        rev = Application.ActiveDocument.Revisions.Item(2)
    rev.Reject();
}
```

## 成员属性

### Revision.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Revision 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Revision.Author

返回或设置完成指定修订的用户姓名

**语法**

**express.Author**

\*express是一个代表 Revision 对象的变量。

**示例**

``` JavaScript
//显示所选部分中的第一节的第一处修订的作者姓名
function test()
{
    let rngSection = Application.Selection.Sections.Item(1).Range
    alert("Revisions made by " + rngSection.Revisions.Item(1).Author)
}
```

### Revision.Cells

返回一个 **Cells** 集合，该集合代表已经用修订标记进行标记的表格单元格。只读。

**语法**

**express.Cells**

\*express是一个代表 Revision 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Revision.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Revision 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Revision.Date

修订的日期和时间

**语法**

**express.Date**

\*express是一个代表 Revision 对象的变量。

### Revision.FormatDescription

返回一个 **String** 类型的值，该值代表对修订中所作格式更改的说明。只读。

**语法**

**express.FormatDescription**

\*express是一个代表 Revision 对象的变量。

**示例**

``` JavaScript
/*本示例显示文档中修订的所有格式更改的说明。*/
function FmtChanges() {
    for(let revFmtRev = 1; revFmtRev <= ActiveDocument.Revisions.Count; revFmtRev++) {
        if(!ActiveDocument.Revisions.Item(revFmtRev).FormatDescription) {
            MsgBox("Format changes made : " + ActiveDocument.Revisions.Item(revFmtRev).FormatDescription)
        }
    }
}
```

### Revision.Index

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Revision 对象的变量。

### Revision.MovedRange

返回一个 **Range** 对象，该对象代表在带修订的文档中从一个位置移至另一个位置的文本区域。只读。

**语法**

**express.MovedRange**

\*express是一个代表 Revision 对象的变量。

### Revision.Parent

返回一个 **Object** 类型值，该值代表指定 **Revision** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Revision 对象的变量。

### Revision.Range

返回一个 **Range** 对象，该对象代表修订标记内包含的文档部分。

**语法**

**express.Range**

\*express是一个代表 Revision 对象的变量。

### Revision.Style

返回一个 **Style** 对象，该对象代表与一个修订标记相关联的样式。

**语法**

**express.Style**

\*express是一个代表 Revision 对象的变量。

### Revision.Type

返回修订的类型。只读 **[WdRevisionType]()** 类型。

**语法**

**express.Type**

\*express是一个代表 Revision 对象的变量。

**示例**

``` JavaScript
//获取当前激活文档的第一个修订的类型
Application.ActiveDocument.Revisions.Item(1).Type
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Revision](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# StoryRanges 对象

## 简介

Range 对象的集合，这些对象代表文档中的 文章 。

## 方法

| 序号 | 名称                      | 说明                                                |
|:----:|:--------------------------|:----------------------------------------------------|
|  1   | [Item](#StoryRanges.Item) | 将区域或选定内容的单个文字部分作为 Range 对象返回。 |

## 属性

| 序号 | 名称                                    | 说明                                                                         |
|:----:|:----------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#StoryRanges.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#StoryRanges.Count)             | 返回一个 Long 类型的值，该值代表集合中存储区域的数量。只读。                 |
|  3   | [Creator](#StoryRanges.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#StoryRanges.Parent)           | 返回一个 Object 类型值，该值代表指定 StoryRanges 对象的父对象。              |

## 成员方法

### StoryRanges.Item

将区域或选定内容的单个文字部分作为 **Range** 对象返回。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 StoryRanges 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型    | 说明                 |
|---------|-----------|-------------|----------------------|
| *Index* | 必选      | WdStoryType | 指定的文字部分类型。 |

**返回值**

Range

## 成员属性

### StoryRanges.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 StoryRanges 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### StoryRanges.Count

返回一个 **Long** 类型的值，该值代表集合中存储区域的数量。只读。

**语法**

**express.Count**

\*express是一个代表 StoryRanges 对象的变量。

### StoryRanges.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 StoryRanges 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### StoryRanges.Parent

返回一个 **Object** 类型值，该值代表指定 **StoryRanges** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 StoryRanges 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/StoryRanges](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TaskPane 对象

## 属性

| 序号 | 名称                         | 说明                                                          |
|:----:|:-----------------------------|:--------------------------------------------------------------|
|  1   | [Visible](#TaskPane.Visible) | 如果指定对象可见，则该属性值为 True 。 Boolean 类型，可读写。 |

## 成员属性

### TaskPane.Visible

如果指定对象可见，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Visible**

\*express是一个代表 TaskPane 对象的变量。

**说明**

对于任何对象，如果其 **Visible** 属性值为 **False** ，则某些方法和属性可能无法使用。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/TaskPane](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# XMLNamespaces 对象

## 简介

XMLNamespace 对象的集合，该集合代表架构库中架构的整个集合。

## 方法

| 序号 | 名称                                              | 说明                                                                                   |
|:----:|:--------------------------------------------------|:---------------------------------------------------------------------------------------|
|  1   | [Add](#XMLNamespaces.Add)                         | 返回一个 XMLNamespace 对象，该对象代表添加到架构库中并在 WPS 中对用户可用的架构。      |
|  2   | [InstallManifest](#XMLNamespaces.InstallManifest) | 将指定的 XML 扩展包安装在用户计算机上，令一个或更多用户能够使用 XML 智能文档解决方案。 |
|  3   | [Item](#XMLNamespaces.Item)                       | 返回集合中的单个 XMLNamespace 对象。                                                   |

## 属性

| 序号 | 名称                                      | 说明                                                                         |
|:----:|:------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#XMLNamespaces.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#XMLNamespaces.Count)             | 返回一个 Long 类型的值，该值代表集合中 XML 命名空间的数量。只读。            |
|  3   | [Creator](#XMLNamespaces.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#XMLNamespaces.Parent)           | 返回一个 Object 类型值，该值代表指定 XMLNamespaces 对象的父对象。            |

## 成员方法

### XMLNamespaces.Add

返回一个 **XMLNamespace** 对象，该对象代表添加到架构库中并在 WPS 中对用户可用的架构。

**语法**

**express.Add ( *Path* , *NamespaceURI* , *Alias* , *InstallForAllUsers* )**

\*express是一个代表 XMLNamespaces 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                                                                                            |
|----------------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------|
| *Path*               | 必选      | String   | 架构的路径和文件名。可以为本地文件路径、网络路径或 Internet 地址。                                              |
| *NamespaceURI*       | 可选      | String   | 架构中指定的命名空间“统一资源标识符”。NamespaceURI 参数区分大小写，因此必须严格按照架构中所指定的那样进行拼写。 |
| *Alias*              | 可选      | String   | 在“模板和加载项”对话框中的“架构”选项卡上显示的架构名称。                                                        |
| *InstallForAllUsers* | 可选      | Boolean  | 如果登录到计算机上的所有用户都能访问和使用新架构，则为 True。默认值为 False。                                   |

**返回值**

XMLNamespace

**示例**

``` JavaScript
/*以下示例将指定的架构添加到架构库，然后将其附加到活动文档中。本示例假设在指定的路径中已有一个名为 simplesample.xsd 的架构。*/
function AddSchema() {
    let objSchema = Application.XMLNamespaces.Add("C:\\schemas\\simplesample.xsd")       
    objSchema.AttachToDocument(ActiveDocument)
}
```

### XMLNamespaces.InstallManifest

将指定的 XML 扩展包安装在用户计算机上，令一个或更多用户能够使用 XML 智能文档解决方案。

**语法**

**express.InstallManifest ( *Path* , *InstallForAllUsers* )**

\*express是一个代表 XMLNamespaces 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                                                                                                          |
|----------------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------|
| *Path*               | 必选      | String   | XML 扩展包的路径和文件名。                                                                                                    |
| *InstallForAllUsers* | 可选      | Boolean  | 如果为 True，则安装 XML 扩展包并允许该计算机上的所有用户使用。如果为 False，则 XML 扩展包仅限于当前用户使用。默认值为 False。 |

**说明**

下面的代码示例将 SimpleSample 智能文档解决方案安装到用户计算机上并使其仅用于当前用户。

| 注释                                                                                                                |
|---------------------------------------------------------------------------------------------------------------------|
| SimpleSample 架构包含在 Smart Document Software Development Kit (SDK) 中。有关详细信息，请参阅 Smart Document SDK。 |

**示例**

``` JavaScript
Application.XMLNamespaces.InstallManifest("http://smartdocuments/simplesample/manifest.xml")
```

### XMLNamespaces.Item

返回集合中的单个 **XMLNamespace** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 XMLNamespaces 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

**返回值**

XMLNamespace

## 成员属性

### XMLNamespaces.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 XMLNamespaces 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### XMLNamespaces.Count

返回一个 **Long** 类型的值，该值代表集合中 XML 命名空间的数量。只读。

**语法**

**express.Count**

\*express是一个代表 XMLNamespaces 对象的变量。

### XMLNamespaces.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 XMLNamespaces 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### XMLNamespaces.Parent

返回一个 **Object** 类型值，该值代表指定 **XMLNamespaces** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 XMLNamespaces 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/XMLNamespaces](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FilterEffect 对象

## 简介

代表动画动作的筛选效果。

## 说明

使用 AnimationBehavior 对象的 FilterEffect 属性返回 FilterEffect 对象。使用 FilterEffect 对象的 Reveal 、 SubType 和 Type 属性可以更改筛选效果。

## 属性

| 序号 | 名称                                     | 说明                                                                                      |
|:----:|:-----------------------------------------|:------------------------------------------------------------------------------------------|
|  1   | [Application](#FilterEffect.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。                                   |
|  2   | [Parent](#FilterEffect.Parent)           | 返回指定对象的父对象。                                                                    |
|  3   | [Reveal](#FilterEffect.Reveal)           | 设置或返回一个 MsoTriState 常量，该常量决定嵌入对象的显示方式。可读/写。                  |
|  4   | [Subtype](#FilterEffect.Subtype)         | 设置或返回一个 MsoAnimFilterEffectSubtype 常量，该常量用于决定筛选效果的子类型。可读/写。 |
|  5   | [Type](#FilterEffect.Type)               | 返回或设置一个 MsoAnimType 常量，该常量代表动画的类型。可读/写。                          |

## 成员属性

### FilterEffect.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 FilterEffect 对象的变量。

**说明**

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function AddAndSave(pptPre) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

``` JavaScript
/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### FilterEffect.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 FilterEffect 对象的变量。

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

### FilterEffect.Reveal

设置或返回一个 **MsoTriState** 常量，该常量决定嵌入对象的显示方式。可读/写。

**语法**

**express.Reveal**

\*express是一个代表 FilterEffect 对象的变量。

**说明**

当筛选效果类型为 **msoAnimFilterEffectTypeWipe** 时，如果将 **Reveal** 属性的值设置为 **msoTrue** ，则会显示形状。设置为 **msoFalse** 值将使对象消失。换言之，如果将筛选设为擦除，当 Reveal 属性为真，将得到向内擦除的效果，当 Reveal 属性为假，就会得到向外擦除的效果。

| MsoTriState 可以是下列 MsoTriState 常量之一。      |
|----------------------------------------------------|
| msoCTrue 不适用于此属性。                          |
| msoFalse 批注、修订和个人信息保留在演示文稿中。    |
| msoTriStateMixed 不适用于此属性。                  |
| msoTriStateToggle 不适用于此属性。                 |
| msoTrue 在保存演示文稿时删除批注、修订和个人信息。 |

**示例**

以下示例在当前演示文稿的第一张幻灯片中添加一个形状并设置一个筛选效果动画动作。

``` JavaScript
function ChangeFilterEffect() {
    let sldFirst = Application.ActivePresentation.Slides.Item(1)
    let shpHeart = sldFirst.Shapes.AddShape(msoShapeHeart,100,100,100,100)
    let effNew = sldFirst.TimeLine.MainSequence.AddEffect(shpHeart,msoAnimEffectChangeFillColor,undefined,msoAnimTriggerAfterPrevious)
    let bhvEffect = effNew.Behaviors.Add(msoAnimTypeFilter)
                                         
    bhvEffect.FilterEffect.Type = msoAnimFilterEffectTypeWipe
    bhvEffect.FilterEffect.Subtype = msoAnimFilterEffectSubtypeUp
    bhvEffect.FilterEffect.Reveal = msoTrue
}
```

### FilterEffect.Subtype

设置或返回一个 **MsoAnimFilterEffectSubtype** 常量，该常量用于决定筛选效果的子类型。可读/写。

**语法**

**express.Subtype**

\*express是一个代表 FilterEffect 对象的变量。

**示例**

``` JavaScript
/*以下示例在当前演示文稿的第一张幻灯片中添加一个形状并设置一个筛选效果动画动作。*/
function ChangeFilterEffect() {
    let sldFirst = Application.ActivePresentation.Slides.Item(1)
    let shpHeart = sldFirst.Shapes.AddShape(msoShapeHeart,100,100,100,100)
    let effNew = sldFirst.TimeLine.MainSequence.AddEffect(shpHeart,msoAnimEffectChangeFillColor,undefined,msoAnimTriggerAfterPrevious)
    let bhvEffect = effNew.Behaviors.Add(msoAnimTypeFilter)
                                         
    bhvEffect.FilterEffect.Type = msoAnimFilterEffectTypeWipe
    bhvEffect.FilterEffect.Subtype = msoAnimFilterEffectSubtypeUp
    bhvEffect.FilterEffect.Reveal = msoTrue
}
```

### FilterEffect.Type

返回或设置一个 **MsoAnimType** 常量，该常量代表动画的类型。可读/写。

**语法**

**express.Type**

\*express是一个代表 FilterEffect 对象的变量。

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

> WPS 开发文档： [WPS 基础接口/演示 API 参考/FilterEffect](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Hyperlink 对象

## 简介

代表与非占位符形状或文本相关联的超链接。

## 说明

可以使用超链接跳转到一个 Internet 或 Intranet 网站、另一个文件或当前演示文稿中的一张幻灯片上。 Hyperlink 对象是 Hyperlinks 集合的成员。 Hyperlinks 集合包含幻灯片或母版中的所有超链接。

## 方法

| 序号 | 名称                                              | 说明                                                                                                                                                                      |
|:----:|:--------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddToFavorites](#Hyperlink.AddToFavorites)       | 在 Program Files 文件夹的 Favorites 文件夹中，添加一个快捷方式，以代表指定演示文稿的当前选择内容（对于 Presentation 对象）或指定超链接的目标文档（对于 Hyperlink 对象）。 |
|  2   | [CreateNewDocument](#Hyperlink.CreateNewDocument) | 新建一个与指定超链接相关联的 Web 演示文稿。                                                                                                                               |
|  3   | [Delete](#Hyperlink.Delete)                       | 删除指定的对象。                                                                                                                                                          |
|  4   | [Follow](#Hyperlink.Follow)                       | 在新的 Web 浏览器窗口中显示与指定超链接相关联的 HTML 文档。                                                                                                               |

## 属性

| 序号 | 名称                                      | 说明                                                                                                                                |
|:----:|:------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Address](#Hyperlink.Address)             | 返回或设置目标文档的 Internet 地址 (URL)。可读/写 String 类型。                                                                     |
|  2   | [Application](#Hyperlink.Application)     | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                             |
|  3   | [EmailSubject](#Hyperlink.EmailSubject)   | 返回或设置超链接主题行的文本字符串。该主题行附加到超链接的 Internet 地址 (URL) 中。可读/写 String 类型。                            |
|  4   | [Parent](#Hyperlink.Parent)               | 返回指定对象的父对象。                                                                                                              |
|  5   | [ScreenTip](#Hyperlink.ScreenTip)         | 返回或设置超链接的屏幕提示文字。可读/写 String 类型。                                                                               |
|  6   | [ShowAndReturn](#Hyperlink.ShowAndReturn) | 决定 WPP是否返回初始的幻灯片放映以及返回的条件。可读/写。 MsoTriState 类型。                                                        |
|  7   | [SubAddress](#Hyperlink.SubAddress)       | 返回或设置与指定超链接相关联的文档中的位置，例如 WPS 文档中的书签、ET 工作表中的区域或WPP 演示文稿中的幻灯片。可读/写 String 类型。 |
|  8   | [TextToDisplay](#Hyperlink.TextToDisplay) | 返回或设置不是与图形相关联的超链接的显示文字。可读/写 String 类型。                                                                 |
|  9   | [Type](#Hyperlink.Type)                   | 返回一个 MsoHyperlinkType 常量，该常量代表超链接的类型。只读。                                                                      |

## 成员方法

### Hyperlink.AddToFavorites

在 Program Files 文件夹的 Favorites 文件夹中，添加一个快捷方式，以代表指定演示文稿的当前选择内容（对于 **Presentation** 对象）或指定超链接的目标文档（对于 **Hyperlink** 对象）。

**语法**

**express.AddToFavorites ()**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

如果文档具有显示名称，则快捷方式的名称是该显示名称；否则，快捷方式名称由 HLINK.DLL 决定。

**示例**

本示例在 Windows 程序文件夹的 Favorites 中添加一个指向当前演示文稿的超链接。

``` JavaScript
Application.ActivePresentation.AddToFavorites()
```

### Hyperlink.CreateNewDocument

新建一个与指定超链接相关联的 Web 演示文稿。

**语法**

**express.CreateNewDocument ( *FileName* , *EditNow* , *Overwrite* )**

\*express是一个代表 Hyperlink 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型    | 说明                                         |
|-------------|-----------|-------------|----------------------------------------------|
| *FileName*  | 必选      | String      | 文档的路径和文件名。                         |
| *EditNow*   | 必选      | MsoTriState | 确定是否在关联的编辑器中立即打开文档。       |
| *Overwrite* | 必选      | MsoTriState | 确定是否覆盖同一文件夹中现有的任何同名文件。 |

**返回值**

Nothing

**示例**

本示例将创建一个新的 Web 演示文稿，该演示文稿将与第一张幻灯片的第一个超链接相关联。新演示文稿的名为“Brittany.ppt”，同时它还会覆盖“HTMLPres”文件夹中的同名文件。此演示文稿还会立即装入到 WPP 中进行编辑。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Hyperlinks.Item(1).CreateNewDocument("C:\\HTMLPres\\Brittany.ppt", msoTrue, msoTrue)
```

### Hyperlink.Delete

删除指定的对象。

**语法**

**express.Delete ()**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

### Hyperlink.Follow

在新的 Web 浏览器窗口中显示与指定超链接相关联的 HTML 文档。

**语法**

**express.Follow ()**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

本示例在一个新的 Web 浏览器实例中加载与第一张幻灯片的第一个超链接相关联的文档。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Hyperlinks.Item(1).Follow()
```

## 成员属性

### Hyperlink.Address

返回或设置目标文档的 Internet 地址 (URL)。可读/写 **String** 类型。

**语法**

**express.Address**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

``` JavaScript
/*本示例浏览第一张幻灯片中的所有形状，以查找指向 网站的 URL。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　for(let i=1;i<=myDocument.Hyperlinks.count;i++) {
　　　　    if(myDocument.Hyperlinks.Item(i).Address == "http://www.wps.cn/") {
 　　　　       alert("You have a link to the  Home Page")
　　　　    }
　　　　}
}
```

### Hyperlink.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Hyperlink 对象的变量。

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

### Hyperlink.EmailSubject

返回或设置超链接主题行的文本字符串。该主题行附加到超链接的 Internet 地址 (URL) 中。可读/写 **String** 类型。

**语法**

**express.EmailSubject**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

此属性通常与电子邮件超链接一起使用。其值优先于在同一 **Hyperlink** 对象的 **Address** 属性中指定的任意电子邮件主题。

**示例**

本示例设置活动演示文稿中第一张幻灯片上第一个超链接的电子邮件主题行。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Hyperlinks.Item(1).EmailSubject 
```

### Hyperlink.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Hyperlink 对象的变量。

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

### Hyperlink.ScreenTip

返回或设置超链接的屏幕提示文字。可读/写 **String** 类型。

**语法**

**express.ScreenTip**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

将演示文稿保存为 HTML、在 Web 浏览器中查看 Web 演示文稿以及将鼠标指针停留在超链接之上时，会显示屏幕提示文字。有些浏览器可能不支持屏幕提示。

**示例**

本示例为第一个超链接设置屏幕提示文字。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Hyperlinks.Item(1).ScreenTip = "Go to the WPS home page"
```

### Hyperlink.ShowAndReturn

决定 WPP是否返回初始的幻灯片放映以及返回的条件。可读/写。 **MsoTriState** 类型。

**语法**

**express.ShowAndReturn**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                                                                                          |
|----------------------------------------------------------------------------------------------------------------------------------------|
| msoCTrue                                                                                                                               |
| msoFalse 默认值。WPP不会从未激活的自定义幻灯片放映返回初始幻灯片放映。                                                                 |
| msoTriStateMixed                                                                                                                       |
| msoTriStateToggle                                                                                                                      |
| msoTrue：WPP从取消激活的自定义幻灯片放映返回初始幻灯片放映，该自定义幻灯片放映通过初始演示文稿的 Hyperlink 或 ActionSetting 对象激活。 |

**示例**

``` JavaScript
/*以下示例设置当前演示文稿中第一张幻灯片上第五个形状的鼠标单击动作为：显示名为“techtalk”的自定义幻灯片放映。当该自定义幻灯片放映结束后，将自动返回演示文稿在鼠标单击前的最初状态。*/
function test(){
　　　　let as = Application.ActivePresentation.Slides.Item(1).Shapes.Item(5).ActionSettings.Item(ppMouseClick)
　　　　as.Action = ppActionNamedSlideShow
　　　　as.SlideShowName = "techtalk"
　　　　as.ShowandReturn = msoTrue
}
```

### Hyperlink.SubAddress

返回或设置与指定超链接相关联的文档中的位置，例如 WPS 文档中的书签、ET 工作表中的区域或WPP 演示文稿中的幻灯片。可读/写 **String** 类型。

**语法**

**express.SubAddress**

\*express是一个代表 Hyperlink 对象的变量。

**示例**

``` JavaScript
/*本示例设置在幻灯片放映过程中，单击当前演示文稿第一张幻灯片第一个形状时，跳转到 Latest Figures.ppt 中名为“Last Quarter”的幻灯片。*/
function test(){
　　　　Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).ActionSettings.Item(ppMouseClick).Action = ppActionHyperlink
　　　　let h = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).ActionSettings.Item(ppMouseClick).Hyperlink
　　　　h.Address = "c:\\sales\\latest figures.ppt"
　　　　h.SubAddress = "last quarter"
}
```

``` JavaScript
/*本示例设置在幻灯片放映过程中，单击当前演示文稿第一张幻灯片第一个形状时，跳转到 Latest.xls 中 A1:B10 单元格区域。*/
function test(){
　　　　ActivePresentation.Slides.Item(1).Shapes.Item(1).ActionSettings.Item(ppMouseClick).Action = ppActionHyperlink
　　　　let h = ActivePresentation.Slides.Item(1).Shapes.Item(1).ActionSettings.Item(ppMouseClick).Hyperlink
　　　　h.Address = "c:\\sales\\latest.xls"
　　　　h.SubAddress = "A1:B10"
}
```

### Hyperlink.TextToDisplay

返回或设置不是与图形相关联的超链接的显示文字。可读/写 **String** 类型。

**语法**

**express.TextToDisplay**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

如果对不是与文字范围相关联的超链接使用此属性，则会产生运行时错误。可以使用以下类似代码检测给定的超链接（此处由 ` myHyperlink ` 代表）是否与某个文字范围相关联。

` If TypeName(myHyperlink.Parent.Parent) = "TextRange" Then strTRtest = "True" End If `

**示例**

``` JavaScript
/*本示例创建一个与第一张幻灯片第二个形状内的文字相关联的超链接。然后将显示文字设置为“WPS Home Page”，并将超链接地址设置为相应的 URL*/
function test(){
　　　　let t = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).TextFrame.TextRange.ActionSettings.Item(ppMouseClick)
　　　　t.Action = ppActionHyperlink
　　　　t.Hyperlink.TextToDisplay = "WPS Home Page"
　　　　t.Hyperlink.Address = "http://www.wps.cn"
}
```

### Hyperlink.Type

返回一个 **MsoHyperlinkType** 常量，该常量代表超链接的类型。只读。

**语法**

**express.Type**

\*express是一个代表 Hyperlink 对象的变量。

**说明**

| MsoHyperlinkType 可以是下列 MsoHyperlinkType 常量之一。 |
|---------------------------------------------------------|
| msoHyperlinkInlineShape 仅用于 WPS。                    |
| msoHyperlinkRange                                       |
| msoHyperlinkShape                                       |

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Hyperlink](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Hyperlinks 对象

## 简介

幻灯片或母版上所有 Hyperlink 对象的集合。

## 方法

| 序号 | 名称                     | 说明                       |
|:----:|:-------------------------|:---------------------------|
|  1   | [Item](#Hyperlinks.Item) | 从指定集合中返回单个对象。 |

## 属性

| 序号 | 名称                                   | 说明                                                    |
|:----:|:---------------------------------------|:--------------------------------------------------------|
|  1   | [Application](#Hyperlinks.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#Hyperlinks.Count)             | 返回指定集合中的对象数目                                |
|  3   | [Parent](#Hyperlinks.Parent)           | 返回指定对象的父对象。                                  |

## 成员方法

### Hyperlinks.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Hyperlinks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明 |
|---------|-----------|----------|------|
| *Index* | 必选      | Long     | Long |

**返回值**

Hyperlink

## 成员属性

### Hyperlinks.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Hyperlinks 对象的变量。

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

### Hyperlinks.Count

返回指定集合中的对象数目

**语法**

**express.Count**

\*express是一个代表 Hyperlinks 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除窗口1外的所有窗口。*/
function test(){
　　　　for(let i=2; i<=Application.Windows.Count; i++) {
　　　　    Application.Windows.Item(i).Close()
　　　　}
}
```

### Hyperlinks.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Hyperlinks 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
　　　　let tf = myShapes.AddShape(msoShapeOval,50,50,300,150).TextFrame
　　　　tf.TextRange.Text = "Test text"
　　　　tf.Parent.Rotation = 45
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Hyperlinks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# NamedSlideShow 对象

## 简介

代表自定义幻灯片放映，该幻灯片放映是演示文稿中幻灯片的一个命名子集。

## 说明

NamedSlideShow 对象是 NamedSlideShows 集合的成员。 NamedSlideShows 集合包含演示文稿中的所有命名幻灯片放映。

## 方法

| 序号 | 名称                             | 说明             |
|:----:|:---------------------------------|:-----------------|
|  1   | [Delete](#NamedSlideShow.Delete) | 删除指定的对象。 |

## 属性

| 序号 | 名称                                       | 说明                                                                                                              |
|:----:|:-------------------------------------------|:------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#NamedSlideShow.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                           |
|  2   | [Count](#NamedSlideShow.Count)             | 返回指定集合中的对象数目。                                                                                        |
|  3   | [Name](#NamedSlideShow.Name)               | 不能使用该属性设置自定义幻灯片放映的名称。使用 Add 方法可以新名称再定义一个自定义幻灯片放映。 String 类型，只读。 |
|  4   | [Parent](#NamedSlideShow.Parent)           | 返回指定对象的父对象。                                                                                            |
|  5   | [SlideIDs](#NamedSlideShow.SlideIDs)       | 返回指定的命名幻灯片放映的幻灯片 ID 数组。只读 Variant 类型。                                                     |

## 成员方法

### NamedSlideShow.Delete

删除指定的对象。

**语法**

**express.Delete ()**

\*express是一个代表 NamedSlideShow 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

## 成员属性

### NamedSlideShow.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 NamedSlideShow 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres){
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

### NamedSlideShow.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 NamedSlideShow 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口。*/
function test(){
　　　　let win = Application.Windows
   　　　　 for(let i = 2; i <= win.Count; i++){
    　　　　    win.Item(i).Close()
 　　　　   }
}
```

### NamedSlideShow.Name

不能使用该属性设置自定义幻灯片放映的名称。使用 **Add** 方法可以新名称再定义一个自定义幻灯片放映。 **String** 类型，只读。

**语法**

**express.Name**

\*express是一个代表 NamedSlideShow 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。例如，如果某个形状的 **Name** 属性值为“Rectangle 2”，则 ` .Shapes("Rectangle 2") ` 将返回对该形状的引用。

### NamedSlideShow.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 NamedSlideShow 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
　　　　let sha = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
　　　　sha.TextRange.Text = "Test text"
　　　　sha.Parent.Rotation = 45
}
```

### NamedSlideShow.SlideIDs

返回指定的命名幻灯片放映的幻灯片 ID 数组。只读 **Variant** 类型。

**语法**

**express.SlideIDs**

\*express是一个代表 NamedSlideShow 对象的变量。

**示例**

``` JavaScript
/*本示例将活动窗口中的当前幻灯片添加到名为“Marketing Short Version”的自定义幻灯片放映中。请注意，要保存修改后的自定义幻灯片放映版本，必须删除原始自定义放映，再使用相同名称重新添加。另请注意，要调整包含于 Variant 变量中的数组大小，必须在调整该数组大小前明确声明该变量。*/
function test(){
//NOTE - The following code line is NOT optional.
//Can't redim array without this
                                    
　　　　customShowName = "Marketing Short Version"
　　　　let customShowToExpand = Application.ActivePresentation.SlideShowSettings.NamedSlideShows.Item(customShowName)
　　　　let slideToAddID = Application.ActiveWindow.View.Slide.SlideID
　　　　let customShowSlideIDs = customShowToExpand.SlideIDs
　　　　let numSlides = UBound(customShowSlideIDs)
                                    
　　　　customShowSlideIDs(numSlides + 1) = slideToAddID
　　　　customShowToExpand.Delete()
　　　　ActivePresentation.SlideShowSettings.NamedSlideShows.Add(customShowName, customShowSlideIDs)
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/NamedSlideShow](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomXMLPrefixMappings 对象

## 简介

代表 CustomXMLPrefixMapping 对象的集合。

## 说明

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

以下示例通过向 CustomXMLPrefixMapping 集合添加命名空间和前缀，创建一个 CustomXMLPrefixMapping 对象。

``` JavaScript
Application.ActivePresentation.CustomXMLParts.Item(1).NamespaceManager.AddNamespace("xs", "urn:invoice:namespace")
```

## 方法

| 序号 | 名称                                                        | 说明                                                    |
|:----:|:------------------------------------------------------------|:--------------------------------------------------------|
|  1   | [AddNamespace](#CustomXMLPrefixMappings.AddNamespace)       | 允许您添加要在查询项目时使用的自定义命名空间/前缀映射。 |
|  2   | [LookupNamespace](#CustomXMLPrefixMappings.LookupNamespace) | 允许您获取对应于指定前缀的命名空间。                    |
|  3   | [LookupPrefix](#CustomXMLPrefixMappings.LookupPrefix)       | 允许您获取对应于指定命名空间的前缀。                    |

## 属性

| 序号 | 名称                                                | 说明                                                                                |
|:----:|:----------------------------------------------------|:------------------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLPrefixMappings.Application) | 获取一个代表 CustomXMLPrefixMappings 对象的容器应用程序的 Application 对象。只读。  |
|  2   | [Count](#CustomXMLPrefixMappings.Count)             | 获取一个 Long 类型的值，指示 CustomXMLPrefixMappings 集合中的项数。只读。           |
|  3   | [Creator](#CustomXMLPrefixMappings.Creator)         | 获取一个 32 位整数，指示创建 CustomXMLPrefixMappings 对象时所使用的应用程序。只读。 |
|  4   | [Item](#CustomXMLPrefixMappings.Item)               | 获取 CustomXMLPrefixMappings 集合中的一个 CustomXMLPrefixMapping 对象。只读。       |
|  5   | [Parent](#CustomXMLPrefixMappings.Parent)           | 获取 CustomXMLPrefixMappings 对象的 Parent 对象。只读。                             |

## 成员方法

### CustomXMLPrefixMappings.AddNamespace

允许您添加要在查询项目时使用的自定义命名空间/前缀映射。

**语法**

**express.AddNamespace ( *Prefix* , *NamespaceURI* )**

\*express是一个代表 CustomXMLPrefixMappings 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                               |
|----------------|-----------|----------|------------------------------------|
| *Prefix*       | 必选      | String   | 包含要添加到前缀映射列表的前缀。   |
| *NamespaceURI* | 必选      | String   | 包含要分配给新添加前缀的命名空间。 |

**说明**

如果前缀在命名空间管理器中已存在，则此方法将覆盖该前缀的含义，除非前缀是数据存储（ **IXMLDataStore** 接口）在内部添加或使用的（在这种情况下它将返回错误）。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMappings.LookupNamespace

允许您获取对应于指定前缀的命名空间。

**语法**

**express.LookupNamespace ( *Prefix* )**

\*express是一个代表 CustomXMLPrefixMappings 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                       |
|----------|-----------|----------|----------------------------|
| *Prefix* | 必选      | String   | 包含前缀映射列表中的前缀。 |

**说明**

如果未将命名空间分配给请求的前缀，则该方法返回空字符串 ("")。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMappings.LookupPrefix

允许您获取对应于指定命名空间的前缀。

**语法**

**express.LookupPrefix ( *NamespaceURI* )**

\*express是一个代表 CustomXMLPrefixMappings 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明               |
|----------------|-----------|----------|--------------------|
| *NamespaceURI* | 必选      | String   | 包含命名空间 URI。 |

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

## 成员属性

### CustomXMLPrefixMappings.Application

获取一个代表 **CustomXMLPrefixMappings** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLPrefixMappings 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMappings.Count

获取一个 **Long** 类型的值，指示 **CustomXMLPrefixMappings** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 CustomXMLPrefixMappings 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMappings.Creator

获取一个 32 位整数，指示创建 **CustomXMLPrefixMappings** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLPrefixMappings 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMappings.Item

获取 **CustomXMLPrefixMappings** 集合中的一个 **CustomXMLPrefixMapping** 对象。只读。

**语法**

**express.Item**

\*express是一个代表 CustomXMLPrefixMappings 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLPrefixMappings.Parent

获取 **CustomXMLPrefixMappings** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLPrefixMappings 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLPrefixMappings](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomXMLSchema 对象

## 简介

代表 CustomXMLSchemaCollection 集合中的一个架构。

## 说明

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

## 方法

| 序号 | 名称                              | 说明                                    |
|:----:|:----------------------------------|:----------------------------------------|
|  1   | [Delete](#CustomXMLSchema.Delete) | 从 CustomXMLSchema 集合中删除指定架构。 |
|  2   | [Reload](#CustomXMLSchema.Reload) | 从文件重新加载架构。                    |

## 属性

| 序号 | 名称                                          | 说明                                                                        |
|:----:|:----------------------------------------------|:----------------------------------------------------------------------------|
|  1   | [Application](#CustomXMLSchema.Application)   | 获取一个代表 CustomXMLSchema 对象的容器应用程序的 Application 对象。只读。  |
|  2   | [Creator](#CustomXMLSchema.Creator)           | 获取一个 32 位整数，指示创建 CustomXMLSchema 对象时所使用的应用程序。只读。 |
|  3   | [Location](#CustomXMLSchema.Location)         | 获取一个 String 类型的值，代表计算机上架构的位置。只读。                    |
|  4   | [NamespaceURI](#CustomXMLSchema.NamespaceURI) | 获取 CustomXMLSchema 对象的命名空间的唯一地址标识符。只读。                 |
|  5   | [Parent](#CustomXMLSchema.Parent)             | 获取 CustomXMLSchema 对象的 Parent 对象。只读。                             |

## 成员方法

### CustomXMLSchema.Delete

从 **CustomXMLSchema** 集合中删除指定架构。

**语法**

**express.Delete ()**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

如果尝试对已经过验证或附加到数据流的集合中的架构进行此操作，则不会执行操作并会显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例将架构添加到集合，然后删除该架构。
function test()
{
     try 
    {
        // Adds a schema to the collection.
    　　Application.ActivePresentation.CustomXMLParts.Item(1).SchemaCollection.Add("urn:invoice:namespace") 
        // Deletes the schema.
        Application.ActivePresentation.CustomXMLParts.Item(1).SchemaCollection.Item(1).Delete()
    }            
    // Exception handling. Show the message and resume.
    catch(exception) 
    {
        alert(exception.Description)
    }
}
```

### CustomXMLSchema.Reload

从文件重新加载架构。

**语法**

**express.Reload ()**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

通常，此方法用于更新架构的位置或确定架构是否仍然有效。它还用于重新加载频繁更改的架构。如果尝试对集合中的已经过验证或绑定到数据流的架构执行此操作，则不执行操作，并显示一条错误消息。

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

``` JavaScript
//以下示例指定了架构的位置，然后重新加载它。
function test()
{
    try{
        let objCustomXMLSchema = Application.ActivePresentation.CustomXMLParts.Item(1).SchemaCollection.Item(1)
        // Reload the schema.
        bjCustomXMLSchema.Reload()
    }catch(exception){
    alert(exception.Description)
    }
}
```

## 成员属性

### CustomXMLSchema.Application

获取一个代表 **CustomXMLSchema** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLSchema.Creator

获取一个 32 位整数，指示创建 **CustomXMLSchema** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLSchema.Location

获取一个 **String** 类型的值，代表计算机上架构的位置。只读。

**语法**

**express.Location**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLSchema.NamespaceURI

获取 **CustomXMLSchema** 对象的命名空间的唯一地址标识符。只读。

**语法**

**express.NamespaceURI**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

### CustomXMLSchema.Parent

获取 **CustomXMLSchema** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomXMLSchema 对象的变量。

**说明**

| 注释                                                                                                                                                                 |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CustomXMLSchema](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PickerField 对象

## 简介

代表 PickerResult 对象中的子项的字段定义。每个 PickerField 对象都代表选取器对话框的一个列定义。

## 属性

| 序号 | 名称                                    | 说明                                                                        |
|:----:|:----------------------------------------|:----------------------------------------------------------------------------|
|  1   | [Application](#PickerField.Application) | 获取一个 Application 对象，该对象代表 PickerField 对象的容器应用程序。只读  |
|  2   | [Creator](#PickerField.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 PickerField 对象的应用程序。只读 |
|  3   | [IsHidden](#PickerField.IsHidden)       | 指定选取器字段是否为隐藏。只读                                              |
|  4   | [Name](#PickerField.Name)               | 检索选取器字段的名称。只读                                                  |
|  5   | [Type](#PickerField.Type)               | 选取器字段的类型。只读                                                      |

## 成员属性

### PickerField.Application

获取一个 **Application** 对象，该对象代表 **PickerField** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 PickerField 对象的变量。

### PickerField.Creator

获取一个 32 位整数，该整数指示在其中创建了 **PickerField** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 PickerField 对象的变量。

### PickerField.IsHidden

指定选取器字段是否为隐藏。只读

**语法**

**express.IsHidden**

\*express是一个代表 PickerField 对象的变量。

### PickerField.Name

检索选取器字段的名称。只读

**语法**

**express.Name**

\*express是一个代表 PickerField 对象的变量。

### PickerField.Type

选取器字段的类型。只读

**语法**

**express.Type**

\*express是一个代表 PickerField 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/PickerField](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PickerProperty 对象

## 简介

代表一个对象，以便传递自定义属性。

## 说明

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

## 属性

| 序号 | 名称                                       | 说明                                                                           |
|:----:|:-------------------------------------------|:-------------------------------------------------------------------------------|
|  1   | [Application](#PickerProperty.Application) | 获取一个 Application 对象，该对象代表 PickerProperty 对象的容器应用程序。只读  |
|  2   | [Creator](#PickerProperty.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 PickerProperty 对象的应用程序。只读 |
|  3   | [Id](#PickerProperty.Id)                   | 检索关联的 PickerProperty 对象的唯一 Id。只读                                  |
|  4   | [Type](#PickerProperty.Type)               | 检索选取器属性的类型。只读                                                     |
|  5   | [Value](#PickerProperty.Value)             | 检索选取器属性的值。只读                                                       |

## 成员属性

### PickerProperty.Application

获取一个 **Application** 对象，该对象代表 **PickerProperty** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 PickerProperty 对象的变量。

### PickerProperty.Creator

获取一个 32 位整数，该整数指示在其中创建了 **PickerProperty** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 PickerProperty 对象的变量。

### PickerProperty.Id

检索关联的 **PickerProperty** 对象的唯一 Id。只读

**语法**

**express.Id**

\*express是一个代表 PickerProperty 对象的变量。

### PickerProperty.Type

检索选取器属性的类型。只读

**语法**

**express.Type**

\*express是一个代表 PickerProperty 对象的变量。

### PickerProperty.Value

检索选取器属性的值。只读

**语法**

**express.Value**

\*express是一个代表 PickerProperty 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/PickerProperty](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PickerResults 对象

## 简介

PickerResult 对象的集合。

## 说明

每个 PickerResult 对象

``` JavaScript
//下面的代码显示选取器对话框，获取结果，然后枚举这些结果。
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
    
    // Enumerate the results.
    for(let index = 1; index <= objPickerResults.Count-1; index++){
        Debug.Print(objPickerResults.Item(index).Id)
        Debug.Print(objPickerResults.Item(index).DisplayName)
        Debug.Print(objPickerResults.Item(index).Type)
        Debug.Print(objPickerResults.Item(index).SIPId)
    }
}
```

都代表一个解析的或选定的项目数据。

## 方法

| 序号 | 名称                      | 说明                                              |
|:----:|:--------------------------|:--------------------------------------------------|
|  1   | [Add](#PickerResults.Add) | 将 PickerResult 对象添加到 PickerResults 集合中。 |

## 属性

| 序号 | 名称                                      | 说明                                                                          |
|:----:|:------------------------------------------|:------------------------------------------------------------------------------|
|  1   | [Application](#PickerResults.Application) | 获取一个 Application 对象，该对象代表 PickerResults 对象的容器应用程序。只读  |
|  2   | [Count](#PickerResults.Count)             | 检索包含在 PickerResults 集合中的 PickerResult 对象数的计数。只读             |
|  3   | [Creator](#PickerResults.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 PickerResults 对象的应用程序。只读 |
|  4   | [Item](#PickerResults.Item)               | 检索位于指定索引处的 PickerResult 对象。只读                                  |

## 成员方法

### PickerResults.Add

将 **PickerResult** 对象添加到 **PickerResults** 集合中。

**语法**

**express.Add ( *Id* , *DisplayName* , *Type* , *SIPId* , *ItemData* , *SubItems* )**

\*express是一个代表 PickerResults 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                                                                               |
|---------------|-----------|----------|------------------------------------------------------------------------------------|
| *Id*          | 必选      | String   | 代表 PickerResult 的标识符。                                                       |
| *DisplayName* | 必选      | String   | 代表 PickerResult 的显示名称。                                                     |
| *Type*        | 必选      | String   | 代表 PickerResult 的类型。                                                         |
| *SIPId*       | 可选      | Boolean  | 当前不支持。SIPId 是 Office Communication Server 的标识符。它仅用于人员选取方案。  |
| *ItemData*    | 可选      | Variant  | 非显示项绑定数据                                                                   |
| *SubItems*    | 可选      | Variant  | 显示用于显示或不用于显示的 PickerResult 的域数据。它用于在选取器对话框中传递列值。 |

**返回值**

PickerResult

**示例**

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

## 成员属性

### PickerResults.Application

获取一个 **Application** 对象，该对象代表 **PickerResults** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 PickerResults 对象的变量。

### PickerResults.Count

检索包含在 **PickerResults** 集合中的 **PickerResult** 对象数的计数。只读

**语法**

**express.Count**

\*express是一个代表 PickerResults 对象的变量。

**示例**

``` JavaScript
//下面的代码显示选取器对话框用户界面，获取结果，然后枚举这些结果。
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
    
    // Enumerate the results.
    for(let index = 1; index <= objPickerResults.Count-1; index++){
        Debug.Print(objPickerResults.Item(index).Id)
        Debug.Print(objPickerResults.Item(index).DisplayName)
        Debug.Print(objPickerResults.Item(index).Type)
        Debug.Print(objPickerResults.Item(index).SIPId)
    }
}
```

### PickerResults.Creator

获取一个 32 位整数，该整数指示在其中创建了 **PickerResults** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 PickerResults 对象的变量。

### PickerResults.Item

检索位于指定索引处的 **PickerResult** 对象。只读

**语法**

**express.Item**

\*express是一个代表 PickerResults 对象的变量。

**示例**

``` JavaScript
//下面的代码显示选取器对话框用户界面，获取结果，然后枚举这些结果。
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
    
    // Enumerate the results.
    for(let index = 1; index <= objPickerResults.Count-1; index++){
        Debug.Print(objPickerResults.Item(index).Id)
        Debug.Print(objPickerResults.Item(index).DisplayName)
        Debug.Print(objPickerResults.Item(index).Type)
        Debug.Print(objPickerResults.Item(index).SIPId)
    }
}
```

> WPS 开发文档： [WPS 基础接口/通用 API 参考/PickerResults](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Signature 对象

## 简介

代表附加到文档的数字签名。 Signature 对象包含在 Document 对象的 SignatureSet 集合中。

## 说明

可以使用 Add 方法向 SignatureSet 集合中添加一个 Signature 对象，并使用 Item 方法返回现有的成员。要从 SignatureSet 集合中删除一个 Signature 对象，请使用 Signature 对象的 Delete 方法。

``` JavaScript
/*下面的示例提示用户选择要用于签署 WPS 中活动文档的数字签名。要使用此示例，请在 WPS 中打开一个文档，并向此函数传递与“数字证书”对话框中数字证书的“颁发者”和“颁发给”字段相匹配的证书颁发者名称和证书签署者名称。此示例将进行测试，以确保在将新签名提交到磁盘之前用户选择的数字签名满足特定条件，例如没有过期。*/
function test(strIssuer, strSigner) {
    var AddSignature;
    try {
        //Display the dialog box that lets the
        //user select a digital signature.
        //If the user selects a signature, then
        //it is added to the Signatures
        //collection. If the user does not, then
        //an error is returned.
        var sig = Application.ActiveWorkbook.Signatures.Add();

        //Test several properties before commiting the Signature object to disk.
        if(sig.Issuer == strIssuer && sig.Signer == strSigner && sig.IsCertificateExpired == false && sig.IsCertificateRevoked == false && sig.IsValid == true) {
            alert("Signed");
            AddSignature = true;
        }
        //Otherwise, remove the Signature object from the SignatureSet collection.
        else {
            sig.Delete();
            alert("Not signed");
            AddSignature = false ;
        }

        //Commit all signatures in the SignatureSet collection to the disk.
        Application.ActiveWorkbook.Signatures.Commit();
    }
    catch(exception) {
        AddSignature = false ;
        alert("Action canceled.");
    }
}
```

## 方法

| 序号 | 名称                                  | 说明                             |
|:----:|:--------------------------------------|:---------------------------------|
|  1   | [Delete](#Signature.Delete)           | 从集合中删除 Signature 对象。    |
|  2   | [ShowDetails](#Signature.ShowDetails) | 显示与签名数据包相关的详细信息。 |
|  3   | [Sign](#Signature.Sign)               | 创建一个签名数据包。             |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                             |
|:----:|:----------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Signature.Application)         | 获取一个 Application 对象，代表 Signature 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [CanSetup](#Signature.CanSetup)               | 获取一个 Boolean 类型的值，指示用户是否能设置 Signature 对象的属性。只读。                                                       |
|  3   | [Creator](#Signature.Creator)                 | 获取一个 32 位整数，指示创建 Signature 对象时所使用的应用程序。只读。                                                            |
|  4   | [Details](#Signature.Details)                 | 获取有关签名的信息。只读。                                                                                                       |
|  5   | [IsSignatureLine](#Signature.IsSignatureLine) | 获取一个指示此行是否为签名行的值。只读。                                                                                         |
|  6   | [IsSigned](#Signature.IsSigned)               | 获取一个 Boolean 类型的值，指示是否已成功签署文档。只读。                                                                        |
|  7   | [Parent](#Signature.Parent)                   | 获取 Signature 对象的 Parent 对象。只读。                                                                                        |
|  8   | [Setup](#Signature.Setup)                     | 获取 SignatureSetup 对象，该对象提供对签名数据包的各种属性的访问。只读。                                                         |
|  9   | [SignatureSetup](#Signature.SignatureSetup)   | 获取与作为签名行的 Signature 对象相关联的 Shape 对象。只读。                                                                     |
|  10  | [SortHint](#Signature.SortHint)               | 获取一个代表数据包（包含多个签名）中的签名排序顺序的值。只读。                                                                   |

## 成员方法

### Signature.Delete

从集合中删除 **Signature** 对象。

**语法**

**express.Delete ()**

\*express是一个代表 Signature 对象的变量。

**说明**

对于 **Scripts** 集合，使用 **Delete** 方法将删除指定的 WPS 文档、ET 工作表或 WPP 幻灯片中的所有脚本。在宿主应用程序中，脚本定位标记以形状表示。因此，与每个 **msoScriptAnchor** 类型的脚本定位标记相关联的 **Shape** 对象将从 ET 和 WPP 的 **Shapes** 集合以及 WPS 的 **InlineShapes** 和 **Shapes** 集合中删除。

### Signature.ShowDetails

显示与签名数据包相关的详细信息。

**语法**

**express.ShowDetails ()**

\*express是一个代表 Signature 对象的变量。

**示例**

``` JavaScript
/*以下示例将调用 ShowDetails 方法来显示 Signature 对象的详细信息。*/function test(objSignature) {
    if(objSignature.IsSigned) {
        alert("The document has been signed with the following details:" + objSignature.ShowDetails());
    }else {
    alert("The document has not been signed.");
    }
}
```

### Signature.Sign

创建一个签名数据包。

**语法**

**express.Sign ( *varSigImg* , *varDelSuggSigner* , *varDelSuggSignerLine2* , *varDelSuggSignerEmail* )**

\*express是一个代表 Signature 对象的变量。

**参数**

| 名称                    | 必选/可选 | 数据类型 | 说明                       |
|-------------------------|-----------|----------|----------------------------|
| *varSigImg*             | 必选      | Variant  | 签名行的图形图像。         |
| *varDelSuggSigner*      | 必选      | Variant  | 建议签署人。               |
| *varDelSuggSignerLine2* | 必选      | Variant  | 附加签名行。               |
| *varDelSuggSignerEmail* | 必选      | Variant  | 建议签署人的电子邮件地址。 |

**说明**

调用 **Sign** 方法时，WPS Office将创建一个清单，并调用签名提供程序为文档中的每个流创建一个哈希。Office 随后会将结果捆绑到一个未签名的 XMLDSIG 模板中，并调用提供程序修改 XMLDSIG（如有必要），然后对其进行签署。生成的已签署签名之后将被退还给 Office 以便存储。

**示例**

``` JavaScript
/*在下面的示例中，将设置代表签名图像、签署人、签署人职务以及签署人电子邮件的变量，然后将调用 Sign 方法来创建并签署签名数据包。*/
function test(sig, img){
    var varSuggestedSigner = "Nancy Davolio";
    var varSignatureTitle = "Sales Represenative";
    var varSignerEmail = "ndavolio@northwindtraders.com";
    sig.Sign(img, varSuggestedSigner, varSignatureTitle, varSignerEmail);
}
```

## 成员属性

### Signature.Application

获取一个 **Application** 对象，代表 **Signature** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 Signature 对象的变量。

### Signature.CanSetup

获取一个 **Boolean** 类型的值，指示用户是否能设置 **Signature** 对象的属性。只读。

**语法**

**express.CanSetup**

\*express是一个代表 Signature 对象的变量。

### Signature.Creator

获取一个 32 位整数，指示创建 **Signature** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 Signature 对象的变量。

### Signature.Details

获取有关签名的信息。只读。

**语法**

**express.Details**

\*express是一个代表 Signature 对象的变量。

**说明**

可读取的信息包括：与签名关联的证书是否已到期、签名是否有效以及签名是否为只读。

### Signature.IsSignatureLine

获取一个指示此行是否为签名行的值。只读。

**语法**

**express.IsSignatureLine**

\*express是一个代表 Signature 对象的变量。

### Signature.IsSigned

获取一个 **Boolean** 类型的值，指示是否已成功签署文档。只读。

**语法**

**express.IsSigned**

\*express是一个代表 Signature 对象的变量。

### Signature.Parent

获取 Signature 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Signature 对象的变量。

### Signature.Setup

获取 **SignatureSetup** 对象，该对象提供对签名数据包的各种属性的访问。只读。

**语法**

**express.Setup**

\*express是一个代表 Signature 对象的变量。

### Signature.SignatureSetup

获取与作为签名行的 **Signature** 对象相关联的 **Shape** 对象。只读。

**语法**

**express.SignatureSetup**

\*express是一个代表 Signature 对象的变量。

### Signature.SortHint

获取一个代表数据包（包含多个签名）中的签名排序顺序的值。只读。

**语法**

**express.SortHint**

\*express是一个代表 Signature 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/Signature](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SignatureProvider 对象

## 简介

代表签名提供程序加载项。

|                                                                     |
|---------------------------------------------------------------------|
| 签名提供程序只能在自定义COM 加载项中实现，而不能在 宏编辑器中实现。 |

## 方法

| 序号 | 名称                                                                        | 说明                                                                                                   |
|:----:|:----------------------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------|
|  1   | [GenerateSignatureLineImage](#SignatureProvider.GenerateSignatureLineImage) | 获取签名行图像。                                                                                       |
|  2   | [GetProviderDetail](#SignatureProvider.GetProviderDetail)                   | 向签名提供程序加载项查询各种详细信息。                                                                 |
|  3   | [HashStream](#SignatureProvider.HashStream)                                 | 允许签名提供程序加载项为文档创建一个哈希值，您可以使用该值来确定在进行数字签名后文档内容是否被篡改过。 |
|  4   | [NotifySignatureAdded](#SignatureProvider.NotifySignatureAdded)             | 用于显示一个对话框，通知用户签署过程已完成并为加载项提供附加功能。                                     |
|  5   | [ShowSignatureDetails](#SignatureProvider.ShowSignatureDetails)             | 使签名提供程序加载项能够显示有关已签署签名行的详细信息，并显示诸如安全时间戳等其他已存储信息。         |
|  6   | [ShowSignatureSetup](#SignatureProvider.ShowSignatureSetup)                 | 使签名提供程序加载项能够向用户显示 “签名设置” 对话框。                                                 |
|  7   | [ShowSigningCeremony](#SignatureProvider.ShowSigningCeremony)               | 使签名提供程序加载项能够向用户显示 “签名” 对话框，从而允许用户指定他们的身份并随后获得验证。           |
|  8   | [SignXmlDsig](#SignatureProvider.SignXmlDsig)                               | 用于签署 XMLDSIG 模板。                                                                                |
|  9   | [VerifyXmlDsig](#SignatureProvider.VerifyXmlDsig)                           | 基于文档的签署状态和用于签名的证书的合法性来验证签名。                                                 |

## 成员方法

### SignatureProvider.GenerateSignatureLineImage

获取签名行图像。

**语法**

**express.GenerateSignatureLineImage ( *siglnimg* , *psigsetup* , *psiginfo* )**

\*express是一个代表 SignatureProvider 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型           | 说明                                 |
|-------------|-----------|--------------------|--------------------------------------|
| *siglnimg*  | 必选      | SignatureLineImage | 如果为签名行图形，则包含相应的名称。 |
| *psigsetup* | 必选      | SignatureSetup     | 指定签名提供程序加载项的初始设置。   |
| *psiginfo*  | 必选      | SignatureInfo      | 指定有关签名提供程序加载项的信息。   |

**说明**

**SignatureProvider** 对象只能在自定义签名提供程序加载项中使用。对于在文档内容中显示的每个图像都调用此方法。可以异步调用该方法。例如，可以在签名设置之后直接为“未签名”图像和“无软件”图像调用该方法。然后可以在为“已签名”图像签名之后调用该方法。使用的四个图像是：

- **siglnimgSoftwareRequired** ：如果在用户计算机上未安装签名提供程序加载项，则向用户显示此图像。如果用户尝试签名或查看签名行，则他们将被重定向到提供程序提供的超链接（在 **GetProviderDetail** 方法中指定）。
- **siglnimgUnsigned** ：对于未签名的签名图像，将显示此图像。基本上，当加载具有未签名的签名行的文档时，签名提供程序将提示最新的签名图像，并显示此图像。
- **siglnimgSignedValid** ：这是在签名行已进行签名且有效的情况下显示的图像（或者，更具体地说，已进行签名且签名没有注册为无效）。在打开文档时，假定在验证过程完成之前所有的已签署签名都是有效的，此时对于无效的签名显示“已签名/无效”图像。由于签名验证需要很长时间，因此签名验证在后台线程上与 Office 并行运行。由于加载项实现签名验证，因此代码将与 Office 并行运行，并且不应该试图在签名或验证期间显示 UI。
- **siglnimgSignedInvalid** ：这是在对签名行进行了签名但签名存在问题（如文档已修改或者用户证书已作废）时显示的图像。由于加载项实现签名验证，因此您可以决定签名无效的方式和时间。

### SignatureProvider.GetProviderDetail

向签名提供程序加载项查询各种详细信息。

**语法**

**express.GetProviderDetail ( *sigprovdet* )**

\*express是一个代表 SignatureProvider 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型                | 说明                                                     |
|--------------|-----------|-------------------------|----------------------------------------------------------|
| *sigprovdet* | 必选      | SignatureProviderDetail | 包含一个枚举值，该枚举值代表要向加载项查询的信息的类型。 |

**返回值**

Variant

**说明**

**SignatureProvider** 对象只能在自定义签名提供程序加载项中使用。此方法用于向加载项查询三段信息：

- 加载项支持哪种哈希算法？
- 加载项是否只是用户界面 (UI)，或者它是否支持哈希处理和验证？如果返回 **TRUE** ，则表明 WPS Office不会调用加载项进行哈希处理或验证，而只是调用加载项来显示 UI。
- 如果用户缺少签名加载项，加载项应向用户提供什么 URL？

### SignatureProvider.HashStream

允许签名提供程序加载项为文档创建一个哈希值，您可以使用该值来确定在进行数字签名后文档内容是否被篡改过。

**语法**

**express.HashStream ( *QueryContinue* , *Stream* )**

\*express是一个代表 SignatureProvider 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型       | 说明                                                                   |
|-----------------|-----------|----------------|------------------------------------------------------------------------|
| *QueryContinue* | 必选      | IQueryContinue | 提供了一种方式，可向主机应用程序进行查询以获得继续进行哈希处理的许可。 |
| *Stream*        | 必选      | IStream        | 包含数据流。                                                           |

**返回值**

Byte

**说明**

**SignatureProvider** 对象只能在自定义签名提供程序加载项中使用。将为文档中的每个签名数据流调用一次此方法。返回值是一个代表使用哈希算法计算出的哈希值的字节数组。

### SignatureProvider.NotifySignatureAdded

用于显示一个对话框，通知用户签署过程已完成并为加载项提供附加功能。

**语法**

**express.NotifySignatureAdded ( *ParentWindow* , *psigsetup* , *psiginfo* )**

\*express是一个代表 SignatureProvider 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型       | 说明                                               |
|----------------|-----------|----------------|----------------------------------------------------|
| *ParentWindow* | 必选      | IOleWindow     | 允许主机应用程序获取包含所显示对话框的窗口的句柄。 |
| *psigsetup*    | 必选      | SignatureSetup | 包含签名提供程序的初始设置。                       |
| *psiginfo*     | 必选      | SignatureInfo  | 包含有关签名提供程序加载项的信息。                 |

**说明**

将在签署过程完成后调用此方法。使签名提供程序加载项能够为加载项增加附加功能。例如，如果想要提供用户可在其中上载已签署文档的存档服务，则可以使用此方法来启动该过程。

### SignatureProvider.ShowSignatureDetails

使签名提供程序加载项能够显示有关已签署签名行的详细信息，并显示诸如安全时间戳等其他已存储信息。

**语法**

**express.ShowSignatureDetails ( *ParentWindow* , *psigsetup* , *psiginfo* , *XmlDsigStream* , *pcontverres* , *pcertverres* )**

\*express是一个代表 SignatureProvider 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型                       | 说明                                       |
|-----------------|-----------|--------------------------------|--------------------------------------------|
| *ParentWindow*  | 必选      | IOleWindow                     | 包含窗口的句柄，该窗口中包含签名详细信息。 |
| *psigsetup*     | 必选      | SignatureSetup                 | 指定签名提供程序的初始设置。               |
| *psiginfo*      | 必选      | SignatureInfo                  | 指定有关已签署签名行的信息。               |
| *XmlDsigStream* | 必选      | IStream                        | 代表数据流或 XML 的二进制大对象。          |
| *pcontverres*   | 必选      | ContentVerificationResults     | 包含一个值，代表验证签名内容的结果。       |
| *pcertverres*   | 必选      | CertificateVerificationResults | 包含一个值，代表验证签名证书的结果。       |

### SignatureProvider.ShowSignatureSetup

使签名提供程序加载项能够向用户显示 **“签名设置”** 对话框。

**语法**

**express.ShowSignatureSetup ( *ParentWindow* , *psigsetup* )**

\*express是一个代表 SignatureProvider 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型       | 说明                                           |
|----------------|-----------|----------------|------------------------------------------------|
| *ParentWindow* | 必选      | IOleWindow     | 包含窗口的句柄，该窗口中包含“签名设置”对话框。 |
| *psigsetup*    | 必选      | SignatureSetup | 指定签名提供程序的初始设置。                   |

**说明**

在插入时配置过程中，以及在用户稍后想要重新配置签名行的情况下，都将使用此方法。将在此回调过程中显示 **“签名设置”** 对话框并等待用户选择 **“确定”** 或 **“取消”** 。除非明确需要作者提供有关签名行的信息，否则不必为签名设置显示对话框。如果可以将所有必要的详细信息提供给 WPS Office而无需用户输入，则不需要对话框。

### SignatureProvider.ShowSigningCeremony

使签名提供程序加载项能够向用户显示 **“签名”** 对话框，从而允许用户指定他们的身份并随后获得验证。

**语法**

**express.ShowSigningCeremony ( *ParentWindow* , *psigsetup* , *psiginfo* )**

\*express是一个代表 SignatureProvider 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型       | 说明                                       |
|----------------|-----------|----------------|--------------------------------------------|
| *ParentWindow* | 必选      | IOleWindow     | 包含窗口的句柄，该窗口中包含“签名”对话框。 |
| *psigsetup*    | 必选      | SignatureSetup | 指定签名提供程序的初始设置。               |
| *psiginfo*     | 必选      | SignatureInfo  | 指定有关签名提供程序的信息。               |

**说明**

当用户尝试在签名行上签名时，或者，如果加载项已在 Office 应用程序的对象模型中对 **SignatureLine** 对象调用了 **Sign** 方法，WPS Office应用程序将在内部调用此方法。

### SignatureProvider.SignXmlDsig

用于签署 XMLDSIG 模板。

**语法**

**express.SignXmlDsig ( *QueryContinue* , *psigsetup* , *psiginfo* , *XmlDsigStream* )**

\*express是一个代表 SignatureProvider 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型       | 说明                                                             |
|-----------------|-----------|----------------|------------------------------------------------------------------|
| *QueryContinue* | 必选      | IQueryContinue | 提供了一种方式，可查询主机应用程序以获得继续进行验证操作的许可。 |
| *psigsetup*     | 必选      | SignatureSetup | 指定有关签名行的配置信息。                                       |
| *psiginfo*      | 必选      | SignatureInfo  | 指定从签署仪式中捕获的信息。                                     |
| *XmlDsigStream* | 必选      | IStream        | 代表一个包含 XML 的数据流，该数据流又代表一个 XMLDSIG 对象。     |

**说明**

XMLDSIG 是一种基于标准的签名格式 (http://www.w3.org/TR/xmldsig-core/)，可由第三方验证。此格式是 WPS Office中的默认签名格式。

### SignatureProvider.VerifyXmlDsig

基于文档的签署状态和用于签名的证书的合法性来验证签名。

**语法**

**express.VerifyXmlDsig ( *QueryContinue* , *psigsetup* , *psiginfo* , *XmlDsigStream* , *pcontverres* , *pcertverres* )**

\*express是一个代表 SignatureProvider 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型                       | 说明                                                             |
|-----------------|-----------|--------------------------------|------------------------------------------------------------------|
| *QueryContinue* | 必选      | IQueryContinue                 | 提供了一种方式，可查询主机应用程序以获得继续进行验证操作的许可。 |
| *psigsetup*     | 必选      | SignatureSetup                 | 指定有关签名行的配置信息。                                       |
| *psiginfo*      | 必选      | SignatureInfo                  | 指定从签署仪式中捕获的信息。                                     |
| *XmlDsigStream* | 必选      | IStream                        | 代表一个包含 XML 的数据流，该数据流又代表一个 XMLDSIG 对象。     |
| *pcontverres*   | 必选      | ContentVerificationResults     | 指定签名验证操作的状态。                                         |
| *pcertverres*   | 必选      | CertificateVerificationResults | 指定签名证书验证的状态。                                         |

**说明**

XMLDSIG 是一种基于标准的签名格式 (http://www.w3.org/TR/xmldsig-core/)，可由第三方验证。此格式是 WPS Office中的默认签名格式。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SignatureProvider](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SmartArtColor 对象

## 简介

为 SmartArt 图表选择配色方案。

## 说明

模拟 WPS Office Fluent 功能区用户界面中“SmartArt 工具”选项卡上“设计”组内的“更改颜色”命令上的命令。

``` JavaScript
/*下面的代码设置 Smart Art 图表的配色方案。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Color = Application.SmartArtColors.Item(1)
```

## 属性

| 序号 | 名称                                      | 说明                                                                           |
|:----:|:------------------------------------------|:-------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtColor.Application) | 获取一个 Application 对象，该对象代表 SmartArtColor 对象的容器应用程序。只读。 |
|  2   | [Category](#SmartArtColor.Category)       | 检索与 SmartArt 颜色样式关联的主类别名称。只读。                               |
|  3   | [Creator](#SmartArtColor.Creator)         | 返回一个 32 位整数，该整数指示在其中创建了此对象的应用程序。只读。             |
|  4   | [Description](#SmartArtColor.Description) | 检索 SmartArt 颜色样式的说明。只读。                                           |
|  5   | [Id](#SmartArtColor.Id)                   | 检索关联的 SmartArt 颜色样式的唯一 Id。只读。                                  |
|  6   | [Name](#SmartArtColor.Name)               | 检索 SmartArt 颜色样式的字符串名称。只读。                                     |
|  7   | [Parent](#SmartArtColor.Parent)           | 返回调用对象。只读。                                                           |

## 成员属性

### SmartArtColor.Application

获取一个 **Application** 对象，该对象代表 **SmartArtColor** 对象的容器应用程序。只读。

**语法**

**express.Application**

\*express是一个代表 SmartArtColor 对象的变量。

### SmartArtColor.Category

检索与 SmartArt 颜色样式关联的主类别名称。只读。

**语法**

**express.Category**

\*express是一个代表 SmartArtColor 对象的变量。

### SmartArtColor.Creator

返回一个 32 位整数，该整数指示在其中创建了此对象的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SmartArtColor 对象的变量。

### SmartArtColor.Description

检索 SmartArt 颜色样式的说明。只读。

**语法**

**express.Description**

\*express是一个代表 SmartArtColor 对象的变量。

### SmartArtColor.Id

检索关联的 SmartArt 颜色样式的唯一 Id。只读。

**语法**

**express.Id**

\*express是一个代表 SmartArtColor 对象的变量。

**说明**

与此属性关联的 ID 区分大小写。

### SmartArtColor.Name

检索 SmartArt 颜色样式的字符串名称。只读。

**语法**

**express.Name**

\*express是一个代表 SmartArtColor 对象的变量。

### SmartArtColor.Parent

返回调用对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SmartArtColor 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtColor](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
