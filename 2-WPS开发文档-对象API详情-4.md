# Application 对象

## 简介

Application 是浏览器 Window 对象的子对象。 Application 对象在浏览器创建时自动注入到浏览器执行上下文环境

## 说明

Wps 对象是访问 Application 对象模型的根对象，它在浏览器创建后动态注入到浏览器执行上下文中，下列代码展示如何从 Wps 对象拿到 Application 对象，然后从 Application 对象拿到当前文档的名字。

``` JavaScript
let app = Application
let doc = app.ActiveDocument
if(doc){
    let docName = doc.Name
}
```

除了原office标准对象外的扩展对象也都是 Application 对象的子对象，具体 Application 对象扩展了哪些对象，可以参阅 Application 对象成员相关说明，下列代码展示了如果从 Application 对象拿到 FileSystem 对象。

``` JavaScript
let fs =  Application.FileSystem
```

Application 对象还包含一些函数，具体 Application 对象支持哪些函数，可以参阅 Application 对像成员说明，下列代码展示了如何从 Application 对象创建一个网页对话框。

``` JavaScript
let width =  300 * window.devicePixelRatio
let hight= 200 * window.devicePixelRatio
let bModal = true
Application.ShowDialog(""http://www.wps.cn"", "wps首页", width, hight, bModal)
```

## 方法

| 序号 | 名称                                          | 说明                                                                         |
|:----:|:----------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [CreateTaskpane](#Application.CreateTaskpane) | 用给定的url创建一个 Taskpane 。                                              |
|  2   | [GetTaskpane](#Application.GetTaskpane)       | 根据给定的ID值，返回对应的 TaskPane 对象，这个ID值与 TaskPane.ID 对应。      |
|  3   | [ShowDialog](#Application.ShowDialog)         | 根据给定的url、标题、宽高等信息创建一个对话框，对话框中的内容是一个web网页。 |
|  4   | [alert](#Application.alert)                   | 弹出alert警告。                                                              |
|  5   | [confirm](#Application.confirm)               | 弹出一个确认框。返回值为一个bool变量，代表对确认框的操作是确认还是取消。     |

## 属性

| 序号 | 名称                                        | 说明                                                                                                                                                                                                                                                      |
|:----:|:--------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ApiEvent](#Application.ApiEvent)           | 返回一个 ApiEvent 对象，有关 ApiEvent 对象的相关说明，请参阅 ApiEvent 对象。                                                                                                                                                                              |
|  2   | [Env](#Application.Env)                     | 返回一个 Env 对象，该对象代表获取系统环境相关的操作对象，有关Env对象的详细属性和方法，可以参考 Env 对象。                                                                                                                                                 |
|  3   | [FileSystem](#Application.FileSystem)       | 返回一个 FileSystem 对象，该对象包含了对文件和文件夹的常见操作，有关FileSystem对象的详细属性和方法，可以参考 FileSystem 对象。                                                                                                                            |
|  4   | [PluginStorage](#Application.PluginStorage) | 返回一个 PluginStorage 对象，该对象与 LocalStorage 的用法类似，但该对象的数据不会被持久化，在wps的jsapi插件中，该对象在一个插件中对应着同一个实例，因此该对象可以在一个jsapi插件的不同网页间共享数据，有关该对象的详细介绍，可以参考 PluginStorage 对象。 |

## 成员方法

### Application.CreateTaskpane

用给定的url创建一个 **Taskpane** 。

**语法**

**express.CreateTaskpane ( *url* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明        |
|-------|-----------|----------|-------------|
| *url* | 必选      | String   | 给定url路径 |

### Application.GetTaskpane

根据给定的ID值，返回对应的 TaskPane 对象，这个ID值与 TaskPane.ID 对应。

**语法**

**express.GetTaskpane ( *ID* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                |
|------|-----------|----------|---------------------|
| *ID* | 必选      | Number   | 表示TaskPane.ID属性 |

### Application.ShowDialog

根据给定的url、标题、宽高等信息创建一个对话框，对话框中的内容是一个web网页。

**语法**

**express.ShowDialog ( *url* , *caption* , *width* , *height* , *bModal* , *hasCaption* , *resizeEdge* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                                                            |
|--------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------|
| *url*        | 必选      | String   | 此参数代表对话框网页的url, 这个url可以是一个http/https的网页，也可以是一个本地资源网页。                                                        |
| *caption*    | 必选      | String   | 对话框的标题。                                                                                                                                  |
| *width*      | 必选      | Long     | 表示对话框的宽度，这个宽度是物理像素值，尤其要注意在普通屏和高分辨率屏下是不一样的，常见的做法是借助于 window.devicePixelRatio 来消除这个影响。 |
| *height*     | 必选      | Long     | 表示对话框的高度，这个高度是物理像素值，尤其要注意在普通屏和高分辨率屏下是不一样的，常见的做法是借助于 window.devicePixelRatio 来消除这个影响。 |
| *bModal*     | 必选      | Bool     | 表示对话框是否模态。                                                                                                                            |
| *hasCaption* | 可选      | Bool     | 默认为true，表示对话框是否有标题栏和边框。为false时为无边框窗口，允许开发者对标题栏，窗口阴影等进行高级自定义                                   |
| *resizeEdge* | 可选      | Bool     | 窗口为无边框窗口时有效。默认为2，表示操作对话框缩放的宽度，单位是像素。                                                                         |

**示例**

``` JavaScript
以下代码示例创建了一个对话框，并让这个对话框模态显示。
let width = 400 * window.devicePixelRatio
let height = 300 * window.devicePixelRatio
Application.ShowDialog("https://www.wps.cn", "wps网站", width, height, true)
```

### Application.alert

弹出alert警告。

**语法**

**express.alert ( *text* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                          |
|--------|-----------|----------|-------------------------------|
| *text* | 必选      | String   | 表示警告框要显示的字符串内容. |

**说明**

alert实际上是浏览器 **window** 对象内置的一个方法，在 **wps** 的js环境中，wps对这个对象进行了拦截，主要原因是由于chrome的技术架构决定的。 wps的js环境是内置了一个chrome的内核来展示网页、执行js脚本，在chrome的技术架构中，chrome的js执行、css及html渲染、网页uil用户操作等都分别是它自己的线程中， 同样的，如果alert不作拦截，当网页前端代码调用alert时，这个alert框仅仅只是阻塞住了chrome的ui线程，而对wps主线程没有影响，这样就会产生alert框弹出来了，但wps仍然能 响应用户操作的奇怪现象，因此，wps对alert框进行了拦截，让alert框在wps主线程中弹出来，当alert弹出时，阻塞的会是wps主线程。

### Application.confirm

弹出一个确认框。返回值为一个bool变量，代表对确认框的操作是确认还是取消。

**语法**

**express.confirm ( *text* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                          |
|--------|-----------|----------|-------------------------------|
| *text* | 必选      | String   | 表示确认框要显示的字符串内容. |

**说明**

confirm与alert一样，实际上也是浏览器 **window** 对象内置的一个方法，在wps的js环境中，wps对这个对象进行了拦截，主要原因是由于chrome的技术架构决定的。 wps的js环境是内置了一个chrome的内核来展示网页、执行js脚本，在chrome的技术架构中，chrome的js执行、css及html渲染、网页uil用户操作等都分别是它自己的线程中， 同样的，如果不作拦截，当网页前端代码调用confirm时，这个confirm框仅仅只是阻塞住了chrome的ui线程，而对wps主线程没有影响，这样就会产生confirm框弹出来了，但wps仍然能 响应用户操作的奇怪现象，因此，wps对confirm框进行了拦截，让confirm框在wps主线程中弹出来，当confirm弹出时，阻塞的会是wps主线程。

## 成员属性

### Application.ApiEvent

返回一个 ApiEvent 对象，有关 **ApiEvent** 对象的相关说明，请参阅 ApiEvent 对象。

**语法**

**express.ApiEvent**

\*express是一个代表 Application 对象的变量。

### Application.Env

返回一个 Env 对象，该对象代表获取系统环境相关的操作对象，有关Env对象的详细属性和方法，可以参考 Env 对象。

**语法**

**express.Env**

\*express是一个代表 Application 对象的变量。

### Application.FileSystem

返回一个 FileSystem 对象，该对象包含了对文件和文件夹的常见操作，有关FileSystem对象的详细属性和方法，可以参考 FileSystem 对象。

**语法**

**express.FileSystem**

\*express是一个代表 Application 对象的变量。

### Application.PluginStorage

返回一个 PluginStorage 对象，该对象与 **LocalStorage** 的用法类似，但该对象的数据不会被持久化，在wps的jsapi插件中，该对象在一个插件中对应着同一个实例，因此该对象可以在一个jsapi插件的不同网页间共享数据，有关该对象的详细介绍，可以参考 PluginStorage 对象。

**语法**

**express.PluginStorage**

\*express是一个代表 Application 对象的变量。

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/Office 全局对象/Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TaskPane 对象

## 简介

TaskPane 对象是一个用来浏览网页的用户界面面板

## 方法

| 序号 | 名称                           | 说明                                                  |
|:----:|:-------------------------------|:------------------------------------------------------|
|  1   | [Delete](#TaskPane.Delete)     | 此方法用于删除当前的 Taskpane 。                      |
|  2   | [Navigate](#TaskPane.Navigate) | 此方法用于在当前的 Taskpane 上重新载入url对应的网页。 |

## 属性

| 序号 | 名称                                   | 说明                                                     |
|:----:|:---------------------------------------|:---------------------------------------------------------|
|  1   | [DockPosition](#TaskPane.DockPosition) | 可读写，设置或返回一个整型，代表 Taskpane 的停靠位置。   |
|  2   | [Height](#TaskPane.Height)             | 返回一个整数，此整数代码 Taskpane 的高度。               |
|  3   | [ID](#TaskPane.ID)                     | 返回一个整数，此整数代码 Taskpane 的ID。                 |
|  4   | [Visible](#TaskPane.Visible)           | 可读写，设置或返回一个bool类型，代表 Taskpane 的可见性。 |
|  5   | [Width](#TaskPane.Width)               | 返回一个整数，此整数代码 Taskpane 的宽度。               |

## 成员方法

### TaskPane.Delete

此方法用于删除当前的 **Taskpane** 。

**语法**

**express.Delete ()**

\*express是一个代表 TaskPane 对象的变量。

### TaskPane.Navigate

此方法用于在当前的 **Taskpane** 上重新载入url对应的网页。

**语法**

**express.Navigate ( *url* )**

\*express是一个代表 TaskPane 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明                                                     |
|-------|-----------|----------|----------------------------------------------------------|
| *url* | 必选      | String   | 这个url可以是一个在线的网址，也可以是一个本地的html资源. |

**示例**

``` JavaScript
let tp = Application.CreateTaskpane("https://www.kingsoft.com") if (tp){ tp.Visible = true tp.Navigate("https://www.wps.cn") }
```

## 成员属性

### TaskPane.DockPosition

可读写，设置或返回一个整型，代表 **Taskpane** 的停靠位置。

**语法**

**express.DockPosition**

\*express是一个代表 TaskPane 对象的变量。

**说明**

{ msoCTPDockPositionLeft : 0, msoCTPDockPositionTop : 1, msoCTPDockPositionRight : 2, msoCTPDockPositionBottom : 3, msoCTPDockPositionFloating : 4 }

### TaskPane.Height

返回一个整数，此整数代码 **Taskpane** 的高度。

**语法**

**express.Height**

\*express是一个代表 TaskPane 对象的变量。

### TaskPane.ID

返回一个整数，此整数代码 **Taskpane** 的ID。

**语法**

**express.ID**

\*express是一个代表 TaskPane 对象的变量。

### TaskPane.Visible

可读写，设置或返回一个bool类型，代表 **Taskpane** 的可见性。

**语法**

**express.Visible**

\*express是一个代表 TaskPane 对象的变量。

### TaskPane.Width

返回一个整数，此整数代码 **Taskpane** 的宽度。

**语法**

**express.Width**

\*express是一个代表 TaskPane 对象的变量。

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/任务窗格/TaskPane](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Label 对象

## 简介

用于显示用户不能编辑的文本或图像，例如图形下的标题文本。

## 说明

用于显示用户不能编辑的文本或图像，例如图形下的标题文本。

## 方法

| 序号 | 名称                | 说明                       |
|:----:|:--------------------|:---------------------------|
|  1   | [Move](#Label.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                | 说明                                                      |
|:----:|:------------------------------------|:----------------------------------------------------------|
|  1   | [AutoSize](#Label.AutoSize)         | 指定该对象是否根据内容自动设置它的大小                    |
|  2   | [BackColor](#Label.BackColor)       | 获取或设置指定对象的内部颜色                              |
|  3   | [BackStyle](#Label.BackStyle)       | 您可以使用该属性指定控件是否透明                          |
|  4   | [BorderColor](#Label.BorderColor)   | 您可以通过该属性指定控件边框的颜色                        |
|  5   | [BorderStyle](#Label.BorderStyle)   | 指定控件边框的显示方式                                    |
|  6   | [Caption](#Label.Caption)           | 获取或设置出现在控件中的文本                              |
|  7   | [Enabled](#Label.Enabled)           | 您可以通过该属性设置或返回是否启用该控件                  |
|  8   | [Font](#Label.Font)                 | 指定对象的字体，只读                                      |
|  9   | [ForeColor](#Label.ForeColor)       | 您可以通过该属性指定控件中文本的颜色                      |
|  10  | [Height](#Label.Height)             | 获取或设置指定对象的高度                                  |
|  11  | [HeightPolicy](#Label.HeightPolicy) | 获取或设置指定对像的高度策略                              |
|  12  | [Left](#Label.Left)                 | 您可以通过该属性指定控件在对象或报表中的位置              |
|  13  | [Name](#Label.Name)                 | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  14  | [TabIndex](#Label.TabIndex)         | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  15  | [TabStop](#Label.TabStop)           | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  16  | [TextAlign](#Label.TextAlign)       | 该属性用于指定对象中文本的对齐方式                        |
|  17  | [Top](#Label.Top)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  18  | [Visible](#Label.Visible)           | 返回或设置该对象是否可见                                  |
|  19  | [Width](#Label.Width)               | 获取或设置指定对象的宽度                                  |
|  20  | [WidthPolicy](#Label.WidthPolicy)   | 获取或设置指定对象的宽度策略                              |
|  21  | [WordWrap](#Label.WordWrap)         | 指定控件的内容是否在行尾自动换行                          |

## 成员方法

### Label.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 Label 对象的变量。

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
/*将Label移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.Label1.Move(100,100,100,100)
}
```

## 成员属性

### Label.AutoSize

指定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 Label 对象的变量。

**说明**

指定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置Label控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.Label1.AutoSize = true
}
```

### Label.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将Label1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.Label1.BackColor = 0x665898
}
```

### Label.BackStyle

您可以使用该属性指定控件是否透明

**语法**

**express.BackStyle**

\*express是一个代表 Label 对象的变量。

**说明**

您可以使用该属性指定控件是否透明

**示例**

``` JavaScript
/*将Label1设置为透明*/
function func1()
{
    UserForm1.Label1.BackStyle = 0
}
```

### Label.BorderColor

您可以通过该属性指定控件边框的颜色

**语法**

**express.BorderColor**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定控件边框的颜色

**示例**

``` JavaScript
/*将Label1的边框设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.Label1.BorderColor = 0x665898
}
```

### Label.BorderStyle

指定控件边框的显示方式

**语法**

**express.BorderStyle**

\*express是一个代表 Label 对象的变量。

**说明**

指定控件边框的显示方式，可设置为以下值

| 设置   | 值  | 描述       |
|--------|-----|------------|
| None   | 0   | 无边框     |
| Single | 1   | 边框为实线 |

**示例**

``` JavaScript
/*将Label1的边框显示方式设置为实线*/
function func1()
{
    UserForm1.Label1.BorderStyle = 1
}
```

### Label.Caption

获取或设置出现在控件中的文本

**语法**

**express.Caption**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置出现在控件中的文本

**示例**

``` JavaScript
/*将Label1中的文本设置为"Label 1"*/
function func1()
{
    UserForm1.Label1.Caption = "Label 1"
}
```

### Label.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.Label1.Enabled = true
}
```

### Label.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 Label 对象的变量。

**说明**

指定对象的字体，只读

### Label.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将Label1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.Label1.ForeColor = 0x658978
}
```

### Label.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将Label1控件的高度设置为100*/
function func1()
{
    UserForm1.Label1.Height = 100
}
```

### Label.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置指定对像的高度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将Label1的高度设置为固定*/
function func1()
{
    UserForm1.Label1.HeightPolicy = 2
}
```

### Label.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将Label1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.Label1.Left = 100
}
```

### Label.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### Label.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将Label1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.Label1.TabIndex = 2
}
```

### Label.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将Label1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.Label1.TabStop = true
}
```

### Label.TextAlign

该属性用于指定对象中文本的对齐方式

**语法**

**express.TextAlign**

\*express是一个代表 Label 对象的变量。

**说明**

该属性用于指定对象中文本的对齐方式，包含以下值

| 设置   | 值  | 描述                     |
|--------|-----|--------------------------|
| Left   | 1   | 文本数字和日期都左对齐   |
| Center | 2   | 文本数字和日期都居中对齐 |
| Right  | 3   | 文本数字和日期都右对齐   |

**示例**

``` JavaScript
/*将Label1中的文本设置为居中对齐*/
function func1()
{
    UserForm1.Label1.TextAlign = 2
}
```

### Label.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 Label 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将Label1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.Label1.Top = 100
}
```

### Label.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 Label 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将Label1设置为可见*/
function func1()
{
    UserForm1.Label1.Visible = true
}
```

### Label.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将Label1的宽度设置为100*/
function func1()
{
    UserForm1.Label1.Width = 100
}
```

### Label.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 Label 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可以设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将Label1的宽度设置为固定*/
function func1()
{
    UserForm1.Label1.WidthPolicy = 2
}
```

### Label.WordWrap

指定控件的内容是否在行尾自动换行

**语法**

**express.WordWrap**

\*express是一个代表 Label 对象的变量。

**说明**

指定控件的内容是否在行尾自动换行

**示例**

``` JavaScript
/*将Label1设置为控件的内容在行尾自动换行*/
function func1()
{
    UserForm1.Label1.WordWrap = true
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/Label](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
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
# AllowEditRanges 对象

## 简介

所有 AllowEditRange 对象的集合，这些对象代表受保护工作表上的可编辑单元格。

## 方法

| 序号 | 名称                        | 说明                                                                 |
|:----:|:----------------------------|:---------------------------------------------------------------------|
|  1   | [Add](#AllowEditRanges.Add) | 在受保护的工作表中添加可编辑的单元格区域。返回 AllowEditRange 对象。 |

## 属性

| 序号 | 名称                            | 说明                                       |
|:----:|:--------------------------------|:-------------------------------------------|
|  1   | [Count](#AllowEditRanges.Count) | 返回一个 Long 值，它代表集合中对象的数量。 |
|  2   | [Item](#AllowEditRanges.Item)   | 从集合中返回一个对象。                     |

## 成员方法

### AllowEditRanges.Add

在受保护的工作表中添加可编辑的单元格区域。返回 AllowEditRange 对象。

**语法**

**express.Add ( *Title* , *Range* , *Password* )**

\*express是一个代表 AllowEditRanges 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                               |
|------------|-----------|----------|------------------------------------|
| *Title*    | 必选      | String   | 单元格区域的标题。                 |
| *Range*    | 必选      | Object   | Range 对象。允许编辑的单元格区域。 |
| *Password* | 可选      | Variant  | 单元格区域的密码。                 |

## 成员属性

### AllowEditRanges.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 AllowEditRanges 对象的变量。

### AllowEditRanges.Item

从集合中返回一个对象。

**语法**

**express.Item**

\*express是一个代表 AllowEditRanges 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/AllowEditRanges](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# AxisTitle 对象

## 简介

代表图表坐标轴标题。

## 说明

使用 AxisTitle 属性可返回 AxisTitle 对象。 只有当坐标轴的 HasTitle 属性为 True 时， AxisTitle 对象才存在，从而才能使用该对象。

示例

``` JavaScript
/*下例激活第一个嵌入式图表，设置其数值轴标题文本，将其字体设为 10 磅的“Bookman”，并将单词“millions”设为倾斜。*/
function test()
{
    Application.Worksheets.Item("sheet1").ChartObjects(1).Activate()
    let axes = Application.ActiveChart.Axes(xlValue)
    axes.HasTitle = true

    let axistitle = axes.AxisTitle
    axistitle.Caption = "Revenue (millions)"
    axistitle.Font.Name = "bookman"
    axistitle.Font.Size = 10
    axistitle.Characters(10, 8).Font.Italic = true
}
```

## 方法

| 序号 | 名称                        | 说明       |
|:----:|:----------------------------|:-----------|
|  1   | [Delete](#AxisTitle.Delete) | 删除对象。 |
|  2   | [Select](#AxisTitle.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                                  | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#AxisTitle.Application)                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Caption](#AxisTitle.Caption)                         | 返回或设置一个 String 值，它代表坐标轴标题文本。                                                                                                                                                                                |
|  3   | [Characters](#AxisTitle.Characters)                   | 返回 Characters 对象，它代表对象文本内某个区域的字符。使用 Characters 对象可为文本字符串内的字符设置格式。                                                                                                                      |
|  4   | [Creator](#AxisTitle.Creator)                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [Format](#AxisTitle.Format)                           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  6   | [Formula](#AxisTitle.Formula)                         | 获取或设置一个 String 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。                                                                                                                                               |
|  7   | [FormulaLocal](#AxisTitle.FormulaLocal)               | 获取或设置一个 String 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。                                                                                                                                         |
|  8   | [FormulaR1C1](#AxisTitle.FormulaR1C1)                 | 获取或设置一个 String 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。                                                                                                                                             |
|  9   | [FormulaR1C1Local](#AxisTitle.FormulaR1C1Local)       | 获取或设置一个 String 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。                                                                                                                                       |
|  10  | [Height](#AxisTitle.Height)                           | 返回对象的高度（以磅为单位）。只读。                                                                                                                                                                                            |
|  11  | [HorizontalAlignment](#AxisTitle.HorizontalAlignment) | 返回或设置一个 Variant 值，它代表指定对象的水平对齐方式。                                                                                                                                                                       |
|  12  | [IncludeInLayout](#AxisTitle.IncludeInLayout)         | 如果在确定图表布局时轴标题将占用图表布局空间，则为 True 。默认值是 True 。可读/写 Boolean 类型。                                                                                                                                |
|  13  | [Left](#AxisTitle.Left)                               | 返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                       |
|  14  | [Name](#AxisTitle.Name)                               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  15  | [Orientation](#AxisTitle.Orientation)                 | 返回或设置一个 Variant 值，它代表文本方向。                                                                                                                                                                                     |
|  16  | [Parent](#AxisTitle.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  17  | [Position](#AxisTitle.Position)                       | 返回或设置图表上轴标题的位置。可读/写 XlChartElementPosition 类型。                                                                                                                                                             |
|  18  | [ReadingOrder](#AxisTitle.ReadingOrder)               | 返回或设置指定对象的阅读次序。可为以下常量之一： xlRTL （从右到左）、 xlLTR （从左到右）或 xlContext 。 Long 类型，可读写。                                                                                                     |
|  19  | [Shadow](#AxisTitle.Shadow)                           | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                               |
|  20  | [Text](#AxisTitle.Text)                               | 返回或设置指定对象中的文本。 String 型，可读写。                                                                                                                                                                                |
|  21  | [Top](#AxisTitle.Top)                                 | 返回或设置一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                      |
|  22  | [VerticalAlignment](#AxisTitle.VerticalAlignment)     | 返回或设置一个 Variant 值，它代表指定对象的垂直对齐方式。                                                                                                                                                                       |
|  23  | [Width](#AxisTitle.Width)                             | 返回对象的宽度（以磅为单位）。只读。                                                                                                                                                                                            |

## 成员方法

### AxisTitle.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

返回值 一个代表所选对象的 Variant 值。

## 成员属性

### AxisTitle.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 AxisTitle 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
  if(Application.ActiveWorkbook.Application.Value == "ET" ){
      alert("This is an ET Application object.")
  }
  else{
      alert("This is not an ET Application object.")
  }
}
```

### AxisTitle.Caption

返回或设置一个 **String** 值，它代表坐标轴标题文本。

**语法**

**express.Caption**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Characters

返回 **Characters** 对象，它代表对象文本内某个区域的字符。使用 **Characters** 对象可为文本字符串内的字符设置格式。

**语法**

**express.Characters**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

参数

| **名称** | **必选/可选** | **数据类型** | **说明**                                                                                      |
|----------|---------------|--------------|-----------------------------------------------------------------------------------------------|
| *Start*  | 可选          | **Variant**  | 要返回的第一个字符。如果此参数是 1 或被省略，则此属性返回一个以第一个字符为开头的字符区域。   |
| *Length* | 可选          | **Variant**  | 要返回的字符数。如果省略此参数，则此属性返回字符串的后半部分（ *Start* 字符之后的所有字符）。 |

**Characters** 对象不是集合。

### AxisTitle.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### AxisTitle.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Formula

获取或设置一个 **String** 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.Formula**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

有关 A1 样式表示法的详细信息，请参阅 <span style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;background-color:#FFFFFF;">操作方法：使用 A1 表示法引用单元格和区域</span> 。

### AxisTitle.FormulaLocal

获取或设置一个 **String** 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaLocal**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

有关 A1 样式表示法的详细信息，请参阅 <span style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;background-color:#FFFFFF;">操作方法：使用 A1 表示法引用单元格和区域</span> 。

### AxisTitle.FormulaR1C1

获取或设置一个 **String** 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.FormulaR1C1Local

获取或设置一个 **String** 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1Local**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Height

返回对象的高度（以磅为单位）。只读。

**语法**

**express.Height**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.HorizontalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

此属性的值可设为以下常量之一：

|                   |
|-------------------|
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlLeft**        |
| **xlRight**       |

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### AxisTitle.IncludeInLayout

如果在确定图表布局时轴标题将占用图表布局空间，则为 **True** 。默认值是 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.IncludeInLayout**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Left

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Orientation

返回或设置一个 **Variant** 值，它代表文本方向。

**语法**

**express.Orientation**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Position

返回或设置图表上轴标题的位置。可读/写 **XlChartElementPosition** 类型。

**语法**

**express.Position**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.ReadingOrder

返回或设置指定对象的阅读次序。可为以下常量之一： **xlRTL** （从右到左）、 **xlLTR** （从左到右）或 **xlContext** 。 **Long** 类型，可读写。

**语法**

**express.ReadingOrder**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.Text

返回或设置指定对象中的文本。 **String** 型，可读写。

**语法**

**express.Text**

\*express是一个代表 AxisTitle 对象的变量。

**示例**

``` JavaScript
/*此示例设置 Chart1 中的分类轴标题文本。*/
function test()
{
  Application.Charts.Item("Chart1").Axes(xlCategory).HasTitle = true
  Application.Charts.Item("Chart1").Axes(xlCategory).AxisTitle.Text = "Month"
}
```

### AxisTitle.Top

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 AxisTitle 对象的变量。

### AxisTitle.VerticalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

此属性的值可设为以下常量之一：

|                   |
|-------------------|
| **xlBottom**      |
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlTop**         |

### AxisTitle.Width

返回对象的宽度（以磅为单位）。只读。

**语法**

**express.Width**

\*express是一个代表 AxisTitle 对象的变量。

**说明**

返回值 Double

> WPS 开发文档： [WPS 基础接口/表格 API 参考/AxisTitle](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Borders 对象

## 简介

由四个 Border 对象组成的集合，它们分别代表 Range 或 Style 对象的四个边框。

## 方法

| 序号 | 名称                  | 说明                                                     |
|:----:|:----------------------|:---------------------------------------------------------|
|  1   | [Item](#Borders.Item) | 返回一个 Border 对象，它代表单元格区域或样式的边框之一。 |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Borders.Application)   | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Color](#Borders.Color)               | 返回或设置对象的主要颜色，如注释部分中的表格所示。 Variant 型，可读写。                                                                                                                                                         |
|  3   | [ColorIndex](#Borders.ColorIndex)     | 返回或设置一个 Variant 值，它代表全部四条边框的颜色。                                                                                                                                                                           |
|  4   | [Count](#Borders.Count)               | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  5   | [Creator](#Borders.Creator)           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  6   | [LineStyle](#Borders.LineStyle)       | 返回或设置边框的线型。 XlLineStyle 、 xlGray25 、 xlGray50 、 xlGray75 或 xlAutomatic 类型，可读写。                                                                                                                            |
|  7   | [Parent](#Borders.Parent)             | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  8   | [ThemeColor](#Borders.ThemeColor)     | 返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 Variant 类型。                                                                                                                                      |
|  9   | [TintAndShade](#Borders.TintAndShade) | 返回或设置一个 Single ，使颜色变深或变浅。                                                                                                                                                                                      |
|  10  | [Value](#Borders.Value)               | 等价于 Borders.LineStyle                                                                                                                                                                                                        |
|  11  | [Weight](#Borders.Weight)             | 返回或设置一个 XlBorderWeight 值，它代表边框的粗细。                                                                                                                                                                            |

## 成员方法

### Borders.Item

返回一个 **Border** 对象，它代表单元格区域或样式的边框之一。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Borders 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型       | 说明                        |
|---------|-----------|----------------|-----------------------------|
| *Index* | 必选      | XlBordersIndex | XlBordersIndex 的常量之一。 |

**说明**

|                                                           |
|-----------------------------------------------------------|
| **XlBordersIndex** 可为下列 **XlBordersIndex** 常量之一。 |
| **xlDiagonalDown**                                        |
| **xlDiagonalUp**                                          |
| **xlEdgeBottom**                                          |
| **xlEdgeLeft**                                            |
| **xlEdgeRight**                                           |
| **xlEdgeTop**                                             |
| **xlInsideHorizontal**                                    |
| **xlInsideVertical**                                      |

**示例**

``` JavaScript
/* 下例设置单元格区域 A1:G1 的底部边界的颜色*/
Application.Worksheets.Item("Sheet1").Range("a1:g1").Borders.Item(xlEdgeBottom).Color = (255, 0, 0)
```

## 成员属性

### Borders.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Borders 对象的变量。

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

### Borders.Color

返回或设置对象的主要颜色，如注释部分中的表格所示。 **Variant** 型，可读写。

**语法**

**express.Color**

\*express是一个代表 Borders 对象的变量。

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
Application.Charts.Item("Chart1").Axes(xlValue).TickLabels.Font.Color = (0, 255, 0)
```

### Borders.ColorIndex

返回或设置一个 **Variant** 值，它代表全部四条边框的颜色。

**语法**

**express.ColorIndex**

\*express是一个代表 Borders 对象的变量。

**说明**

如果全部四条边框不是同一种颜色，此属性返回 **Null** 。

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 **XlColorIndex** 常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

### Borders.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Borders 对象的变量。

### Borders.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Borders 对象的变量。

### Borders.LineStyle

返回或设置边框的线型。 **XlLineStyle** 、 **xlGray25** 、 **xlGray50** 、 **xlGray75** 或 **xlAutomatic** 类型，可读写。

**语法**

**express.LineStyle**

\*express是一个代表 Borders 对象的变量。

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

### Borders.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Borders 对象的变量。

### Borders.ThemeColor

返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。

**语法**

**express.ThemeColor**

\*express是一个代表 Borders 对象的变量。

**说明**

如果对象当前未应用主题颜色，试图访问其主题颜色则会引起无效请求运行时错误。

### Borders.TintAndShade

返回或设置一个 **Single** ，使颜色变深或变浅。

**语法**

**express.TintAndShade**

\*express是一个代表 Borders 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

### Borders.Value

等价于 Borders.LineStyle

**语法**

**express.Value**

\*express是一个代表 Borders 对象的变量。

### Borders.Weight

返回或设置一个 **XlBorderWeight** 值，它代表边框的粗细。

**语法**

**express.Weight**

\*express是一个代表 Borders 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Borders](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Comment 对象

## 简介

代表单元格批注。

## 说明

Comment 对象是 Comments 集合的成员。

使用 Comment 属性可返回 Comment 对象。下例更改单元格 E5 中的批注文本。

``` JavaScript
Application.Worksheets.Item(1).Range("E5").Comment.Text("reviewed on " + Date())
```

使用 Comments ( *index* )（其中 *index* 为批注号）可返回 Comments 集合中的单条批注。下例隐藏第一张工作表中的第二条批注。

``` JavaScript
Application.Worksheets.Item(1).Comments.Item(2).Visible = false
```

使用 AddComment 方法可在区域内添加批注。下例在第一张工作表的单元格 E5 中添加批注。

``` JavaScript
function test() {
    let myComment = Application.Worksheets.Item(1).Range("E5").AddComment()  
    myComment.Visible = false  
    myComment.Text("reviewed on " + Date()) 
} 
```

## 方法

| 序号 | 名称                          | 说明                                          |
|:----:|:------------------------------|:----------------------------------------------|
|  1   | [Delete](#Comment.Delete)     | 删除对象。                                    |
|  2   | [Next](#Comment.Next)         | 返回一个 Comment 对象，该对象代表下一条批注。 |
|  3   | [Previous](#Comment.Previous) | 返回一个 Comment 对象，该对象代表前一条批注。 |
|  4   | [Text](#Comment.Text)         | 设置批注文本，返回 String值。                 |

## 属性

| 序号 | 名称                                | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Comment.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Author](#Comment.Author)           | 返回或设置批注的作者。 String 类型，只读。                                                                                                                                                                                      |
|  3   | [Creator](#Comment.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Comment.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  5   | [Shape](#Comment.Shape)             | 返回一个 Shape 对象，它代表连接到指定批注的形状。                                                                                                                                                                               |
|  6   | [Visible](#Comment.Visible)         | 返回或设置一个 Boolean 值，它确定对象是否可见。可读写。                                                                                                                                                                         |

## 成员方法

### Comment.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Comment 对象的变量。

### Comment.Next

返回一个 **Comment** 对象，该对象代表下一条批注。

**语法**

**express.Next ()**

\*express是一个代表 Comment 对象的变量。

**说明**

本方法仅对单张工作表有效。对工作表中最后一条批注使用本方法可返回 **Null** （不是下一张工作表的第一条批注）。

**示例**

``` JavaScript
/* 本示例隐藏下一条批注*/
/* 请在不带现有批注的新工作簿中进行测试。若要清除工作簿中的所有批注，请在即时窗格中使用 Selection.SpecialCells(xlCellTypeComments).delete*/
//Sets up the comments
function test() {
    for(let xNum = 1;xNum <= 10;xNum++) {
        Application.Range("A" + xNum).AddComment()
        Application.Range("A" + xNum).Comment.Text("Comment " + xNum)
    }

    alert("Comments created... A1:A10")

    //Deletes every second comment in the A1:A10 range
    for(let yNum = 1;yNum <= 10;yNum = yNum + 2) {
        Application.Range("A" + yNum).Comment.Next().Shape.Select(true)
        Application.Selection.Delete()
    }

    alert("Deleted every second comment")
}
```

### Comment.Previous

返回一个 **Comment** 对象，该对象代表前一条批注。

**语法**

**express.Previous ()**

\*express是一个代表 Comment 对象的变量。

**说明**

本方法仅对单张工作表有效。对工作表中第一条批注使用本方法可返回 **Null** （不是前一张工作表的最后一条批注）。

**示例**

``` JavaScript
/* 本示例隐藏前一条批注*/
/* 请在不带现有批注的新工作簿中进行测试。若要清除工作簿中的所有批注，请在即时窗格中使用Selection.SpecialCells(xlCellTypeComments).Delete*/
//Sets up the comments
function test(){
    //Sets up the comments
    for(let xNum = 1;xNum <= 10;xNum++) {
        Application.Range("A" + xNum).AddComment()
        Application.Range("A" + xNum).Comment.Text("Comment " + xNum)
    }

    alert("Comments created... A1:A10")

    //Deletes every second comment in the A1:A10 range
    for(let yNum = 10;yNum >= 1;yNum = yNum - 2) {
        Range("A" + yNum).Comment.Previous().Shape.Select(true)
        Selection.Delete()
    }

    alert("Deleted every second comment")
}
```

### Comment.Text

设置批注文本，返回 String值。

**语法**

**express.Text ( *Text* , *Start* , *Overwrite* )**

\*express是一个代表 Comment 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                         |
|-------------|-----------|----------|------------------------------------------------------------------------------|
| *Text*      | 可选      | Variant  | 要添加的文本。                                                               |
| *Start*     | 可选      | Variant  | 所添加文本的起始位置（字符数）。如果省略此参数，则删除批注中的所有现有文字。 |
| *Overwrite* | 可选      | Variant  | 如果为 True，则覆盖现有文件。默认值是 False（插入文本）。                    |

## 成员属性

### Comment.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Comment 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
function test(){
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### Comment.Author

返回或设置批注的作者。 **String** 类型，只读。

**语法**

**express.Author**

\*express是一个代表 Comment 对象的变量。

**示例**

``` JavaScript
/* 本示例删除活动工作表中由 Jean Selva 添加的所有批注*/
function test(){
    let myComments = Application.ActiveSheet.Comments
    for(let c = 1;c <= myComments.Count;c++) {
        if(myComments.Item(c).Author == "Jean Selva") {
            myComments.Item(c).Delete()
        }
    }
}
```

### Comment.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Comment 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Comment.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Comment 对象的变量。

### Comment.Shape

返回一个 **Shape** 对象，它代表连接到指定批注的形状。

**语法**

**express.Shape**

\*express是一个代表 Comment 对象的变量。

**示例**

``` JavaScript
/* 本示例选定活动工作表上的第二条批注*/
/* Ensure that the comments are not hidden. On the Review Tab, choose Comments, Show All Comments*/
Application.ActiveSheet.Comments.Item(2).Shape.Select()
```

### Comment.Visible

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 Comment 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Comment](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DataLabel 对象

## 简介

代表图表数据点或趋势线上的数据标签。

## 说明

在数据系列上， DataLabel 对象是 DataLabels 集合的成员。 DataLabels 集合包含每个数据点的 DataLabel 对象。对于没有可定义数据点的数据系列（如面积图数据系列）， DataLabels 集合包含单个 DataLabel 对象。

使用 DataLabels ( *index* )（其中 *index* 为数据标签的索引号）可返回单个 DataLabel 对象。

``` JavaScript
/*下例在第一张工作表上嵌入的第一个图表上，设置第一个数据系列中的第五个数据标签的数字格式。*/
Application.Worksheets.Item(1).ChartObjects(1).Chart 
    .SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
```

使用 DataLabel 属性可返回单个数据点的 DataLabel 对象。

``` JavaScript
/*下例打开名为“Chart1”的图表工作表上第一个数据系列中第二个数据点的数据标签，并将数据标签文本设置为“Saturday”。*/
function test(){
let myChart = Application.Charts.Item("Chart1")
    let myPoint = myChart.SeriesCollection(1).Points(2)
        myPoint.HasDataLabel = true
        myPoint.DataLabel.Text = "Saturday"
}
```

在趋势线上， DataLabel 返回与趋势线一起显示的文本。这些文本可能是公式、R 平方值或两者均有（如果两者都出现）。

``` JavaScript
/*下例将趋势线文本设置为仅显示公式，并将数据标签文本置于工作表“Sheet1”上的单元格 A1 中。*/
function test(){
let myChartline = Application.Charts.Item("chart1").SeriesCollection(1).Trendlines(1)
    myChartline.DisplayRSquared = false
    myChartline.DisplayEquation = true
    Worksheets.Item("sheet1").Range("a1").Value2 = myChartline.DataLabel.Text
}
```

## 方法

| 序号 | 名称                        | 说明       |
|:----:|:----------------------------|:-----------|
|  1   | [Delete](#DataLabel.Delete) | 删除对象。 |
|  2   | [Select](#DataLabel.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                                  | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#DataLabel.Application)                 | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [AutoText](#DataLabel.AutoText)                       | 如果对象会根据内容自动生成合适的文字，则为 True 。 Boolean 类型，可读写。                                                                                                                                                       |
|  3   | [Caption](#DataLabel.Caption)                         | 返回或设置一个 String 值，它代表数据标签文本。                                                                                                                                                                                  |
|  4   | [Characters](#DataLabel.Characters)                   | 返回 Characters 对象，它代表对象文本内某个区域的字符。使用 Characters 对象可为文本字符串内的字符设置格式。                                                                                                                      |
|  5   | [Creator](#DataLabel.Creator)                         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  6   | [Format](#DataLabel.Format)                           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  7   | [Formula](#DataLabel.Formula)                         | 获取或设置一个 String 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。                                                                                                                                               |
|  8   | [FormulaLocal](#DataLabel.FormulaLocal)               | 获取或设置一个 String 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。                                                                                                                                         |
|  9   | [FormulaR1C1](#DataLabel.FormulaR1C1)                 | 获取或设置一个 String 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。                                                                                                                                             |
|  10  | [FormulaR1C1Local](#DataLabel.FormulaR1C1Local)       | 获取或设置一个 String 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。                                                                                                                                       |
|  11  | [Height](#DataLabel.Height)                           | 返回对象的高度（以磅为单位）。只读。                                                                                                                                                                                            |
|  12  | [HorizontalAlignment](#DataLabel.HorizontalAlignment) | 返回或设置一个 Variant 值，它代表指定对象的水平对齐方式。                                                                                                                                                                       |
|  13  | [Left](#DataLabel.Left)                               | 返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                       |
|  14  | [Name](#DataLabel.Name)                               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  15  | [NumberFormat](#DataLabel.NumberFormat)               | 返回或设置一个 String 值，它代表对象的格式代码。                                                                                                                                                                                |
|  16  | [NumberFormatLinked](#DataLabel.NumberFormatLinked)   | 如果指定数字格式指向单元格（以便当单元格的格式更改时数据标签的格式也作相应的改动），则为 True 。 Boolean 类型，可读写。                                                                                                         |
|  17  | [NumberFormatLocal](#DataLabel.NumberFormatLocal)     | 以采用用户语言字符串的形式返回或设置一个 Variant 值，它代表对象的格式代码。                                                                                                                                                     |
|  18  | [Orientation](#DataLabel.Orientation)                 | 返回或设置一个 Variant 值，它代表文本方向。                                                                                                                                                                                     |
|  19  | [Parent](#DataLabel.Parent)                           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  20  | [Position](#DataLabel.Position)                       | 返回或设置一个 XlDataLabelPosition 值，它代表数据标签的位置。                                                                                                                                                                   |
|  21  | [ReadingOrder](#DataLabel.ReadingOrder)               | 返回或设置指定对象的阅读次序。可为以下常量之一： xlRTL （从右到左）、 xlLTR （从左到右）或 xlContext 。 Long 类型，可读写。                                                                                                     |
|  22  | [Separator](#DataLabel.Separator)                     | 设置或返回一个 Variant 值，它代表用于图表中数据标签的分隔符。可读/写。                                                                                                                                                          |
|  23  | [Shadow](#DataLabel.Shadow)                           | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                               |
|  24  | [ShowBubbleSize](#DataLabel.ShowBubbleSize)           | 如果为 True ，则在图表中显示数据标签的气泡的大小。如果为 False ，则隐藏气泡的大小。 Boolean 类型，可读写。                                                                                                                      |
|  25  | [ShowCategoryName](#DataLabel.ShowCategoryName)       | 如果为 True ，则在图表中显示数据标签的分类名称。如果为 False ，则隐藏该分类的名称。 Boolean 类型，可读写。                                                                                                                      |
|  26  | [ShowLegendKey](#DataLabel.ShowLegendKey)             | 如果数据标签图例项标示可见，则为 True 。 Boolean 类型，可读写。                                                                                                                                                                 |
|  27  | [ShowPercentage](#DataLabel.ShowPercentage)           | 如果为 True ，则在图表中显示数据标签的百分比数值。如果为 False ，则隐藏该百分比数值。 Boolean 类型，可读写。                                                                                                                    |
|  28  | [ShowSeriesName](#DataLabel.ShowSeriesName)           | 返回或设置一个 Boolean 值，它指明图表上数据标签的系列名称显示方式。如果为 True ，则显示系列名称。如果为 False ，则隐藏系列名称。可读写。                                                                                        |
|  29  | [Text](#DataLabel.Text)                               | 返回或设置指定对象中的文本。 String 型，可读写。                                                                                                                                                                                |
|  30  | [Top](#DataLabel.Top)                                 | 返回或设置一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                      |
|  31  | [VerticalAlignment](#DataLabel.VerticalAlignment)     | 返回或设置一个 Variant 值，它代表指定对象的垂直对齐方式。                                                                                                                                                                       |
|  32  | [Width](#DataLabel.Width)                             | 返回对象的宽度（以磅为单位）。只读。                                                                                                                                                                                            |

## 成员方法

### DataLabel.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 DataLabel 对象的变量。

**返回值**

Variant

### DataLabel.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 DataLabel 对象的变量。

**返回值**

Variant

## 成员属性

### DataLabel.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 DataLabel 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
let myObject = Application.ActiveWorkbook
if(myObject.Application.Value == "ET") {
    alert("This is an ET Application object.")
}
else {
    alert("This is not an ET Application object.")
}
}
```

### DataLabel.AutoText

如果对象会根据内容自动生成合适的文字，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoText**

\*express是一个代表 DataLabel 对象的变量。

**示例**

``` JavaScript
/*此示例对 Chart1 的第一个数据序列的数据标签进行设置，自动生成合适的文字。*/
Application.Charts.Item("Chart1").SeriesCollection(1).DataLabels.AutoText = true
```

### DataLabel.Caption

返回或设置一个 **String** 值，它代表数据标签文本。

**语法**

**express.Caption**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.Characters

返回 **Characters** 对象，它代表对象文本内某个区域的字符。使用 **Characters** 对象可为文本字符串内的字符设置格式。

**语法**

**express.Characters**

\*express是一个代表 DataLabel 对象的变量。

**说明**

**Characters** 对象不是集合。

### DataLabel.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 DataLabel 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### DataLabel.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.Formula

获取或设置一个 **String** 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.Formula**

\*express是一个代表 DataLabel 对象的变量。

**说明**

有关 A1 样式表示法的详细信息，请参阅 操作方法：使用 A1 表示法引用单元格和区域 。

### DataLabel.FormulaLocal

获取或设置一个 **String** 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaLocal**

\*express是一个代表 DataLabel 对象的变量。

**说明**

有关 A1 样式表示法的详细信息，请参阅 操作方法：使用 A1 表示法引用单元格和区域 。

### DataLabel.FormulaR1C1

获取或设置一个 **String** 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.FormulaR1C1Local

获取或设置一个 **String** 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1Local**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.Height

返回对象的高度（以磅为单位）。只读。

**语法**

**express.Height**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.HorizontalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

\*express是一个代表 DataLabel 对象的变量。

**说明**

此属性的值可设为以下常量之一：

|                   |
|-------------------|
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlLeft**        |
| **xlRight**       |

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### DataLabel.Left

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.NumberFormat

返回或设置一个 **String** 值，它代表对象的格式代码。

**语法**

**express.NumberFormat**

\*express是一个代表 DataLabel 对象的变量。

**说明**

格式代码与 **“设置单元格格式”** 对话框中的 **“格式代码”** 选项是同一个字符串。 **Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

### DataLabel.NumberFormatLinked

如果指定数字格式指向单元格（以便当单元格的格式更改时数据标签的格式也作相应的改动），则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.NumberFormatLinked**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.NumberFormatLocal

以采用用户语言字符串的形式返回或设置一个 **Variant** 值，它代表对象的格式代码。

**语法**

**express.NumberFormatLocal**

\*express是一个代表 DataLabel 对象的变量。

**说明**

**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

### DataLabel.Orientation

返回或设置一个 **Variant** 值，它代表文本方向。

**语法**

**express.Orientation**

\*express是一个代表 DataLabel 对象的变量。

**说明**

此属性的值可设为

### DataLabel.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.Position

返回或设置一个 **XlDataLabelPosition** 值，它代表数据标签的位置。

**语法**

**express.Position**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.ReadingOrder

返回或设置指定对象的阅读次序。可为以下常量之一： **xlRTL** （从右到左）、 **xlLTR** （从左到右）或 **xlContext** 。 **Long** 类型，可读写。

**语法**

**express.ReadingOrder**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.Separator

设置或返回一个 **Variant** 值，它代表用于图表中数据标签的分隔符。可读/写。

**语法**

**express.Separator**

\*express是一个代表 DataLabel 对象的变量。

**说明**

如果使用字符串，则用一个字符串作为分隔符。如果使用 **xlDataLabelSeparatorDefault** (= 1)，则用默认数据标签分隔符，即逗号或换行符，具体取决于数据标签。

当返回值“1”时，表明用户未更改默认的分隔符逗号“,”。您也可通过传递值“1”将分隔符更改回默认值。

以程序方式访问数据标签之前必须先激活图表，否则将出现运行错误。

**示例**

``` JavaScript
/*此示例将第一个图表上的第一个系列的数据标签分隔符设置为分号。此示例假定活动工作表上存在图表。*/
 Application.ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1)
        .DataLabels.Separator = ";"
```

### DataLabel.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.ShowBubbleSize

如果为 **True** ，则在图表中显示数据标签的气泡的大小。如果为 **False** ，则隐藏气泡的大小。 **Boolean** 类型，可读写。

**语法**

**express.ShowBubbleSize**

\*express是一个代表 DataLabel 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

``` JavaScript
/*此示例显示第一个图表上第一个系列的数据标签的气泡大小。此示例假定活动工作表上存在图表。*/
function test(){
  Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1)
        .DataLabels.ShowBubbleSize = true

}
```

### DataLabel.ShowCategoryName

如果为 **True** ，则在图表中显示数据标签的分类名称。如果为 **False** ，则隐藏该分类的名称。 **Boolean** 类型，可读写。

**语法**

**express.ShowCategoryName**

\*express是一个代表 DataLabel 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

``` JavaScript
/*此示例显示第一个图表上第一个系列的数据标签的分类名称。此示例假定活动工作表上存在图表。*/
function test(){
 Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1) 
        .DataLabels.ShowCategoryName = true
}
```

### DataLabel.ShowLegendKey

如果数据标签图例项标示可见，则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowLegendKey**

\*express是一个代表 DataLabel 对象的变量。

**示例**

``` JavaScript
/*此示例设置 Chart1 中系列一的数据标签，以显示数值和图例标示.*/
function test(){
let myDataLabel = Application.Charts.Item("Chart1").SeriesCollection(1).DataLabels
    myDataLabel.ShowLegendKey = true
    myDataLabel.Type = xlShowValue
}
```

### DataLabel.ShowPercentage

如果为 **True** ，则在图表中显示数据标签的百分比数值。如果为 **False** ，则隐藏该百分比数值。 **Boolean** 类型，可读写。

**语法**

**express.ShowPercentage**

\*express是一个代表 DataLabel 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

``` JavaScript
/*此示例显示了第一个图表上第一个数据系列的数据标签的百分比数值。此示例假定活动工作表上存在图表。*/
function test(){
Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1) 
        .DataLabels.ShowPercentage = true
}
```

### DataLabel.ShowSeriesName

返回或设置一个 **Boolean** 值，它指明图表上数据标签的系列名称显示方式。如果为 **True** ，则显示系列名称。如果为 **False** ，则隐藏系列名称。可读写。

**语法**

**express.ShowSeriesName**

\*express是一个代表 DataLabel 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

``` JavaScript
/*此示例显示了第一个图表上第一个数据系列的数据标签的系列名称。此示例假定活动工作表上存在图表。*/
function test(){
  Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1) 
        .DataLabels.ShowSeriesName = true
}
```

### DataLabel.Text

返回或设置指定对象中的文本。 **String** 型，可读写。

**语法**

**express.Text**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.Top

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 DataLabel 对象的变量。

### DataLabel.VerticalAlignment

返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

\*express是一个代表 DataLabel 对象的变量。

**说明**

此属性的值可设为以下常量之一：

|                   |
|-------------------|
| **xlBottom**      |
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlTop**         |

### DataLabel.Width

返回对象的宽度（以磅为单位）。只读。

**语法**

**express.Width**

\*express是一个代表 DataLabel 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/DataLabel](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FormatColor 对象

## 简介

代表为色阶条件格式阈值指定的填充色或数据条条件格式的条形颜色。

## 说明

您可以通过传递 Color 属性中的 RGB 值来选择颜色，或者通过使用 ThemeColor 属性在主题调色板中编制索引来指定颜色。

以下代码示例创建了一个数字范围，然后将双色色阶条件格式规则应用于该范围。然后通过在 ColorScaleCriteria 集合中编制索引来设置单独的条件，从而指定最小阈值的颜色为红色，最大阈值的颜色为蓝色。

``` JavaScript
function test() {
    //Fill cells with sample data from 1 to 10
    let sheet = ActiveSheet
    sheet.Range("C1").Value2 = 1
    sheet.Range("C2").Value2 = 2
    sheet.Range("C1:C2").AutoFill(Range("C1:C10"))
    
    Range("C1:C10").Select()
    
    //Create a two-color ColorScale object for the created sample data range
    let cfColorScale = Selection.FormatConditions.AddColorScale(2)
    
    //Set the minimum threshold to red and maximum threshold to blue
    cfColorScale.ColorScaleCriteria.Item(1).FormatColor.Color = (255, 0, 0)
    cfColorScale.ColorScaleCriteria.Item(2).FormatColor.Color = (0, 0, 255)
}
```

## 属性

| 序号 | 名称                                      | 说明                                                                                                                                                               |
|:----:|:------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FormatColor.Application)   | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Color](#FormatColor.Color)               | 返回或设置与数据条或色阶条件格式规则的阈值关联的填充色。                                                                                                           |
|  3   | [ColorIndex](#FormatColor.ColorIndex)     | 返回或设置 XlColorIndex 枚举的常量之一，指定是否以当前调色板中颜色的索引值形式表示填充色。                                                                         |
|  4   | [Creator](#FormatColor.Creator)           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  5   | [Parent](#FormatColor.Parent)             | 返回指定对象的父对象。只读。                                                                                                                                       |
|  6   | [ThemeColor](#FormatColor.ThemeColor)     | 返回或设置一个 XlThemeColor 枚举常量，该值指定数据条或色阶条件格式阈值中所用的主题颜色。                                                                           |
|  7   | [TintAndShade](#FormatColor.TintAndShade) | 返回或设置一个 Single 类型的值，该值使用于数据条或色阶条件格式规则的单元格的填充颜色变浅或变深。                                                                   |

## 成员属性

### FormatColor.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 FormatColor 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

**示例**

``` JavaScript
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

### FormatColor.Color

返回或设置与数据条或色阶条件格式规则的阈值关联的填充色。

**语法**

**express.Color**

\*express是一个代表 FormatColor 对象的变量。

**说明**

格式颜色以 RGB 函数表示。例如，要将颜色设置为红色，请使用 ` RGB(255,0,0) ` 。

### FormatColor.ColorIndex

返回或设置 **XlColorIndex** 枚举的常量之一，指定是否以当前调色板中颜色的索引值形式表示填充色。

**语法**

**express.ColorIndex**

\*express是一个代表 FormatColor 对象的变量。

**说明**

此属性用于色阶或数据条条件格式规则中的每个阈值。

### FormatColor.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FormatColor 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### FormatColor.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FormatColor 对象的变量。

### FormatColor.ThemeColor

返回或设置一个 **XlThemeColor** 枚举常量，该值指定数据条或色阶条件格式阈值中所用的主题颜色。

**语法**

**express.ThemeColor**

\*express是一个代表 FormatColor 对象的变量。

### FormatColor.TintAndShade

返回或设置一个 **Single** 类型的值，该值使用于数据条或色阶条件格式规则的单元格的填充颜色变浅或变深。

**语法**

**express.TintAndShade**

\*express是一个代表 FormatColor 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/FormatColor](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Gridlines 对象

## 简介

代表图表坐标轴的主要和次要网格线。

## 说明

网格线延伸图表坐标轴上的刻度线，以便更容易地分辨与数据标志相关联的数值。此对象不是集合。同时也没有代表单个网格线的对象；或者打开坐标轴上所有的网格线，或者将其全部关闭。

使用 MajorGridlines 属性可返回 GridLines 对象，该对象代表坐标轴的主要网格线。使用 MinorGridlines 属性可返回 GridLines 对象，该对象代表次要网格线。可以同时返回主要网格线和次要网格线。

``` JavaScript
/*下例打开图表工作表 1 上分类轴的主要网格线，并将网格线设置为蓝色虚线。*/
function test(){
let chart1 = Application.Charts.Item("Chart1").Axes(xlCategory)
chart1.HasMajorGridlines = true
chart1.MajorGridlines.Border.Color = (0, 0, 255)
chart1.MajorGridlines.Border.LineStyle = xlDash
}
```

## 方法

| 序号 | 名称                        | 说明       |
|:----:|:----------------------------|:-----------|
|  1   | [Delete](#Gridlines.Delete) | 删除对象。 |
|  2   | [Select](#Gridlines.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Gridlines.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Border](#Gridlines.Border)           | 返回一个 Border 对象，它代表对象的边框。                                                                                                                                                                                        |
|  3   | [Creator](#Gridlines.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Format](#Gridlines.Format)           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  5   | [Name](#Gridlines.Name)               | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  6   | [Parent](#Gridlines.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Gridlines.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Gridlines 对象的变量。

**返回值**

Variant

### Gridlines.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 Gridlines 对象的变量。

**返回值**

Variant

## 成员属性

### Gridlines.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Gridlines 对象的变量。

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

### Gridlines.Border

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

\*express是一个代表 Gridlines 对象的变量。

### Gridlines.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Gridlines 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Gridlines.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 Gridlines 对象的变量。

### Gridlines.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Gridlines 对象的变量。

### Gridlines.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Gridlines 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Gridlines](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# HPageBreak 对象

## 简介

代表水平分页符。

## 说明

HPageBreak 对象是 HPageBreaks 集合的成员。

| 注释                                 |
|--------------------------------------|
| 每张工作表最多有 1026 个水平分页符。 |

使用 HPage Breaks ( *index* )（其中 *index* 是分页符的索引号）可返回一个 HPageBreak 对象。下例更改第一个水平分页符的位置。

``` JavaScript
Application.Worksheets.Item(1).HPageBreaks.Item(1).Location = Worksheets.Item(1).Range("e5")
```

## 方法

| 序号 | 名称                           | 说明                       |
|:----:|:-------------------------------|:---------------------------|
|  1   | [Delete](#HPageBreak.Delete)   | 删除对象。                 |
|  2   | [DragOff](#HPageBreak.DragOff) | 将一个分页符拖出打印区域。 |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#HPageBreak.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#HPageBreak.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Extent](#HPageBreak.Extent)           | 返回指定分页符的类型：全屏或仅在打印区域内。可为以下 任一 XlPageBreakExtent 常量： xlPageBreakFull 或 xlPageBreakPartial 。 Long 类型，只读。                                                                                   |
|  4   | [Location](#HPageBreak.Location)       | 返回或设置定义分页符位置的单元格（ Range 对象）。水平分页符与定位单元格的上边缘对齐；垂直分页符与定位单元格的左边缘对齐。 Range 类型，可读写。                                                                                  |
|  5   | [Parent](#HPageBreak.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  6   | [Type](#HPageBreak.Type)               | 返回或设置一个 XlPage Break 值，它代表分页符类型。                                                                                                                                                                              |

## 成员方法

### HPageBreak.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 HPageBreak 对象的变量。

### HPageBreak.DragOff

将一个分页符拖出打印区域。

**语法**

**express.DragOff ( *Direction* , *RegionIndex* )**

\*express是一个代表 HPageBreak 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型    | 说明                                                                                                                                                           |
|---------------|-----------|-------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Direction*   | 必选      | XlDirection | 分页符拖动方向。                                                                                                                                               |
| *RegionIndex* | 必选      | Long        | 分页符的打印区域索引（当用户按下鼠标按钮拖动分页符时鼠标指针所在的位置）。如果打印区域是连续的，则只有一个打印区域。如果打印区域不是连续的，则有多个打印区域。 |

**说明**

该方法主要用于宏记录器。使用 **Delete** 方法可在 Visual Basic 中删除分页符。

**示例**

``` JavaScript
/* 本示例将活动工作表的第一个垂直分页符拖出第一个打印区域的右边界，以删除该分页符*/
​Application.ActiveSheet.VPageBreaks.Item(1).DragOff(xlToRight, 1)
```

## 成员属性

### HPageBreak.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 HPageBreak 对象的变量。

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

### HPageBreak.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 HPageBreak 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### HPageBreak.Extent

返回指定分页符的类型：全屏或仅在打印区域内。可为以下 任一 **XlPageBreakExtent** 常量： **xlPageBreakFull** 或 **xlPageBreakPartial** 。 **Long** 类型，只读。

**语法**

**express.Extent**

\*express是一个代表 HPageBreak 对象的变量。

**示例**

``` JavaScript
/* 本示例显示全屏水平分页符和打印区水平分页符的总数*/
function test() {
    let cFull = 0
    let cPartial = 0
    for(let pb = 1; pb <=  Application.Worksheets.Item(1).HPageBreaks.Count; pb++) {
        if( Application.Worksheets.Item(1).HPageBreaks.Item(pb).Extent == xlPageBreakFull) {
            cFull++
        }
        else {
            cPartial++
        }
    }
    alert(cFull + " full-screen page breaks, " + cPartial + " print-area page breaks")
}
```

### HPageBreak.Location

返回或设置定义分页符位置的单元格（ **Range** 对象）。水平分页符与定位单元格的上边缘对齐；垂直分页符与定位单元格的左边缘对齐。 **Range** 类型，可读写。

**语法**

**express.Location**

\*express是一个代表 HPageBreak 对象的变量。

**示例**

``` JavaScript
/* 本示例移动水平分页符的位置*/
Application.Worksheets.Item(1).HPageBreaks.Item(1).Location = Application.Worksheets.Item(1).Range("e5")
```

### HPageBreak.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 HPageBreak 对象的变量。

### HPageBreak.Type

返回或设置一个 **XlPage Break** 值，它代表分页符类型。

**语法**

**express.Type**

\*express是一个代表 HPageBreak 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/HPageBreak](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Hyperlinks 对象

## 简介

代表工作表或区域的超链接的集合。

## 说明

每个超链接都由一个 Hyperlink 对象代表。

使用 Hyperlinks 属性可返回 Hyperlinks 集合。下例检查工作表一上的超链接，看是否有包含“Microsoft”一词的链接。

``` JavaScript
function test() {
    for(let h = 1; h <= Application.Worksheets.Item(1).Hyperlinks.Count; h++) {
        if(Application.Worksheets.Item(1).Hyperlinks.Item(h).Name.indexOf("Microsoft") != -1) { 
Application.Worksheets.Item(1).Hyperlinks.Item(h).Follow()
        }
    }
}
```

使用 Add 方法可创建一个超链接并将它添加到 Hyperlinks 集合。下例为单元格 E5 创建一个新的超链接。

``` JavaScript
function test() {
    let hyperlinks2 = Application.Worksheets.Item(1)
    hyperlinks2.Hyperlinks.Add(sheet2.Range("E5"), "http://example.microsoft.com")
} 
```

## 方法

| 序号 | 名称                         | 说明                                                                       |
|:----:|:-----------------------------|:---------------------------------------------------------------------------|
|  1   | [Add](#Hyperlinks.Add)       | 向指定的区域或形状添加超链接。返回 一个 Hyperlink 对象，它代表新的超链接。 |
|  2   | [Delete](#Hyperlinks.Delete) | 删除对象。                                                                 |
|  3   | [Item](#Hyperlinks.Item)     | 从集合中返回一个对象。                                                     |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Hyperlinks.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Hyperlinks.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Hyperlinks.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Hyperlinks.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### Hyperlinks.Add

向指定的区域或形状添加超链接。返回 一个 **Hyperlink** 对象，它代表新的超链接。

**语法**

**express.Add ( *Anchor* , *Address* , *SubAddress* , *ScreenTip* , *TextToDisplay* )**

\*express是一个代表 Hyperlinks 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                         |
|-----------------|-----------|----------|----------------------------------------------|
| *Anchor*        | 必选      | Object   | 超链接的位置。可为 Range 或 Shape 对象。     |
| *Address*       | 必选      | String   | 超链接的地址。                               |
| *SubAddress*    | 可选      | Variant  | 超链接的子地址。                             |
| *ScreenTip*     | 可选      | Variant  | 当鼠标指针停留在超链接上时所显示的屏幕提示。 |
| *TextToDisplay* | 可选      | Variant  | 要显示的超链接的文本。                       |

**说明**

指定 **TextToDisplay** 参数时，文本必须是字符串。

**示例**

``` JavaScript
function test() {
    /* 向单元格 A5 添加超链接。*/
    let sheet2 = Application.Worksheets.Item(1)
    sheet2.Hyperlinks.Add(sheet2.Range("a5"), "http://example.microsoft.com", null, "Microsoft Web Site", "Microsoft")

    /* 向单元格 A5 添加一个电子邮件超链接。*/
    let sheet2 = Application.Worksheets.Item(1)
    sheet2.Hyperlinks.Add(sheet2.Range("a5"), "mailto:someone@example.com?subject=hello", null, "Write us today", "Support")
}
```

### Hyperlinks.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Hyperlinks 对象的变量。

### Hyperlinks.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Hyperlinks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**示例**

``` JavaScript
/*下例激活 E5 单元格的第一个超链接。*/
Application.Worksheets.Item(1).Range("E5").Hyperlinks.Item(1).Follow()
```

## 成员属性

### Hyperlinks.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Hyperlinks 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
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

### Hyperlinks.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Hyperlinks 对象的变量。

### Hyperlinks.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Hyperlinks 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Hyperlinks.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Hyperlinks 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Hyperlinks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# LeaderLines 对象

## 简介

代表图表的引导线。引导线将数据标签连接到数据点。

## 说明

该对象不是一个集合；没有表示单个引导线的对象。

此对象只适用于饼图。

使用 LeaderLines 属性可返回 LeaderLines 对象。下例图表一上的数据系列一添加数据标签和蓝色引导线。如果看不到引导线，此示例代码将会失败。在这种情况下，您可以手动将一个数据标签从饼图上拖离，以便显示出引导线。

``` JavaScript
function test(){
let chartObjects2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
chartObjects2.HasDataLabels = true
chartObjects2.DataLabels.Position = xlLabelPositionBestFit
chartObjects2.HasLeaderLines = true
chartObjects2.LeaderLines.Border.ColorIndex = 5
}
```

## 方法

| 序号 | 名称                          | 说明       |
|:----:|:------------------------------|:-----------|
|  1   | [Delete](#LeaderLines.Delete) | 删除对象。 |
|  2   | [Select](#LeaderLines.Select) | 选择对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#LeaderLines.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Border](#LeaderLines.Border)           | 返回一个 Border 对象，它代表对象的边框。                                                                                                                                                                                        |
|  3   | [Creator](#LeaderLines.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Format](#LeaderLines.Format)           | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  5   | [Parent](#LeaderLines.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### LeaderLines.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 LeaderLines 对象的变量。

### LeaderLines.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 LeaderLines 对象的变量。

## 成员属性

### LeaderLines.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 LeaderLines 对象的变量。

**示例**

\*本示例显示一条有关创建 myObject 的应用程序的消息。\*/ function test() { let myObject = Application.ActiveWorkbook if (myObject.Application.Value == "ET") { alert("This is an ET Application object.") } else { alert("This is not an ET Application object.") } }

### LeaderLines.Border

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

\*express是一个代表 LeaderLines 对象的变量。

### LeaderLines.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 LeaderLines 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### LeaderLines.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 LeaderLines 对象的变量。

### LeaderLines.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 LeaderLines 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/LeaderLines](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ODBCError 对象

## 简介

代表一个由最近的 ODBC 查询生成的 ODBC 错误。

## 说明

ODBCError 对象是 ODBCErrors 集合的成员。如果指定的 ODBC 查询运行时没有出现错误，则 ODBCErrors 集合为空。集合中错误的索引次序与 ODBC 数据源生成它们时的次序相同。

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ODBCError.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#ODBCError.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [ErrorString](#ODBCError.ErrorString) | 返回一个 String 值，它代表 ODBC 错误字符串。                                                                                                                                                                                    |
|  4   | [Parent](#ODBCError.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  5   | [SqlState](#ODBCError.SqlState)       | 返回 SQL 状态错误。 String 型，只读。                                                                                                                                                                                           |

## 成员属性

### ODBCError.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ODBCError 对象的变量。

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

### ODBCError.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ODBCError 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ODBCError.ErrorString

返回一个 **String** 值，它代表 ODBC 错误字符串。

**语法**

**express.ErrorString**

\*express是一个代表 ODBCError 对象的变量。

### ODBCError.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ODBCError 对象的变量。

### ODBCError.SqlState

返回 SQL 状态错误。 **String** 型，只读。

**语法**

**express.SqlState**

\*express是一个代表 ODBCError 对象的变量。

**说明**

有关特定错误的解释，请参阅 SQL 文档。

**示例**

``` JavaScript
/*此示例刷新查询表一，并显示产生的所有 ODBC 错误。*/
function test(){
　　　　let qts = Worksheets.Item(1).QueryTables.Item(1)
　　　　qts.Refresh()
　　　　let errs = Application.ODBCErrors
    if (errs.Count > 0){
　　　　    let r = qts.Destination.Cells.Item(1)
　　　　    r.Value2 = "The following errors occurred:"
　　　　    let c = 0
　　　　    for(let er = 1; er <= errs.Count; er++){
　　　　        c = c + 1
 　　　　       r.offset.Item(c, 0).value2 = errs.Item(er).ErrorString
 　　　　       r.offset.Item(c, 1).value2 = errs.Item(er).SqlState
 　　　　   }
　　　　}  
　　　　else{
　　　　    MsgBox ("Query complete: all records returned.")
　　　　}
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ODBCError](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Phonetic 对象

## 简介

包含单元格中特定拼音文本串的有关信息。

## 说明

在 ET 97 中，此对象包含指定区域中任意拼音文本的格式属性。

## 属性

| 序号 | 名称                                     | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Alignment](#Phonetic.Alignment)         | 返回或设置一个 Long 值，它代表指定的拼音文本或刻度线标签的对齐方式。                                                                                                                                                            |
|  2   | [Application](#Phonetic.Application)     | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  3   | [CharacterType](#Phonetic.CharacterType) | 返回或设置指定单元格中拼音文本的类型。 XlPhoneticCharacterType 类型，可读写。                                                                                                                                                   |
|  4   | [Creator](#Phonetic.Creator)             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [Font](#Phonetic.Font)                   | 返回一个 Font 对象，它代表指定对象的字体。                                                                                                                                                                                      |
|  6   | [Parent](#Phonetic.Parent)               | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  7   | [Text](#Phonetic.Text)                   | 返回或设置指定对象中的文本。 String 型，可读写。                                                                                                                                                                                |
|  8   | [Visible](#Phonetic.Visible)             | 返回或设置一个 Boolean 值，它确定对象是否可见。可读写。                                                                                                                                                                         |

## 成员属性

### Phonetic.Alignment

返回或设置一个 **Long** 值，它代表指定的拼音文本或刻度线标签的对齐方式。

**语法**

**express.Alignment**

\*express是一个代表 Phonetic 对象的变量。

### Phonetic.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Phonetic 对象的变量。

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

### Phonetic.CharacterType

返回或设置指定单元格中拼音文本的类型。 **XlPhoneticCharacterType** 类型，可读写。

**语法**

**express.CharacterType**

\*express是一个代表 Phonetic 对象的变量。

**示例**

本示例将活动单元格中的第一个拼音文本字符串从 Furigana 更改为 Hiragana。

``` JavaScript
ActiveCell.Phonetics.Item(1).CharacterType = xlHiragana
```

### Phonetic.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Phonetic 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Phonetic.Font

返回一个 **Font** 对象，它代表指定对象的字体。

**语法**

**express.Font**

\*express是一个代表 Phonetic 对象的变量。

### Phonetic.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Phonetic 对象的变量。

### Phonetic.Text

返回或设置指定对象中的文本。 **String** 型，可读写。

**语法**

**express.Text**

\*express是一个代表 Phonetic 对象的变量。

**说明**

对于 **Phonetic** 对象，此属性返回或设置其拼音文本。不能将此属性设为 **Null** 。

### Phonetic.Visible

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 Phonetic 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Phonetic](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotFilter 对象

## 简介

透视筛选应用于 PivotField 对象。

## 说明

由于索引不可靠，因此开发人员可以选择指定引用筛选。 DataField 属性指定值筛选基于的透视字段。

## 方法

| 序号 | 名称                          | 说明                                                   |
|:----:|:------------------------------|:-------------------------------------------------------|
|  1   | [Delete](#PivotFilter.Delete) | 删除筛选并从透视字段和数据透视表的筛选集合中将其删除。 |

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
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.Active">Active</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定透视筛选是否是活动的。只读 Boolean 。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 <span>Application</span> 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.DataCubeField">DataCubeField</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>此属性仅适用于 OLAP 数据透视表，可为值筛选提供作为筛选依据的 值 字段（值区域中的透视字段）。可读写 CubeField 。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.DataField">DataField</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>此属性仅适用于非 OLAP 数据透视表，可为值筛选提供作为筛选依据的 值 字段（值区域中的透视字段）。可读写 PivotField 。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.Description">Description</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>为 PivotFilter 对象提供可选说明。只读 String 。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.FilterType">FilterType</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>指定要应用的筛选器类型。只读 xlPivotFilterType 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.IsMemberPropertyFilter">IsMemberPropertyFilter</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>指定标签筛选器是基于字段的成员属性的 PivotItem 标题还是基于透视字段本身的 PivotItem 标题。只读 Boolean 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.MemberPropertyField">MemberPropertyField</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>此属性指定标签筛选器所基于的成员属性 PivotField。可读/写 PivotField 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.Name">Name</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>此属性提供选项可指定引用筛选。索引值可能发生更改，因此不能依赖于这些值获得准确的引用。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.Order">Order</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>指定应用于整个数据透视表的所有“值”筛选器中筛选器的求值顺序。 Integer 类型，可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定 PivotFilter 对象的父对象。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.PivotField">PivotField</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>指定应用筛选器的透视字段。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.Value1">Value1</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>此属性是由用户提供的用于定义透视字段筛选器的参数。可读/写 Variant 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFilter.Value2">Value2</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>此属性是由用户提供的用于定义透视字段筛选器的参数。可读/写 Variant 类型。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PivotFilter.Delete

删除筛选并从透视字段和数据透视表的筛选集合中将其删除。

**语法**

**express.Delete ()**

\*express是一个代表 PivotFilter 对象的变量。

## 成员属性

### PivotFilter.Active

table { word-break:break-all; }

返回指定透视筛选是否是活动的。只读 **Boolean** 。

**语法**

**express.Active**

\*express是一个代表 PivotFilter 对象的变量。

**说明**

table { word-break:break-all; }

当筛选的透视字段在数据透视表中，且数据透视表更新时会计算该筛选时，此属性返回 **True** 。当筛选的透视字段不在数据透视表中，且对数据透视表计算没有影响时，此属性返回 **False** 。

### PivotFilter.Application

table { word-break:break-all; }

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotFilter 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### PivotFilter.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotFilter 对象的变量。

**说明**

table { word-break:break-all;

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotFilter.DataCubeField

table { word-break:break-all; }

此属性仅适用于 OLAP 数据透视表，可为值筛选提供作为筛选依据的 **值** 字段（值区域中的透视字段）。可读写 **CubeField** 。

**语法**

**express.DataCubeField**

\*express是一个代表 PivotFilter 对象的变量。

### PivotFilter.DataField

table { word-break:break-all; }

此属性仅适用于非 OLAP 数据透视表，可为值筛选提供作为筛选依据的 **值** 字段（值区域中的透视字段）。可读写 **PivotField** 。

**语法**

**express.DataField**

\*express是一个代表 PivotFilter 对象的变量。

### PivotFilter.Description

table { word-break:break-all; }

为 **PivotFilter** 对象提供可选说明。只读 **String** 。

**语法**

**express.Description**

\*express是一个代表 PivotFilter 对象的变量。

**说明**

table { word-break:break-all; }

说明字符串的最大长度为 255 个字符。

### PivotFilter.FilterType

table { word-break:break-all; }

指定要应用的筛选器类型。只读 **xlPivotFilterType** 类型。

**语法**

**express.FilterType**

\*express是一个代表 PivotFilter 对象的变量。

**说明**

table { word-break:break-all; }

下表列出了新的筛选器类型，以及对于每一个筛选器类型必须提供的参数和不可提供（不可用）的参数。

| xlPivotFilterType                 | DataField | DataCubeField (OLAP) | Value1 | Value2 |
|-----------------------------------|-----------|----------------------|--------|--------|
| xlTopCount                        | 必需      | 必需                 | 必需   | 不可用 |
| xlBottomCount                     | 必需      | 必需                 | 必需   | 不可用 |
| xlTopPercent                      | 必需      | 必需                 | 必需   | 不可用 |
| xlBottomPercent                   | 必需      | 必需                 | 必需   | 不可用 |
| xlTopSum                          | 必需      | 必需                 | 必需   | 不可用 |
| xlBottomSum                       | 必需      | 必需                 | 必需   | 不可用 |
| xlCaptionEquals                   | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionDoesNotEqual             | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionIsGreaterThan            | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionIsGreaterThanOrEqualTo   | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionIsLessThan               | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionIsLessThanOrEqualTo      | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionBeginsWith               | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionDoesNotBeginWith         | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionEndsWith                 | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionDoesNotEndWith           | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionContains                 | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionDoesNotContain           | 不可用    | 不可用               | 必需   | 不可用 |
| xlCaptionIsBetween                | 不可用    | 不可用               | 必需   | 必需   |
| xlCaptionIsNotBetween             | 不可用    | 不可用               | 必需   | 必需   |
| xlValueEquals                     | 必需      | 必需                 | 必需   | 不可用 |
| xlValueDoesNotEqual               | 必需      | 必需                 | 必需   | 不可用 |
| xlValueIsGreaterThan              | 必需      | 必需                 | 必需   | 不可用 |
| xlValueIsGreaterThanOrEqualTo     | 必需      | 必需                 | 必需   | 不可用 |
| xlValueIsLessThan                 | 必需      | 必需                 | 必需   | 不可用 |
| xlValueIsLessThanOrEqualTo        | 必需      | 必需                 | 必需   | 不可用 |
| xlValueIsBetween                  | 必需      | 必需                 | 必需   | 必需   |
| xlValueIsNotBetween               | 必需      | 必需                 | 必需   | 必需   |
| xlSpecificDate                    | 不可用    | 不可用               | 必需   | 不可用 |
| xlNotSpecificDate                 | 不可用    | 不可用               | 必需   | 不可用 |
| xlBefore                          | 不可用    | 不可用               | 必需   | 不可用 |
| xlBeforeOrEqualTo                 | 不可用    | 不可用               | 必需   | 不可用 |
| xlAfter                           | 不可用    | 不可用               | 必需   | 不可用 |
| xlAfterOrEqualTo                  | 不可用    | 不可用               | 必需   | 不可用 |
| xlBetween                         | 不可用    | 不可用               | 必需   | 不可用 |
| xlNotBetween                      | 不可用    | 不可用               | 必需   | 不可用 |
| xlFilterToday                     | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterYesterday                 | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterTomorrow                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterThisWeek                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterLastWeek                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterNextWeek                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterThisMonth                 | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterLastMonth                 | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterNextMonth                 | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterThisQuarter               | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterLastQuarter               | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterNextQuarter               | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterThisYear                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterLastYear                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterNextYear                  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterYearToDate                | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter1  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter2  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter3  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter4  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodJanuary   | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodFebruary  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodMarch     | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodApril     | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodMay       | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodJune      | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodJuly      | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodAugust    | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodSeptember | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodOctober   | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodNovember  | 不可用    | 不可用               | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodDecember  | 不可用    | 不可用               | 不可用 | 不可用 |

### PivotFilter.IsMemberPropertyFilter

table { word-break:break-all; }

指定标签筛选器是基于字段的成员属性的 PivotItem 标题还是基于透视字段本身的 PivotItem 标题。只读 **Boolean** 类型。

**语法**

**express.IsMemberPropertyFilter**

\*express是一个代表 PivotFilter 对象的变量。

**说明**

table { word-break:break-all; }

此属性的默认值为 **False** 。

如果标签筛选器基于透视字段的成员属性的 PivotItem 标题，则返回 **True** ；或者，如果筛选器基于透视字段的 PivotItem 标题，则返回 **False** 。此属性只对标签筛选器和至少有一个成员属性的 OLAP 透视字段有效。

### PivotFilter.MemberPropertyField

table { word-break:break-all; }

此属性指定标签筛选器所基于的成员属性 PivotField。可读/写 **PivotField** 类型。

**语法**

**express.MemberPropertyField**

\*express是一个代表 PivotFilter 对象的变量。

**说明**

table { word-break:break-all; }

此属性只对标签筛选器和至少有一个成员属性的 OLAP 透视字段有效。

### PivotFilter.Name

table { word-break:break-all; }

此属性提供选项可指定引用筛选。索引值可能发生更改，因此不能依赖于这些值获得准确的引用。

**语法**

**express.Name**

\*express是一个代表 PivotFilter 对象的变量。

### PivotFilter.Order

table { word-break:break-all; }

指定应用于整个数据透视表的所有“值”筛选器中筛选器的求值顺序。 **Integer** 类型，可读/写。

**语法**

**express.Order**

\*express是一个代表 PivotFilter 对象的变量。

**说明**

table { word-break:break-all; }

此属性仅对“值”和“前 *n* 个”类型的 PivotFilters 有效。如果尝试设置或获取“标签”和“日期”筛选器的该属性，则会返回一个运行时错误。1 代表已求值的第一个筛选器，2 代表已求值的下一个筛选器，依此类推，直至到达第 *n* 个值。-1 代表非活动筛选器。

如果在添加新的筛选器时未指定 **EvaluationOrder** 属性，则该属性将被设置为 *N+1* （其中 *N* 为筛选器集合中当前最大的 **EvaluationOrder** 数字）。

可在 **Add** 方法中指定此属性，或者可在稍后通过更改属性来设置字段的此属性。

如果提高某个字段的求值顺序，则会将原来具有该求值顺序值的字段（以及这两个字段之间的所有字段）的求值顺序降低 1。将求值顺序设置为和原来一样将不会产生变化。如果降低某个字段的求值顺序，则会将原来具有该求值顺序值的字段（以及这两个字段之间的所有字段）的求值顺序提高 1。

集合中 PivotFilters 的顺序与对其进行求值的顺序相同。因此开发人员可以更改对透视字段进行求值的顺序。如果将透视字段（非 OLAP 数据透视表）或 CubeField（OLAP 数据透视表）从 **PivotTables** 集合中删除，则会针对应用于字段的“值”或“前 *n* 个”筛选器，将此属性设置为 -1。如果未明确指定值，那么在将删除的字段添加回去时，会针对应用的“值”或“前 *n* 个”筛选器，将 **EvaluationOrder** 属性设置为 *N+1* 。

### PivotFilter.Parent

table { word-break:break-all; }

返回指定 **PivotFilter** 对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotFilter 对象的变量。

### PivotFilter.PivotField

table { word-break:break-all; }

指定应用筛选器的透视字段。只读。

**语法**

**express.PivotField**

\*express是一个代表 PivotFilter 对象的变量。

### PivotFilter.Value1

table { word-break:break-all; }

此属性是由用户提供的用于定义透视字段筛选器的参数。可读/写 **Variant** 类型。

**语法**

**express.Value1**

\*express是一个代表 PivotFilter 对象的变量。

### PivotFilter.Value2

table { word-break:break-all; }

此属性是由用户提供的用于定义透视字段筛选器的参数。可读/写 **Variant** 类型。

**语法**

**express.Value2**

\*express是一个代表 PivotFilter 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotFilter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotLines 对象

## 简介

table { word-break:break-all; }

PivotLines 对象是数据透视表中线条的集合，其中包含数据透视表中行或列上的所有线条。每个线条都是一组 PivotCells。

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
<td style="text-align: left;" data-valian="middle"><a href="#PivotLines.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 <span>Application</span> 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLines.Count">Count</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回 PivotLines 集合中的项数。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLines.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLines.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>根据 PivotLines 集合对象特定元素在集合中的位置返回该特定元素。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLines.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定 PivotLines 对象的父对象。只读。</p></td>
</tr>
</tbody>
</table>

## 成员属性

### PivotLines.Application

table { word-break:break-all; }

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotLines 对象的变量。

**说明**

table { word-break:break-all; }

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### PivotLines.Count

table { word-break:break-all; }

返回 **PivotLines** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 PivotLines 对象的变量。

### PivotLines.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotLines 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotLines.Item

table { word-break:break-all; }

根据 **PivotLines** 集合对象特定元素在集合中的位置返回该特定元素。只读。

**语法**

**express.Item**

\*express是一个代表 PivotLines 对象的变量。

### PivotLines.Parent

table { word-break:break-all; }

返回指定 **PivotLines** 对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotLines 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotLines](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# RecentFiles 对象

## 简介

代表最近使用过的文件的列表。

## 说明

每个文件都由一个 RecentFile 对象代表。

## 方法

| 序号 | 名称                    | 说明                             |
|:----:|:------------------------|:---------------------------------|
|  1   | [Add](#RecentFiles.Add) | 向最近使用的文件列表中添加文件。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#RecentFiles.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#RecentFiles.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#RecentFiles.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Item](#RecentFiles.Item)               | 从集合中返回一个对象。                                                                                                                                                                                                          |
|  5   | [Maximum](#RecentFiles.Maximum)         | 返回或设置最近使用文件清单中文件数目的上限。可为 0（零）到 9 之间的数字。 Long 类型，可读写。                                                                                                                                   |
|  6   | [Parent](#RecentFiles.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### RecentFiles.Add

向最近使用的文件列表中添加文件。

**语法**

**express.Add ( *Name* )**

\*express是一个代表 RecentFiles 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明     |
|--------|-----------|----------|----------|
| *Name* | 必选      | String   | 文件名。 |

**返回值**

包含在集合中的一个 RecentFile 对象。

**示例**

此示例将文件 Oscar.xls 添加到最近使用的文件列表中。

``` JavaScript
Application.RecentFiles.Add("Oscar.xls")
```

## 成员属性

### RecentFiles.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 RecentFiles 对象的变量。

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

### RecentFiles.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 RecentFiles 对象的变量。

### RecentFiles.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 RecentFiles 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### RecentFiles.Item

从集合中返回一个对象。

**语法**

**express.Item**

\*express是一个代表 RecentFiles 对象的变量。

**示例**

此示例打开最近使用过的文件列表中的第二个文件。

``` JavaScript
此示例打开最近使用过的文件列表中的第二个文件。
```

### RecentFiles.Maximum

返回或设置最近使用文件清单中文件数目的上限。可为 0（零）到 9 之间的数字。 **Long** 类型，可读写。

**语法**

**express.Maximum**

\*express是一个代表 RecentFiles 对象的变量。

**示例**

本示例将最近使用的文件列表中的最多文件数设置为 6。

``` JavaScript
Application.RecentFiles.Maximum = 6
```

### RecentFiles.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 RecentFiles 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/RecentFiles](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ShadowFormat 对象

## 简介

代表形状的阴影格式。

## 说明

使用 Shadow 属性可返回一个 ShadowFormat 对象。

## 方法

| 序号 | 名称                                               | 说明                                                                            |
|:----:|:---------------------------------------------------|:--------------------------------------------------------------------------------|
|  1   | [IncrementOffsetX](#ShadowFormat.IncrementOffsetX) | 按指定磅值更改阴影的水平偏移量。可以使用 OffsetX 属性设置阴影的绝对水平偏移量。 |
|  2   | [IncrementOffsetY](#ShadowFormat.IncrementOffsetY) | 按指定磅值更改阴影的垂直偏移量。可以使用 OffsetY 属性设置阴影的绝对垂直偏移量。 |

## 属性

| 序号 | 名称                                             | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ShadowFormat.Application)         | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Blur](#ShadowFormat.Blur)                       | 返回或设置指定底纹的模糊度。可读/写 Single 类型。                                                                                                                                                                               |
|  3   | [Creator](#ShadowFormat.Creator)                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [ForeColor](#ShadowFormat.ForeColor)             | 返回或设置一个 ColorFormat 对象，它代表指定的前景填充色或纯色。                                                                                                                                                                 |
|  5   | [Obscured](#ShadowFormat.Obscured)               | 如果指定形状的阴影是填充的，并且阴影被形状所遮盖（即便该形状没有填充），那么该值为 True 。如果阴影没有填充，并且当形状没有填充时，可透过形状看到阴影的轮廓，那么该值为 False 。 MsoTriState 类型，可读写。                      |
|  6   | [OffsetX](#ShadowFormat.OffsetX)                 | 以磅为单位返回或设置指定形状的阴影的水平偏移量。正偏移值将阴影向右偏移，负偏移值将阴影向左偏移。可读写。 Single 类型。                                                                                                          |
|  7   | [OffsetY](#ShadowFormat.OffsetY)                 | 以磅为单位返回或设置指定形状阴影的垂直偏移。正偏移值将阴影向下偏移，负偏移值将阴影向上偏移。可读写。 Single 类型。                                                                                                              |
|  8   | [Parent](#ShadowFormat.Parent)                   | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  9   | [RotateWithShape](#ShadowFormat.RotateWithShape) | 返回或设置一个 MsoTriState 类型的值，该值表示是否在旋转形状时旋转阴影。可读/写。                                                                                                                                                |
|  10  | [Size](#ShadowFormat.Size)                       | 返回或设置指定底纹的大小。可读/写 Single 类型。                                                                                                                                                                                 |
|  11  | [Style](#ShadowFormat.Style)                     | 返回或设置指定底纹的样式。可读/写 MsoShadowStyle 类型。                                                                                                                                                                         |
|  12  | [Transparency](#ShadowFormat.Transparency)       | 返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。 Double 型，可读写。                                                                                                                                    |
|  13  | [Type](#ShadowFormat.Type)                       | 返回或设置一个代表阴影格式类型的 MsoShadowType 值。                                                                                                                                                                             |
|  14  | [Visible](#ShadowFormat.Visible)                 | 返回或设置一个 MsoTriState 值，它确定对象是否可见。可读写。                                                                                                                                                                     |

## 成员方法

### ShadowFormat.IncrementOffsetX

按指定磅值更改阴影的水平偏移量。可以使用 **OffsetX** 属性设置阴影的绝对水平偏移量。

**语法**

**express.IncrementOffsetX ( *Increment* )**

\*express是一个代表 ShadowFormat 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                     |
|-------------|-----------|----------|--------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定阴影的水平偏移量（以磅为单位）。正值表示阴影右移，负值表示阴影左移。 |

**示例**

``` JavaScript
//本示例使 myDocument 中形状三的阴影向左移动 3 磅。
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.Item(3).Shadow.IncrementOffsetX(-3)
}
```

### ShadowFormat.IncrementOffsetY

按指定磅值更改阴影的垂直偏移量。可以使用 **OffsetY** 属性设置阴影的绝对垂直偏移量。

**语法**

**express.IncrementOffsetY ( *Increment* )**

\*express是一个代表 ShadowFormat 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                     |
|-------------|-----------|----------|--------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定阴影的垂直偏移量（以磅为单位）。正值表示阴影下移，负值表示阴影上移。 |

**示例**

``` JavaScript
//本示例使 myDocument 中形状三的阴影向上移动 3 磅。
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.Item(3).Shadow.IncrementOffsetY(-3)
}
```

## 成员属性

### ShadowFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ShadowFormat 对象的变量。

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

### ShadowFormat.Blur

返回或设置指定底纹的模糊度。可读/写 **Single** 类型。

**语法**

**express.Blur**

\*express是一个代表 ShadowFormat 对象的变量。

### ShadowFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ShadowFormat 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ShadowFormat.ForeColor

返回或设置一个 **ColorFormat** 对象，它代表指定的前景填充色或纯色。

**语法**

**express.ForeColor**

\*express是一个代表 ShadowFormat 对象的变量。

### ShadowFormat.Obscured

如果指定形状的阴影是填充的，并且阴影被形状所遮盖（即便该形状没有填充），那么该值为 **True** 。如果阴影没有填充，并且当形状没有填充时，可透过形状看到阴影的轮廓，那么该值为 **False** 。 **MsoTriState** 类型，可读写。

**语法**

**express.Obscured**

\*express是一个代表 ShadowFormat 对象的变量。

**说明**

|                                                                           |
|---------------------------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。                     |
| **msoCTrue**                                                              |
| **msoFalse** 阴影没有填充；如果形状也没有填充，透过形状可看见阴影的轮廓。 |
| **msoTriStateMixed**                                                      |
| **msoTriStateToggle**                                                     |
| **msoTrue** 指定形状的阴影有填充并被该形状所遮蔽，即使形状也没有填充。    |

**示例**

### ShadowFormat.OffsetX

以磅为单位返回或设置指定形状的阴影的水平偏移量。正偏移值将阴影向右偏移，负偏移值将阴影向左偏移。可读写。 **Single** 类型。

**语法**

**express.OffsetX**

\*express是一个代表 ShadowFormat 对象的变量。

**说明**

使用 **IncrementOffsetX** 方法或 **IncrementOffsetY** 方法可将阴影从当前位置水平或垂直轻微移动而无需指定其绝对位置。

**示例**

``` JavaScript
//本示例为 myDocument 中的形状 3 设置阴影的水平和垂直偏移量。阴影相对形状向右偏移 5 磅，向上偏移 3 磅。如果形状没有阴影，本示例为其添加一个阴影。
function test()
{
    let shadow = Application.Worksheets.Item(1).Shapes.Item(3).Shadow
    shadow.Visible = true
    shadow.OffsetX = 5
    shadow.OffsetY = -3
}
```

### ShadowFormat.OffsetY

以磅为单位返回或设置指定形状阴影的垂直偏移。正偏移值将阴影向下偏移，负偏移值将阴影向上偏移。可读写。 **Single** 类型。

**语法**

**express.OffsetY**

\*express是一个代表 ShadowFormat 对象的变量。

**说明**

使用 **IncrementOffsetX** 方法或 **IncrementOffsetY** 方法可将阴影从当前位置水平或垂直轻微移动而无需指定其绝对位置。

**示例**

``` JavaScript
//本示例为 myDocument 中的形状 3 设置阴影的水平和垂直偏移量。阴影相对形状向右偏移 5 磅，向上偏移 3 磅。如果形状没有阴影，本示例为其添加一个阴影。
function test()
{
    let shadow = Application.Worksheets.Item(1).Shapes.Item(3).Shadow
    shadow.Visible = true
    shadow.OffsetX = 5
    shadow.OffsetY = -3
}
```

### ShadowFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ShadowFormat 对象的变量。

### ShadowFormat.RotateWithShape

返回或设置一个 **MsoTriState** 类型的值，该值表示是否在旋转形状时旋转阴影。可读/写。

**语法**

**express.RotateWithShape**

\*express是一个代表 ShadowFormat 对象的变量。

### ShadowFormat.Size

返回或设置指定底纹的大小。可读/写 **Single** 类型。

**语法**

**express.Size**

\*express是一个代表 ShadowFormat 对象的变量。

### ShadowFormat.Style

返回或设置指定底纹的样式。可读/写 **MsoShadowStyle** 类型。

**语法**

**express.Style**

\*express是一个代表 ShadowFormat 对象的变量。

**说明**

|                                                         |
|---------------------------------------------------------|
| MsoShadowStyle 可以是下列 **MsoShadowStyle** 常量之一。 |
| **msoShadowStyleInnerShadow**                           |
| **msoShadowStyleMixed**                                 |
| **msoShadowStyleOuterShadow**                           |

### ShadowFormat.Transparency

返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。 **Double** 型，可读写。

**语法**

**express.Transparency**

\*express是一个代表 ShadowFormat 对象的变量。

**说明**

此属性的值仅影响均一颜色填充和线条的外观。对图案线条及渐变、图案、图片和纹理填充的外观无效。

**示例**

``` JavaScript
//此示例将第一张工作表上的第三个形状的阴影效果设置为半透明的红色。如果形状没有阴影，此示例将加上阴影效果。
function test()
{
    let shadow = Application.Worksheets.Item(1).Shapes.Item(3).Shadow
    shadow.Visible = true
    shadow.ForeColor.RGB = (255, 0, 0)
    shadow.Transparency = 0.5
}
```

### ShadowFormat.Type

返回或设置一个代表阴影格式类型的 **MsoShadowType** 值。

**语法**

**express.Type**

\*express是一个代表 ShadowFormat 对象的变量。

### ShadowFormat.Visible

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 ShadowFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ShadowFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Sheets 对象

## 简介

指定的或活动工作簿中所有工作表的集合。

## 说明

Sheets 集合可以包含 Chart 或 Worksheet 对象。

如果希望返回所有类型的工作表， Sheets 集合就非常有用。如果仅需使用某一类型的工作表，请参阅该工作表类型的对象主题。

使用 She ets 属性可返回 Sheets 集合。下例打印当前活动工作簿上的所有工作表。

``` JavaScript
Application.Sheets.PrintOut()
```

使用 Add 方法可创建一个新的工作表并将它添加到集合。下例给活动工作簿添加两个图表工作表，将它们放在工作簿中的工作表二之后。

``` JavaScript
Application.Sheets.Add(null, Application.Sheets.Item(2),2,xlChart)
```

使用 Sheets ( *index* )（其中 *index* 是工作表名称或索引号）可返回一个 Chart 或 Worksheet 对象。下例激活名为“Sheet1”的工作表。

``` JavaScript
Application.Sheets.Item("sheet1").Activate()
```

使用 Sheets ( *array* ) 可指定多个工作表。下例将名为“Sheet4”和“Sheet5”的工作表移到工作簿的开头。

``` JavaScript
Application.Sheets.Item(["Sheet4", "Sheet5"]).Move(Application.Sheets.Item(1))
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
<td style="text-align: left;" data-valian="middle"><a href="#Sheets.Add">Add</a></td>
<td style="text-align: left;" data-valian="middle"><p>新建工作表、图表或宏表。新建的工作表将成为活动工作表。</p>
<p>返回值： 一个 Object 值，它代表新的工作表、图表或宏表。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#Sheets.Copy">Copy</a></td>
<td style="text-align: left;" data-valian="middle"><p>将工作表复制到工作簿的另一位置。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#Sheets.Delete">Delete</a></td>
<td style="text-align: left;" data-valian="middle"><p>删除对象。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#Sheets.FillAcrossSheets">FillAcrossSheets</a></td>
<td style="text-align: left;" data-valian="middle"><p>将单元格区域复制到集合中所有其他工作表的同一位置。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#Sheets.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>从集合中返回一个对象。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#Sheets.Move">Move</a></td>
<td style="text-align: left;" data-valian="middle"><p>将工作表移到工作簿中的其他位置。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#Sheets.PrintOut">PrintOut</a></td>
<td style="text-align: left;" data-valian="middle"><p>打印对象。返回Variant值</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#Sheets.PrintPreview">PrintPreview</a></td>
<td style="text-align: left;" data-valian="middle"><p>按对象打印后的外观效果显示对象的预览。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#Sheets.Select">Select</a></td>
<td style="text-align: left;" data-valian="middle"><p>选择对象。</p></td>
</tr>
</tbody>
</table>

## 属性

| 序号 | 名称                               | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Sheets.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#Sheets.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#Sheets.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [HPageBreaks](#Sheets.HPageBreaks) | 返回一个 HPageBreaks 集合，它代表工作表上的水平分页符。只读。                                                                                                                                                                   |
|  5   | [Parent](#Sheets.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  6   | [VPageBreaks](#Sheets.VPageBreaks) | 返回一个 VPageBreaks 集合，它代表工作表上的垂直分页符。只读。                                                                                                                                                                   |
|  7   | [Visible](#Sheets.Visible)         | 返回或设置一个 Variant 值，它确定对象是否可见。                                                                                                                                                                                 |

## 成员方法

### Sheets.Add

新建工作表、图表或宏表。新建的工作表将成为活动工作表。

返回值： 一个 Object 值，它代表新的工作表、图表或宏表。

**语法**

**express.Add ( *Before* , *After* , *Count* , *Type* )**

\*express是一个代表 Sheets 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                        |
|----------|-----------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Before* | 可选      | Variant  | 指定工作表的对象，新建的工作表将置于此工作表之前。                                                                                                                                          |
| *After*  | 可选      | Variant  | 指定工作表的对象，新建的工作表将置于此工作表之后。                                                                                                                                          |
| *Count*  | 可选      | Variant  | 要添加的工作表数。默认值为 1。                                                                                                                                                              |
| *Type*   | 可选      | Variant  | 指定工作表类型。可以为下列 XlSheetType 常量之一：xlWorksheet、xlChart、xlExcel4MacroSheet 或 xlExcel4IntlMacroSheet。如果基于现有模板插入工作表，则指定该模板的路径。默认值为 xlWorksheet。 |

**说明**

如果同时省略 *Before* 和 *After* ，则新工作表插入到活动工作表之前。

**示例**

``` JavaScript
/*本示例将新建工作表插入到活动工作簿的最后一张工作表之前。*/
function test()
{
    Application.ActiveWorkbook.Sheets.Add(Application.Worksheets.Item(Worksheets.Count))
}
```

### Sheets.Copy

将工作表复制到工作簿的另一位置。

**语法**

**express.Copy ( *Before* , *After* )**

\*express是一个代表 Sheets 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                        |
|----------|-----------|----------|-----------------------------------------------------------------------------|
| *Before* | 可选      | Variant  | 将要在其之前放置所复制工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选      | Variant  | 将要在其之后放置所复制工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After* ，则 ET 将新建一个工作簿，其中包含复制的工作表。

**示例**

``` JavaScript
/*此示例复制工作表 Sheet1，并将其放置在工作表 Sheet3 之前。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Copy(Application.Worksheets.Item("Sheet3"))
}
```

### Sheets.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Sheets 对象的变量。

### Sheets.FillAcrossSheets

将单元格区域复制到集合中所有其他工作表的同一位置。

**语法**

**express.FillAcrossSheets ( *Range* , *Type* )**

\*express是一个代表 Sheets 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型   | 说明                                                                       |
|---------|-----------|------------|----------------------------------------------------------------------------|
| *Range* | 必选      | Range      | 要填充到集合中所有工作表上的单元格区域。该区域必须来自集合中的某个工作表。 |
| *Type*  | 可选      | XlFillWith | 指定如何复制区域。                                                         |

**示例**

``` JavaScript
function test(){
  /*本示例用工作表 Sheet1 上 A1:C5 单元格区域的内容填充工作表 Sheet5 和 Sheet7 上的相同区域。*/
  x = ["Sheet1", "Sheet5", "Sheet7"]
  Application.Sheets.Item(x).FillAcrossSheets(Application.Worksheets.Item("Sheet1").Range("A1:C5"))
}
```

### Sheets.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Sheets 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**示例**

``` JavaScript
/*此示例激活工作表 Sheet1。*/
Application.Sheets.Item("sheet1").Activate()
```

### Sheets.Move

将工作表移到工作簿中的其他位置。

**语法**

**express.Move ( *Before* , *After* )**

\*express是一个代表 Sheets 对象的变量。

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
Application.Worksheets.Item("Sheet1").Move(null,Application.Worksheets.Item("Sheet3"))
```

### Sheets.PrintOut

打印对象。返回Variant值

**语法**

**express.PrintOut ( *From* , *To* , *Copies* , *Preview* , *ActivePrinter* , *PrintToFile* , *Collate* , *PrToFileName* , *IgnorePrintAreas* )**

\*express是一个代表 Sheets 对象的变量。

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
/*此示例打印当前活动工作表。*/
Application.ActiveSheet.PrintOut()
```

### Sheets.PrintPreview

按对象打印后的外观效果显示对象的预览。

**语法**

**express.PrintPreview ( *EnableChanges* )**

\*express是一个代表 Sheets 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                          |
|-----------------|-----------|----------|-------------------------------------------------------------------------------|
| *EnableChanges* | 可选      | Variant  | 传递 Boolean 值，以指定用户是否可更改边距和打印预览中可用的其他页面设置选项。 |

**示例**

``` JavaScript
/*此示例在打印预览中显示工作表 Sheet1。*/
Application.Worksheets.Item("Sheet1").PrintPreview()
```

### Sheets.Select

选择对象。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 Sheets 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                                              |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

## 成员属性

### Sheets.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Sheets 对象的变量。

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

### Sheets.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 Sheets 对象的变量。

### Sheets.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Sheets 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Sheets.HPageBreaks

返回一个 **HPageBreaks** 集合，它代表工作表上的水平分页符。只读。

**语法**

**express.HPageBreaks**

\*express是一个代表 Sheets 对象的变量。

**说明**

每张工作表最多有 1026 个水平分页符。

**示例**

``` JavaScript
function test(){
  /*此示例显示全屏水平分页符和打印区水平分页符的个数。*/
  let cFull = 0
  let cPartial = 0
  for(let i = 1; i <= Application.Worksheets.Item(1).HPageBreaks.Count; i++){
      if(Application.Worksheets.Item(1).HPageBreaks.Item(i).Extent == xlPageBreakFull){
          cFull = cFull + 1
      }
      else{
          cPartial = cPartial + 1
      }
  }
  alert(cFull + " full-screen page breaks, " + cPartial + " print-area page breaks")
}
```

### Sheets.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Sheets 对象的变量。

### Sheets.VPageBreaks

返回一个 **VPageBreaks** 集合，它代表工作表上的垂直分页符。只读。

**语法**

**express.VPageBreaks**

\*express是一个代表 Sheets 对象的变量。

### Sheets.Visible

返回或设置一个 **Variant** 值，它确定对象是否可见。

**语法**

**express.Visible**

\*express是一个代表 Sheets 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Sheets](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SheetViews 对象

## 简介

指定的或活动工作簿窗口中所有工作表视图的集合。

## 说明

以下示例将返回活动窗口的视图计数。

``` JavaScript
Application.ActiveWindow.SheetViews.Count
```

## 方法

| 序号 | 名称                     | 说明                                                    |
|:----:|:-------------------------|:--------------------------------------------------------|
|  1   | [Item](#SheetViews.Item) | 返回一个 SheetView 对象，该对象代表工作簿的视图。只读。 |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                               |
|:----:|:---------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#SheetViews.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Count](#SheetViews.Count)             | 返回集合中对象的数目。只读 Long 类型。                                                                                                                             |
|  3   | [Creator](#SheetViews.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  4   | [Parent](#SheetViews.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |

## 成员方法

### SheetViews.Item

返回一个 **SheetView** 对象，该对象代表工作簿的视图。只读。

**语法**

**express.Item ( *index* )**

\*express是一个代表 SheetViews 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *index* | 必选      | Variant  | 视图的索引值。 |

**示例**

## 成员属性

### SheetViews.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 SheetViews 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### SheetViews.Count

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

\*express是一个代表 SheetViews 对象的变量。

### SheetViews.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 SheetViews 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### SheetViews.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SheetViews 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/SheetViews](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TextFrame 对象

## 简介

代表 Shape 对象中的文本框架。包含文本框架中的文本以及控制文本框架的对齐和定位的属性和方法。

## 说明

使用 TextFrame 属性可返回一个 TextFrame 对象。

下例向 *myDocument* 中添加一个矩形，向矩形中添加文本，然后设置文本框的边距。

``` JavaScript
let myDocument = Application.Worksheets.Item(1)
let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame

rng.Characters().Text = "Here is some test text"
rng.MarginBottom = 10
rng.MarginLeft = 10
rng.MarginRight = 10
rng.MarginTop = 10
```

## 方法

| 序号 | 名称                                | 说明                                                                                                                                |
|:----:|:------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Characters](#TextFrame.Characters) | 返回一个 Ch aracters 对象，该对象表示某个形状的文本框架中的字符区域。可以使用 Characters 对象向文本框架中添加字符和设置字符的格式。 |

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
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.AutoMargins">AutoMargins</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 ET 是否自动计算文本框边距。可读/写</p>
<p>返回Boolean值</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.AutoSize">AutoSize</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定对象能自动调整大小，以适应其中所包含的文字，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.HorizontalAlignment">HorizontalAlignment</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 XlHAlign 值，它代表指定对象的水平对齐方式。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.HorizontalOverflow">HorizontalOverflow</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定对象的水平溢出设置。可读/写</p>
<p>返回 XlOartHorizontalOverflow 值</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.MarginBottom">MarginBottom</a></td>
<td style="text-align: left;" data-valian="middle"><p>以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读/写。 Single 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.MarginLeft">MarginLeft</a></td>
<td style="text-align: left;" data-valian="middle"><p>以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。 Single 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.MarginRight">MarginRight</a></td>
<td style="text-align: left;" data-valian="middle"><p>以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。 Single 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.MarginTop">MarginTop</a></td>
<td style="text-align: left;" data-valian="middle"><p>以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。 Single 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.Orientation">Orientation</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Long 值，它代表文本框的方向。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.ReadingOrder">ReadingOrder</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定对象的阅读次序。可为以下常量之一： xlRTL （从右到左）、 xlLTR （从左到右）或 xlContext 。 Long 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.VerticalAlignment">VerticalAlignment</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 XlVAlign 值，它代表指定对象的垂直对齐方式。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#TextFrame.VerticalOverflow">VerticalOverflow</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定对象的垂直溢出设置。可读/写</p>
<p>返回 XlOartVerticalOverflow 值</p></td>
</tr>
</tbody>
</table>

## 成员方法

### TextFrame.Characters

返回一个 **Ch aracters** 对象，该对象表示某个形状的文本框架中的字符区域。可以使用 **Characters** 对象向文本框架中添加字符和设置字符的格式。

**语法**

**express.Characters ( *Start* , *Length* )**

\*express是一个代表 TextFrame 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                                                    |
|----------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------|
| *Start*  | 可选      | Variant  | 要返回的第一个字符。如果此参数设置为 1 或被省略，则 Characters 方法将返回以第一个字符为起始字符的字符区域。             |
| *Length* | 可选      | Variant  | 要返回的字符个数。如果省略此参数，则 Characters 方法将返回该字符串的剩余部分（由 Start 参数设置的字符以后的所有字符）。 |

**说明**

**Characters** 对象不是集合。

**示例**

``` JavaScript
/*本示例将活动工作表中第一个形状的文本框架中第三个字符的格式设置为加粗。*/
let rng = Application.ActiveSheet.Shapes.Item(1).TextFrame
rng.Characters().Text = "abcdefg"
rng.Characters(3, 1).Font.Bold = true
```

## 成员属性

### TextFrame.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
let myObject = ActiveWorkbook
if(myObject.Application.Value == "ET"){
    MsgBox("This is an ET Application object.")
}
else{
    MsgBox("This is not an ET Application object.")
}
```

### TextFrame.AutoMargins

返回或设置 ET 是否自动计算文本框边距。可读/写

返回Boolean值

**语法**

**express.AutoMargins**

\*express是一个代表 TextFrame 对象的变量。

**说明**

如果 ET 自动计算文本框边距，则为 **True** ；否则为 **False** 。当此属性为 **True** 时， **MarginLeft** , **MarginRight** 、 **MarginTop** 和 **MarginBottom** 属性将被忽略。

### TextFrame.AutoSize

如果指定对象能自动调整大小，以适应其中所包含的文字，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutoSize**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例使第一个形状的文本框能自动调整大小，以适应其中所包含的文字。*/
Application.Worksheets.Item(1).Shapes.Item(1).TextFrame.AutoSize = true
```

### TextFrame.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 TextFrame 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### TextFrame.HorizontalAlignment

返回或设置一个 **XlHAlign** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

\*express是一个代表 TextFrame 对象的变量。

**说明**

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### TextFrame.HorizontalOverflow

返回或设置指定对象的水平溢出设置。可读/写

返回 **XlOartHorizontalOverflow 值**

**语法**

**express.HorizontalOverflow**

\*express是一个代表 TextFrame 对象的变量。

**说明**

此属性只在 **Word Wrap** 属性为 **msoFalse** (0) 时有效。

### TextFrame.MarginBottom

以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读/写。 **Single** 类型。

**语法**

**express.MarginBottom**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
let myDocument = Worksheets.Item(1)
let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
rng.Characters().Text = "Here is some test text"
rng.MarginBottom = 0
rng.MarginLeft = 100
rng.MarginRight = 0
rng.MarginTop = 20
```

### TextFrame.MarginLeft

以磅为单位返回或设置从文本框左边界到包含文本的形状中内接矩形左边界的距离。可读写。 **Single** 类型。

**语法**

**express.MarginLeft**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
let myDocument = Worksheets.Item(1)
let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
rng.Characters().Text = "Here is some test text"
rng.MarginBottom = 0
rng.MarginLeft = 100
rng.MarginRight = 0
rng.MarginTop = 20
```

### TextFrame.MarginRight

以磅为单位返回或设置从文本框右边界到包含文本的形状中内接矩形右边界的距离。可读写。 **Single** 类型。

**语法**

**express.MarginRight**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
let myDocument = Worksheets.Item(1)
let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
rng.Characters().Text = "Here is some test text"
rng.MarginBottom = 0
rng.MarginLeft = 100
rng.MarginRight = 0
rng.MarginTop = 20
```

### TextFrame.MarginTop

以磅为单位返回或设置从文本框架顶端到包含文本的形状中内接矩形顶端的距离。可读写。 **Single** 类型。

**语法**

**express.MarginTop**

\*express是一个代表 TextFrame 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
let myDocument = Worksheets.Item(1)
let rng = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
rng.Characters().Text = "Here is some test text"
rng.MarginBottom = 0
rng.MarginLeft = 100
rng.MarginRight = 0
rng.MarginTop = 20
```

### TextFrame.Orientation

返回或设置一个 **Long** 值，它代表文本框的方向。

**语法**

**express.Orientation**

\*express是一个代表 TextFrame 对象的变量。

**说明**

此属性值可设为 -90 到 90 度之间的整数值或以下 **<span "="" r="" rlidawscontentredir?assetid="HV100730982052&amp;CTT=11&amp;Origin=HV103325142052&quot;" style="color: rgb(0, 0, 0); text-decoration-line: none;" target="_new">MsoTextOrientation</span>** 常量之一。

### TextFrame.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.ReadingOrder

返回或设置指定对象的阅读次序。可为以下常量之一： **xlRTL** （从右到左）、 **xlLTR** （从左到右）或 **xlContext** 。 **Long** 类型，可读写。

**语法**

**express.ReadingOrder**

\*express是一个代表 TextFrame 对象的变量。

### TextFrame.VerticalAlignment

返回或设置一个 **XlVAlign** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

\*express是一个代表 TextFrame 对象的变量。

**说明**

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

### TextFrame.VerticalOverflow

返回或设置指定对象的垂直溢出设置。可读/写

返回 **XlOartVerticalOverflow** 值

**语法**

**express.VerticalOverflow**

\*express是一个代表 TextFrame 对象的变量。

**说明**

此属性只在 **AutoS ize** 属性为 **False** 时有效。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/TextFrame](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Window 对象

## 简介

代表窗口。

## 说明

许多工作表特征（如滚动条和标尺）实际上是窗口的属性。 Window 对象是 Windows 集合的成员。 Application 对象的 Windows 集合包含应用程序中的所有窗口，而 Workbook 对象的 Windows 集合只包含指定工作簿中的窗口。

## 示例

使用 Windows ( *index* )（其中 *index* 是窗口名称或索引号）可返回一个 Window 对象。下例最大化活动窗口。

``` JavaScript
Application.Windows.Item(1).WindowState = xlMaximized
```

注意，活动窗口总是 ` Windows(1) ` 。

窗口题主是窗口未处于最大化时窗口最上方标题栏中显示的文本。该题注还会显示在 “窗口” 菜单最下方的打开文件列表中。使用 Caption 属性可设置或返回窗口题注。更改窗口题注不会更改工作簿的名称。下例为“Book1.xls:1”窗口中显示的工作表关闭单元格网格线。

``` JavaScript
Application.Windows.Item("book1.xls":1).DisplayGridlines = false
```

## 方法

| 序号 | 名称                                         | 说明                                                                    |
|:----:|:---------------------------------------------|:------------------------------------------------------------------------|
|  1   | [Activate](#Window.Activate)                 | 将窗口提到 z-次序的最前面。 返回值类型为 Variant。                      |
|  2   | [ActivateNext](#Window.ActivateNext)         | 激活指定窗口，并将其移到窗口 z-次序的末尾。 返回值类型为 Variant。      |
|  3   | [ActivatePrevious](#Window.ActivatePrevious) | 激活指定窗口，并将其移到窗口 z-次序的末尾。 返回值类型为Variant。       |
|  4   | [Close](#Window.Close)                       | 关闭对象。 返回值： 如果该方法成功关闭对象，则为 True ，否则为 False 。 |
|  5   | [NewWindow](#Window.NewWindow)               | 新建一个窗口或者创建指定窗口的副本。 返回值类型为 Window。              |
|  6   | [RangeFromPoint](#Window.RangeFromPoint)     | 新建一个窗口或者创建指定窗口的副本。 返回值类型为 Object。              |

## 属性

| 序号 | 名称                                     | 说明                                                                                                                                                                                                                                                                                                                                                                                   |
|:----:|:-----------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ActiveCell](#Window.ActiveCell)         | 返回一个 Range 对象，它代表活动窗口（最上方的窗口）或指定窗口中的活动单元格。如果窗口中没有显示工作表，此属性无效。只读。                                                                                                                                                                                                                                                              |
|  2   | [ActiveSheet](#Window.ActiveSheet)       | 返回一个对象，它代表活动工作簿中或指定的窗口或工作簿中的活动工作表（最上面的工作表）。如果没有活动的工作表，则返回 Nothing 。                                                                                                                                                                                                                                                          |
|  3   | [Application](#Window.Application)       | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。                                                                                                                                                        |
|  4   | [Caption](#Window.Caption)               | 返回或设置一个 Variant 值，它代表文档窗口标题栏中显示的名称。                                                                                                                                                                                                                                                                                                                          |
|  5   | [Height](#Window.Height)                 | 返回或设置一个 Double 值，它代表窗口的高度（以磅为单位）。                                                                                                                                                                                                                                                                                                                             |
|  6   | [Index](#Window.Index)                   | 返回 Long 值，它代表对象在其同类对象所组成的集合内的索引号。                                                                                                                                                                                                                                                                                                                           |
|  7   | [Left](#Window.Left)                     | 返回一个 Double 值，它代表从客户区左边缘到窗口左边缘的距离（以磅为单位）。                                                                                                                                                                                                                                                                                                             |
|  8   | [RangeSelection](#Window.RangeSelection) | 返回一个 Range 对象，该对象表示指定窗口中工作表上的选定单元格，即使工作表上一个图形对象是活动或选定的。只读。                                                                                                                                                                                                                                                                          |
|  9   | [Selection](#Window.Selection)           | 对于 Windows 对象，返回一个指定的窗口。                                                                                                                                                                                                                                                                                                                                                |
|  10  | [SheetViews](#Window.SheetViews)         | 返回指定窗口的 SheetViews 对象。只读。                                                                                                                                                                                                                                                                                                                                                 |
|  11  | [Top](#Window.Top)                       | 返回或设置一个 Double 值，它代表从窗口上边缘到使用区域（在菜单、任何停放在顶端的工具栏和编辑栏下方）上边缘的距离（以磅为单位）。                                                                                                                                                                                                                                                       |
|  12  | [Type](#Window.Type)                     | 返回或设置一个代表窗口类型的 XlWindowType 值。                                                                                                                                                                                                                                                                                                                                         |
|  13  | [UsableHeight](#Window.UsableHeight)     | 返回在应用程序窗口区域中一个窗口能占有的最大高度（以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" darkreader-inline-color="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal; --darkreader-inline-color:#e8e6e3;">磅</span> 为单位）。 Double 型，只读。 |
|  14  | [UsableWidth](#Window.UsableWidth)       | 返回在应用程序窗口区域中一个窗口能占有的最大宽度（以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位）。 Double 型，只读。                                                               |
|  15  | [Visible](#Window.Visible)               | 返回或设置一个 Boolean 值，它确定对象是否可见。可读写。                                                                                                                                                                                                                                                                                                                                |
|  16  | [VisibleRange](#Window.VisibleRange)     | 返回一个 Range 对象，它代表显示在窗口或窗格中的单元格区域。如果列或行只显示了一部分，则说明它是包括在区域之内的。只读。                                                                                                                                                                                                                                                                |
|  17  | [Width](#Window.Width)                   | 返回或设置一个 Double 值，它代表窗口的宽度（以磅为单位）。                                                                                                                                                                                                                                                                                                                             |
|  18  | [WindowState](#Window.WindowState)       | 返回或设置窗口的状态。 XlWindowState 类型，可读写。                                                                                                                                                                                                                                                                                                                                    |
|  19  | [Zoom](#Window.Zoom)                     | 返回或设置一个 Variant 值，它代表窗口的显示大小，以百分数表示（100 表示正常大小，200 表示双倍大小，以此类推）。                                                                                                                                                                                                                                                                        |

## 成员方法

### Window.Activate

将窗口提到 z-次序的最前面。 返回值类型为 Variant。

**语法**

**express.Activate ()**

\*express是一个代表 Window 对象的变量。

**说明**

这样不会运行可能附加在工作簿上的 Auto_Activate 或 Auto_Deactivate 宏（使用 **RunAutoMacros** 方法可运行这些宏）。

### Window.ActivateNext

激活指定窗口，并将其移到窗口 z-次序的末尾。 返回值类型为 Variant。

**语法**

**express.ActivateNext ()**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*本示例将活动窗口移到窗口 z-次序的末尾。*/
Application.ActiveWindow.ActivateNext()
```

### Window.ActivatePrevious

激活指定窗口，并将其移到窗口 z-次序的末尾。 返回值类型为Variant。

**语法**

**express.ActivatePrevious ()**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*本示例激活窗口 z-次序末尾的窗口。*/
Application.ActiveWindow.ActivatePrevious()
```

### Window.Close

关闭对象。 返回值： 如果该方法成功关闭对象，则为 **True** ，否则为 **False** 。

**语法**

**express.Close ( *SaveChanges* , *Filename* , *RouteWorkbook* )**

\*express是一个代表 Window 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                                                                                                                              |
|-----------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *SaveChanges*   | 可选      | Variant  | 如果工作簿中没有改动，则忽略此参数。如果工作簿中有改动但工作簿显示在其他打开的窗口中，则忽略此参数。如果工作簿中有改动且工作簿未显示在任何其他打开的窗口中，则由此参数指定是否应保存更改。如果设为 True，则保存对工作簿所做的更改。如果工作簿尚未命名，则使用 Filename。如果省略 Filename，则要求用户提供文件名。 |
| *Filename*      | 可选      | Variant  | 以此文件名保存所做的更改。                                                                                                                                                                                                                                                                                        |
| *RouteWorkbook* | 可选      | Variant  | 如果工作簿不需要传送给下一个收件人（没有传送名单或已经传送），则忽略此参数。否则，ET 根据此参数的值传送工作簿。如果设为 True，则将工作簿传送给下一个收件人。如果设为 False，则不发送工作簿。如果忽略，则要求用户确认是否发送工作簿。                                                                              |

**说明**

使用 **RunAutoMac ros** 方法可运行自动关闭宏。

### Window.NewWindow

新建一个窗口或者创建指定窗口的副本。 返回值类型为 Window。

**语法**

**express.NewWindow ()**

\*express是一个代表 Window 对象的变量。

### Window.RangeFromPoint

新建一个窗口或者创建指定窗口的副本。 返回值类型为 Object。

**语法**

**express.RangeFromPoint ( *x* , *y* )**

\*express是一个代表 Window 对象的变量。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                                                       |
|------|-----------|----------|------------------------------------------------------------|
| *x*  | 必选      | Long     | 表示从顶部开始到屏幕左边缘的水平距离的值（以像素为单位）。 |
| *y*  | 必选      | Long     | 表示从左侧开始到屏幕顶部的垂直距离的值（以像素为单位）。   |

## 成员属性

### Window.ActiveCell

返回一个 **Range** 对象，它代表活动窗口（最上方的窗口）或指定窗口中的活动单元格。如果窗口中没有显示工作表，此属性无效。只读。

**语法**

**express.ActiveCell**

\*express是一个代表 Window 对象的变量。

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
/*此示例在消息框中显示活动单元格的值。由于如果活动表不是工作表则 ActiveCell 属性无效，所以此示例使用 ActiveCell 属性之前先激活 Sheet1。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    alert(ActiveCell.Value())
}

/*此示例更改活动单元格的字体格式设置。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    let activeCell = ActiveCell.Font
    activeCell.Bold = true
    activeCell.Italic = true
}
```

### Window.ActiveSheet

返回一个对象，它代表活动工作簿中或指定的窗口或工作簿中的活动工作表（最上面的工作表）。如果没有活动的工作表，则返回 **Nothing** 。

**语法**

**express.ActiveSheet**

\*express是一个代表 Window 对象的变量。

**说明**

如果不指定对象识别符，则此属性返回活动工作簿中的活动工作表。

如果某个工作簿出现在若干个窗口中，那么该工作簿的 **ActiveSheet** 属性在不同窗口中可能不同。

**示例**

``` JavaScript
/*此示例显示活动工作表的名称。*/
alert("The name of the active sheet is " + Application.ActiveSheet.Name)
```

### Window.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test()
{
    if (Application.ActiveWorkbook.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### Window.Caption

返回或设置一个 **Variant** 值，它代表文档窗口标题栏中显示的名称。

**语法**

**express.Caption**

\*express是一个代表 Window 对象的变量。

**说明**

如果设置了该名称，就可以使用该名称作为 **Windows** 集合的索引（如示例中的所述）。

**示例**

``` JavaScript
/*此示例将活动工作簿中第一个窗口的名称设置为“Consolidated Balance Sheet”。然后该名称将用作 Windows 集合中该窗口的索引。*/
function test()
{
    Application.ActiveWorkbook.Windows.Item(1).Caption = "Consolidated Balance Sheet"
    Application.ActiveWorkbook.Windows.Item("Consolidated Balance Sheet").ActiveSheet.Calculate()
}
```

### Window.Height

返回或设置一个 Double 值，它代表窗口的高度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 Window 对象的变量。

**说明**

使用 **UsableHeight** 属性可确定窗口的最大尺寸。如果窗口已被最大化或最小化，则不能设置此属性。使用 **WindowState** 属性可确定窗口的状态。

### Window.Index

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

\*express是一个代表 Window 对象的变量。

### Window.Left

返回一个 **Double** 值，它代表从客户区左边缘到窗口左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 Window 对象的变量。

### Window.RangeSelection

返回一个 **Range** 对象，该对象表示指定窗口中工作表上的选定单元格，即使工作表上一个图形对象是活动或选定的。只读。

**语法**

**express.RangeSelection**

\*express是一个代表 Window 对象的变量。

**说明**

当工作表中已选定一个图形对象， **Selection** 属性返回的是一个图形对象而不是一个 **Range** 对象； **RangeSelection** 属性将返回在图形对象被选定之前选定的单元格区域。

当工作表中选定的是一个单元格区域时（不是一个图形对象），该属性和 **Selection** 属性将返回相同的值。

如果指定窗口中的活动表不是一个工作表，则该属性无效。

**示例**

``` JavaScript
/*本示例显示在当前窗口的工作表中选定的单元格区域的地址。*/
alert(Application.ActiveWindow.RangeSelection.Address())
```

### Window.Selection

对于 **Windows** 对象，返回一个指定的窗口。

**语法**

**express.Selection**

\*express是一个代表 Window 对象的变量。

**说明**

返回的对象类型取决于当前所选内容（例如，如果选择了单元格，此属性将返回 **Range** 对象）。如果未选择任何内容， **Selection** 属性将返回 **Nothing** 。

在不使用对象识别符的情况下，使用此属性等效于使用 ` Application.Selection ` 。

**示例**

``` JavaScript
/*本示例清空 Sheet1 的选定对象（假定选定对象为单元格区域）。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.Selection.Clear()
}

/*本示例显示选定对象的 JSAPI 对象类型。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    alert("The selection object type is "+ typeof(Application.Selection))
}
```

### Window.SheetViews

返回指定窗口的 **SheetViews** 对象。只读。

**语法**

**express.SheetViews**

\*express是一个代表 Window 对象的变量。

### Window.Top

返回或设置一个 **Double** 值，它代表从窗口上边缘到使用区域（在菜单、任何停放在顶端的工具栏和编辑栏下方）上边缘的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 Window 对象的变量。

**说明**

无法对最大化窗口设置此属性。使用 **WindowState** 属性可返回或设置窗口的状态。

**示例**

``` JavaScript
/*本示例水平排列第一个窗口和第二个窗口，也就是说，每个窗口占用可用的垂直空间的一半，占用应用程序窗口的工作区域的所有水平空间。若要运行该示例，必须只打开两个工作表窗口。*/
function test()
{
    Application.Windows.Arrange(-4128)
    let ah = Application.Windows.Item(1).Height                      // available height
    let aw = Application.Windows.Item(1).Width + Application.Windows.Item(2).Width    // available width

    let wintop1 = Application.Windows.Item(1)
    wintop1.Width = aw
    wintop1.Height = ah / 2
    wintop1.Left = 0

    let wintop2 = Application.Windows.Item(2)
    wintop2.Width = aw
    wintop2.Height = ah / 2
    wintop2.Top = ah / 2
    wintop2.Left = 0
}
```

### Window.Type

返回或设置一个代表窗口类型的 **XlWindowType** 值。

**语法**

**express.Type**

\*express是一个代表 Window 对象的变量。

### Window.UsableHeight

返回在应用程序窗口区域中一个窗口能占有的最大高度（以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" darkreader-inline-color="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal; --darkreader-inline-color:#e8e6e3;">磅</span> 为单位）。 **Double** 型，只读。

**语法**

**express.UsableHeight**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*此示例将活动窗口扩展为可用的最大值（假定指定窗口尚未最大化）。*/
function test()
{
    Application.ActiveWindow.WindowState = xlNormal
    Application.ActiveWindow.Top = 1
    Application.ActiveWindow.Left = 1
    Application.ActiveWindow.Height = Application.UsableHeight
    Application.ActiveWindow.Width = Application.UsableWidth
}
```

### Window.UsableWidth

返回在应用程序窗口区域中一个窗口能占有的最大宽度（以 <span class="glossary" "appendpopup(this,'ofdefpoint_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">磅</span> 为单位）。 **Double** 型，只读。

**语法**

**express.UsableWidth**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*此示例将活动窗口扩展为可用的最大值（假定指定窗口尚未最大化）。*/
function test()
{
    Application.ActiveWindow.WindowState = xlNormal
    Application.ActiveWindow.Top = 1
    Application.ActiveWindow.Left = 1
    Application.ActiveWindow.Height = Application.UsableHeight
    Application.ActiveWindow.Width = Application.UsableWidth
}
```

### Window.Visible

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 Window 对象的变量。

### Window.VisibleRange

返回一个 **Range** 对象，它代表显示在窗口或窗格中的单元格区域。如果列或行只显示了一部分，则说明它是包括在区域之内的。只读。

**语法**

**express.VisibleRange**

\*express是一个代表 Window 对象的变量。

### Window.Width

返回或设置一个 **Double** 值，它代表窗口的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 Window 对象的变量。

**说明**

使用 **UsableWidth** 属性可确定窗口的最大尺寸。如果窗口已被最大化或最小化，则不能设置此属性。使用 **<span style="color:#000000;text-decoration-line:none;">WindowState</span>** 属性可确定窗口的状态。

### Window.WindowState

返回或设置窗口的状态。 **XlWindowState** 类型，可读写。

**语法**

**express.WindowState**

\*express是一个代表 Window 对象的变量。

**示例**

``` JavaScript
/*本示例将 ET 应用程序窗口最大化。*/
Application.WindowState = xlMaximized

/*本示例将活动窗口扩展为可用的最大值（假定指定窗口尚未最大化）。*/
function test()
{
    Application.ActiveWindow.WindowState = xlNormal
    Application.ActiveWindow.Top = 1
    Application.ActiveWindow.Left = 1
    Application.ActiveWindow.Height = Application.UsableHeight
    Application.ActiveWindow.Width = Application.UsableWidth
}
```

### Window.Zoom

返回或设置一个 **Variant** 值，它代表窗口的显示大小，以百分数表示（100 表示正常大小，200 表示双倍大小，以此类推）。

**语法**

**express.Zoom**

\*express是一个代表 Window 对象的变量。

**说明**

将此属性设为 **True** ，可将窗口大小设置成与当前选定区域相适应的大小。

本功能仅对窗口中当前的活动工作表起作用。若要对其他工作表使用此属性，必须先激活工作表。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Window](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomProperties 对象

## 简介

CustomProperty 对象的集合，这些对象代表与智能标记相关的属性。 CustomProperties 集合包括文档中的所有智能标记自定义属性。

## 方法

| 序号 | 名称                           | 说明                                                                 |
|:----:|:-------------------------------|:---------------------------------------------------------------------|
|  1   | [Add](#CustomProperties.Add)   | 返回一个 CustomProperty 对象，该对象代表添加到智能标记的自定义属性。 |
|  2   | [Item](#CustomProperties.Item) | 返回集合中的一个 CustomProperty 对象。                               |

## 属性

| 序号 | 名称                                         | 说明                                                                         |
|:----:|:---------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#CustomProperties.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#CustomProperties.Count)             | 返回一个 Long 类型的值，该值代表集合中的项目数。只读。                       |
|  3   | [Creator](#CustomProperties.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#CustomProperties.Parent)           | 返回一个 Object 类型值，该值代表指定 CustomProperties 对象的父对象。         |

## 成员方法

### CustomProperties.Add

返回一个 **CustomProperty** 对象，该对象代表添加到智能标记的自定义属性。

**语法**

**express.Add ( *Name* , *Value* )**

\*express是一个代表 CustomProperties 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                       |
|---------|-----------|----------|----------------------------|
| *Name*  | 必选      | String   | 自定义智能标记属性的名称。 |
| *Value* | 必选      | String   | 自定义智能标记属性的值     |

### CustomProperties.Item

返回集合中的一个 **CustomProperty** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 CustomProperties 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

## 成员属性

### CustomProperties.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 CustomProperties 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### CustomProperties.Count

返回一个 **Long** 类型的值，该值代表集合中的项目数。只读。

**语法**

**express.Count**

\*express是一个代表 CustomProperties 对象的变量。

### CustomProperties.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CustomProperties 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### CustomProperties.Parent

返回一个 **Object** 类型值，该值代表指定 **CustomProperties** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 CustomProperties 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/CustomProperties](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomProperty 对象

## 简介

代表智能标记的自定义属性的单个实例。 CustomProperty 对象是 CustomProperties 集合的一个成员。

## 说明

使用 CustomProperties 集合的 Item 方法或 Properties ( *Index* ) 属性（其中 *Index* 是属性编号）返回一个 CustomProperty 对象。

使用 Nam e 和 Value 属性返回与智能标记的自定义属性相关的信息。本示例显示一条消息，其中包含当前文档中第一个智能标记的第一个自定义属性的名称和值。本示例假定当前文档包含至少一个智能标记，并且第一个智能标记至少具有一个自定义属性。

``` JavaScript
function SmartTagsProps(){
  let rng = Application.ActiveDocument.SmartTags.Item(1).Properties.Item(1)
  alert("Smart Tag Name: " + rng.Name + "\n" + "Smart Tag Value: " + rng.Value)
}
```

## 方法

| 序号 | 名称                             | 说明                 |
|:----:|:---------------------------------|:---------------------|
|  1   | [Delete](#CustomProperty.Delete) | 删除指定自定义属性。 |

## 属性

| 序号 | 名称                                       | 说明                                                                         |
|:----:|:-------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#CustomProperty.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Creator](#CustomProperty.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  3   | [Name](#CustomProperty.Name)               | 返回指定对象的名称。 String 类型，只读。                                     |
|  4   | [Parent](#CustomProperty.Parent)           | 返回一个 Object 类型值，该值代表指定 CustomProperty 对象的父对象。           |
|  5   | [Value](#CustomProperty.Value)             | 返回或设置自定义属性的值。可读/写 String 类型。                              |

## 成员方法

### CustomProperty.Delete

删除指定自定义属性。

**语法**

**express.Delete ()**

\*express是一个代表 CustomProperty 对象的变量。

## 成员属性

### CustomProperty.Application

返回一个代表 WPS 应用程序的 Application 对象。

**语法**

**express.Application**

\*express是一个代表 CustomProperty 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### CustomProperty.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CustomProperty 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### CustomProperty.Name

返回指定对象的名称。 **String** 类型，只读。

**语法**

**express.Name**

\*express是一个代表 CustomProperty 对象的变量。

### CustomProperty.Parent

返回一个 **Object** 类型值，该值代表指定 **CustomProperty** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 CustomProperty 对象的变量。

### CustomProperty.Value

返回或设置自定义属性的值。可读/写 **String** 类型。

**语法**

**express.Value**

\*express是一个代表 CustomProperty 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/CustomProperty](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Dialogs 对象

## 简介

WPS 中的 Dialog 对象的集合。每个 Dialog 对象代表一个内置的 WPS 对话框。

## 方法

| 序号 | 名称                  | 说明                      |
|:----:|:----------------------|:--------------------------|
|  1   | [Item](#Dialogs.Item) | 返回 WPS 中的一个对话框。 |

## 属性

| 序号 | 名称                                | 说明                                                                         |
|:----:|:------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Dialogs.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Dialogs.Count)             | 返回一个 Number 类型的值，该值代表集合中的对话框数。只读。                   |
|  3   | [Creator](#Dialogs.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#Dialogs.Parent)           | 返回一个 Object 类型值，该值代表指定 Dialogs 对象的父对象。                  |

## 成员方法

### Dialogs.Item

返回 WPS 中的一个对话框。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Dialogs 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型     | 说明               |
|---------|-----------|--------------|--------------------|
| *Index* | 必选      | WdWordDialog | 指定对话框的常量。 |

**示例**

``` JavaScript
/* 本示例显示“页面设置”对话框 */
function test(){
  Application.Dialogs.Item(wdDialogFileDocumentLayout).Display()
}
```

## 成员属性

### Dialogs.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Dialogs 对象的变量。

### Dialogs.Count

返回一个 **Number** 类型的值，该值代表集合中的对话框数。只读。

**语法**

**express.Count**

\*express是一个代表 Dialogs 对象的变量。

### Dialogs.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Dialogs 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Dialogs.Parent

返回一个 **Object** 类型值，该值代表指定 **Dialogs** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Dialogs 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Dialogs](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FileConverters 对象

## 简介

FileConverter 对象的集合，这些对象代表可用于打开和保存文件的所有文件转换器。

## 方法

| 序号 | 名称                         | 说明                                  |
|:----:|:-----------------------------|:--------------------------------------|
|  1   | [Item](#FileConverters.Item) | 返回集合中的单个 FileConverter 对象。 |

## 属性

| 序号 | 名称                                                             | 说明                                                                         |
|:----:|:-----------------------------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#FileConverters.Application)                       | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [ConvertMacWordChevrons](#FileConverters.ConvertMacWordChevrons) | 控制是否将 V 型字符 (? ?) 中的文字转化成合并域。 Long 类型，可读写。         |
|  3   | [Count](#FileConverters.Count)                                   | 返回一个 Long 类型的值，该值代表集合中文件转换器的数量。只读。               |
|  4   | [Creator](#FileConverters.Creator)                               | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  5   | [Parent](#FileConverters.Parent)                                 | 返回一个 Object 类型值，该值代表指定 FileConverters 对象的父对象。           |

## 成员方法

### FileConverters.Item

返回集合中的单个 **FileConverter** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 FileConverters 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

**返回值**

FileConverter

## 成员属性

### FileConverters.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 FileConverters 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### FileConverters.ConvertMacWordChevrons

控制是否将 V 型字符 (? ?) 中的文字转化成合并域。 **Long** 类型，可读写。

**语法**

**express.ConvertMacWordChevrons**

\*express是一个代表 FileConverters 对象的变量。

**说明**

**ConvertMacWordChevrons** 属性可以是任何 **WdChevronConvertRule** 常量。

适用于 Macintosh 的 WPS 4.0 和 5.x 版本的文档用 V 型字符指出邮件合并域。

**示例**

``` JavaScript
/*本示例设置 ConvertMacWordChevrons 属性，将 V 型字符中的文本转换成邮件合并域，然后打开名为“Mac WPS Document”的文档。*/
function test()
{
FileConverters.ConvertMacWordChevrons = wdAlwaysConvert
Documents.Open("C:\\Documents\\Mac  WPS Document")
}
```

### FileConverters.Count

返回一个 **Long** 类型的值，该值代表集合中文件转换器的数量。只读。

**语法**

**express.Count**

\*express是一个代表 FileConverters 对象的变量。

### FileConverters.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FileConverters 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### FileConverters.Parent

返回一个 **Object** 类型值，该值代表指定 **FileConverters** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 FileConverters 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/FileConverters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Templates 对象

## 简介

Template 对象的集合，这些对象代表当前可用的所有模板。该集合包括打开的模板、附加到打开的文档的模板和在 “模板和加载项” 对话框中加载的全局模板。

## 方法

| 序号 | 名称                                                | 说明                                |
|:----:|:----------------------------------------------------|:------------------------------------|
|  1   | [Item](#Templates.Item)                             | 返回集合中的单个 Template 对象。    |
|  2   | [LoadBuildingBlocks](#Templates.LoadBuildingBlocks) | 将所有模板的构建基块加载到 WPS 中。 |

## 属性

| 序号 | 名称                                  | 说明                                                                         |
|:----:|:--------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Templates.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Templates.Count)             | 返回一个 Long 类型的值，该值代表指定集合中模板的数量。只读。                 |
|  3   | [Creator](#Templates.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#Templates.Parent)           | 返回一个 Object 类型值，该值代表指定 Templates 对象的父对象。                |

## 成员方法

### Templates.Item

返回集合中的单个 **Template** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Templates 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                     |
|---------|-----------|----------|------------------------------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

### Templates.LoadBuildingBlocks

将所有模板的构建基块加载到 WPS 中。

**语法**

**express.LoadBuildingBlocks ()**

\*express是一个代表 Templates 对象的变量。

**说明**

WPS 通常在第一次需要构建基块时加载这些构建基块，例如，当用户使用功能区显示库时。此方法强制 WPS 立即加载所有构建基块。

## 成员属性

### Templates.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Templates 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Templates.Count

返回一个 **Long** 类型的值，该值代表指定集合中模板的数量。只读。

**语法**

**express.Count**

\*express是一个代表 Templates 对象的变量。

### Templates.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Templates 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Templates.Parent

返回一个 **Object** 类型值，该值代表指定 **Templates** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Templates 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Templates](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# View 对象

## 简介

包含窗口或窗格的视图属性（如全部显示、域底纹和网格线）。

## 说明

使用 View 属性可返回 View 对象。以下示例设置活动窗口的视图选项。

``` JavaScript
let view = Application.ActiveDocument.ActiveWindow.View
view.ShowAll = true
view.TableGridlines = true
view.WrapToWindow = false
```

使用 Type 属性可更改视图。以下示例将活动窗口切换至普通视图。

``` JavaScript
Application.ActiveDocument.ActiveWindow.View.Type = wdNormalView
```

使用 Percentage 属性可更改屏幕上文本的尺寸。以下示例将屏幕上文本的显示比例放大到 120%。

``` JavaScript
Application.ActiveDocument.ActiveWindow.View.Zoom.Percentage = 120
```

使用 SeekView 属性可查看批注、尾注、脚注或者文档页眉或页脚。以下示例在页面视图中显示活动窗口中的当前页脚。

``` JavaScript
Application.ActiveDocument.ActiveWindow.View.Type = wdPrintView 
Application.ActiveDocument.ActiveWindow.View.SeekView = wdSeekCurrentPageFooter
```

## 方法

| 序号 | 名称                                               | 说明                                                   |
|:----:|:---------------------------------------------------|:-------------------------------------------------------|
|  1   | [CollapseOutline](#View.CollapseOutline)           | 将选定内容或指定范围内的正文折叠一个标题级别。         |
|  2   | [ExpandOutline](#View.ExpandOutline)               | 按照一个标题级别展开选定内容下的文本。                 |
|  3   | [PreviousHeaderFooter](#View.PreviousHeaderFooter) | 根据视图中显示的是页眉还是页脚，移动到前一页眉或页脚。 |
|  4   | [ShowAllHeadings](#View.ShowAllHeadings)           | 在显示所有文本（包括标题和正文）和仅显示标题之间切换。 |
|  5   | [ShowHeading](#View.ShowHeading)                   | 显示所有高于指定标题级别的标题，隐藏子标题和正文。     |

## 属性

| 序号 | 名称                                                                             | 说明                                                                                                                                                          |
|:----:|:---------------------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#View.Application)                                                 | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                |
|  2   | [ConflictMode](#View.ConflictMode)                                               | 如果文档位于冲突模式视图中，则该属性值为 True 。可读/写。                                                                                                     |
|  3   | [Creator](#View.Creator)                                                         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                  |
|  4   | [DisplayBackgrounds](#View.DisplayBackgrounds)                                   | 返回或设置一个 Boolean 类型的值，该值代表在页面视图中显示文档时，是否显示背景色和图像。                                                                       |
|  5   | [DisplayPageBoundaries](#View.DisplayPageBoundaries)                             | 如果为 True ，则显示文档顶部和底部页边距（空白区域）并在页面间显示灰色区域（灰色间隔）。 Boolean 类型，可读写。                                               |
|  6   | [DisplaySmartTags](#View.DisplaySmartTags)                                       | 如果为 True ，则 WPS 在文档中的智能标记下方显示一条下划线。 Boolean 类型，可读写。                                                                            |
|  7   | [Draft](#View.Draft)                                                             | 如果窗口中的所有文本以具有最少格式信息的相同的 sans-serif 字体显示以加速显示，则该属性值为 True 。 Boolean 类型，可读写。                                     |
|  8   | [FieldShading](#View.FieldShading)                                               | 返回或设置域的屏幕底纹。 WdFieldShading 类型，可读写。                                                                                                        |
|  9   | [FullScreen](#View.FullScreen)                                                   | 如果窗口为全屏显示，则该属性值为 True 。 Boolean 类型，可读写。                                                                                               |
|  10  | [Magnifier](#View.Magnifier)                                                     | 如果为 True ，则在打印预览中将鼠标指针显示为放大镜形状，表示用户可以通过单击来放大页面的某个区域或通过缩小显示比例查看整页或展开多页。 Boolean 类型，可读写。 |
|  11  | [MailMergeDataView](#View.MailMergeDataView)                                     | 如果在指定窗口中显示邮件合并数据，而不显示邮件合并域，则该属性值为 True 。 Boolean 类型，可读写。                                                             |
|  12  | [MarkupMode](#View.MarkupMode)                                                   | 返回或设置一个 WdRevisionsMode 常量，该常量代表修订的显示模式。可读写。                                                                                       |
|  13  | [Panning](#View.Panning)                                                         | 返回或设置 Boolean 类型的值，该值表示 WPS 是否处于全景模式。可读写。                                                                                          |
|  14  | [Parent](#View.Parent)                                                           | 返回一个 Object 类型值，该值代表指定 View 对象的父对象。                                                                                                      |
|  15  | [ReadingLayout](#View.ReadingLayout)                                             | 设置或返回一个 Boolean 类型的值，该值表示是否在阅读版式视图中查看文档。                                                                                       |
|  16  | [ReadingLayoutActualView](#View.ReadingLayoutActualView)                         | 设置或返回一个 Boolean 类型的值，该值代表是否使用与打印页面相同的版式在阅读版式视图中显示页面。                                                               |
|  17  | [ReadingLayoutAllowEditing](#View.ReadingLayoutAllowEditing)                     | 返回 Boolean 值，该值表示是否允许在阅读布局模式下编辑文本。可读/写。                                                                                          |
|  18  | [ReadingLayoutAllowMultiplePages](#View.ReadingLayoutAllowMultiplePages)         | 设置或返回一个 Boolean 类型的值，该值代表阅读版式视图是否并排显示两页。                                                                                       |
|  19  | [ReadingLayoutTruncateMargins](#View.ReadingLayoutTruncateMargins)               | 返回或设置一个 WdReadingLayoutMargin 常量，该常量代表当在阅读版式中查看文档时是显示还是隐藏边距。可读写。                                                     |
|  20  | [Reviewers](#View.Reviewers)                                                     | 返回一个 Reviewers 对象，该对象代表所有审阅者。                                                                                                               |
|  21  | [RevisionsBalloonShowConnectingLines](#View.RevisionsBalloonShowConnectingLines) | 如果为 True ，则 WPS 显示从文本到修订和批注气球之间的连接线。 Boolean 类型，可读写。                                                                          |
|  22  | [RevisionsBalloonSide](#View.RevisionsBalloonSide)                               | 设置或返回一个 WdRevisionsBalloonMargin 常量，指定 WPS 是否在文档的左边距或右边距中显示修订气球。                                                             |
|  23  | [RevisionsBalloonWidth](#View.RevisionsBalloonWidth)                             | 设置或返回一个 Single 类型的值，该值代表 WPS 中用于指定修订气球宽度的全局设置。可读写。                                                                       |
|  24  | [RevisionsBalloonWidthType](#View.RevisionsBalloonWidthType)                     | 设置或返回一个 WdRevisionsBalloonWidthType 常量，该常量代表指定 WPS 如何测量修订气球的宽度的全局设置。可读写。                                                |
|  25  | [RevisionsView](#View.RevisionsView)                                             | 设置或返回一个代表全局选项的 WdRevisionsView 常量，该选项用于指定 WPS 是显示文档的原始版本还是应用了修订和格式更改的版本。可读写。                            |
|  26  | [SeekView](#View.SeekView)                                                       | 返回或设置在页面视图中显示的文档元素。 WdSeekView 类型，可读写。                                                                                              |
|  27  | [ShadeEditableRanges](#View.ShadeEditableRanges)                                 | 返回或设置一个 Long 类型的值，该值代表是否将底纹应用于文档中用户有权修改的区域。                                                                              |
|  28  | [ShowAll](#View.ShowAll)                                                         | 如果显示所有非打印字符（如隐藏文字、制表符、空格和段落标记），则该属性值为 True 。可读写 Boolean 类型。                                                       |
|  29  | [ShowAnimation](#View.ShowAnimation)                                             | 如果显示文本动画，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                 |
|  30  | [ShowBookmarks](#View.ShowBookmarks)                                             | 如果在每个书签的首尾显示方括号，则该属性值为 True 。 Boolean 类型，可读写。                                                                                   |
|  31  | [ShowComments](#View.ShowComments)                                               | 如果为 True ，则 WPS 在文档中显示批注。 Boolean 类型，可读写。                                                                                                |
|  32  | [ShowCropMarks](#View.ShowCropMarks)                                             | 返回或设置 Boolean 值，该值表示是否在页面的角显示裁剪标记，以指示页边距的位置。可读/写。                                                                      |
|  33  | [ShowDrawings](#View.ShowDrawings)                                               | 如果在页面视图中显示由绘图工具创建的对象，则该属性值为 True 。 Boolean 类型，可读写。                                                                         |
|  34  | [ShowFieldCodes](#View.ShowFieldCodes)                                           | 如果显示域代码，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                   |
|  35  | [ShowFirstLineOnly](#View.ShowFirstLineOnly)                                     | 如果在大纲视图中仅显示正文每段的第一行，则该属性值为 True 。 Boolean 类型，可读写。                                                                           |
|  36  | [ShowFormat](#View.ShowFormat)                                                   | 如果在大纲视图中显示字符格式，则该属性值为 True 。 Boolean 类型，可读写。                                                                                     |
|  37  | [ShowFormatChanges](#View.ShowFormatChanges)                                     | 如果为 True ，则在启用“修订”功能后，WPS 显示对文档所作的格式更改。 Boolean 类型，可读写。                                                                     |
|  38  | [ShowHiddenText](#View.ShowHiddenText)                                           | 如果显示具有隐藏格式的文本，则该属性值为 True 。 Boolean 类型，可读写。                                                                                       |
|  39  | [ShowHighlight](#View.ShowHighlight)                                             | 如果显示和打印文档中突出显示格式的文本，则该属性值为 True 。 Boolean 类型，可读写。                                                                           |
|  40  | [ShowHyphens](#View.ShowHyphens)                                                 | 如果为 True ，则显示可选连字符，该连字符指明在何处断开处于行尾的字。 Boolean 类型，可读写。                                                                   |
|  41  | [ShowInkAnnotations](#View.ShowInkAnnotations)                                   | 返回或设置一个 Boolean 类型的值，该值表示显示还是隐藏手写墨迹注释。如果值为 True ，则显示墨迹注释。如果值为 False ，则隐藏墨迹注释。                          |
|  42  | [ShowInsertionsAndDeletions](#View.ShowInsertionsAndDeletions)                   | 如果为 True ，则 WPS 显示用“修订”功能在文档中插入和删除的文本。 Boolean 类型，可读写。                                                                        |
|  43  | [ShowMainTextLayer](#View.ShowMainTextLayer)                                     | 如果在显示页眉和页脚区域时同时显示指定文档的文字，则该属性值为 True 。该属性等效于 “页眉和页脚” 工具栏上的 “显示/隐藏文档文字” 按钮。 Boolean 类型，可读写。  |
|  44  | [ShowMarkupAreaHighlight](#View.ShowMarkupAreaHighlight)                         | 返回或设置 Boolean 值，该值表示显示修订和注释气球的标记区域是否带底纹。可读/写。                                                                              |
|  45  | [ShowObjectAnchors](#View.ShowObjectAnchors)                                     | 如果在可在页面视图中定位的项目附近显示对象锁定标记，则该属性值为 True 。 Boolean 类型，可读写。                                                               |
|  46  | [ShowOptionalBreaks](#View.ShowOptionalBreaks)                                   | 如果 WPS 显示可选换行符，则该属性值为 True 。 Boolean 类型，可读写                                                                                            |
|  47  | [ShowOtherAuthors](#View.ShowOtherAuthors)                                       | 如果应在文档中显示其他作者，则该属性值为 True 。可读/写 Boolean 类型。                                                                                        |
|  48  | [ShowParagraphs](#View.ShowParagraphs)                                           | 如果显示段落标记，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                 |
|  49  | [ShowPicturePlaceHolders](#View.ShowPicturePlaceHolders)                         | 如果将空白方框作为图片的占位符进行显示，则该属性值为 True 。 Boolean 类型，可读写。                                                                           |
|  50  | [ShowRevisionsAndComments](#View.ShowRevisionsAndComments)                       | 如果为 True ，则 WPS 显示使用“修订”功能对文档所作的修订和批注。 Boolean 类型，可读写。                                                                        |
|  51  | [ShowSpaces](#View.ShowSpaces)                                                   | 如果显示空格字符，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                 |
|  52  | [ShowTabs](#View.ShowTabs)                                                       | 如果显示制表符，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                   |
|  53  | [ShowTextBoundaries](#View.ShowTextBoundaries)                                   | 如果在页面视图中显示页边距、文本列、对象和框架周围的虚线，则该属性值为 True 。 Boolean 类型，可读写。                                                         |
|  54  | [ShowXMLMarkup](#View.ShowXMLMarkup)                                             | 返回一个 Long 类型的值，该值表示是否在文档中显示 XML 标记。                                                                                                   |
|  55  | [SplitSpecial](#View.SplitSpecial)                                               | 返回或设置活动窗口的窗格。 WdSpecialPane 类型，可读写。                                                                                                       |
|  56  | [TableGridlines](#View.TableGridlines)                                           | 如果显示表格虚框，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                 |
|  57  | [Type](#View.Type)                                                               | 返回或设置视图的类型。可读/写 WdViewType 类型。                                                                                                               |
|  58  | [WrapToWindow](#View.WrapToWindow)                                               | 如果文字在文档窗口的右边缘换行，而不是在右边距或文本栏的右边界，则该属性值为 True 。 Boolean 类型，可读写。                                                   |
|  59  | [Zoom](#View.Zoom)                                                               | 返回一个 Zoom 对象，该对象代表指定视图的显示比例。                                                                                                            |

## 成员方法

### View.CollapseOutline

将选定内容或指定范围内的正文折叠一个标题级别。

**语法**

**express.CollapseOutline ( *Range* )**

\*express是一个代表 View 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型     | 说明                                                     |
|---------|-----------|--------------|----------------------------------------------------------|
| *Range* | 可选      | Range object | 需要折叠的段落范围。如果省略此参数，则折叠全部选定内容。 |

**说明**

如果文档未处于大纲或主控文档视图，将导致出错。

**示例**

``` JavaScript
/*本示例为活动文档的第二段设置“标题 2”样式，将活动窗口切换至大纲视图，并折叠第二段的正文。*/
function test()
{
ActiveDocument.Paragraphs.Item(2).Style = wdStyleHeading2

ActiveDocument.ActiveWindow.View.Type = wdOutlineView
ActiveDocument.ActiveWindow.View.CollapseOutline(ActiveDocument.Paragraphs.Item(2).Range)
}
```

``` JavaScript
/*本示例将文档中每个标题折叠一个级别。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdOutlineView
ActiveDocument.ActiveWindow.View.CollapseOutline(ActiveDocument.Content)
}
```

### View.ExpandOutline

按照一个标题级别展开选定内容下的文本。

**语法**

**express.ExpandOutline ( *Range* )**

\*express是一个代表 View 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                     |
|---------|-----------|----------|----------------------------------------------------------|
| *Range* | 可选      | Range    | 要展开的段落的范围。如果省略此参数，则展开全部选定内容。 |

**说明**

如果文档未处于大纲或主控文档视图，将导致出错。

**示例**

``` JavaScript
/*本示例将文档中每个标题展开一个级别。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdOutlineView
ActiveDocument.ActiveWindow.View.ExpandOutline(ActiveDocument.Content)
}
```

``` JavaScript
/*本示例展开 Document2 窗口中的活动段落。*/
function test()
{
let myDoc = Windows.Item("Document2")

myDoc.Activate()
myDoc.View.Type = wdOutlineView
myDoc.View.ExpandOutline()
}
```

### View.PreviousHeaderFooter

根据视图中显示的是页眉还是页脚，移动到前一页眉或页脚。

**语法**

**express.PreviousHeaderFooter ()**

\*express是一个代表 View 对象的变量。

**说明**

如果视图显示的是页眉，则该方法移至当前节的前一个页眉（例如，从偶数页眉到奇数页眉）或者前一节的最后一个页眉。如果视图显示页脚，则移至前一个页脚。

| 注释                                                                                   |
|----------------------------------------------------------------------------------------|
| 如果视图显示文档的第一节的第一个页眉或页脚，或者如果根本不显示页眉或页脚，则导致出错。 |

**示例**

``` JavaScript
/*本示例插入一个偶数页分节符，将活动窗口切换到页面视图，显示当前页眉，然后切换到前一个页眉。*/
function test()
{
Selection.Collapse(wdCollapseStart)
Selection.InsertBreak(wdSectionBreakEvenPage)

ActiveDocument.ActiveWindow.View.Type = wdPrintView
ActiveDocument.ActiveWindow.View.SeekView = wdSeekCurrentPageHeader
ActiveDocument.ActiveWindow.View.PreviousHeaderFooter()
}
```

### View.ShowAllHeadings

在显示所有文本（包括标题和正文）和仅显示标题之间切换。

**语法**

**express.ShowAllHeadings ()**

\*express是一个代表 View 对象的变量。

**说明**

如果不是大纲视图或者主控文档视图，此方法将导致出错。

**示例**

``` JavaScript
/*本示例在大纲视图中使用 ShowHeading 方法显示所有标题（不显示正文），然后切换显示所有文本（包括标题和正文）。*/
function test()
{
let view = ActiveDocument.ActiveWindow.View

view.Type = wdOutlineView
view.ShowHeading(9)
view.ShowAllHeadings()
}　　　　
```

### View.ShowHeading

显示所有高于指定标题级别的标题，隐藏子标题和正文。

**语法**

**express.ShowHeading ( *Level* )**

\*express是一个代表 View 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                       |
|---------|-----------|----------|--------------------------------------------|
| *Level* | 必选      | Long     | 指大纲标题的级别（可取 1 到 9 之间数值）。 |

**说明**

如果不是大纲视图或者主控文档视图，此方法将导致出错。

**示例**

``` JavaScript
/*本示例将活动窗口切换到大纲视图，并显示所有以“标题 1”样式为格式的文本。隐藏正文和其他类型的标题。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdOutlineView
ActiveDocument.ActiveWindow.View.ShowHeading(1)
}
```

``` JavaScript
/*本示例将 Document1 窗口切换到大纲视图，并显示所有以“标题 1”、“标题 2”或“标题 3”样式为格式的文本。*/
function test()
{
let myDoc = Windows.Item("Document1").View
myDoc.Type = wdOutlineView
myDoc.ShowHeading(3)
}
```

## 成员属性

### View.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 View 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### View.ConflictMode

如果文档位于冲突模式视图中，则该属性值为 **True** 。可读/写。

**语法**

**express.ConflictMode**

\*express是一个代表 View 对象的变量。

### View.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 View 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### View.DisplayBackgrounds

返回或设置一个 **Boolean** 类型的值，该值代表在页面视图中显示文档时，是否显示背景色和图像。

**语法**

**express.DisplayBackgrounds**

\*express是一个代表 View 对象的变量。

**说明**

对应于 **“选项”** 对话框的 **“视图”** 选项卡上的 **“背景色和图像(仅页面视图)”** 选项。

**示例**

``` JavaScript
/*以下示例在页面视图中显示活动文档时，隐藏背景色和图像。*/
ActiveDocument.ActiveWindow.View.DisplayBackgrounds = false
```

### View.DisplayPageBoundaries

如果为 **True** ，则显示文档顶部和底部页边距（空白区域）并在页面间显示灰色区域（灰色间隔）。 **Boolean** 类型，可读写。

**语法**

**express.DisplayPageBoundaries**

\*express是一个代表 View 对象的变量。

**说明**

如果为 **False** ，则隐藏这些白色和灰色区域，使页面形成很长的一页。默认值为 **True** 。

该功能仅在页面视图中可用，并仅影响页面顶部和底部的灰色间隔，而不影响页面的左右两边。该设置影响指定窗口中的文档。保存该文档时，同时保存该设置的状态。

**示例**

``` JavaScript
/*本示例将当前视图更改为“页面视图”并隐藏文档页面之间的白色和灰色间隔。*/
function WhiteSpace(){
    ActiveWindow.View.Type = wdPrintView
    ActiveWindow.View.DisplayPageBoundaries = false
}
```

### View.DisplaySmartTags

如果为 **True** ，则 WPS 在文档中的智能标记下方显示一条下划线。 **Boolean** 类型，可读写。

**语法**

**express.DisplaySmartTags**

\*express是一个代表 View 对象的变量。

**说明**

在文档中用下划线标记智能标记。将 **DisplaySmartTags** 属性设为 **False** 不能删除智能标记；该操作仅关闭下划线的显示。

**示例**

``` JavaScript
/*本示例在活动视图中关闭智能标记下方下划线的显示。*/
function DontShowSmartTags(){
    ActiveWindow.View.DisplaySmartTags = false
}
```

### View.Draft

如果窗口中的所有文本以具有最少格式信息的相同的 sans-serif 字体显示以加速显示，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Draft**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例以草稿字体显示 Document1 窗口的内容。*/
Windows.Item("Document1").View.Draft = true
```

``` JavaScript
/*本示例切换活动窗口的草稿字体选项。*/
ActiveDocument.ActiveWindow.View.Draft = !ActiveDocument.ActiveWindow.View.Draft
```

### View.FieldShading

返回或设置域的屏幕底纹。 **WdFieldShading** 类型，可读写。

**语法**

**express.FieldShading**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例激活活动窗口中所有窗体域的域底纹。*/
ActiveDocument.ActiveWindow.View.FieldShading = wdFieldShadingAlways
```

### View.FullScreen

如果窗口为全屏显示，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.FullScreen**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例将活动窗口切换为全屏显示*/
Application.ActiveDocument.ActiveWindow.View.FullScreen = true

/*本示例激活 Sales.doc 窗口，并切换到非全屏显示*/
Application.Windows.Item("Sales.doc").Activate()
Application.Windows.Item("Sales.doc").View.FullScreen = false
```

### View.Magnifier

如果为 **True** ，则在打印预览中将鼠标指针显示为放大镜形状，表示用户可以通过单击来放大页面的某个区域或通过缩小显示比例查看整页或展开多页。 **Boolean** 类型，可读写。

**语法**

**express.Magnifier**

\*express是一个代表 View 对象的变量。

**说明**

如果不是在打印预览视图中，则该属性将产生错误。

**示例**

``` JavaScript
/*本示例将视图切换到打印预览，并把鼠标指针改为插入点。*/
function test()
{
PrintPreview = true
ActiveDocument.ActiveWindow.View.Magnifier = false
}
```

### View.MailMergeDataView

如果在指定窗口中显示邮件合并数据，而不显示邮件合并域，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.MailMergeDataView**

\*express是一个代表 View 对象的变量。

**说明**

如果指定窗口不是主控文档，则会产生一个错误。

**示例**

``` JavaScript
/*如果活动文档至少包含一个邮件合并域，则本示例显示附加数据源的第一个记录的邮件合并数据。*/
function test()
{
if(ActiveDocument.MailMerge.Fields.Count >= 1){
    ActiveDocument.MailMerge.DataSource.ActiveRecord = 1
    ActiveDocument.ActiveWindow.View.ShowFieldCodes = false
    ActiveDocument.ActiveWindow.View.MailMergeDataView = true
}
}
```

``` JavaScript
/*本示例在查看邮件合并域和查看结果数据之间切换。*/
function test()
{
ActiveDocument.ActiveWindow.View.ShowFieldCodes = false
ActiveDocument.ActiveWindow.View.MailMergeDataView = !ActiveDocument.ActiveWindow.View.MailMergeDataView
}
```

### View.MarkupMode

返回或设置一个 **WdRevisionsMode** 常量，该常量代表修订的显示模式。可读写。

**语法**

**express.MarkupMode**

\*express是一个代表 View 对象的变量。

### View.Panning

返回或设置 **Boolean** 类型的值，该值表示 WPS 是否处于全景模式。可读写。

**语法**

**express.Panning**

\*express是一个代表 View 对象的变量。

**说明**

全景模式将鼠标指针变为手形，使用户能够通过按住鼠标来拖动文档，从而滚动查看文档页。

### View.Parent

返回一个 **Object** 类型值，该值代表指定 **View** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 View 对象的变量。

### View.ReadingLayout

设置或返回一个 **Boolean** 类型的值，该值表示是否在阅读版式视图中查看文档。

**语法**

**express.ReadingLayout**

\*express是一个代表 View 对象的变量。

**说明**

如果为 **True** ，则将文档切换到阅读版式视图。如果为 **False** ，则关闭阅读版式视图。

**示例**

``` JavaScript
/*以下示例关闭阅读版式视图。*/
ActiveDocument.ActiveWindow.View.ReadingLayout = false
```

### View.ReadingLayoutActualView

设置或返回一个 **Boolean** 类型的值，该值代表是否使用与打印页面相同的版式在阅读版式视图中显示页面。

**语法**

**express.ReadingLayoutActualView**

\*express是一个代表 View 对象的变量。

**说明**

在阅读版式视图中，页面中不显示（能在普通视图和页面视图看到的）包含在打印页面中的全部内容。这些内容将在屏幕中予以显示。当 **ReadingLayoutActualView** 属性设为 **True** 时，文档显示将与打印出的页面一致。在较小的显示器上，会使用一定比例进行缩放，这将使文档不易阅读，在较大的显示器上显示效果就很好。

**示例**

``` JavaScript
/*以下示例在阅读版式视图中显示的页面与打印出来的页面相同。*/
function test()
{
ActiveWindow.View.ReadingLayout = true
ActiveWindow.View.ReadingLayoutActualView = true
}
```

### View.ReadingLayoutAllowEditing

返回 **Boolean** 值，该值表示是否允许在阅读布局模式下编辑文本。可读/写。

**语法**

**express.ReadingLayoutAllowEditing**

\*express是一个代表 View 对象的变量。

**说明**

此属性对应于阅读布局模式下 **“视图选项”** 菜单上的 **“允许键入** / **禁止键入”** 选项。

| 注释                                                   |
|--------------------------------------------------------|
| 如果在阅读布局模式以外的其他视图中，则不能设置此属性。 |

**示例**

``` JavaScript
/*以下示例使 WPS 禁止在阅读布局模式下编辑文本。*/
ActiveDocument.ActiveWindow.View.ReadingLayoutAllowEditing = false
```

### View.ReadingLayoutAllowMultiplePages

设置或返回一个 **Boolean** 类型的值，该值代表阅读版式视图是否并排显示两页。

**语法**

**express.ReadingLayoutAllowMultiplePages**

\*express是一个代表 View 对象的变量。

**说明**

WPS 可能允许、也可能不允许在阅读版式视图中显示两页。如果 WPS 不能保持合理的纵横比，则 WPS 将只显示一页。因此，如果将 **ReadingLayoutAllowMultiplePages** 属性设置为 **True** 并且 WPS 无法显示两页，则该属性将设为 **False** ，并且 WPS 将仅显示单页。

**示例**

``` JavaScript
/*以下示例在版式允许的情况下在阅读版式视图中并排显示两页。*/
function test()
{
ActiveWindow.View.ReadingLayout = true
ActiveWindow.View.ReadingLayoutAllowMultiplePages = true
}
```

### View.ReadingLayoutTruncateMargins

返回或设置一个 **WdReadingLayoutMargin** 常量，该常量代表当在阅读版式中查看文档时是显示还是隐藏边距。可读写。

**语法**

**express.ReadingLayoutTruncateMargins**

\*express是一个代表 View 对象的变量。

### View.Reviewers

返回一个 **Reviewers** 对象，该对象代表所有审阅者。

**语法**

**express.Reviewers**

\*express是一个代表 View 对象的变量。

**说明**

**Reviewers** 对象是所有审阅者的全局列表，无论审阅者是否审阅指定窗口中显示的文档。

**示例**

``` JavaScript
/*本示例隐藏文档中的所有修订和批注，只显示由“Jeff Smith”所作的修订和批注。*/
function HideRevisions(){
    let view = ActiveWindow.View

    view.ShowRevisionsAndComments = false
    view.ShowFormatChanges = true
    view.ShowInsertionsAndDeletions = true

    for(let i = 1; i <= view.Reviewers.Count; i++){
        view.Reviewers.Item(i).Visible = true
    }

    view.Reviewers.Item("Jeff Smith").Visible = true
}
```

### View.RevisionsBalloonShowConnectingLines

如果为 **True** ，则 WPS 显示从文本到修订和批注气球之间的连接线。 **Boolean** 类型，可读写。

**语法**

**express.RevisionsBalloonShowConnectingLines**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例隐藏文档文本和相应的修订或批注气球之间的连接线。*/
function ShowConnectingLines(){
    ActiveWindow.View.RevisionsBalloonShowConnectingLines = false
}
```

### View.RevisionsBalloonSide

设置或返回一个 **WdRevisionsBalloonMargin** 常量，指定 WPS 是否在文档的左边距或右边距中显示修订气球。

**语法**

**express.RevisionsBalloonSide**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例在左右边距之间切换修订气球。本示例假定活动窗口中的文档包含由一个或多个审阅者所做的修订，并且修订显示于气球中。*/
function ToggleRevisionBalloons(){
    if(ActiveWindow.View.RevisionsBalloonSide == wdLeftMargin){
        ActiveWindow.View.RevisionsBalloonSide = wdRightMargin
    }
    else{
        ActiveWindow.View.RevisionsBalloonSide = wdLeftMargin
    }
}
```

### View.RevisionsBalloonWidth

设置或返回一个 **Single** 类型的值，该值代表 WPS 中用于指定修订气球宽度的全局设置。可读写。

**语法**

**express.RevisionsBalloonWidth**

\*express是一个代表 View 对象的变量。

**说明**

修订气球的宽度包括文档边距与气球边缘之间 0.5 英寸宽的填充距离以及气球边缘与纸张边缘之间八分之一英寸的间距。 WPS 会在纸的左右边缘处添加空白。此宽度会扩展到边距，但并不更改文档的宽度或纸张大小。可以使用 **RevisionsBalloonWidthType** 属性指定在设置 **RevisionsBalloonWidth** 属性时的度量单位。

**示例**

``` JavaScript
/*本示例将修订气球的宽度设为 1 英寸并在左边距显示修订气球。本示例假定活动窗口中的文档包含由一个或多个审阅者所做的修订，并且修订显示于气球中。*/
function BalloonWidth(){
    ActiveWindow.View.RevisionsBalloonWidthType = wdBalloonWidthPoints
    ActiveWindow.View.RevisionsBalloonWidth = InchesToPoints(1)
    ActiveWindow.View.RevisionsBalloonSide = wdLeftMargin
}
```

### View.RevisionsBalloonWidthType

设置或返回一个 **WdRevisionsBalloonWidthType** 常量，该常量代表指定 WPS 如何测量修订气球的宽度的全局设置。可读写。

**语法**

**express.RevisionsBalloonWidthType**

\*express是一个代表 View 对象的变量。

**说明**

**RevisionsBalloonWidthType** 属性设置用于设置 **RevisionsBalloonWidth** 属性的度量单位。

**示例**

``` JavaScript
/*本示例将修订气球的宽度设为文档宽度的 25%。本示例假定活动窗口中的文档包含多个审阅者所作的修订，并且修订显示于气球中。*/
function BalloonWidthType(){
    ActiveWindow.View.RevisionsBalloonWidthType = wdBalloonWidthPercent
    ActiveWindow.View.RevisionsBalloonWidth = 25
}
```

### View.RevisionsView

设置或返回一个代表全局选项的 **WdRevisionsView** 常量，该选项用于指定 WPS 是显示文档的原始版本还是应用了修订和格式更改的版本。可读写。

**语法**

**express.RevisionsView**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例在显示文档的原始版本和最终版本之间切换。本示例假定活动窗口中的文档包含由一个或多个审阅者所做的修订，并且修订显示于气球中。*/
function ToggleRevView(){
    if(ActiveWindow.View.RevisionsMode == wdBalloonRevisions){
        if(ActiveWindow.View.RevisionsView == wdRevisionsViewFinal){
            ActiveWindow.View.RevisionsView = wdRevisionsViewOriginal
        }
        else{
            ActiveWindow.View.RevisionsView = wdRevisionsViewFinal
        }
    }
}
```

### View.SeekView

返回或设置在页面视图中显示的文档元素。 **WdSeekView** 类型，可读写。

**语法**

**express.SeekView**

\*express是一个代表 View 对象的变量。

**说明**

如果当前视图不是页面视图，该属性将产生一个错误。

**示例**

``` JavaScript
/*如果活动文档有脚注，则本示例在页面视图中显示它。*/
function test()
{
if(ActiveDocument.Footnotes.Count >= 1){
    ActiveDocument.ActiveWindow.View.Type = wdPrintView
    ActiveDocument.ActiveWindow.View.SeekView = wdSeekFootnotes
}
}
```

``` JavaScript
/*本示例显示当前节中第一页的页脚。*/
function test()
{
ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = true
ActiveDocument.ActiveWindow.View.Type = wdPrintView
ActiveDocument.ActiveWindow.View.SeekView = wdSeekFirstPageFooter
}
```

``` JavaScript
/*如果在页面视图中选定脚注或尾注区域，本示例可切换到主控文档。*/
function test()
{
let myView = ActiveDocument.ActiveWindow.View

if(myView.SeekView == wdSeekFootnotes || myView.SeekView == wdSeekEndnotes){
    myView.SeekView = wdSeekMainDocument
}
}
```

### View.ShadeEditableRanges

返回或设置一个 **Long** 类型的值，该值代表是否将底纹应用于文档中用户有权修改的区域。

**语法**

**express.ShadeEditableRanges**

\*express是一个代表 View 对象的变量。

**说明**

如果为 **True** ，则将底纹应用于文档中用户可以修改的区域。默认情况下，区域底纹处于打开状态。当区域底纹处于打开状态，或将该属性设置为 **True** 时， **ShadeEditableRanges** 属性返回的值为 65535。将 **ShadeEditableRanges** 属性设置为 **False** 时，返回的值为 0。该值仅用于代表属性值为 **True** 或 **False** ，没有其他意义。

**示例**

``` JavaScript
/*以下示例将底纹应用于用户有权修改的所有区域。*/
ActiveWindow.View.ShadeEditableRanges = true
```

### View.ShowAll

如果显示所有非打印字符（如隐藏文字、制表符、空格和段落标记），则该属性值为 **True** 。可读写 **Boolean** 类型。

**语法**

**express.ShowAll**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动窗口中显示所有非打印字符。*/
ActiveDocument.ActiveWindow.View.ShowAll = true
```

``` JavaScript
/*以下示例在第一个窗口中切换非打印字符的显示方式。*/
Windows.Item(1).View.ShowAll = !Windows.Item(1).View.ShowAll
```

### View.ShowAnimation

如果显示文本动画，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowAnimation**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例先激活活动窗口的文本动画，然后对选定内容应用闪烁文本动画。*/
function test()
{
ActiveDocument.ActiveWindow.View.ShowAnimation = true
Selection.Font.Animation = wdAnimationSparkleText
}
```

``` JavaScript
/*本示例关闭所有打开窗口中的字体动画。*/
function test()
{
for(let i = 1; i <= Windows.Count; i++){
    Windows.Item(i).View.ShowAnimation = false
}
}
```

### View.ShowBookmarks

如果在每个书签的首尾显示方括号，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowBookmarks**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例在所有窗口中的书签周围显示方括号。*/
function test()
{
for(let i = 1; i <= Windows.Count; i++){
    Windows.Item(i).View.ShowBookmarks = true
}
}
```

``` JavaScript
/*本示例用书签标记选定内容，在活动文档的每个书签周围显示方括号，然后折叠选定内容。*/
function test()
{
ActiveDocument.Bookmarks.Add("temp", Selection.Range)
ActiveDocument.ActiveWindow.View.ShowBookmarks = true
Selection.Collapse(wdCollapseStart)
}
```

### View.ShowComments

如果为 **True** ，则 WPS 在文档中显示批注。 **Boolean** 类型，可读写。

**语法**

**express.ShowComments**

\*express是一个代表 View 对象的变量。

**说明**

如果修订标记显示在右边距或左边距的气球中，则备注也显示在气球中。如果修订标记显示在文本中，则应用备注的文本被括在方括号中。当将鼠标指针置于括号中的文本时，相关的备注就会显示在鼠标指针上方的方形气球中。

**示例**

``` JavaScript
/*本示例隐藏活动文档中的所有备注。本示例假定活动窗口中的文档包括一个或多个备注。*/
function HideComments(){
    ActiveWindow.View.ShowComments = false
}
```

### View.ShowCropMarks

返回或设置 **Boolean** 值，该值表示是否在页面的角显示裁剪标记，以指示页边距的位置。可读/写。

**语法**

**express.ShowCropMarks**

\*express是一个代表 View 对象的变量。

**说明**

如果显示裁减标记，将禁止用户通过拖动裁减标记来更改边距。显示裁减标记只是为了指示边距在页面上的位置。此属性对应于 **“ WPS 选项”** 对话框的 **“高级”** 选项卡中的 **“裁减标记”** 复选框。

| 注释                                                               |
|--------------------------------------------------------------------|
| 在东亚语言中默认显示裁减标记，在其他所有语言中默认情况下将关闭它。 |

### View.ShowDrawings

如果在页面视图中显示由绘图工具创建的对象，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowDrawings**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例将活动窗口切换到页面视图，并显示由绘图工具创建的对象。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdPrintView
ActiveDocument.ActiveWindow.View.ShowDrawings = true
}
```

### View.ShowFieldCodes

如果显示域代码，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowFieldCodes**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例隐藏 Document1 窗口的域代码。*/
Windows.Item("Document1").View.ShowFieldCodes = false
```

``` JavaScript
/*本示例显示第一个窗口的域代码。*/
Windows.Item(1).View.ShowFieldCodes = true
```

``` JavaScript
/*本示例切换活动窗口的域代码。*/
ActiveDocument.ActiveWindow.View.ShowFieldCodes = !ActiveDocument.ActiveWindow.View.ShowFieldCodes
```

### View.ShowFirstLineOnly

如果在大纲视图中仅显示正文每段的第一行，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowFirstLineOnly**

\*express是一个代表 View 对象的变量。

**说明**

如果不是大纲或主控文档视图，该属性将产生一个错误。

**示例**

``` JavaScript
/*本示例将活动窗口切换到大纲视图，并且隐藏除了正文每段的第一行之外所有的文本。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdOutlineView
ActiveDocument.ActiveWindow.View.ShowFirstLineOnly = true
}
```

### View.ShowFormat

如果在大纲视图中显示字符格式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowFormat**

\*express是一个代表 View 对象的变量。

**说明**

如果不是大纲或主控文档视图，该属性将产生一个错误。

**示例**

``` JavaScript
/*本示例将活动窗口切换到大纲视图，并显示字符格式。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdOutlineView
ActiveDocument.ActiveWindow.View.ShowFormat = true
}
```

### View.ShowFormatChanges

如果为 **True** ，则在启用“修订”功能后，WPS 显示对文档所作的格式更改。 **Boolean** 类型，可读写。

**语法**

**express.ShowFormatChanges**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例隐藏对活动文档所作的格式更改。本示例假定在启用“修订”功能后，以对文档进行了格式更改。*/
function HideFormattingChanges(){
    ActiveWindow.View.ShowFormatChanges = false
}
```

### View.ShowHiddenText

如果显示具有隐藏格式的文本，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowHiddenText**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例隐藏每个窗口中具有隐藏格式的文本。*/
function test()
{
for(let i = 1; i <= Windows.Count; i++){
    Windows.Item(i).View.ShowHiddenText = false
}
}
```

``` JavaScript
/*本示例切换对隐藏文本的显示。*/
ActiveDocument.ActiveWindow.View.ShowHiddenText = !ActiveDocument.ActiveWindow.View.ShowHiddenText
```

### View.ShowHighlight

如果显示和打印文档中突出显示格式的文本，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowHighlight**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例切换对活动文档中突出显示文本的显示。*/
ActiveDocument.ActiveWindow.View.ShowHighlight = !ActiveDocument.ActiveWindow.View.ShowHighlight
```

``` JavaScript
/*本示例打印活动文档，忽略突出显示格式。*/
function test()
{
ActiveDocument.ActiveWindow.View.ShowHighlight = false
ActiveDocument.PrintOut()
}
```

### View.ShowHyphens

如果为 **True** ，则显示可选连字符，该连字符指明在何处断开处于行尾的字。 **Boolean** 类型，可读写。

**语法**

**express.ShowHyphens**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例在选定内容之前插入一个可选连字符，然后在活动窗口显示可选连字符。*/
function test()
{
Selection.InsertBefore(String.fromCharCode(31))
ActiveDocument.ActiveWindow.View.ShowHyphens = true
}
```

### View.ShowInkAnnotations

返回或设置一个 **Boolean** 类型的值，该值表示显示还是隐藏手写墨迹注释。如果值为 **True** ，则显示墨迹注释。如果值为 **False** ，则隐藏墨迹注释。

**语法**

**express.ShowInkAnnotations**

\*express是一个代表 View 对象的变量。

**说明**

若要处理墨迹注释，必须在 Tablet 计算机上运行 WPS 。有关向文档添加手写墨迹注释的详细信息，请参阅“WPS 帮助”中的“使用墨迹注释标记文档”。

**示例**

``` JavaScript
/*以下示例显示活动文档中的所有手写墨迹注释。*/
ActiveDocument.ActiveWindow.View.ShowInkAnnotations = true
```

### View.ShowInsertionsAndDeletions

如果为 **True** ，则 WPS 显示用“修订”功能在文档中插入和删除的文本。 **Boolean** 类型，可读写。

**语法**

**express.ShowInsertionsAndDeletions**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例隐藏对文档所作的插入和删除修订。本示例假定活动窗口中的文档包含一个或多个审阅者所作的修订。*/
function HideInsertDelete(){
    ActiveWindow.View.ShowInsertionsAndDeletions = false
}
```

### View.ShowMainTextLayer

如果在显示页眉和页脚区域时同时显示指定文档的文字，则该属性值为 **True** 。该属性等效于 **“页眉和页脚”** 工具栏上的 **“显示/隐藏文档文字”** 按钮。 **Boolean** 类型，可读写。

**语法**

**express.ShowMainTextLayer**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例在活动窗口显示文档页眉并隐藏文档正文。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdPrintView
ActiveDocument.ActiveWindow.View.SeekView = wdSeekCurrentPageHeader
ActiveDocument.ActiveWindow.View.ShowMainTextLayer = false
}
```

### View.ShowMarkupAreaHighlight

返回或设置 **Boolean** 值，该值表示显示修订和注释气球的标记区域是否带底纹。可读/写。

**语法**

**express.ShowMarkupAreaHighlight**

\*express是一个代表 View 对象的变量。

**说明**

该属性对应于 **“审阅”** 功能区上 **“显示标记”** 菜单中的 **“标记区域突出显示”** 选项。

### View.ShowObjectAnchors

如果在可在页面视图中定位的项目附近显示对象锁定标记，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowObjectAnchors**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例给选定内容添加一个框架，将活动窗口切换到页面视图，并显示框架对象的锁定标记。*/
function test()
{
Selection.Frames.Add(Selection.Range).LockAnchor = true
ActiveDocument.ActiveWindow.View.Type = wdPrintView
ActiveDocument.ActiveWindow.View.ShowObjectAnchors = true
}
```

### View.ShowOptionalBreaks

如果 WPS 显示可选换行符，则该属性值为 **True** 。 **Boolean** 类型，可读写

**语法**

**express.ShowOptionalBreaks**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例显示活动窗口中的可选换行符。*/
ActiveDocument.ActiveWindow.View.ShowOptionalBreaks = true
```

### View.ShowOtherAuthors

如果应在文档中显示其他作者，则该属性值为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.ShowOtherAuthors**

\*express是一个代表 View 对象的变量。

**说明**

默认情况下，使用 WPS 2015 时，其他作者正在进行编辑或者已经从文档的某一部分隔离。在进行共同创作时使用 **ShowOtherAuthors** 属性可切换显示其他作者。

### View.ShowParagraphs

如果显示段落标记，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowParagraphs**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例隐藏活动窗口中的段落标记。*/
ActiveDocument.ActiveWindow.View.ShowParagraphs = false
```

### View.ShowPicturePlaceHolders

如果将空白方框作为图片的占位符进行显示，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowPicturePlaceHolders**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例在活动文档中插入图片，并在活动窗口中显示图片占位符。*/
function test()
{
Selection.Collapse(wdCollapseStart)
ActiveDocument.InlineShapes.AddPicture("C:\\anime.jpg", null, null, Selection.Range)
ActiveDocument.ActiveWindow.View.ShowPicturePlaceHolders = true
}
```

### View.ShowRevisionsAndComments

如果为 **True** ，则 WPS 显示使用“修订”功能对文档所作的修订和批注。 **Boolean** 类型，可读写。

**语法**

**express.ShowRevisionsAndComments**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例隐藏文档中的修订和备注。本示例假定活动窗口中的文档包括一个或多个审阅者所作的修订。*/
function ShowRevsComments(){
    ActiveWindow.View.ShowRevisionsAndComments = false
}
```

### View.ShowSpaces

如果显示空格字符，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowSpaces**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例在选定内容之前插入空格，并显示活动窗口中的空格字符。*/
function test()
{
Selection.InsertBefore("    ")
ActiveDocument.ActiveWindow.View.ShowSpaces = true
}
```

### View.ShowTabs

如果显示制表符，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowTabs**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例在选定内容之前插入一个制表符，并显示 Document2 窗口中的制表符。*/
function test()
{
Windows.Item("Document2").Activate()
Windows.Item("Document2").View.ShowTabs = true
Selection.InsertBefore("\t")
Selection.Collapse(wdCollapseEnd)
}
```

``` JavaScript
/*本示例拆分活动窗口，显示第一个窗格中的制表符，隐藏第二个窗格中的制表符。*/
function test()
{
ActiveDocument.ActiveWindow.Split = true
ActiveDocument.ActiveWindow.Panes.Item(1).View.ShowTabs = true
ActiveDocument.ActiveWindow.Panes.Item(2).View.ShowTabs = false
}
```

### View.ShowTextBoundaries

如果在页面视图中显示页边距、文本列、对象和框架周围的虚线，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowTextBoundaries**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例将活动窗口切换到页面视图，并显示文本边界线。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdPrintView
ActiveDocument.ActiveWindow.View.ShowTextBoundaries = true
}
```

### View.ShowXMLMarkup

返回一个 **Long** 类型的值，该值表示是否在文档中显示 XML 标记。

**语法**

**express.ShowXMLMarkup**

\*express是一个代表 View 对象的变量。

**说明**

如果为 **True** ，则表示显示标记。如果为 **False** ，则表示隐藏标记。使用 **wdToggle** ，可在显示和隐藏 XML 标记之间切换。

**示例**

``` JavaScript
/*以下示例在活动文档中切换显示和隐藏 XML 标记。*/
ActiveDocument.ActiveWindow.View.ShowXMLMarkup = wdToggle
```

### View.SplitSpecial

返回或设置活动窗口的窗格。 **WdSpecialPane** 类型，可读写。

**语法**

**express.SplitSpecial**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例将在活动窗口的一个单独窗格中显示主页脚。*/
ActiveDocument.ActiveWindow.View.SplitSpecial = wdPanePrimaryFooter
```

``` JavaScript
/*本示例将一个脚注添至活动文档，并在活动窗口的一个单独的窗格中显示所有脚注。*/
function test()
{
ActiveDocument.Footnotes.Add(Selection.Range,"1","Footnote text")
ActiveDocument.ActiveWindow.View.Type = wdNormalView
ActiveDocument.ActiveWindow.View.SplitSpecial = wdPaneFootnotes
}
```

### View.TableGridlines

如果显示表格虚框，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.TableGridlines**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例显示活动窗口中的表格虚框。*/
ActiveDocument.ActiveWindow.View.TableGridlines = true
```

``` JavaScript
/*本示例显示与 Windows 集合中窗口 1 相关的各个窗格的表格虚框。*/
function test()
{
let myPanes = Windows.Item(1).Panes
for(let i=1;i<=myPanes.Count;i++){
    myPanes.Item(i).View.TableGridlines = true
}
}
```

### View.Type

返回或设置视图的类型。可读/写 WdViewType 类型。

**语法**

**express.Type**

\*express是一个代表 View 对象的变量。

**说明**

**Type** 属性为当前处于大纲视图或主控文档视图的所有文档返回 wdMasterView 。除非预先在代码中明确设置，否则当前视图不会返回 **wdOutlineView** 。

要检查当前文档是否为大纲视图，请使用 **Type** 属性和 **Subdocuments** 集合的 **Count** 属性。如果 **Type** 属性返回 wdOutlineView 或 wdMasterView ，并且 **Count** 属性返回 0（零），则该文档为大纲视图。例如：

``` JavaScript
function VerifyOutlineView() {
    if(Application.ActiveWindow.View.Type == wdOutlineView || Application.ActiveWindow.View.Type ==wdMasterView) {
        if(Application.ActiveDocument.Subdocuments.Count == 0) {
            MsgBox("当前是大纲视图")
        }
    }
}
```

### View.WrapToWindow

如果文字在文档窗口的右边缘换行，而不是在右边距或文本栏的右边界，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.WrapToWindow**

\*express是一个代表 View 对象的变量。

**说明**

该属性在页面视图或 Web 版式视图中无效。

**示例**

``` JavaScript
/*本示例使文本换行，以容纳在活动窗口中。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdNormalView
ActiveDocument.ActiveWindow.View.WrapToWindow = true
}
```

### View.Zoom

返回一个 Zoom 对象，该对象代表指定视图的显示比例。

**语法**

**express.Zoom**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*本示例将所有打开窗口的显示比例更改为 125%*/
function wndBig(){
    for(let i = 1;i <= Application.Windows.Count;i++){
        Application.Windows.Item(i).View.Zoom.Percentage = 125
    }
}

/*本示例改变活动窗口的显示比例，以显示文本的全部宽度*/
Application.ActiveDocument.ActiveWindow.View.Zoom.PageFit = wdPageFitBestFit
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/View](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# XSLTransform 对象

## 简介

代表一个已注册的可扩展样式表语言转换 (XSLT)。

## 方法

| 序号 | 名称                           | 说明                                                      |
|:----:|:-------------------------------|:----------------------------------------------------------|
|  1   | [Delete](#XSLTransform.Delete) | 删除可用的可扩展样式表语言转换 (XSLT) 列表中指定的 XSLT。 |

## 属性

| 序号 | 名称                                     | 说明                                                                                      |
|:----:|:-----------------------------------------|:------------------------------------------------------------------------------------------|
|  1   | [Alias](#XSLTransform.Alias)             | 返回一个 String 类型的值，该值代表指定对象的显示名称。                                    |
|  2   | [Application](#XSLTransform.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                                            |
|  3   | [Creator](#XSLTransform.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。              |
|  4   | [ID](#XSLTransform.ID)                   | 返回 String 类型的值，该值包含分配给当前 XSLTransform 对象的 GUID。只读。                 |
|  5   | [Location](#XSLTransform.Location)       | 返回或设置 String 值，该值表示 XML 架构命名空间的 XSL 转换在架构库中的物理位置。可读/写。 |
|  6   | [Parent](#XSLTransform.Parent)           | 返回一个 Object 类型值，该值代表指定 XSLTransform 对象的父对象。                          |

## 成员方法

### XSLTransform.Delete

删除可用的可扩展样式表语言转换 (XSLT) 列表中指定的 XSLT。

**语法**

**express.Delete ()**

\*express是一个代表 XSLTransform 对象的变量。

## 成员属性

### XSLTransform.Alias

返回一个 **String** 类型的值，该值代表指定对象的显示名称。

**语法**

**express.Alias**

\*express是一个代表 XSLTransform 对象的变量。

**示例**

``` JavaScript
/*以下示例显示活动文档所附加的第一个架构的显示名称。*/
MsgBox(Application.XMLNamespaces.Item(1).Alias)
```

### XSLTransform.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 XSLTransform 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### XSLTransform.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 XSLTransform 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### XSLTransform.ID

返回 **String** 类型的值，该值包含分配给当前 XSLTransform 对象的 GUID。只读。

**语法**

**express.ID**

\*express是一个代表 XSLTransform 对象的变量。

### XSLTransform.Location

返回或设置 **String** 值，该值表示 XML 架构命名空间的 XSL 转换在架构库中的物理位置。可读/写。

**语法**

**express.Location**

\*express是一个代表 XSLTransform 对象的变量。

### XSLTransform.Parent

返回一个 **Object** 类型值，该值代表指定 **XSLTransform** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 XSLTransform 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/XSLTransform](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Fonts 对象

## 简介

指定的演示文稿中所有 Font 对象的集合。

## 说明

每个 Font 对象代表演示文稿中使用的一种字体。

“Genigraphics 向导”使用 Fonts 集合来决定，当 Genigraphics 转换幻灯片时是否在指定演示文稿中有任何字体不受支持。如果只需要设置特定项目符号或文本范围的字符格式，请使用 Font 属性为项目符号或文本范围返回 Font 对象。

“Genigraphics 向导”使用户可直接将演示文稿传输到 Genigraphics，以便转换成胶片幻灯片、投影机透明胶片或其他专用媒体格式。有关 Genigraphics 所提供的服务的详细信息，请访问 Genigraphics 网站：http://www.genigraphics.com/。此服务仅在美国国内可用。

``` JavaScript
/*使用 Fonts 属性返回 Fonts 集合。以下示例显示活动演示文稿中使用的字体种数。*/
alert(Application.ActivePresentation.Fonts.Count)
```

``` JavaScript
/*使用 Fonts(index) 返回单个 Font 对象，其中 index 是字体名称或索引号。以下示例检查活动演示文稿的第一种字体是否已嵌入。*/
function test() {
    if(Application.ActivePresentation.Fonts.Item(1).Embedded == true) {
        alert("Font 1 is embedded")
    }
}
```

## 方法

| 序号 | 名称                      | 说明                                       |
|:----:|:--------------------------|:-------------------------------------------|
|  1   | [Item](#Fonts.Item)       | 从指定集合中返回单个对象。返回值类型为Font |
|  2   | [Replace](#Fonts.Replace) | 替换 Fonts 集合中的一种字体。              |

## 属性

| 序号 | 名称                              | 说明                                                    |
|:----:|:----------------------------------|:--------------------------------------------------------|
|  1   | [Application](#Fonts.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#Fonts.Count)             | 返回指定集合中的对象数目。                              |

## 成员方法

### Fonts.Item

从指定集合中返回单个对象。返回值类型为Font

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Fonts 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

### Fonts.Replace

替换 **Fonts** 集合中的一种字体。

**语法**

**express.Replace ( *Original* , *Replacement* )**

\*express是一个代表 Fonts 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型 | 说明                 |
|---------------|-----------|----------|----------------------|
| *Original*    | 必选      | String   | 要替换的字体的名称。 |
| *Replacement* | 必选      | String   | 替换字体的名称。     |

**示例**

``` JavaScript
/*本示例将当前演示文稿中的 Times New Roman 字体替换为 Courier 字体。*/
function test() {
    Application.ActivePresentation.Fonts.Replace("Times New Roman"，"Courier")
}
```

## 成员属性

### Fonts.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Fonts 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(pptPre) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test() {
  let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
  for(let i =1;i<= Application.ActivePresentation.Slides.Item(1).Shapes.Count;i++) {
      if(shpOle.Item(i).Type == msoLinkedOLEObject) {
          alert(shpOle.Item(i).OLEFormat.Application.Name)
      }                                 
  }
}
```

### Fonts.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 Fonts 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除窗口1外的所有窗口。*/
function test() {
  for(let i=2; i <= Application.Windows.Count; i++){
      Application.Windows.Item(i).Close()
  }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Fonts](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Selection 对象

## 简介

代表指定文档窗口中的选定范围。每次在幻灯片视图中更改幻灯片时，Selection 对象将被删除（Type 属性将返回 ppSelectionNone）。

## 方法

| 序号 | 名称                            | 说明                             |
|:----:|:--------------------------------|:---------------------------------|
|  1   | [Copy](#Selection.Copy)         | 将指定对象复制到剪贴板。         |
|  2   | [Cut](#Selection.Cut)           | 删除指定对象并将其放到剪贴板中。 |
|  3   | [Delete](#Selection.Delete)     | 删除指定的对象。                 |
|  4   | [Unselect](#Selection.Unselect) | 取消当前选择。                   |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                                                                           |
|:----:|:----------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Selection.Application)         | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                        |
|  2   | [ChildShapeRange](#Selection.ChildShapeRange) | 返回一个代表选定范围的子形状的 ShapeRange 对象。                                                                                                                                               |
|  3   | [ShapeRange](#Selection.ShapeRange)           | 返回一个 ShapeRange 对象，该对象代表指定幻灯片上所有选定的幻灯片对象。其范围可包括幻灯片上的图形、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号占位符以及日期和时间对象。只读。 |
|  4   | [SlidRange](#Selection.SlidRange)             | 返回一个 SlideRange 对象，该对象代表选定的幻灯片范围。只读。                                                                                                                                   |
|  5   | [TextRange](#Selection.TextRange)             | 返回一个 TextRange 对象，该对象代表选定文本（ Selection 对象）或指定文本框（ TextFrame 对象）中的文本。                                                                                        |
|  6   | [Type](#Selection.Type)                       | 返回一个 PpSelectionType 常量，该常量代表选定范围中对象的类型。只读。                                                                                                                          |

## 成员方法

### Selection.Copy

将指定对象复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 Selection 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

``` JavaScript
/*本示例在第一个窗口中拷贝选定幻灯片到剪贴板然后将其粘贴到第二窗口的视图中。如果剪贴板的内容不能粘贴到第二个窗口的视图中（例如，如果试图将一个形状粘贴到幻灯片浏览视图中），则本示例失败*/
function test(){
    Application.Windows.Item(1).Selection.Copy()
    Application.Windows.Item(2).View.Paste()
}

/*本示例将当前演示文稿的第一张幻灯片复制到剪贴板*/
Application.ActivePresentation.Slides.Item(1).Copy()


/*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy()
```

### Selection.Cut

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut ()**

\*express是一个代表 Selection 对象的变量。

**示例**

``` JavaScript
/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中*/
 Application.Windows.Item(1).Selection.Cut()

/*本示例从当前演示文稿的第一张幻灯片中删除第一个和第二个形状，并将它们的副本放到剪贴板，然后将该副本粘贴到第二张幻灯片*/
function test(){
    let active = Application.ActivePresentation
        active.Slides.Item(1).Shapes.Range([1, 2]).Cut()
        active.Slides.Item(2).Shapes.Paste()
}


/*本示例从当前演示文稿中删除第一张幻灯片并将它的一个副本放到剪贴板中*/
Application.ActivePresentation.Slides.Item(1).Cut()


/*本示例删除当前演示文稿中第一张幻灯片的第一个形状中的文本，并将它的一个副本放到剪贴板中*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Cut()
```

### Selection.Delete

删除指定的对象。

**语法**

**express.Delete ()**

\*express是一个代表 Selection 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**示例**

``` JavaScript
/*本示例删除第一个窗口中选中的对象
Application.Windows.Item(1).Selection.Delete()
```

### Selection.Unselect

取消当前选择。

**语法**

**express.Unselect ()**

\*express是一个代表 Selection 对象的变量。

**示例**

``` JavaScript
/*本示例取消对第一个窗口中当前选定内容的选定*/
 Application.Windows.Item(1).Selection.Unselect()
```

## 成员属性

### Selection.Application

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 Selection 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/
function test(){
    let pptPres = Application.ActivePresentation
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称*/
function test(){
    let shpole =Application.ActivePresentation.Slides.Item(1).Shapes
    for(let i = 1; i <= shpole.Count; i++){
        if(shpole.Item(i).Type == msoLinkedOLEObject){
            MsgBox(shpole.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### Selection.ChildShapeRange

返回一个代表选定范围的子形状的 ShapeRange 对象。

**语法**

**express.ChildShapeRange**

\*express是一个代表 Selection 对象的变量。

### Selection.ShapeRange

返回一个 ShapeRange 对象，该对象代表指定幻灯片上所有选定的幻灯片对象。其范围可包括幻灯片上的图形、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号占位符以及日期和时间对象。只读。

**语法**

**express.ShapeRange**

\*express是一个代表 Selection 对象的变量。

**说明**

当演示文稿为普通视图、幻灯片视图或任一母版视图时，可以从选定范围中返回一个形状范围。

**示例**

``` JavaScript
/*本示例设置第一个窗口中所有选定形状的前景填充颜色*/
Application.Windows.Item(1).Selection.ShapeRange.Fill.ForeColor.RGB = (255,0,255)
```

### Selection.SlidRange

返回一个 SlideRange 对象，该对象代表选定的幻灯片范围。只读。

**语法**

**express.SlidRange**

\*express是一个代表 Selection 对象的变量。

**说明**

可以在幻灯片视图、幻灯片浏览视图、普通视图、备注页视图或大纲视图中构造幻灯片范围。在幻灯片视图中，SlideRange 返回当前放映的幻灯片。

**示例**

``` JavaScript
/*本示例设置第一个窗口中所有选定幻灯片的数量*/
Application.Windows.Item(1).Selection.SlideRange.Count
```

### Selection.TextRange

返回一个 TextRange 对象，该对象代表选定文本（ Selection 对象）或指定文本框（ TextFrame 对象）中的文本。

**语法**

**express.TextRange**

\*express是一个代表 Selection 对象的变量。

**说明**

当演示文稿在幻灯片视图、普通视图、大纲视图、备注页视图或任何母版视图时，可从选定的内容中创建文本范围。

### Selection.Type

返回一个 PpSelectionType 常量，该常量代表选定范围中对象的类型。只读。

**语法**

**express.Type**

\*express是一个代表 Selection 对象的变量。

**说明**

PpSelectionType 可以是下列 PpSelectionType 常量之一。 ppSelectionNone ppSelectionShapes ppSelectionSlides ppSelectionText

**示例**

``` JavaScript
/*以下示例将显示选中对象的类型*/
Application.Windows.Item(1).Selection.Type
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Selection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CommandBarComboBox 对象

## 简介

代表命令栏中的一个组合框控件。

## 说明

使用 Controls(index) 可返回一个 CommandBarComboBox 对象，其中 *index* 为该控件的索引号。请注意，该控件的 Type 属性必须是 msoControlEdit 、 msoControlDropdown 、 msoControlComboBox 、 msoControlButtonDropdown 、 msoControlSplitDropdown 、 msoControlOCXDropdown 、 msoControlGraphicCombo 或 msoControlGraphicDropdown 。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

## 方法

| 序号 | 名称                                         | 说明                                                                                                                                                        |
|:----:|:---------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddItem](#CommandBarComboBox.AddItem)       | 向指定命令栏组合框控件添加列表项。该组合框控件必须是自定义控件，并且必须是下拉列表框或组合框。                                                              |
|  2   | [Clear](#CommandBarComboBox.Clear)           | 从命令栏组合框控件（下拉列表框或组合框）中删除所有列表项。                                                                                                  |
|  3   | [Copy](#CommandBarComboBox.Copy)             | 将一个命令栏组合框控件复制到已有的命令栏中。                                                                                                                |
|  4   | [Delete](#CommandBarComboBox.Delete)         | 将 CommandBarCombo 控件对象从其集合中删除。在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |
|  5   | [Execute](#CommandBarComboBox.Execute)       | 运行分配给指定 CommandBarComboBox 控件的过程或内置命令。                                                                                                    |
|  6   | [Move](#CommandBarComboBox.Move)             | 将指定的控件移动到已有的命令栏。                                                                                                                            |
|  7   | [RemoveItem](#CommandBarComboBox.RemoveItem) | 从 CommandBarComboBox 控件中删除一项。                                                                                                                      |
|  8   | [Reset](#CommandBarComboBox.Reset)           | 将内置命令栏重置为其默认配置，或将内置 CommandBarComboBox 控件重置为其初始功能和图符。                                                                      |
|  9   | [SetFocus](#CommandBarComboBox.SetFocus)     | 将键盘的焦点移到指定 CommandBarComboBox 控件。如果该控件被禁用或不可见，那么此方法将失败。                                                                  |

## 属性

| 序号 | 名称                                                       | 说明                                                                                                                                                 |
|:----:|:-----------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CommandBarComboBox.Application)             | 获取一个 Application 对象，代表 CommandBarComboBox 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。            |
|  2   | [BeginGroup](#CommandBarComboBox.BeginGroup)               | 如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 True。可读写。                                                                              |
|  3   | [BuiltIn](#CommandBarComboBox.BuiltIn)                     | 如果指定的命令栏控件是容器应用程序的内置控件，则获取 True 。如果是自定义控件或是已设置了 OnAction 属性的内置控件，则返回 False 。只读。              |
|  4   | [Caption](#CommandBarComboBox.Caption)                     | 获取或设置命令栏控件的标题文字。可读写。                                                                                                             |
|  5   | [Creator](#CommandBarComboBox.Creator)                     | 获取一个 32 位整数，指示创建 CommandBarComboBox 对象时所使用的应用程序。只读。                                                                       |
|  6   | [DescriptionText](#CommandBarComboBox.DescriptionText)     | 获取或设置命令栏组合框控件的说明。可读写。                                                                                                           |
|  7   | [DropDownLines](#CommandBarComboBox.DropDownLines)         | 获取或设置命令栏组合框控件中的行数。该组合框控件必须为自定义控件并且是下拉列表框或组合框。可读写。                                                   |
|  8   | [DropDownWidth](#CommandBarComboBox.DropDownWidth)         | 获取或设置指定命令栏组合框控件的列表宽度（以像素为单位）。可读写。                                                                                   |
|  9   | [Enabled](#CommandBarComboBox.Enabled)                     | 获取或设置用于指定是否启用 CommandBarComboBox 的 Boolean 值。可读写。                                                                                |
|  10  | [Height](#CommandBarComboBox.Height)                       | 获取或设置 CommandBarComboBox 控件的高度。可读写。                                                                                                   |
|  11  | [HelpContextId](#CommandBarComboBox.HelpContextId)         | 获取或设置附加于 CommandBarComboBox 控件的“帮助”主题的帮助上下文 ID 号。可读写。                                                                     |
|  12  | [HelpFile](#CommandBarComboBox.HelpFile)                   | 获取或设置附加于 CommandBarComboBox 控件的“帮助”主题的文件名。可读写。                                                                               |
|  13  | [Id](#CommandBarComboBox.Id)                               | 获取内置的 CommandBarComboBox 控件的 ID。只读。                                                                                                      |
|  14  | [Index](#CommandBarComboBox.Index)                         | 获取一个 Long 类型的值，该值代表集合中 CommandBarComboBox 对象的索引号。只读。                                                                       |
|  15  | [IsPriorityDropped](#CommandBarComboBox.IsPriorityDropped) | 如果当前已根据使用统计信息和布局空间将控件从菜单或工具栏中删除，则返回 True 。（请注意，这与通过 Visible 属性设置的控件可见性不同）。只读。          |
|  16  | [Left](#CommandBarComboBox.Left)                           | 获取 CommandBarComboBox 控件相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。                                           |
|  17  | [List](#CommandBarComboBox.List)                           | 获取或设置 CommandBarComboBox 控件中的某一项。可读/写。                                                                                              |
|  18  | [ListCount](#CommandBarComboBox.ListCount)                 | 获取 CommandBarComboBox 控件中列表项的数目。只读。                                                                                                   |
|  19  | [ListHeaderCount](#CommandBarComboBox.ListHeaderCount)     | 获取或设置出现在分隔行之上的 CommandBarComboBox 控件中的列表项数目。可读写。                                                                         |
|  20  | [ListIndex](#CommandBarComboBox.ListIndex)                 | 获取或设置 CommandBarComboBox 控件的列表部分中选定项的索引号。如果没有选择列表中的任何项，此属性将返回 0。可读写。                                   |
|  21  | [OLEUsage](#CommandBarComboBox.OLEUsage)                   | 获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 CommandBarComboBox 控件。可读写。                  |
|  22  | [OnAction](#CommandBarComboBox.OnAction)                   | 获取或设置一个 Visual Basic 过程的名称，该过程在用户单击或更改 CommandBarComboBox 控件的值时运行。可读写。                                           |
|  23  | [Parameter](#CommandBarComboBox.Parameter)                 | 获取或设置一个应用程序可用于执行 CommandBarComboBox 控件中命令的字符串。可读写。                                                                     |
|  24  | [Parent](#CommandBarComboBox.Parent)                       | 获取 CommandBarComboBox 对象的 Parent 对象。只读。                                                                                                   |
|  25  | [Priority](#CommandBarComboBox.Priority)                   | 获取或设置 CommandBarComboBox 控件的优先级。对于固定命令栏，当一行中无法容纳其中的所有控件时，控件的优先级将决定该控件是否会从命令栏中删除。可读写。 |
|  26  | [Style](#CommandBarComboBox.Style)                         | 获取或设置 CommandBarComboBox 控件的显示方式。该属性值可设置为以下 MsoComboStyle 常量之一： msoComboLabel 或 msoComboNormal 。可读写。               |
|  27  | [Tag](#CommandBarComboBox.Tag)                             | 获取或设置有关 CommandBarComboBox 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。                                             |
|  28  | [Text](#CommandBarComboBox.Text)                           | 获取或设置 CommandBarComboBox 控件的显示或编辑区中的文字。可读写。                                                                                   |
|  29  | [TooltipText](#CommandBarComboBox.TooltipText)             | 获取或设置 CommandBarComboBox 的 “屏幕提示” 中显示的文本。可读写。                                                                                   |
|  30  | [Top](#CommandBarComboBox.Top)                             | 获取指定的 CommandBarComboBox 控件顶边到屏幕顶边的距离（以像素为单位）。只读。                                                                       |
|  31  | [Type](#CommandBarComboBox.Type)                           | 获取 CommandBarComboBox 控件的类型。只读。                                                                                                           |
|  32  | [Visible](#CommandBarComboBox.Visible)                     | 获取或设置 CommandBarComboBox 控件的 Visible 属性。如果 CommandBarControl 可见，则此属性为 True 。可读写。                                           |
|  33  | [Width](#CommandBarComboBox.Width)                         | 获取或设置指定的 CommandBarComboBox 控件的宽度（以像素为单位）。可读写。                                                                             |

## 成员方法

### CommandBarComboBox.AddItem

向指定命令栏组合框控件添加列表项。该组合框控件必须是自定义控件，并且必须是下拉列表框或组合框。

**语法**

**express.AddItem ( *Text* , *Index* )**

\*express是一个代表 CommandBarComboBox 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                       |
|---------|-----------|----------|------------------------------------------------------------|
| *Text*  | 必选      | String   | 要添加到控件的文字。                                       |
| *Index* | 可选      | Variant  | 项在列表中的位置。如果省略该参数，项会被添加到列表的尾部。 |

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

| 注释                                     |
|------------------------------------------|
| 此方法不能用于编辑框或内置的组合框控件。 |

**示例**

``` JavaScript
//本示例在命令栏中添加一个组合框控件，然后在该控件中添加两项。本示例还设置项的行数和组合框的宽度。 
function test()
{
    let myBar = Application.CommandBars.Item("Custom")
    let myControl = myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)
    myControl.AddItem("First Item", 1)
    myControl.AddItem("Second Item", 2)
    myControl.DropDownLines = 3
    myControl.DropDownWidth = 75
    myControl.ListHeaderCount = 0
}
```

### CommandBarComboBox.Clear

从命令栏组合框控件（下拉列表框或组合框）中删除所有列表项。

**语法**

**express.Clear ()**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                     |
|--------------------------------------------------------------------------|
| 如果将此方法应用于内置命令栏控件（WPS Office附带的控件），此方法将失败。 |

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例将检查名为“自定义栏”的命令栏上组合框控件中的项目数。如果组合框的列表中的项目数少于 3，那么本示例清除该列表，在其中添加新的第一项，然后将该新项目显示为组合框控件的默认值。
function test()
{
    let myBar = Application.CommandBars.Item("Custom Bar")
    let myControl = myBar.Controls.Item(Application.Enum.msoControlComboBox)
    myControl.AddItem("First Item", 1)
    myControl.AddItem("Second Item", 2)
    if(myControl.ListCount < 3)
    {
        myControl.Clear()
        myControl.AddItem("New Item", 1)
    }
}
```

### CommandBarComboBox.Copy

将一个命令栏组合框控件复制到已有的命令栏中。

**语法**

**express.Copy ( *Bar* , *Before* )**

\*express是一个代表 CommandBarComboBox 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                                                 |
|----------|-----------|----------|----------------------------------------------------------------------------------------------------------------------|
| *Bar*    | 可选      | Variant  | 一个代表目标命令栏的 CommandBar 对象。如果省略此参数，控件将被复制到它原来所在的命令栏中。                           |
| *Before* | 可选      | Variant  | 一个指示新控件在命令栏上位置的数值。新控件将插入到位于此位置的控件之前。如果省略此参数，控件将被复制到命令栏的末尾。 |

**返回值**

CommandBarControl

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Delete

将 **CommandBarCombo** 控件对象从其集合中删除。在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。

**语法**

**express.Delete ( *Temporary* )**

\*express是一个代表 CommandBarComboBox 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                          |
|-------------|-----------|----------|-------------------------------------------------------------------------------|
| *Temporary* | 可选      | Variant  | 如果为 True，则从当前会话中删除此控件。应用程序在下次会话中将再次显示该控件。 |

### CommandBarComboBox.Execute

运行分配给指定 **CommandBarComboBox** 控件的过程或内置命令。

**语法**

**express.Execute ()**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本 ET 示例创建一个命令栏，然后向其中添加内置命令栏按钮控件。该按钮将执行 ET AutoSum 函数。本示例使用 Execute 方法在显示该命令栏时计算选定单元格区域的总计。
function test()
{
    let cbrCustBar = Application.CommandBars.Add("Custom")
    let ctlAutoSum = cbrCustBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("AutoSum").Id)
    cbrCustBar.Visible = true 
    ctlAutoSum.Execute()
}
```

### CommandBarComboBox.Move

将指定的控件移动到已有的命令栏。

**语法**

**express.Move ( *Bar* , *Before* )**

\*express是一个代表 CommandBarComboBox 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                          |
|----------|-----------|----------|-----------------------------------------------------------------------------------------------|
| *Bar*    | 可选      | Variant  | 一个代表控件的目标命令栏的 Command 对象。如果忽略该参数，则控件将移动到当前所在命令栏的末端。 |
| *Before* | 可选      | Variant  | 表示控件位置的数字。控件将插到该位置的控件之前。如果忽略该参数，控件插入到同一命令栏。        |

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例将名为 Custom 的命令栏上的第一个组合框控件移动到该命令栏上的第七个控件之前。本示例将 Tag 设置为“Selection box”并赋予控件较低的优先级，以便在一行容纳不下所有控件时将其隐藏。
function test()
{
    let allcontrols = Application.CommandBars.Item("Custom").Controls
    for(let i = 1; i <= allControls.Count; i++)
    {
        if(allControls.Item(i).Type == Application.Enum.msoControlComboBox)
        {
            allControls.Item(i).Move(undefined,7)
            allControls.Item(i).Tag = "Selection box"
            allControls.Item(i).Priority = 5
        }
    }
}
```

### CommandBarComboBox.RemoveItem

从 **CommandBarComboBox** 控件中删除一项。

**语法**

**express.RemoveItem ()**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

| 注释                       |
|----------------------------|
| 本属性只能用于列表框控件。 |

**示例**

``` JavaScript
//本示例确定指定组合框所含的列表项是否多于三项。如果多于三项，那么本示例将删除第二项，更改样式并设置一个新值。此外，本示例还设置了父对象（即 CommandBarControl 对象）的 Tag 属性以表明该列表发生了变化。
function test()
{
    let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop,undefined, true)
    myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)
    myBar.Visible = true 
    let myControls = Application.CommandBars.Item("Custom").Controls.Item(1)
    myControls.AddItem("Get Stock Quote", 1)
    myControls.AddItem("View Chart", 2)
    myControls.AddItem("View Fundamentals", 3)
    myControls.AddItem("View News", 4)
    myControls.Caption = "Stock Data"
    myControls.DescriptionText = "View Data For Stock"
    let myControl = myBar.Controls.Item(1)
    if(myControl.ListCount > 3)
    {
       myControl.RemoveItem(2)
       myControl.Style = Application.Enum.msoComboNormal
       myControl.Text = "New Default"
       let ctrl = myControl.Parent
    }
}
```

### CommandBarComboBox.Reset

将内置命令栏重置为其默认配置，或将内置 **CommandBarComboBox** 控件重置为其初始功能和图符。

**语法**

**express.Reset ()**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

重置一个内置控件将恢复该控件原来对应的动作，并将各属性恢复成初始状态。重置一个内置命令栏将删除其中的自定义控件并恢复其内置控件。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例自定义一个命令栏组合框。首先，将组合框重置为其默认状态。然后将两个行项添加到组合框，并设置各种属性。
function test()
{
    let combo = Application.CommandBars.Item("Custom").Controls.Item(2)
    combo.Reset()
    combo.AddItem("First Item", 1)
    combo.AddItem("Second Item", 2)
    combo.DropDownLines = 3
    combo.DropDownWidth = 75
    combo.ListIndex = 0

}
```

### CommandBarComboBox.SetFocus

将键盘的焦点移到指定 **CommandBarComboBox** 控件。如果该控件被禁用或不可见，那么此方法将失败。

**语法**

**express.SetFocus ()**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例创建一个名为“Custom”的命令栏，并在其中添加一个 ComboBox 控件和一个 Button 控件，然后使用 SetFocus 方法为 ComboBox 控件设置焦点。
function test()
{
    let focusBar = Application.CommandBars.Add("Custom")
    CommandBars.Item("Custom").Visible = true 
    CommandBars.Item("Custom").Position = Application.Enum.msoBarTop
                                        
    let testComboBox = Application.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlComboBox, 1)
    testComboBox.AddItem("First Item", 1)
    testComboBox.AddItem("Second Item", 2)
    let testButton = Application.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlButton)
    testButton.FaceId = 17
    // Set the focus to the combo box.
    testComboBox.SetFocus()
}
```

## 成员属性

### CommandBarComboBox.Application

获取一个 **Application** 对象，代表 **CommandBarComboBox** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.BeginGroup

如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 True。可读写。

**语法**

**express.BeginGroup**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.BuiltIn

如果指定的命令栏控件是容器应用程序的内置控件，则获取 **True** 。如果是自定义控件或是已设置了 **OnAction** 属性的内置控件，则返回 **False** 。只读。

**语法**

**express.BuiltIn**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Caption

获取或设置命令栏控件的标题文字。可读写。

**语法**

**express.Caption**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                     |
|------------------------------------------|
| 控件的标题也可作为默认的“屏幕提示”显示。 |

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Creator

获取一个 32 位整数，指示创建 **CommandBarComboBox** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.DescriptionText

获取或设置命令栏组合框控件的说明。可读写。

**语法**

**express.DescriptionText**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

该说明不显示给用户，但有助于其他开发者将控件行为归档。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.DropDownLines

获取或设置命令栏组合框控件中的行数。该组合框控件必须为自定义控件并且是下拉列表框或组合框。可读写。

**语法**

**express.DropDownLines**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

如果将此属性设置为 0（零），则控件中的行数将取决于列表中的项数。

| 注释                                                                   |
|------------------------------------------------------------------------|
| 如果试图设置该属性的组合框控件是编辑框或内置组合框控件，则会出现错误。 |

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在命令栏“Custom”中添加一个包含两项的组合框控件。此外，本示例还设置项的行数，组合框的宽度，并将该组合框的默认值设置为空。 
function test()
{
    let myBar = Application.CommandBars.Item("Custom")
    let myControl = myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)                                  
    myControl.AddItem ("First Item", 1)
    myControl.AddItem ("Second Item", 2)
    myControl.DropDownLines = 3
    myControl.DropDownWidth = 75
    myControl.ListHeaderCount = 0
}
```

### CommandBarComboBox.DropDownWidth

获取或设置指定命令栏组合框控件的列表宽度（以像素为单位）。可读写。

**语法**

**express.DropDownWidth**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

如果将该属性设置为 -1，那么将根据组合框列表中最长的项来设置列表的宽度。如果将该属性设置为 0，那么将根据控件的宽度来设置列表的宽度。

| 注释                                             |
|--------------------------------------------------|
| 如果为内置的组合框控件设置该属性，则会出现错误。 |

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在命令栏“Custom”中添加一个包含两项的组合框控件。此外，本示例还设置项的行数，组合框的宽度，并将该组合框的默认值设置为空。 
function test()
{
    let myBar = Application.CommandBars.Item("Custom")
    let myControl = myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)                                  
    myControl.AddItem ("First Item", 1)
    myControl.AddItem ("Second Item", 2)
    myControl.DropDownLines = 3
    myControl.DropDownWidth = 75
    myControl.ListHeaderCount = 0
}
```

### CommandBarComboBox.Enabled

获取或设置用于指定是否启用 **CommandBarComboBox** 的 **Boolean** 值。可读写。

**语法**

**express.Enabled**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

对于命令栏，如果将该属性设置为 **True** ，则该命令栏的名称将出现在可用命令栏列表中。

对于内置控件，如果将 **Enabled** 属性设置为 **True** ，则将由应用程序确定其状态。但如果将该属性设置为 **False** ，则强行禁用该控件。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Height

获取或设置 **CommandBarComboBox** 控件的高度。可读写。

**语法**

**express.Height**

\*express是一个代表 CommandBarComboBox 对象的变量。

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
    let myControl = myBar.Controls.Add(Application.Enum.msoControlButton,CommandBars.Item("Standard").Controls.Item("Save").Id,null,null,true)
    myControl.Height = barHeight * 2
    myControl.Width = 50
    myBar.Visible = true
}
```

### CommandBarComboBox.HelpContextId

获取或设置附加于 **CommandBarComboBox** 控件的“帮助”主题的帮助上下文 ID 号。可读写。

**语法**

**express.HelpContextId**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

若要使用此属性，必须同时设置 HelpFile 属性。帮助主题响应 Shift+F1 按键操作。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//下面的示例添加了一个自定义命令栏以及跟踪股票的组合框。该示例也指定了在用户按下 Shift+F1 时所显示的组合框“帮助”主题。 
function test()
{
    let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)
    myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)
    myBar.Visible = true 
                      
    let cont = CommandBars.Item("Custom").Controls.Item(1)
    cont.AddItem("Get Stock Quote", 1)
    cont.AddItem("View Chart", 2)
    cont.AddItem("View Fundamentals", 3)
    cont.AddItem("View News", 4)
    cont.Caption = "Stock Data"
    cont.DescriptionText = "View Data For Stock"
    cont.HelpFile = "C:\\corphelp\\custom.hlp"
    cont.HelpContextID = 47
}
```

### CommandBarComboBox.HelpFile

获取或设置附加于 **CommandBarComboBox** 控件的“帮助”主题的文件名。可读写。

**语法**

**express.HelpFile**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

若要使用该属性，必须同时设置 HelpContextID 属性。帮助主题响应 Shift+F1 用户按键操作。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//下面的示例添加了一个自定义命令栏以及跟踪股票的组合框。该示例也指定了在用户按下 Shift+F1 时所显示的组合框“帮助”主题。 
function test()
{
    let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)
    myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)
    myBar.Visible = true 
                      
    let cont = CommandBars.Item("Custom").Controls.Item(1)
    cont.AddItem("Get Stock Quote", 1)
    cont.AddItem("View Chart", 2)
    cont.AddItem("View Fundamentals", 3)
    cont.AddItem("View News", 4)
    cont.Caption = "Stock Data"
    cont.DescriptionText = "View Data For Stock"
    cont.HelpFile = "C:\\corphelp\\custom.hlp"
    cont.HelpContextID = 47
}
```

### CommandBarComboBox.Id

获取内置的 **CommandBarComboBox** 控件的 ID。只读。

**语法**

**express.Id**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

控件的 ID 号决定了该控件的内置操作。所有自定义控件的 **Id** 属性的值都是 1。

| 注释                                                          |
|---------------------------------------------------------------|
| WPS Office Assistant 在 WPS Office system 2015 版中已被淘汰。 |

**示例**

``` JavaScript
//下面的示例将名为“Standard”的工具栏上每个控件的标题更改为该控件当前的 Id 属性值。 
function test()
{
    let ctl = Application.CommandBars.Item("Standard").Controls
    for(let i=1; i<=ctl.Count; i++)
    {
        ctl.Item(i).Caption = String(ctl.Id)
    }
}
```

### CommandBarComboBox.Index

获取一个 **Long** 类型的值，该值代表集合中 **CommandBarComboBox** 对象的索引号。只读。

**语法**

**express.Index**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

第一个命令栏控件的位置为 1。 **CommandBarControls** 集合中的分隔符不计算在内。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在名为“Custom”的命令栏中搜索 Id 值为 23 的控件。如果能找到这样的控件，并且其索引号大于 5，则将该控件放置在命令栏上的第一个位置上。
function test()
{
    let myBar = CommandBars.Item("Custom")
    let ctrl1 = myBar.FindControl(null,23)
    if (ctrl1 && ctrl1.Index > 5)
    {
        ctrl1.Move (null,1)
    }
}
```

### CommandBarComboBox.IsPriorityDropped

如果当前已根据使用统计信息和布局空间将控件从菜单或工具栏中删除，则返回 **True** 。（请注意，这与通过 **Visible** 属性设置的控件可见性不同）。只读。

**语法**

**express.IsPriorityDropped**

\*express是一个代表 CommandBarComboBox 对象的变量。

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

### CommandBarComboBox.Left

获取 **CommandBarComboBox** 控件相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。

**语法**

**express.Left**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.List

获取或设置 **CommandBarComboBox** 控件中的某一项。可读/写。

**语法**

**express.List**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

| 注释                                 |
|--------------------------------------|
| 对于内置的组合框控件，该属性为只读。 |

**示例**

### CommandBarComboBox.ListCount

获取 **CommandBarComboBox** 控件中列表项的数目。只读。

**语法**

**express.ListCount**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

### CommandBarComboBox.ListHeaderCount

获取或设置出现在分隔行之上的 **CommandBarComboBox** 控件中的列表项数目。可读写。

**语法**

**express.ListHeaderCount**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

**ListHeaderCount** 属性值为 -1 表示组合框控件中没有分隔行。

| 注释                                 |
|--------------------------------------|
| 对于内置的组合框控件，该属性为只读。 |

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

### CommandBarComboBox.ListIndex

获取或设置 **CommandBarComboBox** 控件的列表部分中选定项的索引号。如果没有选择列表中的任何项，此属性将返回 0。可读写。

**语法**

**express.ListIndex**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

设置 **ListIndex** 属性会使指定控件选择给定的项并在应用程序中执行与相应的操作。

| 注释                     |
|--------------------------|
| 此属性只能用于列表控件。 |

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

### CommandBarComboBox.OLEUsage

获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 **CommandBarComboBox** 控件。可读写。

**语法**

**express.OLEUsage**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

本属性允许用户在某 Office 应用程序与另一个 Office 应用程序合并时，指定加载应用程序的各个命令栏控件在该 Office 应用程序中的表示方法。如果客户端和服务器均提供命令栏，那么命令栏控件将逐个嵌入客户端。系统将从服务器中略去标记为只用于客户端（或既不用于客户端也不用于服务器）的自定义控件，而从客户端中略去标记为只用于服务器（或既不用于客户端也不用于服务器）的控件。剩下的控件再进行合并。

如果要合并的应用程序之一不是 Office 应用程序，则使用常规 OLE 菜单合并，该过程将由 OLEMenuGroup 属性控制。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.OnAction

获取或设置一个 Visual Basic 过程的名称，该过程在用户单击或更改 **CommandBarComboBox** 控件的值时运行。可读写。

**语法**

**express.OnAction**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

容器应用程序将确定该值是否是一个合法的宏名。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Parameter

获取或设置一个应用程序可用于执行 **CommandBarComboBox** 控件中命令的字符串。可读写。

**语法**

**express.Parameter**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

如果为内置控件设置一个指定的参数，那么应用程序将修改其默认的操作，前提是该应用程序能分析和使用这个新值。如果为自定义控件设置参数，那么该参数可用来为 Visual Basic 程序传送信息，或保存关于该控件的信息（类似于另一个 Tag 属性值）。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Parent

获取 **CommandBarComboBox** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Priority

获取或设置 **CommandBarComboBox** 控件的优先级。对于固定命令栏，当一行中无法容纳其中的所有控件时，控件的优先级将决定该控件是否会从命令栏中删除。可读写。

**语法**

**express.Priority**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

有效的优先级数为从 0（零）至 7 之间的数字，默认值为 3。优先级为 1 表示不能从工具栏中删除该控件。而优先级为其他值的控件将被忽略。

菜单项的命令栏控件不能使用 **Priority** 属性。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Style

获取或设置 **CommandBarComboBox** 控件的显示方式。该属性值可设置为以下 **MsoComboStyle** 常量之一： **msoComboLabel** 或 **msoComboNormal** 。可读写。

**语法**

**express.Style**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Tag

获取或设置有关 **CommandBarComboBox** 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。

**语法**

**express.Tag**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Text

获取或设置 **CommandBarComboBox** 控件的显示或编辑区中的文字。可读写。

**语法**

**express.Text**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例创建一个名为“Custom”的新命令栏，并向其添加一个包含有四个列表项的组合框。然后用 Text 属性将“Item 3”设置为默认列表项。 
function test()
{
    let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)
    myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)
    myBar.Visible = true
    let testComboBox = Application.CommandBars.Item("Custom").Controls.Item(1)
    testComboBox.Text = "Item 3"

}
```

### CommandBarComboBox.TooltipText

获取或设置 **CommandBarComboBox** 的 **“屏幕提示”** 中显示的文本。可读写。

**语法**

**express.TooltipText**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

默认情况下， **Caption** 属性的值将用作 **“屏幕提示”** 。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Top

获取指定的 **CommandBarComboBox** 控件顶边到屏幕顶边的距离（以像素为单位）。只读。

**语法**

**express.Top**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Type

获取 **CommandBarComboBox** 控件的类型。只读。

**语法**

**express.Type**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Visible

获取或设置 **CommandBarComboBox** 控件的 **Visible** 属性。如果 **CommandBarControl** 可见，则此属性为 **True** 。可读写。

**语法**

**express.Visible**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBarComboBox.Width

获取或设置指定的 **CommandBarComboBox** 控件的宽度（以像素为单位）。可读写。

**语法**

**express.Width**

\*express是一个代表 CommandBarComboBox 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CommandBarComboBox](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CommandBars 对象

## 简介

代表容器应用程序中命令栏的 CommandBar 对象的集合。

## 说明

在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。

使用 CommandBars 属性可返回 CommandBars 集合。下面的示例在 “立即” 窗口中显示每个菜单栏和工具栏的名称和本地名称，并显示一个值以指示该菜单栏或工具栏是否可见。

``` JavaScript
function test()
{
    for(let i=1;i<=CommandBars.Count;i++){
        let cbar = CommandBars.Item(i)
        alert(cbar.Name + "\t" + cbar.NameLocal + "\t" + cbar.Visible)
    }   
}
```

使用 Add 方法可向集合中添加一个新的命令栏。下面的示例创建一个名为“Custom1”的自定义工具栏并将其显示为浮动工具栏。

``` JavaScript
function test()
{
    let cbar1 = Application.ActiveDocument.CommandBars.Add("Custom1", msoBarFloating)
    cbar1.Visible = true
}
```

使用 enumName 可返回一个 CommandBar 对象，其中 *index* 是该命令栏的名称或索引号。下面的示例将名为“Custom1”的工具栏停靠在应用程序窗口的底部。

``` JavaScript
Application.ActiveDocument.CommandBars.Item("Custom1").Position = msoBarBottom
```

| 注释                                                                                                                                                                                                                                                                                                                                                                               |
|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 可以使用名称或索引号指定容器应用程序的可用菜单栏和工具栏列表中的菜单栏或工具栏，但必须使用名称指定菜单、快捷菜单或子菜单（所有这些均由 CommandBar 对象代表）。如果两个或两个以上的自定义菜单或子菜单同名，则 enumName 将返回第一个具有该名称的菜单。为确保返回正确的菜单或子菜单，可以找到显示该菜单的弹出式控件，然后对该弹出式控件应用 CommandBar 属性以返回代表该菜单的命令栏。 |

## 方法

| 序号 | 名称                                                                  | 说明                                                                                |
|:----:|:----------------------------------------------------------------------|:------------------------------------------------------------------------------------|
|  1   | [Add](#CommandBars.Add)                                               | 创建一个新的命令栏并将其添加到命令栏集合中。                                        |
|  2   | [CommitRenderingTransaction](#CommandBars.CommitRenderingTransaction) | 提交呈现事务                                                                        |
|  3   | [ExecuteMso](#CommandBars.ExecuteMso)                                 | 执行由 idMso 参数标识的控件。                                                       |
|  4   | [FindControl](#CommandBars.FindControl)                               | 获取一个符合指定条件的 CommandBarControl 对象。                                     |
|  5   | [FindControls](#CommandBars.FindControls)                             | 获取符合指定条件的 CommandBarControls 集合。                                        |
|  6   | [GetEnabledMso](#CommandBars.GetEnabledMso)                           | 如果启用了 idMso 参数标识的控件，则返回 True。                                      |
|  7   | [GetImageMso](#CommandBars.GetImageMso)                               | 返回由 idMso 参数标识的控件图像（缩放到指定的宽度和高度尺寸）的 IPictureDisp 对象。 |
|  8   | [GetLabelMso](#CommandBars.GetLabelMso)                               | 将 idMso 参数标识的控件的标签作为 String 类型的值返回。                             |
|  9   | [GetPressedMso](#CommandBars.GetPressedMso)                           | 返回一个值，该值指示是否已按下 idMso 参数标识的 toggleButton 控件。                 |
|  10  | [GetScreentipMso](#CommandBars.GetScreentipMso)                       | 将 idMso 参数标识的控件的屏幕提示作为 String 类型的值返回。                         |
|  11  | [GetSupertipMso](#CommandBars.GetSupertipMso)                         | 将 idMso 参数标识的控件的超级提示作为 String 类型的值返回。                         |
|  12  | [GetVisibleMso](#CommandBars.GetVisibleMso%20)                        | 如果 idMso 参数标识的控件可见，则返回 true。                                        |
|  13  | [ReleaseFocus](#CommandBars.ReleaseFocus)                             | 释放所有命令栏的用户界面焦点。                                                      |
|  14  | [SetVisibleMso](#CommandBars.SetVisibleMso)                           | 通过控件对应的idMso控制该控件的显示状态(显示/隐藏）。                               |

## 属性

| 序号 | 名称                                                                    | 说明                                                                                                                               |
|:----:|:------------------------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ActionControl](#CommandBars.ActionControl)                             | 获取一个 CommandBarControl 对象，该对象的 OnAction 属性设置为当前正在运行的过程。只读。                                            |
|  2   | [ActiveMenuBar](#CommandBars.ActiveMenuBar)                             | 获取一个 CommandBar 对象，该对象代表容器应用程序中的活动菜单栏。只读。                                                             |
|  3   | [AdaptiveMenus](#CommandBars.AdaptiveMenus)                             | 此属性可选中或取消选中指定 WPS Office中菜单是完全显示还是按个性化方式显示的选项的复选框控件。可读写。                              |
|  4   | [Application](#CommandBars.Application)                                 | 获取一个 Application 对象，代表 CommandBars 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  5   | [Count](#CommandBars.Count)                                             | 获取宿主应用程序中的命令栏数量。只读。                                                                                             |
|  6   | [Creator](#CommandBars.Creator)                                         | 获取一个 32 位整数，指示创建 CommandBars 对象时所使用的应用程序。只读。                                                            |
|  7   | [DisableAskAQuestionDropdown](#CommandBars.DisableAskAQuestionDropdown) | 如果启用了 “应答向导” 下拉菜单，则为 True 。可读写。                                                                               |
|  8   | [DisableCustomize](#CommandBars.DisableCustomize)                       | 如果禁用了工具栏的自定义功能，则为 True 。可读写。                                                                                 |
|  9   | [DisplayFonts](#CommandBars.DisplayFonts)                               | 如果在 “字体” 框中以实际字体显示字体名称，则为 True 。可读写。                                                                     |
|  10  | [DisplayKeysInTooltips](#CommandBars.DisplayKeysInTooltips)             | 如果每个命令栏控件的快捷键都显示在 “工具提示” 中，则为 True 。可读写。                                                             |
|  11  | [DisplayTooltips](#CommandBars.DisplayTooltips)                         | 如果只要用户将指针放在命令栏控件上方就显示“屏幕提示”，则为 True 。可读写。                                                         |
|  12  | [Item](#CommandBars.Item)                                               | 获取 CommandBars 集合中的 CommandBar 对象。只读。                                                                                  |
|  13  | [LargeButtons](#CommandBars.LargeButtons)                               | 如果显示的工具栏按钮比常规尺寸要大，则为 True 。可读写。                                                                           |
|  14  | [MenuAnimationStyle](#CommandBars.MenuAnimationStyle)                   | 获取或设置一个代表命令栏的动画方式的 MsoMenuAnimation 。可读写。                                                                   |
|  15  | [Parent](#CommandBars.Parent)                                           | 获取 CommandBars 对象的 Parent 对象。只读。                                                                                        |

## 成员方法

### CommandBars.Add

创建一个新的命令栏并将其添加到命令栏集合中。

**语法**

**express.Add ( *Name* , *Position* , *MenuBar* , *Temporary* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                              |
|-------------|-----------|----------|---------------------------------------------------------------------------------------------------|
| *Name*      | 可选      | Variant  | 新命令栏的名称。如果省略此参数，则为命令栏指定默认名称（如 Custom 1）。                           |
| *Position*  | 可选      | Variant  | 新命令栏的位置或类型。可以是 MsoBarPosition 常量之一。                                            |
| *MenuBar*   | 可选      | Variant  | 设置为 true 将以新命令栏替换活动菜单栏。默认值为 false。                                          |
| *Temporary* | 可选      | Variant  | 如果为 true，则将使新命令栏变成临时命令栏。临时命令栏将在容器应用程序关闭时删除。默认值为 false。 |

**示例**

``` JavaScript
function test()
{
    let cbar1 = Application.ActiveDocument.CommandBars.Add("Custom1", msoBarFloating)
    cbar1.Visible = true
}
```

### CommandBars.CommitRenderingTransaction

提交呈现事务

**语法**

**express.CommitRenderingTransaction ( *hwnd* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                               |
|--------|-----------|----------|------------------------------------|
| *hwnd* | 必选      | Long     | 要在其中提交呈现事务的窗口的句柄。 |

### CommandBars.ExecuteMso

执行由 **idMso** 参数标识的控件。

**语法**

**express.ExecuteMso ( *idMso* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *idMso* | 必选      | String   | 控件的标识符。 |

**说明**

| 注释                                                   |
|--------------------------------------------------------|
| WPS Office Assistant 在 WPS Office 2015 版中已被淘汰。 |

在某个命令没有对象模型的情况下，可以使用此方法。此方法可用于内置按钮、toggleButton 和 splitButton 类型的控件。此方法执行失败时将返回 E_InvalidArg（如果 **IdMso** 参数无效）或 E_Fail（若未启用控件或控件不可见）。

**示例**

``` JavaScript
//下面的示例执行“Copy”按钮。
Application.ActiveDocument.CommandBars.ExecuteMso("Copy")
```

### CommandBars.FindControl

获取一个符合指定条件的 **CommandBarControl** 对象。

**语法**

**express.FindControl ( *Type* , *Id* , *Tag* , *Visible* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                                                       |
|-----------|-----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------|
| *Type*    | 可选      | Variant  | 控件的类型。                                                                                                                               |
| *Id*      | 可选      | Variant  | 控件的标识符。                                                                                                                             |
| *Tag*     | 可选      | Variant  | 控件的标记值。                                                                                                                             |
| *Visible* | 可选      | Variant  | 如果为 true，则只能在搜索中包含可见的命令栏控件。默认值为 false。可见的命令栏包括所有可见的工具栏和执行 FindControl 方法时打开的所有菜单。 |

**说明**

如果 **CommandBars** 集合中有两个或者更多的控件符合搜索条件，那么 FindControl 返回找到的第一个控件。如果没有控件符合搜索条件，那么 **FindControl** 返回 Nothing。

### CommandBars.FindControls

获取符合指定条件的 **CommandBarControls** 集合。

**语法**

**express.FindControls ( *Type* , *Id* , *Tag* , *Visible* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                              |
|-----------|-----------|----------|-------------------------------------------------------------------|
| *Type*    | 可选      | Variant  | 为指定控件类型的 MsoControlType 常量之一。                        |
| *Id*      | 可选      | Variant  | 控件的标识符。                                                    |
| *Tag*     | 可选      | Variant  | 控件的标记值。                                                    |
| *Visible* | 可选      | Variant  | 如果为 true，则只能在搜索中包含可见的命令栏控件。默认值为 false。 |

**说明**

|                                                                                                                  |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

如果未找到符合条件的控件， **FindControls** 方法将返回 **Nothing**

**示例**

``` JavaScript
//本示例用 FindControls 方法返回 CommandBars 集合中所有 ID 为 18 的成员，并在消息框中显示符合指定条件的控件的个数。
function test()
{
    let  myControls = CommandBars.FindControls(msoControlButton,18)
    alert("There are " + myControls.Count + " controls that meet the search criteria.")
}
```

### CommandBars.GetEnabledMso

如果启用了 **idMso** 参数标识的控件，则返回 True。

**语法**

**express.GetEnabledMso ( *idMso* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *idMso* | 必选      | String   | 控件的标识符。 |

**示例**

``` JavaScript
//下面的示例在“加粗”按钮已启用的情况下返回 True。
Application.CommandBars.GetEnabledMso("Bold")
```

### CommandBars.GetImageMso

返回由 **idMso** 参数标识的控件图像（缩放到指定的宽度和高度尺寸）的 **IPictureDisp** 对象。

**语法**

**express.GetImageMso ( *idMso* , *Width* , *Height* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明           |
|----------|-----------|----------|----------------|
| *idMso*  | 必选      | String   | 控件的标识符。 |
| *Width*  | 必选      | Number   | 图像的宽度     |
| *Height* | 必选      | Number   | 图像的高度。   |

**说明**

**Width** 和 **Height** 参数的值必须介于 16 和 128 之间。

**示例**

``` JavaScript
//下面的示例将 32x32 尺寸版本的“Paste”图标作为 IPictureDisp 对象返回。
Application.CommandBars.GetImageMso("Paste", 32, 32)
```

### CommandBars.GetLabelMso

将 **idMso** 参数标识的控件的标签作为 String 类型的值返回。

**语法**

**express.GetLabelMso ( *idMso* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *idMso* | 必选      | String   | 控件的标识符。 |

**示例**

``` JavaScript
//下面的示例返回 String 类型的值“Paste”。

Application.CommandBars.GetLabelMso("Paste")
```

### CommandBars.GetPressedMso

返回一个值，该值指示是否已按下 **idMso** 参数标识的 toggleButton 控件。

**语法**

**express.GetPressedMso ( *idMso* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *idMso* | 必选      | String   | 控件的标识符。 |

**返回值**

Boolean

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//下面的示例在“加粗”按钮已按下的情况下返回 True。
Application.CommandBars.GetPressedMso("Bold") 
```

### CommandBars.GetScreentipMso

将 **idMso** 参数标识的控件的屏幕提示作为 String 类型的值返回。

**语法**

**express.GetScreentipMso ( *idMso* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *idMso* | 必选      | String   | 控件的标识符。 |

**返回值**

String

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//下面的示例返回 String 类型的值“Paste”。
Application.CommandBars.GetScreentipMso("Paste")
```

### CommandBars.GetSupertipMso

将 **idMso** 参数标识的控件的超级提示作为 String 类型的值返回。

**语法**

**express.GetSupertipMso ( *idMso* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *idMso* | 必选      | String   | 控件的标识符。 |

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//下面的示例返回 String 类型的值“从文档中剪切所选内容，并将其放入剪贴板”。
Application.CommandBars.GetSupertipMso("Cut")
```

### CommandBars.GetVisibleMso

如果 **idMso** 参数标识的控件可见，则返回 true。

**语法**

**express.GetVisibleMso ( *idMso* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *idMso* | 必选      | String   | 控件的标识符。 |

**示例**

``` JavaScript
//下面的示例在“加粗”按钮可见的情况下返回 True。
Application.CommandBars.GetVisibleMso("Bold")
```

### CommandBars.ReleaseFocus

释放所有命令栏的用户界面焦点。

**语法**

**express.ReleaseFocus ()**

\*express是一个代表 CommandBars 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例向命令栏“Custom”中添加三个空白按钮，并将焦点设置在中间的按钮上。等待五秒钟后，释放所有命令栏的用户界面焦点。
function test()
{
    let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop,undefined,true)
    myBar.Controls.Add(Application.Enum.msoControlButton)
    myBar.Controls.Add(Application.Enum.msoControlButton)
    myBar.Controls.Add(Application.Enum.msoControlButton)
    myBar.Visible = true 
    let myControl = Application.CommandBars.Item("Custom").Controls.Item(2)
    myControl.SetFocus()
    let PauseTime = 5   // Set duration.
    let Start = Timer   // Set start time.
    while(Timer(Start + PauseTime))
    {
        DoEvents()          // Yield to other processes.
        let Finish = Timer
        Application.CommandBars.ReleaseFocus()
    }
}
```

### CommandBars.SetVisibleMso

通过控件对应的idMso控制该控件的显示状态(显示/隐藏）。

**语法**

**express.SetVisibleMso ( *idMso* , *Enabled* )**

\*express是一个代表 CommandBars 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明           |
|-----------|-----------|----------|----------------|
| *idMso*   | 必选      | String   | 控件的标识符。 |
| *Enabled* | 必选      | Boolean  | 是否隐藏。     |

## 成员属性

### CommandBars.ActionControl

获取一个 **CommandBarControl** 对象，该对象的 **OnAction** 属性设置为当前正在运行的过程。只读。

**语法**

**express.ActionControl**

\*express是一个代表 CommandBars 对象的变量。

**示例**

``` JavaScript
//本示例创建命令栏“Custom”，向命令栏中添加三个按钮，然后用 ActionControl 属性和 Tag 属性确定最后一次单击的是哪一个命令栏按钮。
function test()
{
    let myBar = Application.ActiveDocument.CommandBars.Add("Custom", Application.Enum.msoBarTop,undefined,true)
    let buttonOne = myBar.Controls.Add(Application.Enum.msoControlButton)
    buttonOne.FaceId = 133
    buttonOne.Tag = "RightArrow"
    buttonOne.OnAction = "whichButton"
    let buttonTwo = myBar.Controls.Add(Application.Enum.msoControlButton)
    buttonTwo.FaceId = 134
    buttonTwo.Tag = "UpArrow"
    buttonTwo.OnAction = "whichButton"
    let buttonThree = myBar.Controls.Add(Application.Enum.msoControlButton)
    buttonThree.FaceId = 135
    buttonThree.Tag = "DownArrow"
    buttonThree.OnAction = "whichButton"
    myBar.Visible = true
}
//下面的子例程响应 OnAction 方法并确定最后单击的是哪一个命令栏按钮。
 function whichButton()
 {
    switch(Application.ActiveDocument.CommandBars.ActionControl.Tag) {
        case "RightArrow":
            alert("Right Arrow button clicked.")
            break
        case "UpArrow":
            alert ("Up Arrow button clicked.")
            break
        case "DownArrow":
            alert ("Down Arrow button clicked.")
            break
    }

}
```

### CommandBars.ActiveMenuBar

获取一个 **CommandBar** 对象，该对象代表容器应用程序中的活动菜单栏。只读。

**语法**

**express.ActiveMenuBar**

\*express是一个代表 CommandBars 对象的变量。

**示例**

``` JavaScript
/*
本示例在活动菜单栏的末尾暂时添加一个弹出式控件“Custom”，然后向该弹出式控件中添加一个名为“Import”的控件。
*/
function test()
{
    let myMenuBar = Application.ActiveDocument.CommandBars.ActiveMenuBar
    let newMenu = myMenuBar.Controls.Add(Application.Enum.msoControlPopup,undefined,undefined,undefined,true)
    newMenu.Caption = "Custom"
    let ctrl1 = newMenu.CommandBar.Controls.Add(Application.Enum.msoControlButton,1)
    ctrl1.Caption = "Import"
    ctrl1.TooltipText = "Import"
    ctrl1.Style = Application.Enum.msoButtonCaption
}
```

### CommandBars.AdaptiveMenus

此属性可选中或取消选中指定 WPS Office中菜单是完全显示还是按个性化方式显示的选项的复选框控件。可读写。

**语法**

**express.AdaptiveMenus**

\*express是一个代表 CommandBars 对象的变量。

**说明**

如果对所有 WPS Office应用程序都启用了自适应菜单，则为 **True** 。可读写 **Boolean** 类型。

在任意应用程序中执行下列操作均可设置此控件：

1.  在 **“工具”** 菜单上，选择 **“自定义”** 。
2.  选择 **“选项”** 选项卡。
3.  **“始终显示整个菜单”** 选项位于 **“个性化菜单和工具栏”** 部分中。

**示例**

``` JavaScript
//本示例为 WPS Office中的所有命令栏（其中包括自定义命令栏及其上的自定义命令控件）设置了三个选项。
Application.ActiveDocument.CommandBars.LargeButtons = true 
Application.ActiveDocument.CommandBars.DisplayFonts = true
Application.ActiveDocument.CommandBars.AdaptiveMenus = true
```

### CommandBars.Application

获取一个 **Application** 对象，代表 **CommandBars** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 CommandBars 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBars.Count

获取宿主应用程序中的命令栏数量。只读。

**语法**

**express.Count**

\*express是一个代表 CommandBars 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBars.Creator

获取一个 32 位整数，指示创建 **CommandBars** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 CommandBars 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

### CommandBars.DisableAskAQuestionDropdown

如果启用了 **“应答向导”** 下拉菜单，则为 **True** 。可读写。

**语法**

**express.DisableAskAQuestionDropdown**

\*express是一个代表 CommandBars 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//下面的示例打开或关闭 DisableAskAQuestionDropdown 属性。
function test()
{
    let comBars = Application.CommandBars
    if(comBars.DisableAskAQuestionDropdown == true) {
        comBars.DisableAskAQuestionDropdown = false 
    }
    else {
        comBars.DisableAskAQuestionDropdown = true 
    }

}
```

### CommandBars.DisableCustomize

如果禁用了工具栏的自定义功能，则为 **True** 。可读写。

**语法**

**express.DisableCustomize**

\*express是一个代表 CommandBars 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//下面的示例打开或关闭 DisableCustomize 属性。
function test()
{
    let comBars = Application.CommandBars
    if(comBars.DisableCustomize == true)
    {
        comBars.DisableCustomize = false
    }
    else{
        comBars.DisableCustomize = true
    }

}
```

### CommandBars.DisplayFonts

如果在 **“字体”** 框中以实际字体显示字体名称，则为 **True** 。可读写。

**语法**

**express.DisplayFonts**

\*express是一个代表 CommandBars 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例为 WPS Office中的所有命令栏（其中包括自定义命令栏及其上的自定义命令控件）设置了三个选项。 
function test()
{
    Application.CommandBars.LargeButtons = true 
    Application.CommandBars.DisplayFonts = true
    Application.CommandBars.AdaptiveMenus = true

}
```

### CommandBars.DisplayKeysInTooltips

如果每个命令栏控件的快捷键都显示在 **“工具提示”** 中，则为 **True** 。可读写。

**语法**

**express.DisplayKeysInTooltips**

\*express是一个代表 CommandBars 对象的变量。

**说明**

要在 **“工具提示”** 中显示快捷键，还必须将 **DisplayTooltips** 属性设置为 **True** 。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例为 WPS Office中所有命令栏设置选项。 
function test()
{
    Application.CommandBars.LargeButtons = true
    Application.CommandBars.DisplayTooltips = true 
    Application.CommandBars.DisplayKeysInTooltips = true
    Application.CommandBars.MenuAnimationStyle = Application.Enum.msoMenuAnimationUnfold

}
```

### CommandBars.DisplayTooltips

如果只要用户将指针放在命令栏控件上方就显示“屏幕提示”，则为 **True** 。可读写。

**语法**

**express.DisplayTooltips**

\*express是一个代表 CommandBars 对象的变量。

**说明**

在容器应用程序中设置 **DisplayTooltips** 属性会立刻影响当时正在运行的每个 WPS Office应用程序中的所有命令栏和设置该属性后打开的所有 Office 应用程序中的所有命令栏。

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在所有命令栏上显示大控件和工具提示。  
function test()
{
    Application.CommandBars.LargeButtons = true 
    Application.CommandBars.DisplayTooltips = true
}
```

### CommandBars.Item

获取 **CommandBars** 集合中的 **CommandBar** 对象。只读。

**语法**

**express.Item**

\*express是一个代表 CommandBars 对象的变量。

**示例**

``` JavaScript
//Item 为对象或集合的默认成员。下面语句将 CommandBar 对象指定给 cmdBar。
let cmdBar = Application.ActiveDocument.CommandBars.Item("Standard")
```

### CommandBars.LargeButtons

如果显示的工具栏按钮比常规尺寸要大，则为 **True** 。可读写。

**语法**

**express.LargeButtons**

\*express是一个代表 CommandBars 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例在所有命令栏上的工具栏按钮的显示大小间切换。  
function test()
{
    if(Application.CommandBars.LargeButtons)
    {
        Application.CommandBars.LargeButtons = false
    }
    else
    {
        Application.CommandBars.LargeButtons = true
    }
}
```

### CommandBars.MenuAnimationStyle

获取或设置一个代表命令栏的动画方式的 **MsoMenuAnimation** 。可读写。

**语法**

**express.MenuAnimationStyle**

\*express是一个代表 CommandBars 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

``` JavaScript
//本示例为 WPS Office中所有命令栏设置选项。  
function test()
{
    Application.CommandBars.LargeButtons = true
    Application.CommandBars.DisplayTooltips = true 
    Application.CommandBars.DisplayKeysInTooltips = true 
    Application.CommandBars.MenuAnimationStyle = Application.Enum.msoMenuAnimationUnfold
}
```

### CommandBars.Parent

获取 **CommandBars** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CommandBars 对象的变量。

**说明**

| 注释                                                                                                             |
|------------------------------------------------------------------------------------------------------------------|
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CommandBars](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ODSOColumns 对象

## 简介

代表邮件合并数据源中的数据字段的 ODSOColumn 对象的集合。

## 说明

## 方法

| 序号 | 名称                      | 说明                                        |
|:----:|:--------------------------|:--------------------------------------------|
|  1   | [Item](#ODSOColumns.Item) | 指定 ODSOColumns 集合中的 ODSOColumn 对象。 |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                               |
|:----:|:----------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ODSOColumns.Application) | 获取一个 Application 对象，代表 ODSOColumns 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Count](#ODSOColumns.Count)             | 获取一个 Long 类型的值，指示 ODSOColumns 集合中的项数。只读。                                                                      |
|  3   | [Creator](#ODSOColumns.Creator)         | 获取一个 32 位整数，指示创建 ODSOColumns 对象时所使用的应用程序。只读。                                                            |
|  4   | [Parent](#ODSOColumns.Parent)           | 获取 ODSOColumns 对象的 Parent 对象。只读。                                                                                        |

## 成员方法

### ODSOColumns.Item

指定 **ODSOColumns** 集合中的 **ODSOColumn** 对象。

**语法**

**express.Item ( *varIndex* )**

\*express是一个代表 ODSOColumns 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明           |
|------------|-----------|----------|----------------|
| *varIndex* | 必选      | Variant  | 项目的索引号。 |

**示例**

## 成员属性

### ODSOColumns.Application

获取一个 **Application** 对象，代表 **ODSOColumns** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 ODSOColumns 对象的变量。

### ODSOColumns.Count

获取一个 **Long** 类型的值，指示 **ODSOColumns** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 ODSOColumns 对象的变量。

### ODSOColumns.Creator

获取一个 32 位整数，指示创建 **ODSOColumns** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 ODSOColumns 对象的变量。

### ODSOColumns.Parent

获取 **ODSOColumns** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ODSOColumns 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/ODSOColumns](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SmartArtNode 对象

## 简介

Smart Art 图形的数据模型内部的单个语义节点。

## 说明

``` JavaScript
/*下面的代码返回 Smart Art 图表中的节点数目。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.AllNodes.Count
```

## 方法

| 序号 | 名称                                     | 说明                                                                                                        |
|:----:|:-----------------------------------------|:------------------------------------------------------------------------------------------------------------|
|  1   | [AddNode](#SmartArtNode.AddNode)         | 通过SmartArtNodeType类型的SmartArtNodePosition值指定的方式，将新的SmartArtNode添加到数据模型。              |
|  2   | [Delete](#SmartArtNode.Delete)           | 删除当前的 SmartArt 节点。                                                                                  |
|  3   | [Demote](#SmartArtNode.Demote)           | 将当前节点在数据模型中降低一级。                                                                            |
|  4   | [Larger](#SmartArtNode.Larger)           | 增加 SmartArt 节点的大小。模拟 SmartArt 的 WPS Office Fluent 功能区“格式”选项卡上的“增大”按钮的行为。       |
|  5   | [Promote](#SmartArtNode.Promote)         | 将当前节点（以及它的所有子级）在数据模型中提升一级。                                                        |
|  6   | [ReorderDown](#SmartArtNode.ReorderDown) | 与项目符号列表中的下一个节点交换节点。此方法会对整个节点系列进行重新排序。                                  |
|  7   | [ReorderUp](#SmartArtNode.ReorderUp)     | 与项目符号列表中的上一个节点交换节点。此方法会对整个节点系列进行重新排序。                                  |
|  8   | [Smaller](#SmartArtNode.Smaller)         | 减小 SmartArt 的大小。模拟 SmartArt 的 WPS Office Fluent 功能区用户界面的“格式”选项卡上的“减小”按钮的行为。 |

## 属性

| 序号 | 名称                                           | 说明                                                                           |
|:----:|:-----------------------------------------------|:-------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtNode.Application)       | 获取一个 Application 对象，该对象代表 SmartArtNode 对象的容器应用程序。只读。  |
|  2   | [Creator](#SmartArtNode.Creator)               | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtNode 对象的应用程序。只读。 |
|  3   | [Hidden](#SmartArtNode.Hidden)                 | 如果此节点在数据模型中是一个隐藏节点，则返回 true 。只读。                     |
|  4   | [Level](#SmartArtNode.Level)                   | 检索节点在层次结构中的级别。只读。                                             |
|  5   | [Nodes](#SmartArtNode.Nodes)                   | 检索与此 Smart Art 节点关联的子节点。只读。                                    |
|  6   | [OrgChartLayout](#SmartArtNode.OrgChartLayout) | 如果有一个节点，则检索或设置与此节点关联的 MsoOrgChartLayoutType 。可读/写。   |
|  7   | [Parent](#SmartArtNode.Parent)                 | 返回调用对象。只读。                                                           |
|  8   | [ParentNode](#SmartArtNode.ParentNode)         | 检索此 SmartArtNode 的父 SmartArtNode。只读。                                  |
|  9   | [Shapes](#SmartArtNode.Shapes)                 | 返回与此 SmartArtNode 对象关联的形状范围。只读。                               |
|  10  | [TextFrame2](#SmartArtNode.TextFrame2)         | 返回与该 SmartArtNode 对象关联的文本。只读。                                   |
|  11  | [Type](#SmartArtNode.Type)                     | 检索 SmartArt 节点的类型。只读。                                               |

## 成员方法

### SmartArtNode.AddNode

通过SmartArtNodeType类型的SmartArtNodePosition值指定的方式，将新的SmartArtNode添加到数据模型。

**语法**

**express.AddNode ( *Position* , *Type* )**

\*express是一个代表 SmartArtNode 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型                | 说明                                                                                               |
|------------|-----------|-------------------------|----------------------------------------------------------------------------------------------------|
| *Position* | 可选      | MsoSmartArtNodePosition | 指定 SmartArtNode 在数据模型中的位置。例如，msoSmartArtNodeAbove 或 msoSmartArtNodeAfter。         |
| *Type*     | 可选      | MsoSmartArtNodeType     | 指定添加的 SmartArtNode 的类型。例如，msoSmartArtNodeTypeAssistant 或 msoSmartArtNodeTypeDefault。 |

**返回值**

SmartArtNode

**示例**

``` JavaScript
/*下面的代码将默认的 SmartArtNode 添加到当前节点下。*/
function test(){
    var saNode = Application.ActiveWorkbook.ActiveChart.Shapes.Item(1).SmartArt.AllNodes.Item(1);
    saNode = saNode.AddNode(Application.Enum.msoSmartArtNodeBelow, Application.Enum.msoSmartArtNodeTypeDefault);
}
```

### SmartArtNode.Delete

删除当前的 SmartArt 节点。

**语法**

**express.Delete ()**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

删除节点时，第一个子级得到提升。在下面的数据模型中：

- A

- - B

  - - C

- D

如果删除了 B，则该数据模型的外观如下所示：

- A

- - C

- D

### SmartArtNode.Demote

将当前节点在数据模型中降低一级。

**语法**

**express.Demote ()**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

在内容窗格中工作时，此功能模拟 WPS Office Fluent 功能区用户界面中的“降级”按钮。例如，以下面的数据模型为例：

- A

- B

- - C

- D

如果使 B 降级，产生的数据模型的外观如下所示：

- A

- - B
  - C

- D

### SmartArtNode.Larger

增加 SmartArt 节点的大小。模拟 SmartArt 的 WPS Office Fluent 功能区“格式”选项卡上的“增大”按钮的行为。

**语法**

**express.Larger ()**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

实际的增加量取决于节点的布局。

### SmartArtNode.Promote

将当前节点（以及它的所有子级）在数据模型中提升一级。

**语法**

**express.Promote ()**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

在内容窗格中工作时，此功能模拟 WPS Office Fluent 功能区用户界面中的“提升”按钮。例如，以下面的数据模型为例：

- A

- - B

  - - C

- D

如果使 B 提升，产生的数据模型的外观如下所示：

- A

- B

- - C

- D

### SmartArtNode.ReorderDown

与项目符号列表中的下一个节点交换节点。此方法会对整个节点系列进行重新排序。

**语法**

**express.ReorderDown ()**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

此方法模拟单击 WPS Office Fluent 功能区用户界面上的“向下重新排序”按钮，该按钮位于“向下重新排序”的“设计”组的“SmartArt 工具”选项卡上。

**示例**

``` JavaScript
/*下面的代码将第一个节点与下一个节点交换，并对其所有后代进行重新排序。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Item(1).ReorderDown()
```

### SmartArtNode.ReorderUp

与项目符号列表中的上一个节点交换节点。此方法会对整个节点系列进行重新排序。

**语法**

**express.ReorderUp ()**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

此方法模拟单击 WPS Office Fluent 功能区用户界面上的“向上重新排序”按钮，该按钮位于“向上重新排序”的“设计”组的“SmartArt 工具”选项卡上。

**示例**

``` JavaScript
/*下面的代码将第一个节点与下一个节点交换，并对父节点进行重新排序。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Item(1).ReorderUp()
```

### SmartArtNode.Smaller

减小 SmartArt 的大小。模拟 SmartArt 的 WPS Office Fluent 功能区用户界面的“格式”选项卡上的“减小”按钮的行为。

**语法**

**express.Smaller ()**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

减小量取决于节点的布局大小。

## 成员属性

### SmartArtNode.Application

获取一个 **Application** 对象，该对象代表 **SmartArtNode** 对象的容器应用程序。只读。

**语法**

**express.Application**

\*express是一个代表 SmartArtNode 对象的变量。

### SmartArtNode.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtNode** 对象的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SmartArtNode 对象的变量。

### SmartArtNode.Hidden

如果此节点在数据模型中是一个隐藏节点，则返回 **true** 。只读。

**语法**

**express.Hidden**

\*express是一个代表 SmartArtNode 对象的变量。

### SmartArtNode.Level

检索节点在层次结构中的级别。只读。

**语法**

**express.Level**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

起始级别为 1 级并且向上递增。如果某个节点没有级别，则返回 0。例如，在下面的数据模型中，A 和 F 为 1 级，B 和 D 为 2 级，C 和 E 为 3 级。

- A

- - B

  - - C

  - D

- - E

- F

### SmartArtNode.Nodes

检索与此 Smart Art 节点关联的子节点。只读。

**语法**

**express.Nodes**

\*express是一个代表 SmartArtNode 对象的变量。

**示例**

``` JavaScript
/*下面的代码返回 Smart Art 图表中的节点数目。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Item(1).Nodes.Count
```

### SmartArtNode.OrgChartLayout

如果有一个节点，则检索或设置与此节点关联的 **MsoOrgChartLayoutType** 。可读/写。

**语法**

**express.OrgChartLayout**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

以下是可能的成员：

- msoOrgChartLayoutBothHanging
- msoOrgChartLayoutDefault
- msoOrgChartLayoutLeftHanging
- msoOrgChartLayoutMixed
- msoOrgChartLayoutRightHanging
- msoOrgChartLayoutStandard

**示例**

``` JavaScript
/*下面的代码将 OrgChartLayout 属性设置为默认布局。*/
function test(){
    var saNode =  Application.ActiveWorkbook.ActiveChart.Shapes.Item(1);
    saNode.OrgChartLayout = Application.Enum.msoOrgChartLayoutDefault;
}
```

### SmartArtNode.Parent

返回调用对象。只读。

**语法**

**express.Parent**

\*express是一个代表 SmartArtNode 对象的变量。

### SmartArtNode.ParentNode

检索此 SmartArtNode 的父 SmartArtNode。只读。

**语法**

**express.ParentNode**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

如果此节点是根节点，则返回一个错误。

### SmartArtNode.Shapes

返回与此 **SmartArtNode** 对象关联的形状范围。只读。

**语法**

**express.Shapes**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

若要查找该范围，请使用 Count 属性。

**示例**

``` JavaScript
/*下面的代码返回形状范围。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Item(1).Shapes.Count
```

### SmartArtNode.TextFrame2

返回与该 **SmartArtNode** 对象关联的文本。只读。

**语法**

**express.TextFrame2**

\*express是一个代表 SmartArtNode 对象的变量。

**示例**

``` JavaScript
/*下面的示例设置第一个节点内部的文本。*/
function test(){
    var smartart = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt
    smartart.AllNodes.Item(1).TextFrame2.TextRange.Text="Node 1"
}
```

### SmartArtNode.Type

检索 SmartArt 节点的类型。只读。

**语法**

**express.Type**

\*express是一个代表 SmartArtNode 对象的变量。

**说明**

可能值为 msoSmartArtNodeTypeAssistant 或 msoSmartArtNodeTypeDefault。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtNode](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SmartArtNodes 对象

## 简介

代表 Smart Art 图表内部的节点的集合。

## 说明

这些节点直接对应于包含在图形的数据模型内部的语义元素。

``` JavaScript
/*下面的代码返回 Smart Art 图表中的节点数目。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.AllNodes.Count
```

## 方法

| 序号 | 名称                        | 说明                                                         |
|:----:|:----------------------------|:-------------------------------------------------------------|
|  1   | [Add](#SmartArtNodes.Add)   | 将新的 SmartArtNode 对象添加到具有指定文本的图表中。         |
|  2   | [Item](#SmartArtNodes.Item) | 在指定的索引处或通过指定的唯一 ID 来检索 SmartArtNode 对象。 |

## 属性

| 序号 | 名称                                      | 说明                                                                          |
|:----:|:------------------------------------------|:------------------------------------------------------------------------------|
|  1   | [Application](#SmartArtNodes.Application) | 获取一个 Application 对象，该对象代表 SmartArtNodes 对象的容器应用程序。只读  |
|  2   | [Count](#SmartArtNodes.Count)             | 检索包含在 SmartArtNodes 集合中的 SmartArtNode 对象数。只读                   |
|  3   | [Creator](#SmartArtNodes.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 SmartArtNodes 对象的应用程序。只读 |
|  4   | [Parent](#SmartArtNodes.Parent)           | 返回调用对象。只读                                                            |

## 成员方法

### SmartArtNodes.Add

将新的 **SmartArtNode** 对象添加到具有指定文本的图表中。

**语法**

**express.Add ()**

\*express是一个代表 SmartArtNodes 对象的变量。

**返回值**

SmartArtNode

**说明**

这个新节点将添加到位于此节点集合最顶层的数据模型的底部。如果最高级别是 2，则这个新节点的级别也将是 2。

**示例**

``` JavaScript
/*下面的代码将一个 SmartArtNode 添加到 SmartArtNodes 集合中。*/
function test(){
    var saNodes = Application.ActiveWorkbook.ActiveChart.Shapes.Item(1).SmartArt.AllNodes.Item(1);
    saNodes.AddNode();
}
```

### SmartArtNodes.Item

在指定的索引处或通过指定的唯一 ID 来检索 **SmartArtNode** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SmartArtNodes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                             |
|---------|-----------|----------|------------------------------------------------------------------|
| *Index* | 必选      | Variant  | 指定一个代表索引的整数或一个代表 SmartArtNode 对象位置的字符串。 |

**返回值**

SmartArtNode

## 成员属性

### SmartArtNodes.Application

获取一个 **Application** 对象，该对象代表 **SmartArtNodes** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 SmartArtNodes 对象的变量。

### SmartArtNodes.Count

检索包含在 **SmartArtNodes** 集合中的 **SmartArtNode** 对象数。只读

**语法**

**express.Count**

\*express是一个代表 SmartArtNodes 对象的变量。

### SmartArtNodes.Creator

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtNodes** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 SmartArtNodes 对象的变量。

### SmartArtNodes.Parent

返回调用对象。只读

**语法**

**express.Parent**

\*express是一个代表 SmartArtNodes 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartArtNodes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DataSource 对象

## 简介

代表容器应用程序的一个数据源对象。

## 方法

| 序号 | 名称                                                   | 说明                                      |
|:----:|:-------------------------------------------------------|:------------------------------------------|
|  1   | [Activate](#DataSource.Activate)                       | 激活数据源，跨进程表现为激活ET文档。      |
|  2   | [ApplyDemoRange](#DataSource.ApplyDemoRange)           | 往活动单元格填充示例数据。返回DataRange。 |
|  3   | [CreateDataRange](#DataSource.CreateDataRange)         | 为选定的单元格创建DataRange对象。         |
|  4   | [GetAllRangeInfo](#DataSource.GetAllRangeInfo)         | 获取所有已注册的DataRange。               |
|  5   | [GetDataRange](#DataSource.GetDataRange)               | 返回已注册的对应DataRange。               |
|  6   | [InitDemoRange](#DataSource.InitDemoRange)             | 初始化示例数据的行列数                    |
|  7   | [RegisterDataRange](#DataSource.RegisterDataRange)     | 向当前DataSource对象注册DataRange对象。   |
|  8   | [SetDemoRangeItem](#DataSource.SetDemoRangeItem)       | 填充示例数据。                            |
|  9   | [UnRegisterDataRange](#DataSource.UnRegisterDataRange) | 反注册DataRange。                         |

## 属性

| 序号 | 名称                     | 说明                         |
|:----:|:-------------------------|:-----------------------------|
|  1   | [Type](#DataSource.Type) | 返回DataSource的类型。只读。 |

## 成员方法

### DataSource.Activate

激活数据源，跨进程表现为激活ET文档。

**语法**

**express.Activate ()**

\*express是一个代表 DataSource 对象的变量。

**返回值**

Boolean

**示例**

``` JavaScript
//以下示例为WPP激活数据源对象，调起ET文档
Application.ActivePresentation.Slides(1).Shapes.Item(3).WebShape.DataSource.Activate()
```

### DataSource.ApplyDemoRange

往活动单元格填充示例数据。返回DataRange。

**语法**

**express.ApplyDemoRange ()**

\*express是一个代表 DataSource 对象的变量。

**返回值**

Object

### DataSource.CreateDataRange

为选定的单元格创建DataRange对象。

**语法**

**express.CreateDataRange ( *Range* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                              |
|---------|-----------|----------|-----------------------------------|
| *Range* | 可选      | Variant  | 代表单元格的字符串或者Range对象。 |

**返回值**

Object

**说明**

参数Range可以为代表单元格的字符串或者Range对象，当为空或字符串时，会弹出对话框输入单元格地址。形如“sheet1!\$a\$1”的字符串表示sheet1中A1单元格。返回值为DataRange。

**示例**

``` JavaScript
//下列示例为Sheet1中的A1单元格创建DataRange对象
Application.ActiveWorkbook.Sheets.Item(1).Shapes.Item(1).WebShape.DataSource.CreateDataRange("sheet1!$a$1")
```

### DataSource.GetAllRangeInfo

获取所有已注册的DataRange。

**语法**

**express.GetAllRangeInfo ()**

\*express是一个代表 DataSource 对象的变量。

**返回值**

String

**说明**

返回的字符串以JSON模式记录所有DataRange的Key。

### DataSource.GetDataRange

返回已注册的对应DataRange。

**语法**

**express.GetDataRange ( *Index* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明       |
|---------|-----------|----------|------------|
| *Index* | 必选      | Variant  | 对象的索引 |

**返回值**

Object

**说明**

目前Index仅支持字符串，代表DataRange对象的键。当输入的Index为数字时，不会有类型检查报错，但是也不会有对象返回。

### DataSource.InitDemoRange

初始化示例数据的行列数

**语法**

**express.InitDemoRange ( *rows* , *cols* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明 |
|--------|-----------|----------|------|
| *rows* | 必选      | Number   | 行数 |
| *cols* | 必选      | Number   | 列数 |

### DataSource.RegisterDataRange

向当前DataSource对象注册DataRange对象。

**语法**

**express.RegisterDataRange ( *Key* , *Value* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明              |
|---------|-----------|----------|-------------------|
| *Key*   | 必选      | String   | DataRange对象的键 |
| *Value* | 必选      | Object   | DataRange对象     |

**返回值**

Boolean

### DataSource.SetDemoRangeItem

填充示例数据。

**语法**

**express.SetDemoRangeItem ( *rowsIdx* , *colsIdx* , *value* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明       |
|-----------|-----------|----------|------------|
| *rowsIdx* | 必选      | Number   | 行数索引   |
| *colsIdx* | 必选      | Number   | 列数索引   |
| *value*   | 必选      | Variant  | 要填充的值 |

### DataSource.UnRegisterDataRange

反注册DataRange。

**语法**

**express.UnRegisterDataRange ( *Key* )**

\*express是一个代表 DataSource 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明           |
|-------|-----------|----------|----------------|
| *Key* | 必选      | String   | DataRange的Key |

**返回值**

Boolean

## 成员属性

### DataSource.Type

返回DataSource的类型。只读。

**语法**

**express.Type**

\*express是一个代表 DataSource 对象的变量。

**类型**

Number

> WPS 开发文档： [WPS 基础接口/WPS 扩展 API/公共/DataSource](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
