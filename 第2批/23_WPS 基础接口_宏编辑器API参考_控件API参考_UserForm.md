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
