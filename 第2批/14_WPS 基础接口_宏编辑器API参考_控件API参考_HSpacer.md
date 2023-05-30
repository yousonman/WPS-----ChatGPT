# HSpacer 对象

## 简介

填充水平方向上无用的空隙。

## 说明

填充水平方向上无用的空隙。

## 属性

| 序号 | 名称                                  | 说明                                                 |
|:----:|:--------------------------------------|:-----------------------------------------------------|
|  1   | [Height](#HSpacer.Height)             | 获取或设置指定对象的高度                             |
|  2   | [HeightPolicy](#HSpacer.HeightPolicy) | 获取或设置指定对像的高度策略                         |
|  3   | [Left](#HSpacer.Left)                 | 您可以通过该属性指定控件在对象或报表中的位置         |
|  4   | [Name](#HSpacer.Name)                 | 您可以通过该属性指定或确定表示对象名称的字符串，只读 |
|  5   | [Top](#HSpacer.Top)                   | 您可以通过该属性指定控件在对象或报表中的位置         |
|  6   | [Width](#HSpacer.Width)               | 获取或设置指定对象的宽度                             |
|  7   | [WidthPolicy](#HSpacer.WidthPolicy)   | 获取或设置指定对象的宽度策略                         |

## 成员属性

### HSpacer.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 HSpacer 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将HSpacer1控件的高度设置为100*/
function func1()
{
    UserForm1.HSpacer1.Height = 100
}
```

### HSpacer.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 HSpacer 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将HSpacer1的高度设置为固定*/
function func1()
{
    UserForm1.HSpacer1.HeightPolicy = 2
}
```

### HSpacer.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 HSpacer 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将HSpacer1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.HSpacer1.Left = 100
}
```

### HSpacer.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 HSpacer 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### HSpacer.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 HSpacer 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将HSpacer1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.HSpacer1.Top = 100
}
```

### HSpacer.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 HSpacer 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将HSpacer1的宽度设置为100*/
function func1()
{
    UserForm1.HSpacer1.Width = 100
}
```

### HSpacer.WidthPolicy

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolicy**

\*express是一个代表 HSpacer 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将HSpacer1的宽度设置为固定*/
function func1()
{
    UserForm1.HSpacer1.WidthPolicy = 2
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/HSpacer](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
