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
