# HeaderFooter 对象

## 简介

代表单一页眉或页脚。 HeaderFooter 对象是 HeadersFooters 集合的一个成员。 HeadersFooters 集合包含指定工作簿节中的所有页眉和页脚。

## 说明

您也可以使用 HeaderFooter 属性和 Selection 对象来返回单个 HeaderFooter 对象。

| 注释                                                 |
|------------------------------------------------------|
| 不能将 HeaderFooter 对象添加到 HeadersFooters 集合。 |

使用 PageSetup 对象的 DifferentFirstPageHeaderFooter 属性可指定不同的首页。

使用 PageNumbers 对象的 Add 方法可在页眉或页脚中添加页码。以下示例在活动工作簿第一节的主页脚中添加页码。

``` JavaScript
function test() {
    let sheet1 = Application.ActiveSheet.PageSetup
    sheet1.CenterHeader = "&D&T"
    sheet1.OddAndEvenPagesHeaderFooter = false
    sheet1.DifferentFirstPageHeaderFooter = false
    sheet1.ScaleWithDocHeaderFooter = true
    sheet1.AlignMarginsHeaderFooter = true
} 
```

## 属性

| 序号 | 名称                             | 说明                                                                      |
|:----:|:---------------------------------|:--------------------------------------------------------------------------|
|  1   | [Picture](#HeaderFooter.Picture) | 返回一个 Picture 对象，该对象代表指定页眉或页脚中包含的图片字段。只读。   |
|  2   | [Text](#HeaderFooter.Text)       | 返回或设置一个 Text 对象，该对象代表指定页眉或页脚中包含的文本。可读/写。 |

## 成员属性

### HeaderFooter.Picture

返回一个 **Picture** 对象，该对象代表指定页眉或页脚中包含的图片字段。只读。

**语法**

**express.Picture**

\*express是一个代表 HeaderFooter 对象的变量。

### HeaderFooter.Text

返回或设置一个 **Text** 对象，该对象代表指定页眉或页脚中包含的文本。可读/写。

**语法**

**express.Text**

\*express是一个代表 HeaderFooter 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/HeaderFooter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
