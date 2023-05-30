# PageNumber 对象

## 简介

代表页眉或页脚中的页码。 PageNumber 对象是 PageNumbers 集合的一个成员。 PageNumbers 集合包含单个页眉或页脚中的所有页码。

## 说明

使用 PageNumbers ( *Index* ) 可返回一个 PageNumber 对象，其中 *Index* 为索引号。在大多数情况下，一个页眉或页脚只包含一个页码（索引号为 1）。以下示例将活动文档第一节的主页眉中的起始页码居中。

``` JavaScript
Application.ActiveDocument.Sections.Item(1).Headers(wdHeaderFooterPrimary).PageNumbers.Item(1).Alignment = wdAlignPageNumberCenter
```

使用 Add 方法可在页眉或页脚中添加页码（PAGE 域）。以下示例向第一节和任何后续节中的主页脚添加一个页码。

``` JavaScript
function test()
{
    let sect = Application.ActiveDocument.Sections.Item(1)
    sect.Footers(wdHeaderFooterPrimary).PageNumbers.Add(wdAlignPageNumberLeft, true)
}
```

## 属性

| 序号 | 名称                               | 说明                                                                          |
|:----:|:-----------------------------------|:------------------------------------------------------------------------------|
|  1   | [Alignment](#PageNumber.Alignment) | 返回或设置一个 WdPageNumberAlignment 常量，该常量代表页码的对齐方式，可读写。 |

## 成员属性

### PageNumber.Alignment

返回或设置一个 **WdPageNumberAlignment** 常量，该常量代表页码的对齐方式，可读写。

**语法**

**express.Alignment**

\*express是一个代表 PageNumber 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/PageNumber](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
