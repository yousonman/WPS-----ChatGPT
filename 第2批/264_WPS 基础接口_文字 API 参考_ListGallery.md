# ListGallery 对象

## 简介

代表单个列表格式库。 ListGallery 对象是 ListGalleries 集合的一个成员。

## 说明

每个 ListGallery 对象代表 “项目符号和编号” 对话框中的三个选项卡之一。

使用 ListGalleries ( *Index* ) 可返回一个 ListGallery 对象，其中 *Index* 为 wdBulletGallery 、 wdNumberGallery 或 wdOutlineNumberGallery 。

以下示例返回 “项目符号和编号” 对话框中 “项目符号” 选项卡上的第三种列表格式（不包括 “无” ），并将其应用于所选内容。

``` JavaScript
function test() {
    let temp3 = Application.ListGalleries.Item(wdBulletGallery).ListTemplates.Item(3)
    Application.Selection.Range.ListFormat.ApplyListTemplate(temp3)
}
```

要查看指定列表模板是否包含 WPS 内置格式，请对 ListGallery 对象使用 Modified 属性。要将格式重新设置为原始列表格式，请对 ListGallery 对象使用 Reset 方法。

## 属性

| 序号 | 名称                                        | 说明                                                                    |
|:----:|:--------------------------------------------|:------------------------------------------------------------------------|
|  1   | [ListTemplates](#ListGallery.ListTemplates) | 返回一个 ListTemplates 集合，该集合代表指定列表库的所有列表格式。只读。 |

## 成员属性

### ListGallery.ListTemplates

返回一个 **ListTemplates** 集合，该集合代表指定列表库的所有列表格式。只读。

**语法**

**express.ListTemplates**

\*express是一个代表 ListGallery 对象的变量。

**示例**

``` JavaScript
/* 以下代码将选区设置为项目符号的第3钟符号模板 */
function test() {
    let temp3 = Application.ListGalleries.Item(wdBulletGallery).ListTemplates.Item(3)
    Application.Selection.Range.ListFormat.ApplyListTemplate(temp3)
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/ListGallery](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
