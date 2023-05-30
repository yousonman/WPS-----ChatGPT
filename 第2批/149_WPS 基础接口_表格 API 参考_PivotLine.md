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
