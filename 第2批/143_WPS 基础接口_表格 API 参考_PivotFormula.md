# PivotFormula 对象

## 简介

代表在数据透视表中用于计算的公式。

## 说明

本对象及其相关属性和方法对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效，这是因为它不支持计算字段和计算项。

使用 PivotFormulas ( *index* )（其中 *index* 是公式左侧的公式号或字符串）可返回 PivotFormula 对象。下例将更改公式一（第一张工作表的第一个数据透视表中）的索引号，使其在公式二计算完毕后再进行计算。

``` JavaScript
Worksheets.Item(1).PivotTables(1).PivotFormulas.Item(1).Index = 2
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
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Delete">Delete</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>删除对象。</p></td>
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
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 <span>Application</span> 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Formula">Formula</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 String 值，它代表 A1 样式表示法和宏语言中的对象的公式。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Index">Index</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 Long 值，它代表 <span>PivotFormula</span> 对象在 <span>PivotFormulas</span> 集合中的索引号。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.StandardFormula">StandardFormula</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 String 值，该值指定使用标准英语（美国）格式的公式。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotFormula.Value">Value</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 String 值，它代表数据透视表中指定的公式的名称。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PivotFormula.Delete

table { word-break:break-all; }

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 PivotFormula 对象的变量。

## 成员属性

### PivotFormula.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotFormula 对象的变量。

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

### PivotFormula.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotFormula 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotFormula.Formula

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

**语法**

**express.Formula**

\*express是一个代表 PivotFormula 对象的变量。

**说明**

table { word-break:break-all; }

此属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回空字符串。如果单元格包含公式， **Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

### PivotFormula.Index

table { word-break:break-all; }

返回或设置一个 **Long** 值，它代表 **PivotFormula** 对象在 **PivotFormulas** 集合中的索引号。

**语法**

**express.Index**

\*express是一个代表 PivotFormula 对象的变量。

### PivotFormula.Parent

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotFormula 对象的变量。

### PivotFormula.StandardFormula

返回或设置一个 **String** 值，该值指定使用标准英语（美国）格式的公式。可读写。

**语法**

**express.StandardFormula**

\*express是一个代表 PivotFormula 对象的变量。

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

### PivotFormula.Value

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表数据透视表中指定的公式的名称。

**语法**

**express.Value**

\*express是一个代表 PivotFormula 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotFormula](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
