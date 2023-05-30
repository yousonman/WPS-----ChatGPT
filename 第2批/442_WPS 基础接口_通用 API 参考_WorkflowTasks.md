# WorkflowTasks 对象

## 简介

代表 WorkflowTask 对象的集合。

## 说明

``` JavaScript
/*以下示例显示当前文档中每个工作流任务的名称，然后显示特定任务的工作流任务编辑用户界面。应该注意到，调用 GetWorkflowTasks 方法涉及到服务器的往返行程。*/
function test(){
    const objWorkflowTasks = Application.ActiveWorkbook.GetWorkflowTasks();
    for(let i = 1; i <= objWorkflowTasks.Count; i++)
        Debug.Print(objWorkflowTasks.Item(i).Name);
        
    var objWorkflowTask = objWorkflowTasks.Item(1);
    objWorkflowTask.Show();
}
```

## 方法

| 序号 | 名称                        | 说明                                                |
|:----:|:----------------------------|:----------------------------------------------------|
|  1   | [Item](#WorkflowTasks.Item) | 获取 WorkflowTasks 集合中的一个 WorkflowTask 对象。 |

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
<td style="text-align: left;" data-valian="middle"><a href="#WorkflowTasks.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取一个代表 WorkflowTasks 对象的容器应用程序的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorkflowTasks.Count">Count</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; } &lt;STYLE TYPE="text/CSS"&gt; .ACICollapsed { display: inline; } .ACECollapsed { display: block; } &lt;/STYLE&gt;</p>
<p>获取一个 Long 类型的值，指示 WorkflowTasks 集合中的项数。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#WorkflowTasks.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取一个 32 位整数，指示创建 WorkflowTasks 对象时所使用的应用程序。只读。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### WorkflowTasks.Item

获取 **WorkflowTasks** 集合中的一个 **WorkflowTask** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 WorkflowTasks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                 |
|---------|-----------|----------|--------------------------------------|
| *Index* | 必选      | Long     | 要返回的 WorkflowTask 对象的索引号。 |

**返回值**

WorkflowTask

## 成员属性

### WorkflowTasks.Application

获取一个代表 **WorkflowTasks** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 WorkflowTasks 对象的变量。

### WorkflowTasks.Count

table { word-break:break-all; } \<STYLE TYPE="text/CSS"\> .ACICollapsed { display: inline; } .ACECollapsed { display: block; } \</STYLE\>

获取一个 **Long** 类型的值，指示 **WorkflowTasks** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 WorkflowTasks 对象的变量。

### WorkflowTasks.Creator

获取一个 32 位整数，指示创建 **WorkflowTasks** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 WorkflowTasks 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/WorkflowTasks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
