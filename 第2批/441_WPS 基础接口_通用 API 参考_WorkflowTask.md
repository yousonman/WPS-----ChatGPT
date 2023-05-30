# WorkflowTask 对象

## 简介

代表 WorkflowTasks 集合中的单个工作流任务。

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

| 序号 | 名称                       | 说明                                                 |
|:----:|:---------------------------|:-----------------------------------------------------|
|  1   | [Show](#WorkflowTask.Show) | 显示指定 WorkflowTask 对象的工作流任务编辑用户界面。 |

## 属性

| 序号 | 名称                                     | 说明                                                                         |
|:----:|:-----------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#WorkflowTask.Application) | 获取一个代表 WorkflowTemplate 对象的容器应用程序的 Application 对象。只读。  |
|  2   | [AssignedTo](#WorkflowTask.AssignedTo)   | 获取向其分配了工作流任务的人员的姓名。只读。                                 |
|  3   | [CreatedBy](#WorkflowTask.CreatedBy)     | 获取创建工作流任务的人员的姓名。只读                                         |
|  4   | [CreatedDate](#WorkflowTask.CreatedDate) | 获取工作流任务的创建日期。只读。                                             |
|  5   | [Creator](#WorkflowTask.Creator)         | 获取一个 32 位整数，指示创建 WorkflowTemplate 对象时所使用的应用程序。只读。 |
|  6   | [Description](#WorkflowTask.Description) | 获取工作流模板的说明。只读。                                                 |
|  7   | [DueDate](#WorkflowTask.DueDate)         | 获取工作流任务的截止日期。只读                                               |
|  8   | [Id](#WorkflowTask.Id)                   | 获取用于创建工作流实例的模板的 ID。只读。                                    |
|  9   | [ListID](#WorkflowTask.ListID)           | 获取包含工作流任务的列表的 ID。只读。                                        |
|  10  | [Name](#WorkflowTask.Name)               | 获取 WorkflowTemplate 对象的名称。只读。                                     |
|  11  | [WorkflowID](#WorkflowTask.WorkflowID)   | 获取与工作流任务关联的工作流的 ID。只读。                                    |

## 成员方法

### WorkflowTask.Show

显示指定 **WorkflowTask** 对象的工作流任务编辑用户界面。

**语法**

**express.Show ()**

\*express是一个代表 WorkflowTask 对象的变量。

**返回值**

Integer

**示例**

``` JavaScript
/*以下示例显示当前文档中每个工作流任务的名称，然后显示特定任务的工作流任务编辑用户界面。*/
function test(){
    const objWorkflowTasks = Application.ActiveWorkbook.GetWorkflowTasks();
    for(let i = 1; i <= objWorkflowTasks.Count; i++)
        Debug.Print(objWorkflowTasks.Item(i).Name);

    var objWorkflowTask = objWorkflowTasks.Item(1);
    objWorkflowTask.Show();
}
```

## 成员属性

### WorkflowTask.Application

获取一个代表 **WorkflowTemplate** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 WorkflowTask 对象的变量。

**说明**

### WorkflowTask.AssignedTo

获取向其分配了工作流任务的人员的姓名。只读。

**语法**

**express.AssignedTo**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.CreatedBy

获取创建工作流任务的人员的姓名。只读

**语法**

**express.CreatedBy**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.CreatedDate

获取工作流任务的创建日期。只读。

**语法**

**express.CreatedDate**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.Creator

获取一个 32 位整数，指示创建 **WorkflowTemplate** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.Description

获取工作流模板的说明。只读。

**语法**

**express.Description**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.DueDate

获取工作流任务的截止日期。只读

**语法**

**express.DueDate**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.Id

获取用于创建工作流实例的模板的 ID。只读。

**语法**

**express.Id**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.ListID

获取包含工作流任务的列表的 ID。只读。

**语法**

**express.ListID**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.Name

获取 **WorkflowTemplate** 对象的名称。只读。

**语法**

**express.Name**

\*express是一个代表 WorkflowTask 对象的变量。

### WorkflowTask.WorkflowID

获取与工作流任务关联的工作流的 ID。只读。

**语法**

**express.WorkflowID**

\*express是一个代表 WorkflowTask 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/WorkflowTask](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
