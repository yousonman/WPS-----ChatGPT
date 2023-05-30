# 表格事件

## 事件列表

| 名称                           | 级别        | 类型   | 触发时机|
|--------------------------------|-------------|--------|-----------|
| WorkbookOpen                   | Application | 通知型 | 当打开一个工作簿时触发此事件。                                   |
| NewWorkbook                    | Application | 通知型 | 当新建工作簿时触发此事件。                                       |
| WorkbookBeforeSave             | Application | 询问型 | 任一工作簿被保存之前触发此事件。                                 |
| WorkbookAfterSave              | Application | 通知型 | 任一工作簿被保存之后触发此事件。                                 |
| WorkbookBeforeClose            | Application | 询问型 | 任一打开的工作簿关闭之前触发此事件。                             |
| WorkbookBeforePrint            | Application | 询问型 | 在打印任一打开的工作簿之前触发此事件。                           |
| SheetSelectionChange           | Application | 通知型 | 该工作簿任一工作表上的选定区域发生更改时，将触发此事件。         |
| SheetChange                    | Application | 通知型 | 当用户或外部链接更改了该工作簿任一工作表中的单元格时触发此事件。 |
| SheetActivate                  | Application | 通知型 | 当激活任一工作表时触发此事件。                                   |
| SheetDeactivate                | Application | 通知型 | 当该工作簿任一工作表被切换到非激活状态时触发此事件。             |
| WindowActivate                 | Application | 通知型 | 任一工作簿窗口被激活时，将触发此事件。                           |
| WindowDeactivate               | Application | 通知型 | 任一工作簿窗口被切换到非激活状态时触发此事件。                   |
| SheetBeforeDoubleClick         | Application | 询问型 | 当双击任一工作表之前触发此事件。                                 |
| SheetBeforeRightClick          | Application | 询问型 | 当右击任一工作表之前触发此事件 。                                |
| SheetCalculate                 | Application | 通知型 | 在任一工作表进行计算时触发此事件。                               |
| WorkbookActivate               | Application | 通知型 | 任一工作簿被激活时，将触发此事件。                               |
| WorkbookDeactivate             | Application | 通知型 | 任一工作簿被切换到非激活状态时触发此事件。                       |
| WorkbookNewSheet               | Application | 通知型 | 在任一打开的工作簿中创建新工作表时触发此事件。                   |
| WindowResize                   | Application | 通知型 | 任一工作簿窗口调整大小时将触发此事件。                           |
| SheetFollowHyperlink           | Application | 通知型 | 单击任一工作表的超链接时触 发此事件。                            |
| AfterCalculate                 | Application | 通知型 | 当计算完成后触发此事件。                                         |
| SheetBeforeDelete              | Application | 通知型 | 当删除任一工作表之前触发此事件。                                 |
| DocumentRightsInfo             | Application | 通知型 | 当操作文档权限时会触发此事件。                                   |
| AfterLogin                     | Application | 通知型 | 当用户登陆之后触发此事件。                                       |
| AfterLogout                    | Application | 通知型 | 当用户登出之后触发此事件。                                       |
| AfterVIPInfoChanged            | Application | 通知型 | 当用户的会员信息变化之后触发此事件。                             |
| AfterWebExtWatchingDataUpdated | Application | 通知型 | 当网页拓展监控的数据发生变化之后触发此事件                       |
| AfterTaskPaneShow              | Application | 通知型 | 当任务窗格显示之后触发此事件。                                   |
| AfterTaskPaneHidden            | Application | 通知型 | 当任务窗格隐藏之后触发此事件                                     |
| DocumentBeforeNew              | Application | 询问型 | 在任一文档新建之前触发该事件。                                   |
| DocumentBeforeOpen             | Application | 询问型 | 在任一文档打开之前触发该事件。                                   |
| DocumentBeforeCopy             | Application | 询问型 | 在任一文档复制之前触发该事件。                                   |
| DocumentBeforePaste            | Application | 询问型 | 在任一文档粘贴之前触发该事件。                                   |
| FileAfterSave                  | Application | 通知型 | 在任一文档保存之后触发该事件。                                   |
| LinkedDataTypeConvert          | Application | 通知型 | 在连接数据类型转变时触发该事件。                                 |
| LinkedDataTypeChange           | Application | 通知型 | 在连接数据类型改变时触发该事件。                                 |
| LinkedDataTypeRefresh          | Application | 通知型 | 在连接数据类型刷新时触发该事件。                                 |
| LinkedDataTypeCancel           | Application | 通知型 | 在连接数据类型取消时触发该事件。                                 |
| DocumentAfterPrint             | Application | 通知型 | 该文档打印之后触发该事件。                                       |

## 事件

### WorkbookOpen

当打开一个工作簿时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Wb   | 必选      | Workbook | 打开的工作簿对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( "WorkbookOpen" , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("打开工作簿")
}                                       
Application.ApiEvent.AddApiEventListener("WorkbookOpen", func) 
```

### NewWorkbook

当新建工作簿时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Wb   | 必选      | Workbook | 新建的工作簿对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " NewWorkbook " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("新建工作簿")
}                                       
Application.ApiEvent.AddApiEventListener("NewWorkbook", func) 
```

### WorkbookBeforeSave

任一工作簿被保存之前触发此事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明|
|----------|-----------|----------|--------------------------------------------------------------|
| Wb       | 必选      | Workbook | 工作簿对象。                                                 |
| SaveAsUI | 必选      | Boolean  | 如果为 true 则表示此次保存操作将会弹出保存或是另存为对话框。 |
| Cancel   | 必选      | Object   | 如果设置其属性Value为 true ，则不关闭工作簿。                |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookBeforeSave " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, cancal) {
    var msg = "您保存后要关闭吗？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};
Application.ApiEvent.AddApiEventListener("WorkbookBeforeSave", func) 
```

### WorkbookAfterSave

任一工作簿被保存之后触发此事件。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                      |
|---------|-----------|----------|-------------------------------------------|
| Wb      | 必选      | Workbook | 工作簿对象。                              |
| Success | 必选      | Boolean  | 如果保存操作成功，则为true; 否则为false。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookAfterSave " , callbackFunc )`

**示例**

``` JavaScript
function func(Success) {
    if(!Success)
        alert("保存成功")
};  
Application.ApiEvent.AddApiEventListener("WorkbookAfterSave", func) 
```

### WorkbookBeforeClose

任一打开的工作簿关闭之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Wb     | 必选      | Workbook | 工作簿对象。                                  |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不关闭工作簿。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookBeforeClose " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, cancal) {
    var msg = "您确定要关闭吗？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};  
Application.ApiEvent.AddApiEventListener("WorkbookBeforeClose", func) 
```

### WorkbookBeforePrint

在打印任一打开的工作簿之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Wb     | 必选      | Workbook | 工作簿对象。                                  |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不打印工作簿。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookBeforePrint " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, cancal) {
    var msg = "您打印后要关闭吗？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};
Application.ApiEvent.AddApiEventListener("WorkbookBeforePrint", func) 
```

### SheetSelectionChange

该工作簿任一工作表上的选定区域发生更改时，将触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                   |
|--------|-----------|----------|------------------------|
| Sh     | 必选      | Object   | 选区改变的工作表对象。 |
| Target | 必选      | Range    | 新选定的区域。         |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetSelectionChange " , callbackFunc )`

**示例**

``` JavaScript
function func(Sh, Target) {
    var msg = "工作表" + Sh + "选中的区域是：" + Target.Address();
    alert(msg);
}
Application.ApiEvent.AddApiEventListener("SheetSelectionChange", func) 
```

### SheetChange

当用户或外部链接更改了该工作簿任一工作表中的单元格时触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                   |
|--------|-----------|----------|------------------------|
| Sh     | 必选      | Object   | 选区改变的工作表对象。 |
| Target | 必选      | Range    | 新选定的区域。         |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetChange " , callbackFunc )`

**示例**

``` JavaScript
function func(Sh, Target) {
    var msg = "工作表" + Sh + "选中的区域是：" + Target.Address();
    alert(msg);
}
Application.ApiEvent.AddApiEventListener("SheetChange", func) 
```

### SheetActivate

当激活任一工作表时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明         |
|------|-----------|----------|--------------|
| Sh   | 必选      | Object   | 工作表对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetActivate " , callbackFunc )`

**示例**

``` JavaScript
let func = (Sh)=>{
    alert("工作表激活")
}                                       
Application.ApiEvent.AddApiEventListener("SheetActivate", func) 
```

### SheetDeactivate

当该工作簿任一工作表被切换到非激活状态时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明         |
|------|-----------|----------|--------------|
| Sh   | 必选      | Object   | 工作表对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetDeactivate " , callbackFunc )`

**示例**

``` JavaScript
let func = (Sh)=>{
    alert("工作表切换为未激活")
}                                       
Application.ApiEvent.AddApiEventListener("SheetDeactivate", func) 
```

### WindowActivate

任一工作簿窗口被激活时，将触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |
| Wn   | 必选      | Window   | 激活窗口对象。   |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowActivate " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, wn) {
    var msg = "您激活的窗口是: " + wn.Caption;
    alret(msg);
}           
Application.ApiEvent.AddApiEventListener("WindowActivate", func) 
```

### WindowDeactivate

任一工作簿窗口被切换到非激活状态时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |
| Wn   | 必选      | Window   | 激活窗口对象。   |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowDeactivate " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, wn) {
    var msg = "您取消激活的窗口是: " + wn.Caption;
    alret(msg);
}                                       
Application.ApiEvent.AddApiEventListener("WindowDeactivate", func) 
```

### SheetBeforeDoubleClick

当双击任一工作表之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sh     | 必选      | Object    | 双击的工作表对象。                            |
| Target | 必选      | Range对象 | 双击区域所在的单元格对象。                    |
| Cancel | 必选      | Object    | 如果设置其属性Value为 true ，则取消此次双击。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetBeforeDoubleClick " , callbackFunc )`

**示例**

``` JavaScript
function func(Sh, Target, Cancal) {
    var msg = "您双击的是：" + Target.Address() +"\n是否取消此次双击？"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("SheetBeforeDoubleClick", func) 
```

### SheetBeforeRightClick

当右击任一工作表之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sh     | 必选      | Object    | 单击的工作表对象。                            |
| Target | 必选      | Range对象 | 单击区域所在的单元格对象。                    |
| Cancel | 必选      | Object    | 如果设置其属性Value为 true ，则取消此次单击。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetBeforeRightClick " , callbackFunc )`

**示例**

``` JavaScript
function func(Sh, Target, Cancal) {
    var msg = "您单击的是：" + Target.Address() +"\n是否取消此次单击？"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("SheetBeforeRightClick", func) 
```

### SheetCalculate

在任一工作表进行计算时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Sh   | 必选      | Object   | 激活的工作表对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetCalculate " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("进行计算")
}                                       
Application.ApiEvent.AddApiEventListener("SheetCalculate", func) 
```

### WorkbookActivate

任一工作簿被激活时，将触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookActivate " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("工作簿激活")
}                                       
Application.ApiEvent.AddApiEventListener("WorkbookActivate", func) 
```

### WorkbookDeactivate

任一工作簿被切换到非激活状态时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookDeactivate " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("工作簿切换为未激活")
}                                       
Application.ApiEvent.AddApiEventListener("WorkbookDeactivate", func) 
```

### WorkbookNewSheet

在任一打开的工作簿中创建新工作表时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |
| Sh   | 必选      | Object   | 激活工作表对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookNewSheet " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("新建工作表")
}                                       
Application.ApiEvent.AddApiEventListener("WorkbookNewSheet", func) 
```

### WindowResize

任一工作簿窗口调整大小时将触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |
| Wn   | 必选      | Window   | 激活窗口对象。   |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowResize " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("窗口调整大小")
}                                       
Application.ApiEvent.AddApiEventListener("WindowResize", func) 
```

### SheetFollowHyperlink

单击任一工作表的超链接时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Sh   | 必选      | Object   | 激活工作表对象。 |
| Hl   | 必选      | Object   | 超链接对象。     |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetFollowHyperlink " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("跳转超链接")
}                                       
Application.ApiEvent.AddApiEventListener("SheetFollowHyperlink", func) 
```

### AfterCalculate

当计算完成后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterCalculate " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("计算完成")
}                                       
Application.ApiEvent.AddApiEventListener("AfterCalculate", func) 
```

### SheetBeforeDelete

当删除任一工作表之前触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Sh   | 必选      | Object   | 激活工作表对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetBeforeDelete " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("准备删除工作表")
}                                       
Application.ApiEvent.AddApiEventListener("SheetBeforeDelete", func) 
```

### DocumentRightsInfo

当操作文档权限时会触发此事件。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明         |
|------------|-----------|----------|--------------|
| Wb         | 必选      | Workbook | 工作簿对象。 |
| RightsInfo | 必选      | Int      | 权限信息。   |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentRightsInfo " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("操作文档权限")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentRightsInfo", func) 
```

### AfterLogin

当用户登陆之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterLogin " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("用户已登陆")
}                                       
Application.ApiEvent.AddApiEventListener("AfterLogin", func) 
```

### AfterLogout

当用户登出之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterLogout " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("用户已登出")
}                                       
Application.ApiEvent.AddApiEventListener("AfterLogout", func) 
```

### AfterVIPInfoChanged

当用户的会员信息变化之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterVIPInfoChanged " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("用户会员信息已变化")
}                                       
Application.ApiEvent.AddApiEventListener("AfterVIPInfoChanged", func) 
```

### AfterWebExtWatchingDataUpdated

当网页拓展监控的数据发生变化之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterWebExtWatchingDataUpdated " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("网页拓展监控的数据发生变化")
}                                       
Application.ApiEvent.AddApiEventListener("AfterWebExtWatchingDataUpdated", func) 
```

### AfterTaskPaneShow

当任务窗格显示之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterTaskPaneShow " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("任务窗格已显示")
}                                       
Application.ApiEvent.AddApiEventListener("AfterTaskPaneShow", func) 
```

### AfterTaskPaneHidden

当任务窗格隐藏之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterTaskPaneHidden " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("任务窗格已隐藏")
}                                       
Application.ApiEvent.AddApiEventListener("AfterTaskPaneHidden", func) 
```

### DocumentBeforeNew

在任一文档新建之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不新建工作簿。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeNew " , callbackFunc )`

**示例**

``` JavaScript
function func(Cancal) {
    var msg = "是否要新建工作簿"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                   
Application.ApiEvent.AddApiEventListener("DocumentBeforeNew", func) 
```

### DocumentBeforeOpen

在任一文档打开之前触发该事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|-----------------------------------------------|
| FilePath | 必选      | String   | 工作簿路径                                    |
| Cancel   | 必选      | Object   | 如果设置其属性Value为 true ，则不打开工作簿。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeOpen " , callbackFunc )`

**示例**

``` JavaScript
function func(FilePath, Cancal) {
    var msg = "是否要打开文档：" + FilePath
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeOpen", func) 
```

### DocumentBeforeCopy

在任一文档复制之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| Wb     | 必选      | Workbook | 工作簿对象。                            |
| Type   | 必选      | Long     | 复制类型。                              |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不复制。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeCopy " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("准备复制")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeCopy", func) 
```

### DocumentBeforePaste

在任一文档粘贴之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| Wb     | 必选      | Workbook | 工作簿对象。                            |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不粘贴。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforePaste " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("准备粘贴")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforePaste", func) 
```

### FileAfterSave

在任一文档保存之后触发该事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明         |
|----------|-----------|----------|--------------|
| Wb       | 必选      | Workbook | 工作簿对象。 |
| FilePath | 必选      | String   | 保存路径。   |
| Format   | 必选      | String   | 保存格式。   |

**语法**

`Application.ApiEvent.AddApiEventListener( " FileAfterSave " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("文档已保存")
}                                       
Application.ApiEvent.AddApiEventListener("FileAfterSave", func) 
```

### LinkedDataTypeConvert

在连接数据类型转变时触发该事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " LinkedDataTypeConvert " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("连接数据类型转变")
}                                       
Application.ApiEvent.AddApiEventListener("LinkedDataTypeConvert", func) 
```

### LinkedDataTypeChange

在连接数据类型改变时触发该事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " LinkedDataTypeChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("连接数据类型改变")
}                                       
Application.ApiEvent.AddApiEventListener("LinkedDataTypeChange", func) 
```

### LinkedDataTypeRefresh

在连接数据类型刷新时触发该事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " LinkedDataTypeRefresh " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("连接数据类型刷新")
}                                       
Application.ApiEvent.AddApiEventListener("LinkedDataTypeRefresh", func) 
```

### LinkedDataTypeCancel

在连接数据类型取消时触发该事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " LinkedDataTypeCancel " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("连接数据类型取消")
}                                       
Application.ApiEvent.AddApiEventListener("LinkedDataTypeCancel", func) 
```

### DocumentAfterPrint

该文档打印之后触发该事件。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明           |
|--------------|-----------|----------|----------------|
| Wb           | 必选      | Workbook | 工作簿对象。   |
| PrintOutPage | 必选      | Object   | 打印页面对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentAfterPrint " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("文档已打印")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentAfterPrint", func) 
```

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/事件/表格事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E8%A1%A8%E6%A0%BC%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
