# 文字事件

## 事件列表

| 名称                              | 级别        | 类型   | 触发时机                                              |
|-----------------------------------|-------------|--------|-------------------------------------------------------|
| DocumentOpen                      | Application | 通知型 | 文档打开时触发。                                      |
| DocumentBeforeClose               | Application | 询问型 | 任一打开的文档关闭之前触发此事件。                    |
| DocumentBeforeSave                | Application | 询问型 | 任一打开的文档保存之前触发此事件。                    |
| DocumentAfterClose                | Application | 通知型 | 任一打开的文档关闭之后触发此事件。                    |
| DocumentChange                    | Application | 通知型 | 任一打开的文档内容发生变化时触发此事件。              |
| DocumentBeforePrint               | Application | 询问型 | 任一打开的文档打印之前触发此事件。                    |
| DocumentNew                       | Application | 通知型 | 新建任一文档时触发此事件。                            |
| DocumentRightsInfo                | Application | 通知型 | 当操作文档权限时会触发此事件。                        |
| DocumentBeforeCopy                | Application | 询问型 | 在任一文档复制之前触发该事件。                        |
| DocumentBeforePaste               | Application | 询问型 | 在任一文档粘贴之前触发该事件。                        |
| DocumentBeforeNew                 | Application | 询问型 | 在任一文档新建之前触发该事件。                        |
| DocumentBeforeOpen                | Application | 询问型 | 在任一文档打开之前触发该事件。                        |
| DocumentAfterSave                 | Application | 通知型 | 在任一文档保存之后触发该事件。                        |
| DocumentViewFocusIn               | Application | 通知型 | 在任一文档获取到焦点之后触发该事件。                  |
| DocumentViewFocusOut              | Application | 通知型 | 在任一文档失去焦点之后触发该事件。                    |
| DocumentAfterPrint                | Application | 通知型 | 在任一文档打印之后触发该事件。                        |
| WindowActivate                    | Application | 通知型 | 当激活文档窗口时触发此事件。                          |
| WindowDeactivate                  | Application | 通知型 | 任一文档窗口被切换到非激活状态时触发此事件。          |
| WindowSelectionChange             | Application | 通知型 | 活动窗口所选内容更改时将触发此事件。                  |
| WindowBeforeRightClick            | Application | 询问型 | 当右击任一文档窗口之前触发此事件。                    |
| WindowBeforeDoubleClick           | Application | 询问型 | 当双击任一文档窗口之前触发此事件。                    |
| ApplicationQuit                   | Application | 通知型 | 当退出文字组件进程时触发此事件。                      |
| AfterLogin                        | Application | 通知型 | 在用户登陆之后触发此事件。                            |
| AfterLogout                       | Application | 通知型 | 在用户登出之后触发此事件。                            |
| AfterVIPInfoChanged               | Application | 通知型 | 在用户会员信息变化之后触发此事件。                    |
| AfterWebExtensionDataRangeChange  | Application | 通知型 | 在网络扩展数据范围变化之后触发此事件。                |
| AfterTaskPaneShow                 | Application | 通知型 | 当任务窗格显示之后触发此事件。                        |
| AfterTaskPaneHidden               | Application | 通知型 | 当任务窗格隐藏之后触发此事件。                        |
| QueryDocumentPrintCopies          | Application | 询问型 | 当打印复本排队时触发此事件。                          |
| ContentChange                     | Application | 通知型 | 当文档内容发生变化时触发此事件。                      |
| SelectionAfterStyleChange         | Application | 通知型 | 当选择的内容字体变化之后触发此事件。                  |
| FileAfterSave                     | Application | 通知型 | 在文件保存之后触发该事件。                            |
| ContentControlAfterAdd            | Application | 通知型 | 当内容控件被添加到文档之后触发此事件。                |
| ContentControlBeforeDelete        | Application | 通知型 | 当内容控件被从文档删除之前触发此事件。                |
| ContentControlOnExit              | Application | 通知型 | 当退出内容控件时触发此事件。                          |
| ContentControlOnEnter             | Application | 通知型 | 当进入内容控件时触发此事件。                          |
| ContentControlBeforeStoreUpdate   | Application | 通知型 | 当内容控件的内容去更新XML映射内容时触发此事件。       |
| ContentControlBeforeContentUpdate | Application | 通知型 | 当内容控件的内容发生来自XML映射内容更新时触发此事件。 |

## 事件

### DocumentOpen

文档打开时触发。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Doc  | 必选      | Document | 打开的文档对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentOpen " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("打开文档")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentOpen", func) 
```

### DocumentBeforeClose

任一打开的文档关闭之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                        |
|--------|-----------|----------|---------------------------------------------|
| Doc    | 必选      | Document | 文档对象。                                  |
| Cancel | 必选      | Boolean  | 如果设置其属性Value为 true ，则不关闭文档。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeClose " , callbackFunc )`

**示例**

``` JavaScript
function func(doc, cancal) {
    var msg = "是否要关闭文档？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};                                      
Application.ApiEvent.AddApiEventListener("DocumentBeforeClose", func) 
```

### DocumentBeforeSave

任一打开的文档保存之前触发此事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明|
|----------|-----------|----------|--------------------------------------------------------------|
| Doc      | 必选      | Document | 关闭的文档对象。                                             |
| SaveAsUI | 必选      | Boolean  | 如果为 true 则表示此次保存操作将会弹出保存或是另存为对话框。 |
| Cancel   | 必选      | Object   | 如果设置其属性Value为 true ，则不关闭文档。                  |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeSave " , callbackFunc )`

**示例**

``` JavaScript
function func(doc, cancal) {
    var msg = "您保存后要关闭吗？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};                                      
Application.ApiEvent.AddApiEventListener("DocumentBeforeSave", func) 
```

### DocumentAfterClose

任一打开的文档关闭之后触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明       |
|------|-----------|----------|------------|
| Doc  | 必选      | Document | 文档对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentAfterClose " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档已关闭")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentAfterClose", func) 
```

### DocumentChange

任一打开的文档关闭之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档内容发生变化")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentChange", func) 
```

### DocumentBeforePrint

任一打开的文档打印之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                        |
|--------|-----------|----------|---------------------------------------------|
| Doc    | 必选      | Document | 文档对象。                                  |
| Cancel | 必选      | Boolean  | 如果设置其属性Value为 true ，则不打印文档。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforePrint " , callbackFunc )`

**示例**

``` JavaScript
function func(doc, cancal) {
    var msg = "您要打印文档吗吗？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};                                      
Application.ApiEvent.AddApiEventListener("DocumentBeforePrint", func) 
```

### DocumentNew

新建任一文档时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Doc  | 必选      | Document | 新建的文档对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentNew " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("新建文档")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentNew", func) 
```

### DocumentRightsInfo

当操作文档权限时会触发此事件。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明             |
|------------|-----------|----------|------------------|
| Doc        | 必选      | Document | 新建的文档对象。 |
| RightsInfo | 必选      | Int      | 权限信息。       |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentRightsInfo " , callbackFunc )`

**示例**

``` JavaScript
function func(doc, info) {
    var msg = "文档的权限为：" + info;
    alert(msg);
}                                       
Application.ApiEvent.AddApiEventListener("DocumentRightsInfo", func) 
```

### DocumentBeforeCopy

在任一文档复制之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| Doc    | 必选      | Document | 文档对象。                              |
| Type   | 必选      | Long     | 复制类型。                              |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不复制。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeCopy " , callbackFunc )`

**示例**

``` JavaScript
function func(Doc, Type, Cancal) {
    var msg = "是否要复制文档内容"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeCopy", func) 
```

### DocumentBeforePaste

在任一文档粘贴之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| Doc    | 必选      | Document | 文档对象。                              |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不粘贴。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforePaste " , callbackFunc )`

**示例**

``` JavaScript
function func(Doc, Cancal) {
    var msg = "是否要粘贴文档内容"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforePaste", func) 
```

### DocumentBeforeNew

在任一文档新建之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                        |
|--------|-----------|----------|---------------------------------------------|
| Cancel | 必选      | Boolean  | 如果设置其属性Value为 true ，则不新建文档。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeNew " , callbackFunc )`

**示例**

``` JavaScript
function func(Cancal) {
    var msg = "是否要新建文档"
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

| 名称     | 必选/可选 | 数据类型 | 说明                                        |
|----------|-----------|----------|---------------------------------------------|
| FilePath | 必选      | String   | 文档路径                                    |
| Cancel   | 必选      | Object   | 如果设置其属性Value为 true ，则不打开文档。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeOpen " , callbackFunc )`

**示例**

``` JavaScript
function func(FilePath, Cancal) {
    var msg = "是否要打开文档" + FilePath
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeOpen", func) 
```

### DocumentAfterSave

在任一文档保存之后触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                      |
|--------|-----------|----------|-------------------------------------------|
| Doc    | 必选      | Document | 文档对象。                                |
| Result | 必选      | Boolean  | 如果保存操作成功，则为true; 否则为false。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentAfterSave " , callbackFunc )`

**示例**

``` JavaScript
function func(Result) {
    if(!Result)
        alert("保存成功")
};                                      
Application.ApiEvent.AddApiEventListener("DocumentAfterSave", func) 
```

### DocumentViewFocusIn

在任一文档获取到焦点之后触发该事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明       |
|------|-----------|----------|------------|
| Doc  | 必选      | Document | 文档对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentViewFocusIn " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档获取到焦点")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentViewFocusIn", func) 
```

### DocumentViewFocusOut

在任一文档失去焦点之后触发该事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明       |
|------|-----------|----------|------------|
| Doc  | 必选      | Document | 文档对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentViewFocusOut " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档失去焦点")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentViewFocusOut", func) 
```

### DocumentAfterPrint

在任一文档打印之后触发该事件。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明           |
|--------------|-----------|----------|----------------|
| Doc          | 必选      | Document | 文档对象。     |
| PrintOutPage | 必选      | Object   | 打印页面对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentAfterPrint " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档已打印")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentAfterPrint", func) 
```

### WindowActivate

当激活文档窗口时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明           |
|------|-----------|----------|----------------|
| Doc  | 必选      | Document | 激活文档对象。 |
| Wn   | 必选      | Window   | 激活窗口对象。 |

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

任一文档窗口被切换到非激活状态时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明           |
|------|-----------|----------|----------------|
| Doc  | 必选      | Document | 激活文档对象。 |
| Wn   | 必选      | Window   | 激活窗口对象。 |

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

### WindowSelectionChange

活动窗口所选内容更改时将触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明       |
|------|-----------|----------|------------|
| Sel  | 必选      | Object   | 选择区域。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowSelectionChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("窗口所选内容更改")
}                                       
Application.ApiEvent.AddApiEventListener("WindowSelectionChange", func) 
```

### WindowBeforeRightClick

当右击任一文档窗口之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                      |
|--------|-----------|----------|-------------------------------------------|
| Sel    | 必选      | Object   | 选择区域。                                |
| Cancel | 必选      | Boolean  | 如果设置其属性Value为 true ，则取消单击。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowBeforeRightClick " , callbackFunc )`

**示例**

``` JavaScript
function func(Cancal) {
    var msg = "是否要取消单击"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("WindowBeforeRightClick", func) 
```

### WindowBeforeDoubleClick

当双击任一文档窗口之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                      |
|--------|-----------|----------|-------------------------------------------|
| Sel    | 必选      | Object   | 选择区域。                                |
| Cancel | 必选      | Boolean  | 如果设置其属性Value为 true ，则取消双击。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowBeforeDoubleClick " , callbackFunc )`

**示例**

``` JavaScript
function func(Cancal) {
    var msg = "是否要取消双击"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("WindowBeforeDoubleClick", func) 
```

### ApplicationQuit

当退出文字组件进程时触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ApplicationQuit " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文字组件退出")
}                                       
Application.ApiEvent.AddApiEventListener("ApplicationQuit", func) 
```

### AfterLogin

在用户登陆之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterLogin " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("用户已登陆")
}                                       
Application.ApiEvent.AddApiEventListener("AfterLogin", func) 
```

### AfterLogout

在用户登出之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterLogout " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("用户已登出")
}                                       
Application.ApiEvent.AddApiEventListener("AfterLogout", func) 
```

### AfterVIPInfoChanged

在用户会员信息变化之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterVIPInfoChanged " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("用户会员信息已变化")
}                                       
Application.ApiEvent.AddApiEventListener("AfterVIPInfoChanged", func) 
```

### AfterWebExtensionDataRangeChange

在网络扩展数据范围变化之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterWebExtensionDataRangeChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("网络扩展数据范围已变化")
}                                       
Application.ApiEvent.AddApiEventListener("AfterWebExtensionDataRangeChange", func) 
```

### AfterTaskPaneShow

当任务窗格显示之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterTaskPaneShow " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
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
let func = (doc)=>{
    alert("任务窗格已隐藏")
}                                       
Application.ApiEvent.AddApiEventListener("AfterTaskPaneHidden", func) 
```

### QueryDocumentPrintCopies

当打印复本排队时触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| Doc    | 必选      | Document | 文档对象。                              |
| Copies | 必选      | Int      | 要打印的份数。                          |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不打印。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " QueryDocumentPrintCopies " , callbackFunc )`

**示例**

``` JavaScript
function func(Doc, Copies, Cancal) {
    var msg = "是否要打印文档"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("QueryDocumentPrintCopies", func) 
```

### ContentChange

当文档内容发生变化时触发此事件。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明       |
|-------|-----------|----------|------------|
| Doc   | 必选      | Document | 文档对象。 |
| Range | 必选      | Object   | 范围对象。 |
| Type  | 必选      | Int      | 变化类型。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档内容发生变化")
}                                       
Application.ApiEvent.AddApiEventListener("ContentChange", func) 
```

### SelectionAfterStyleChange

当选择的内容字体变化之后触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明       |
|------|-----------|----------|------------|
| Sel  | 必选      | Object   | 选择区域。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SelectionAfterStyleChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("内容字体变化")
}                                       
Application.ApiEvent.AddApiEventListener("SelectionAfterStyleChange", func) 
```

### FileAfterSave

在文件保存之后触发该事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明       |
|----------|-----------|----------|------------|
| Doc      | 必选      | Document | 文档对象。 |
| FilePath | 必选      | String   | 保存路径。 |
| Format   | 必选      | String   | 保存格式。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " FileAfterSave " , callbackFunc )`

**示例**

``` JavaScript
function func(Doc, FilePath, Format) {
    var msg = "文档以" + Format + "格式保存在" + FilePath
    alret(msg);
}                                       
Application.ApiEvent.AddApiEventListener("FileAfterSave", func) 
```

### ContentControlAfterAdd

当内容控件被添加到文档之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlAfterAdd " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("内容控件被添加")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlAfterAdd", func) 
```

### ContentControlBeforeDelete

当内容控件被从文档删除之前触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlBeforeDelete " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("内容控件被从文档删除")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlBeforeDelete", func) 
```

### ContentControlOnExit

当退出内容控件时触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlOnExit " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("退出内容控件")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlOnExit", func) 
```

### ContentControlOnEnter

当进入内容控件时触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlOnEnter " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("进入内容控件")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlOnEnter", func) 
```

### ContentControlBeforeStoreUpdate

当内容控件的内容去更新XML映射内容时触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlBeforeStoreUpdate " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("内容控件的内容更新XML映射内容")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlBeforeStoreUpdate", func) 
```

### ContentControlBeforeContentUpdate

当内容控件的内容发生来自XML映射内容更新时触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlBeforeContentUpdate " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("内容控件的内容发生来自XML映射内容更新")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlBeforeContentUpdate", func) 
```

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/事件/文字事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%96%87%E5%AD%97%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
