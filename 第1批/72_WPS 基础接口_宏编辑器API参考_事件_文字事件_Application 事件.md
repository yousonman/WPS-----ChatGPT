# Application 事件

## 事件列表

| 名称                    | 触发时机                                                    |
|-------------------------|-------------------------------------------------------------|
| DocumentBeforeClose     | 当文档关闭前时触发此事件。                                  |
| DocumentBeforePrint     | 当文档打印之前时 触发 此事件。                              |
| DocumentChange          | 当创建新文档、打开已有文档或者切换激活文档 时 触发 此事件。 |
| DocumentOpen            | 当任一文档打开时 触发 此事件。                              |
| NewDocument             | 当 创建新文档 时 触发 此事件 。                             |
| Quit                    | 当关闭退出文字组件进程时触发此事件。                        |
| WindowActivate          | 当激活文档窗口 时 触发 此事件。                             |
| WindowBeforeDoubleClick | 当双击任一 文档窗口 之前 触发 此事件。                      |
| WindowBeforeRightClick  | 当右击任一 文档窗口 之前 触发 此事件 。                     |
| WindowDeactivate        | 任一 文档 窗口被切换到非激活状态时触发此事件。              |
| WindowSelectionChange   | 活动窗口所选内容更改时将 触发 此事件。                      |
| WindowSize              | 任一 文档 窗口调整大小时将 触发 此事件。                    |
| XMLSelectionChange      | 当前所选内容的 XML 父节点更改时将 触发 此事件。             |
| XMLValidationError      | 文档中存在验证错误时 时，将 触发 此事件。                   |

## 事件

### DocumentBef oreClose

当文档关闭前时触发此事件。

**语法**

`function Application_DocumentBef oreClose ( Doc, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型     | 说明                                        |
|--------|-----------|--------------|---------------------------------------------|
| Doc    | 必选      | Document对象 | 关闭的文档对象                              |
| Cancel | 必选      | Object       | 如果设置其属性Value为 true ，则不关闭文稿。 |

**示例**

当关闭一个 文档 时，弹消息框提醒用户即将关闭 文档 ， 询问 是否取消。用户点击“是”就会取消关闭 文档 。

``` JavaScript
function Application_DocumentBeforeClose(Doc, Cancel)
{
    var ret = MsgBox("文档\"" + Doc.Name + "\"" +"正在关闭。是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### DocumentBeforePrint

当文档打印之前时 触发 此事件。

**语法**

`function Application_DocumentBeforePrint ( Doc, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型     | 说明                                          |
|--------|-----------|--------------|-----------------------------------------------|
| Doc    | 必选      | Document对象 | 关闭的文档对象                                |
| Cancel | 必选      | Object       | 如果设置其属性Value为 true ，则不 打印文档 。 |

**示例**

当打印一个 文档 时，弹消息框提醒用户即将打印 文档 ， 询问 是否取消。用户点击“是”就会取消打印 文档。

``` JavaScript
function Application_DocumentBeforePrint(Doc, Cancel)
{
var ret = MsgBox("即将对文档\"" + Doc.Name + "\"" +"进行打印。是否取消？", jsYesNo);
if (ret == jsResultYes)
    Cancel.Value = true;
}
```

### DocumentChange

当创建新文档、打开已有文档或者切换激活文档 时 触发 此事件。

**语法**

`function Application_DocumentChange () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**示例**

当触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_DocumentChange()
{
    MsgBox("文档切换了！");
}
```

### DocumentOpen

当任一文档打开时 触发 此事件。

**语法**

`function Application_DocumentOpen ( Doc ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明           |
|------|-----------|--------------|----------------|
| Doc  | 必选      | Document对象 | 关闭的文档对象 |

**示例**

当触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_DocumentOpen(Doc)
{
    MsgBox("打开了文档:" + Doc.Name);
}
```

### NewDocument

当创建新文档 时 触发 此事件 。

**语法**

`function Application_NewDocument ( Doc ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明           |
|------|-----------|--------------|----------------|
| Doc  | 必选      | Document对象 | 关闭的文档对象 |

**示例**

当触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_NewDocument(Doc)
{
    MsgBox("新建了文档:" + Doc.Name);
}
```

### Quit

当关闭退出文字组件进程时触发此事件。

**语法**

`function Application_Quit () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

### WindowActivate

当激活文档窗口 时 触发 此事件。

**语法**

`function Application_WindowActivate ( Doc, Wn ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明            |
|------|-----------|--------------|-----------------|
| Doc  | 必选      | Document对象 | 激活 的文档对象 |
| Wn   | 必选      | Window 对象  | 激活 的窗口对象 |

**示例**

当切换文档窗口时，输出激活的窗口标题。

``` JavaScript
function Application_WindowActivate(Doc,  Wn)
{
    Debug.Print("当前激活窗口是："+ Wn.Caption)
}
```

### WindowBeforeDoubleClick

当双击任一文档窗口之前 触发 此事件。

**语法**

`function Application_WindowBeforeDoubleClick ( Sel , Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sel    | 必选      | Selection | 当前所选内容。                                |
| Cancel | 必选      | Object    | 如果设置其属性Value为 true ，则取消此次双击。 |

**示例**

当在一个窗口内双击 时，弹消息框提醒用户 ， 询问 是否取消。用户点击“是”就会取消该操作 。

``` JavaScript
function Application_WindowBeforeDoubleClick(Sel, Cancel)
{
    var ret = MsgBox("您进行了双击，当前选中内容是：" + Sel +"是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### WindowBeforeRightClick

当右击任一 文档窗口之前 触发 此事件 。

**语法**

`function Application_WindowBeforeRightClick ( Sel , Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sel    | 必选      | Selection | 当前所选内容。                                |
| Cancel | 必选      | Object    | 如果设置其属性Value为 true ，则取消此次右击。 |

**示例**

当在一个窗口内右击时，弹消息框提醒用户，询问是否取消。用户点击“是”就会取消该操作 。

``` JavaScript
function Application_WindowBeforeRightClick(Sel, Cancel)
{
    var ret = MsgBox("您进行了右击，当前选中内容是：" + Sel +"是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### WindowDeactivate

任一 文档 窗口被切换到非激活状态时触发此事件。

**语法**

`function Application_WindowDeactivate ( Doc, Wn ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明                          |
|------|-----------|--------------|-------------------------------|
| Doc  | 必选      | Document对象 | 被切换到非激活状态 的文档对象 |
| Wn   | 必选      | Window 对象  | 被切换到非激活状态 的窗口对象 |

**示例**

当切换文档窗口时，输出 被切换到非激活状态 的窗口 标题。

``` JavaScript
function Application_WindowDeactivate(Doc,  Wn)
{
    Debug.Print("当前被切换到非激活窗口是："+ Wn.Caption)
}
```

### WindowSelectionChange

活动窗口所选内容更改时将 触发 此事件。

**语法**

`function Application_WindowSelectionChange ( Sel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型      | 说明           |
|------|-----------|---------------|----------------|
| Sel  | 必选      | Selection对象 | 当前所选内容。 |

**示例**

当 活动窗口所选内容更改时 触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_WindowSelectionChange(Sel)
{
    MsgBox("当前选中内容是：" + Sel);
}
```

### WindowSize

任一文档窗口调整大小时将 触发 此事件。

**语法**

`function Application_WindowSize ( Doc, Wn ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明                    |
|------|-----------|--------------|-------------------------|
| Doc  | 必选      | Document对象 | 调整窗口大小 的文档对象 |
| Wn   | 必选      | Window 对象  | 调整窗口大小 的窗口对象 |

**示例**

当触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_WindowSize(Doc, Wn)
{
    MsgBox("窗口\""+ Wn.Caption + "\"大小发生了改变。")
}
```

### XMLSelectionChange

当前所选内容的 XML 父节点更改时将 触发 此事件。

**语法**

`function Application_XMLSelectionChange ( Sel, OldXMLNode, NewXMLNode, Reason ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称       | 必选/可选 | 数据类型       | 说明                        |
|------------|-----------|----------------|-----------------------------|
| Sel        | 必选      | Selection 对象 | 当前所选内容。              |
| OldXMLNode | 必选      | XMLNode 对象   | 插入区域移动前的XML 节点。  |
| NewXMLNode | 必选      | XMLNode 对象   | 插入区域移动后的 XML 节点。 |
| Reason     | 必选      | long           | 更改的操作                  |

**示例**

以下示例在将新元素插入到文档中时，对新添加的 XML 元素进行验证。

``` JavaScript
function Application_XMLSelectionChange(Sel, OldXMLNode, NewXMLNode, Reason)
{
    if (Reason == wdXMLSelectionChangeReasonInsert)
    {
        if (OldXMLNode !== undefined && OldXMLNode !== null)
            NewXMLNode.Validate()
    }
}
```

### XMLValidationError

文档中存在验证错误时 时，将触发此事件。

**语法**

`function Application_XMLValidationError ( XMLNode ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型     | 说明           |
|---------|-----------|--------------|----------------|
| XMLNode | 必选      | XMLNode 对象 | 无效的 XML节点 |

**示例**

当触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_XMLValidationError(XMLNode)
{
    MsgBox("XML 节点"  + XMLNode.BaseName +" 无效:" + XMLNode.ValidationErrorText)
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/文字事件/Application 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%96%87%E5%AD%97%E4%BA%8B%E4%BB%B6/Application%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
