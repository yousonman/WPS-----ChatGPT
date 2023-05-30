# Workbook 事件

## 事件列表

| 名称                   | 触发时机      |
|------------------------|------------------------------------------------------------------------|
| Activate               | 该工 作簿被激活时，将触发此事件。                                      |
| AfterSave              | 该工作簿被保存之后触发此事件。                                         |
| BeforeClose            | 该工作簿关闭之前触发此事件。                                           |
| BeforePrint            | 该工作簿打印之前触发此事件。                                           |
| BeforeSave             | 该工作簿保存之前触发此事件。                                           |
| Deactivate             | 该 工 作簿被切换到非激活状态时触发此事件。                             |
| NewSheet               | 该 工作簿中创建新工作表时 触发此事件。                                 |
| Open                   | 该 工作簿打开时 触发 此事件。                                          |
| SheetActivate          | 当激活 该 工作簿 任一工作表时 触发 此事件。                            |
| SheetBeforeDelete      | 删除 该 工作簿 任 一 工作表之前 触发 此事件。                          |
| SheetBeforeDoubleClick | 双击 该 工作簿 任一工作表之前 触发 此事件。                            |
| SheetBeforeRightClick  | 右击 该 工作簿 任一 工作表之前 触发 此事件 。                          |
| SheetCalculate         | 在 该 工作簿 任一工作表进行 计算时 触发此事件。                        |
| SheetChange            | 当用户或外部链接更改了 该 工作簿 任一 工作表中的单元格时 触发 此事件。 |
| SheetDeactivate        | 当 该 工作簿 任 一 工作表被切换到非激活状态时 触发 此事件。            |
| SheetFollowHyperlink   | 单击 该 工作簿 任一 工作表的超链接 时 触 发 此事 件。                  |
| SheetSelectionChange   | 该 工作簿 任一工作表上的选定区域发生更改时，将触发此事件。             |

## 事件

### Activate

该工 作簿被激活时，将触发此事件。

**语法**

`function Workbook_Activate () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**示例**

该工 作簿被激活时 ，弹消息框提醒用户。

``` JavaScript
function Workbook_Activate()
{
    MsgBox("您激活了当前工作簿")
} 
```

### AfterSave

该工作簿被保存之后触发此事件。

**语法**

`function Workbook_AfterSave ( Success ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                      |
|---------|-----------|----------|-------------------------------------------|
| Success | 必选      | Boolean  | 如果保存操作成功，则为true; 否则为false。 |

**示例**

该工 作簿被保存后， 弹消息框提醒用户。

``` JavaScript
function Workbook_AfterSave(Success)
{
    if (Success)
    {
        MsgBox("文档被成功保存！")
    }
} 
```

### BeforeClose

该工作簿关闭之前触发此事件。

**语法**

`function Workbook_BeforeClose ( Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则取消此次关闭。 |

**示例**

该工作簿被关闭时，弹消息框询问用户是否取消，用户点击“是”就会取消关闭 工作簿 。

``` JavaScript
function Workbook_BeforeClose(Cancel)
{
    var ret = MsgBox("工作簿正在关闭。是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### BeforePrint

该工作簿打印之前触发此事件

**语法**

`function Workbook_BeforePrint ( Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则取消此次打印。 |

**示例**

该工作簿打印前 时，弹消息框询问 用 户是否取消， 用户点击“是”就会取消打印 工作簿 。

``` JavaScript
function Workbook_BeforePrint(Cancel)
        {
        var ret = MsgBox("工作簿即将打印。是否取消？", jsYesNo);
        if (ret == jsResultYes) Cancel.Value = true; }
        
        
```

### BeforeSave

该工作簿保存之前触发此事件。

**语法**

`function Workbook_BeforeSave ( S aveAsUI ， Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明|
|-----------|-----------|----------|--------------------------------------------------------------|
| S aveAsUI | 必选      | Boolean  | 如果为 true 则表示此次保存操作将会弹出保存或是另存为对话框。 |
| Cancel    | 必选      | Object   | 如果设置其属性Value为 true ，则取消此次保存。                |

**示例**

对该工作簿 保存之前 ， 如果 会弹出 保存或是另存为 对话框 的情况， 弹消息框询问用户是否取消，用户点击“是”就会取消保存工作簿。

``` JavaScript
function Workbook_BeforeSave(SaveAsUI, Cancel)
{
    if (SaveAsUI){
        var ret = MsgBox("工作簿即将保存。是否取消？", jsYesNo);
        if (ret == jsResultYes)
            Cancel.Value = true;
    }
} 
```

### Deactivate

该 工 作簿被切换到非激活 （非活动 ） 状态时触发此事件。

**语法**

`function Workbook_Deactivate () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**示例**

该工作簿被 切换到非激活（非活动）状态时 ，弹消息框提醒用户。

``` JavaScript
function Workbook_Deactivate()
{
    MsgBox("工作簿被切换到非激活了。");
} 
```

### NewSheet

该 工作簿中创建新工作表时 触发此事件。

**语法**

`function Workbook_NewSheet ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明         |
|------|-----------|----------|--------------|
| Sh   | 必选      | Object   | 新建的工作表 |

**示例**

该工作簿 创建新工作表 时，弹消息框提醒用户。

``` JavaScript
function Workbook_NewSheet(Sh)
{
    MsgBox("新建了工作表：" + Sh.Name)
} 
```

### Open

该 工作簿打开时 触发 此事件

**语法**

`function Workbook_Open () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**示例**

该工作簿被打开时，弹消息框提醒用户。

``` JavaScript
function Workbook_Open()
{
    MsgBox("工作簿被打开了。");
}
```

### SheetActivate

当激活 该 工作簿 任一工作表时 触发 此事件。

**语法**

`function Workbook_SheetActivate ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Sh   | 必选      | Object   | 激活的工作表对象。 |

**示例**

当 激活 了 工作表 时，弹消息框提醒用户。

``` JavaScript
function Application_SheetActivate(Sh)
{
    MsgBox("您激活了工作表："+Sh.Name)
}
```

### SheetBeforeDelete

删除该工作簿任 一 工作表之前 触发 此事件。

**语法**

`function Workbook_SheetBeforeDelete ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Sh   | 必选      | Object   | 删除的工作表对象。 |

**示例**

删除该工作簿任 一 工作表之前 ，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetBeforeDelete(Sh)
{
    MsgBox("您删除了工作表："+Sh.Name)
}
```

### SheetBeforeDoubleClick

双击该工作簿任一工作表之前 触发 此事件。

**语法**

`function Workbook_SheetBeforeDoubleClick ( Sh, Target, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sh     | 必选      | Object    | 双击的工作表对象。                            |
| Target | 必选      | Range对象 | 双击区域所在的单元格对象。                    |
| Cancel | 必选      | O bject   | 如果设置其属性Value为 true ，则取消此次双击。 |

**示例**

双击该工作簿任一工作表之前 ，如果双击操作在A1单元格上，则不响应双击操作（默认是进入编辑），否则则继续响应。

``` JavaScript
function Workbook_SheetBeforeDoubleClick(Sh, Target, Cancel)
{
    if (Target.Address() == "$A$1")
        Cancel.Value = true;
} 
```

### SheetBeforeRightClick

右击该工作簿任一 工作表之前 触发 此事件 。

**语法**

`function Workbook_SheetBeforeRightClick ( Sh, Target, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                            |
|--------|-----------|-----------|-------------------------------------------------|
| Sh     | 必选      | Object    | 右击 的工作表对象。                             |
| Target | 必选      | Range对象 | 右击 区域所在的单元格对象。                     |
| Cancel | 必选      | Object    | 如果设置其属性Value为 true ，则取消此次 右击 。 |

**示例**

右击 该工作簿任一工作表之前 ，如果右击操作在A1单元格上，则不响应 右击 操作（默认是弹出右键菜单），否则则继续 右击 。

``` JavaScript
function Workbook_SheetBeforeRightClick(Sh, Target, Cancel)
        {
        if (Target.Address() == "$A$1") Cancel.Value = true; } 
        
```

### SheetCalculate

该工作簿任一工作表进行 计算时 触发此事件。

**语法**

`function Workbook_SheetCalculate ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Sh   | 必选      | Object   | 计算的工作表对象。 |

**示例**

当 工作表进行 计算时 ，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetCalculate(Sh)
{
    MsgBox("工作表正在计算："+Sh.Name)
} 
```

### SheetChange

当用户或外部链接更改了该工作簿任一 工作表中的单元格时 触发 此事件。

**语法**

`function Workbook_SheetChange ( Sh, Target ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                  |
|--------|-----------|-----------|-----------------------|
| Sh     | 必选      | Object    | 修改的工作表对象。    |
| Target | 必 选     | Range对象 | 修改的单元格Range对象 |

**示例**

工作表中的单元格被修改 时，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetChange(Sh, Target)
{
    MsgBox("工作表:" + Sh.Name + "。区域："+Target.Address() + "。发生了修改")
}
```

### SheetDeactivate

当该工作簿任 一 工作表被切换到非激活状态时 触发 此事件。

**语法**

`function Workbook_SheetDeactivate ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                             |
|------|-----------|----------|----------------------------------|
| Sh   | 必选      | Object   | 被切换到非激活状态的工作表对象。 |

**示例**

任 一 工作表被切换到非激活状态时 ，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetDeactivate(Sh)
{
    MsgBox("被切换到非激活的工作表："+Sh.Name)
} 
```

### SheetFollowHyperlink

单击该工作簿 任一 工作表的超链接 时 触 发 此事 件。

**语法**

`function Workbook_SheetFollowHyperlink ( Sh, Target ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型      | 说明                              |
|--------|-----------|---------------|-----------------------------------|
| Sh     | 必选      | Object        | 超链接所在的工作表对象。          |
| Target | 必选      | Hyperlink对象 | Hyperlink对象，代表超链接的目标。 |

**示例**

单击该工作簿 任一 工作表的超链接 时，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetFollowHyperlink(Sh, Target)
{
    MsgBox("你点击了工作表\"" + Sh.Name + "\"上的超链接："+Target.Address)
} 
```

### SheetSelectionChange

该 工作簿 任一工作表上的选定区域发生更改时，将触发此事件。

**语法**

`function Workbook_SheetSelectionChange ( Sh, Target ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                   |
|--------|-----------|----------|------------------------|
| Sh     | 必选      | Object   | 选区改变的工作表对象。 |
| Target | 必选      | Range    | 新选定的区域。         |

**示例**

当 工作表上的选定区域发生更改时 ，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetSelectionChange(Sh, Target)
{
    MsgBox("工作表\"" + Sh.Name + "\"选中的区域是："+Target.Address() + "。")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/表格事件/Workbook 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E8%A1%A8%E6%A0%BC%E4%BA%8B%E4%BB%B6/Workbook%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
