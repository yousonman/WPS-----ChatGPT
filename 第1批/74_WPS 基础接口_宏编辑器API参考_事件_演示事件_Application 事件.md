# Application 事件

## 事件列表

| 名称                    | 触发时机                                     |
|-------------------------|----------------------------------------------|
| NewPresentation         | 当新建演示文稿时触发此事件。                 |
| PresentationBeforeClose | 任一打开的 演示文稿 关闭之前触发此事件。     |
| PresentationBeforeSave  | 任一 演示文稿 被保存 之前 触发 此事件。      |
| PresentationOpen        | 当打开任一演示文稿时 触发 此事件。           |
| PresentationSave        | 任一 演示文稿被 保存 时 触发 此事件。        |
| WindowActivate          | 任一 演示文稿 窗口被激活时，将 触发 此事件。 |

## 事件

### NewPresentation

当新建演示文稿时触发此事件 。

**语法**

`function Application_NewPresentation ( Pres ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型         | 说明                   |
|------|-----------|------------------|------------------------|
| Pres | 必选      | Presentation对象 | 新建的 演示文稿 对象。 |

**示例**

当新建一个 演示文稿 时，弹消息框提醒用户新建了 演示文稿 。

``` JavaScript
function Application_NewPresentation(Pres)
{
    MsgBox("您新建了演示文稿："+Pres.Name)
}
```

### PresentationBeforeClose

任一打开的 演示文稿 关闭之前触发此事件。

**语法**

`function Application_PresentationBeforeClose ( Pres, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型         | 说明                                            |
|--------|-----------|------------------|-------------------------------------------------|
| Pres   | 必选      | Presentation对象 | 关闭的 演示文稿 对象。                          |
| Cancel | 必选      | Object           | 如果设置其属性Value为 true ，则不关闭演示文稿。 |

**示例**

当关闭一个演示文稿时，弹消息框提醒用户即将关闭演示文稿，是否取消。用户点击“是”就会取消关闭 演示文稿。

``` JavaScript
function Application_PresentationBeforeClose(pres, Cancel)
{
    var ret = MsgBox("演示文稿\"" + pres.Name + "\"" +"正在关闭。是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### PresentationBeforeSave

任一 演示文稿被保存 之前 触发 此事件。

**语法**

`function Application_PresentationBeforeSave ( Pres, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型         | 说明                                            |
|--------|-----------|------------------|-------------------------------------------------|
| Pres   | 必选      | Presentation对象 | 保存的 演示文稿 对象。                          |
| Cancel | 必选      | Object           | 如果设置其属性Value为 true ，则不保存演示文稿。 |

**示例**

当保存一个演示文稿时，弹消息框提醒用户即将保存演示文稿，是否取消。用户点击“是”就会取消保存 演示文稿。

``` JavaScript
function Application_PresentationBeforeSave(pres, Cancel)
{
    var ret = MsgBox("演示文稿\"" + pres.Name + "\"" +"即将保存。是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### PresentationOpen

当打开任一演示文稿时 触发 此事件。

**语法**

`function Application_PresentationOpen ( Pres ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型         | 说明                   |
|------|-----------|------------------|------------------------|
| Pres | 必选      | Presentation对象 | 新建的 演示文稿 对象。 |

**示例**

当 打开任一演示文稿 时，弹消息框提醒用户打开了演示文稿。

``` JavaScript
function Application_PresentationOpen(Pres)
{
    MsgBox("您打开了演示文稿。")
}
```

### PresentationSave

任一 演示文稿被保存 时 触发 此事件。

**语法**

`function Application_PresentationSave ( Pres ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型         | 说明                   |
|------|-----------|------------------|------------------------|
| Pres | 必选      | Presentation对象 | 保存的 演示文稿 对象。 |

**示例**

当保存演示文稿时，弹消息框提醒用户 正在保存演示文稿 。

``` JavaScript
function Application_PresentationSave(pre)
{
    MsgBox("正在保存演示文稿："+ pre.Name)
}
```

### WindowActivate

任一 演示文稿 窗口被激活时，将 触发 此事件。

**语法**

`function Application_WindowActivate ( Pres, Wn ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型           | 说明                   |
|------|-----------|--------------------|------------------------|
| Pres | 必选      | Presentation对象   | 新建的 演示文稿 对象。 |
| Wn   | 必选      | DocumentWindow对象 | 激活的文档窗口对象。   |

**示例**

当切换演示文稿窗口时，输出激活的窗口标题。

``` JavaScript
function Application_WindowActivate(pre, win)
{
    Debug.Print("当前激活窗口是："+ win.Caption)
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/演示事件/Application 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%BC%94%E7%A4%BA%E4%BA%8B%E4%BB%B6/Application%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
