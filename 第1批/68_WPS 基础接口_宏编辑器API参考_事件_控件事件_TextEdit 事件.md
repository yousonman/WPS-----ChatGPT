# TextEdit 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Change    | 当控件的内容发生变化时触发此事件                                                  |
| Enter     | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit      | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### DblClick

**语法**

`function UserForm1_TextEdit1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_TextEdit1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_TextEdit1_MouseDown(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_TextEdit1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_TextEdit1_MouseUp(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_TextEdit1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_TextEdit1_MouseMove(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_TextEdit1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_TextEdit1_KeyUp(keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_TextEdit1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}               
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_TextEdit1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_TextEdit1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm 1_TextEdit1_KeyPress(keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_TextEdit1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Change

当控件的内容发生变化时触发此事件

**语法**

`function UserForm1_TextEdit1_Change () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_TextEdit1_Change()
{
    MsgBox("控件内容发生改变")
} 
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm 1_TextEdit1_Enter() { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_TextEdit1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm 1_TextEdi t1_Exit(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_TextEdit1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/TextEdit 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/TextEdit%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
