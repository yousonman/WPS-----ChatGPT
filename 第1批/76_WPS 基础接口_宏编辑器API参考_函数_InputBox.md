# InputBox

弹出 显示自定义提示信息 的输入对话框，等待用户在输入框中输入文本或单击按钮，然后返回输入框的内容。

## 语法

**InputBox** ( prompt, \[ *title* \], \[ *default* \], \[ *xpos* \], \[ *ypos* \] )

## 参数

| 名字      | 说明                                         |
|-----------|-------------------------------------------------------------------|
| *prompt*    | 必需项。 在对话框中显示的 消息的 字符串表达式 。                              |
| *title*   | 可选。 对话框标题栏中显示的字符串表达式。 如果省略 *title* ，则标题栏中将显示应用程序名称。                            |
| *default* | 可选。 文本框中显示的字符串表达式，在未提供其他输入时作为默认响应。 如果省略了 *default* ，文本框将显示为空。          |
| *xpos*    | 可选。 指定对话框的左边缘与屏幕的左边缘的水平距离（以缇为单位）的数值表达式。 如果省略了 *xpos* ，对话框将水平居中。   |
| *ypos*    | 可选。 指定对话框的上边缘与屏幕的顶部的垂直距离（以缇为单位）的数值表达式。 如果省略了 *ypos* ，对话框将 将垂直居中 。 |

## 返回值

如果用户选择"**确定**" 或按 **Enter**，**InputBox** 函数将返回文本框中的内容。 如果用户选择"**取消**"，函数将返回一个空字符串 。

## 示例

下面的代码是就是 **InputBox** 的一个示例。

``` JavaScript
function abcd()
{
    var strMessage = "这是一个输入对话框，请输入？";
    var strTitle = "我是标题";
    var strDefault = "我是标题";
    var ret = InputBox(strMessage, strTitle, strDefault, 3000, 4000)
    Debug.Print("输入对话框返回值是：" + ret)
}
```

![](Base64图像/Base64图像45来自_WPS%20基础接口_宏编辑器API参考_函数_InputBox.png)

![](Base64图像/Base64图像46来自_WPS%20基础接口_宏编辑器API参考_函数_InputBox.png)

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/函数/InputBox](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E5%87%BD%E6%95%B0/InputBox.html)

------------------------------------------------------------------------
