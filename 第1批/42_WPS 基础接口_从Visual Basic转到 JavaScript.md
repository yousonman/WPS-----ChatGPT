# 从Visual Basic转到 JavaScript

## JSAPI接口的差异

### 方法的差异

vb的方法可以不加括号，但jsapi中所有的方法都需要加括号，如果方法不加括号会被js语法判定为属性。

``` vb
Application.Workbooks(1).Close
```

``` JavaScript
Application.Workbooks.Item(1).Close()
```

vb的方法支持给部分参数赋值。但js对缺省的参数需要用undefined占位补齐，如下面例子中为Find方法的第一和第三个参数赋值，在js中Find方法第二个参数需要用undefined占位补齐

``` vb
Range("A1:N29").Find("香港特别", LookIn:=xlValues).Select
```

``` JavaScript
Application.ActiveSheet.Range("A1:N29").Find("我",undefined, xlValues).Select();
```

vb可通过数组方式取集合中的对象，jsapi必须通过Item方法获取集合中的对象，其中要注意二维数组取值时需要将value要改成Value2

``` vb
Application.Workbooks(1).Close
```

``` JavaScript
Application.Workbooks.Item(1).Close()
```

以下是二维数组的取值方法

``` vb
cells(2,3).value
```

``` JavaScript
Application.Cells.Item(2,3).Value2
```

### 属性的差异

vb中调用书写错误的属性会报错，js不会报错，这是一个bug，所以特别注意。

``` JavaScript
Application.ActiveDocument.Name111 // 不报错
```

vb支持thisdocument对象，jsapi暂时不支持该对象，如果遇到thisdocument对象，可以用ActiveDocument代替

``` vb
MsgBox (Application.ThisDocument.Name)
```

``` JavaScript
MsgBox (Application.ActiveDocument.Name)
```

### 事件的差异

jside的导航栏展示和自动补全的事件比较少，但js支持的事件基本和vb一致，下面给出一个vb事件转换为jside中事件的例子

``` vb
Private Sub Workbook_SheetSelectionChange(ByVal Target As Range)
Dim i As Integer
IF Target.Column>6 And Target.Column>37 && Target.Row>6 And Target.Row Mod 2 = 1 And Cells(Target.Row -1, 5).Value <> “”
Then
    With Target.Validation
        .Delete
        .Add Tpye:=xlValidateList,AlertStyle:=xlValidAlertStop,Operator:=xlBetween,Formular1:="=区域城市！$A2:$A30000")
        .IncellDropdown = true
```

``` JavaScript
function Workbook_SheetSelectionChange(Sh, Target) {
    if(Target.Column>6 && Target.Column>37 && Target.Row>6 && (Target.Row)%2n==1 && Cells.Item(Target.Row -1n, 5).Value2 !==””) {
        var temptarget = Target.Validation
        temptarget.Delete()
        temptarget.Add(xlValidateList,xlValidAlertStop,xlBetween,"=区域城市！$A2:$A30000")
        temptarget.IncellDropdown = true
    }
}
```

vb支持thisdocument对象，jsapi暂时不支持该对象，如果遇到thisdocument对象，可以用ActiveDocument代替

``` vb
MsgBox (Application.ThisDocument.Name)
```

``` JavaScript
MsgBox (Application.ActiveDocument.Name)
```

## 函数定义

vb用sub ... end sub关键字定义函数，js用funtion{}关键字定义函数

``` vb
Public(Private) Sub AddTemplate()
    MsgBox "d"
End Sub
```

``` JavaScript
function AddTemplate() {
    MsgBox（"d"）
}
```

## 数据类型

vb在数据定义时需要指明类型，但是js是动态类型，赋值后才有类型,js包括的基本类型：字符串（String）、数字(Number)、布尔(Boolean)、对空（Null）、未定义（Undefined）、Symbol，声明这些类型都用关键字var

``` vb
Dim i As Integer
```

``` JavaScript
var i
```

``` vb
Dim j As String
```

``` JavaScript
var j
```

**基本类型对比**

| js类型  | vb类型                                  |
|---------|-----------------------------------------|
| boolean | Boolean                                 |
| number  | Integer，Long，LongLong，Single，Double |
| string  | String                                  |

vb的bool类型首字母大写，js中都是小写

``` vb
True, False
```

``` JavaScript
true, false
```

**字符串**

js字符串使用反斜杠\\转义字符把特殊字符转换为字符串字符:

| 代码 | 结果 | 描述   |
|------|------|--------|
| \\   | '    | 单引号 |
| \\   | "    | 双引号 |
| \\   | \\   | 反斜杠 |

因此，在常见的文件路径的场景，需要采用“/”或者“\\”(Windows)来作为路径分隔符：

``` JavaScript
function test()
{
    Workbooks.Open("H:\\新建文件夹\\工作簿1.xlsx")
}
```

## 运算符

算术运算符的差异，如vb中的mod关键字应改为%

``` vb
Target.Row Mod 2 = 1
```

``` JavaScript
(Target.Row)%2==1
```

“+”可以用来拼接字符串

``` vb
AddIns.Add FileName:="C:\Program Files\Microsoft Office" _"\Templates\Letters Faxes\MyFax.dot", Install:=True
```

``` JavaScript
AddIns.Add("C:/Program Files/Microsoft Office" + "/Templates/Letters Faxes/MyFax.dot",true)
```

逻辑运算符

``` vb
1.逻辑与（AND）2.逻辑或（OR） 3.逻辑非（<>）
```

``` JavaScript
逻辑与（&&）2.逻辑或（||） 3.逻辑非（！）
```

注意js和vb判断空的逻辑不一样，vb内部会智能一点，遇到判断空的语句，尽量用"!"

``` vb
if Cells(1,2).value == ""  //空
```

``` JavaScript
if (Cells.Item(1,2).Value2)   //空
```

``` vb
if Cells(1,2).value <> ""  //非空
```

``` JavaScript
if (!Cells.Item(1,2).Value2)  //非空
```

## 枚举

vb和js枚举的使用上存在的差异如下，在vb中wdAlignParagraphLeft和WdParagraphAlignment.wdAlignParagraphLeft语法都正确，但在js中只有wdAlignParagraphLeft是正确用法

``` vb
//vb用以上两种方式的枚举都可以正常使用
Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
Application.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
```

``` JavaScript
//js只支持这一种
Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
```

## With关键字

vb有with关键字，js没有该关键字，转换时可以将with后面的对象赋值给一个变量，例子如下

``` vb
With Selection.Range.PageSetup.TextColumns
    .SetCount NumColumns:=2
    .EvenlySpaced = 0
End With
```

``` JavaScript
Var columens = Selection.Range.PageSetup.TextColumns
columens.SetCount(2)
columens.EvenlySpaced = 0
//注意TextColumns后面没有括号因为这是一个属性
```

## 条件语句

VB中的常用条件语句{IF 条件 Then 表达式 Endif}，在JS中应该转换为 if(条件) {表达式}

``` vb
IF Target.Column>6 And Target.Column>37 And Target.Row>6 And Target.Row Mod 2 = 1 And Cells(Target.Row -1, 5).Value <> ""
Then
    With Target.Validation
        .Delete
        .Add
Endif
```

``` JavaScript
if(Target.Column > 6 && Target.Column > 37 && Target.Row > 6 && (Target.Row) % 2 == 1 && Cells.Item(Target.Row -1n, 5).Value2 !== null) {
    var temptarget = Target.Validation
    temptarget.Delete()
    temptarget.Add(xlValidateList,xlValidAlertStop,xlBetween,"=区域城市！$A2:$A30000")
}
```

## 路径

VB中的常用条件语句{IF 条件 Then 表达式 Endif}，在JS中应该转换为 if(条件) {表达式}

``` vb
IF Target.Column>6 And Target.Column>37 And Target.Row>6 And Target.Row Mod 2 = 1 And Cells(Target.Row -1, 5).Value <> ""
Then
    With Target.Validation
        .Delete
        .Add
Endif
```

``` JavaScript
if(Target.Column > 6 && Target.Column > 37 && Target.Row > 6 && (Target.Row) % 2 == 1 && Cells.Item(Target.Row -1n, 5).Value2 !== null) {
    var temptarget = Target.Validation
    temptarget.Delete()
    temptarget.Add(xlValidateList,xlValidAlertStop,xlBetween,"=区域城市！$A2:$A30000")
}
```

> WPS 开发文档： [WPS 基础接口/从Visual Basic转到 JavaScrip](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E4%BB%8EVisual%20Basic%E8%BD%AC%E5%88%B0%20JavaScript.html)

------------------------------------------------------------------------
