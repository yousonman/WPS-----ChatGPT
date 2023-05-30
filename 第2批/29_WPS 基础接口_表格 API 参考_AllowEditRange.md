# AllowEditRange 对象

## 简介

代表受保护的工作表上可进行编辑的单元格。

## 说明

使用 Add 方法或 AllowEditRanges 集合的 Item 属性可返回 AllowEditRange 对象。 返回 AllowEditRange 对象后，可使用 ChangePa ssword 方法更改访问受保护工作表上可编辑区域的密码。

此示例中，ET 允许用户在活动工作表上编辑单元格区域 A1:A4，并通知用户，然后更改此指定区域的密码并将所做的更改通知用户。

``` JavaScript
function test() {
    let wksOne
    let wksPassword
    wksOne = Application.ActiveSheet
　　wksPassword = prompt("Enter password for the worksheet")
    // Establish a range that can allow edits on the protected worksheet.
    wksOne.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), "123")

    alert("Cells A1 to A4 can be edited on the protected worksheet.")

    // Change the password.　　wksPassword = prompt("Enter the new password for the worksheet")
    wksOne.Protection.AllowEditRanges.Item(1).ChangePassword("456")

    alert("The password for these cells has been changed.")
} 
```

## 方法

| 序号 | 名称                                             | 说明                                                                         |
|:----:|:-------------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [ChangePassword](#AllowEditRange.ChangePassword) | 更改受保护的工作表中可以进行编辑的区域的密码。                               |
|  2   | [Delete](#AllowEditRange.Delete)                 | 删除对象。                                                                   |
|  3   | [Unprotect](#AllowEditRange.Unprotect)           | 取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。 |

## 属性

| 序号 | 名称                           | 说明                                                                   |
|:----:|:-------------------------------|:-----------------------------------------------------------------------|
|  1   | [Range](#AllowEditRange.Range) | 返回一个 Range 代表，它代表在受保护工作表上可编辑的区域的子集。        |
|  2   | [Title](#AllowEditRange.Title) | 返回或设置受保护的工作上可编辑的单元格区域的标题。 String 型，可读写。 |
|  3   | [Users](#AllowEditRange.Users) | 返回工作表上受保护区域的一个 UserAccessList 对象。                     |

## 成员方法

### AllowEditRange.ChangePassword

更改受保护的工作表中可以进行编辑的区域的密码。

**语法**

**express.ChangePassword ( *Password* )**

\*express是一个代表 AllowEditRange 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明     |
|------------|-----------|----------|----------|
| *Password* | 必选      | String   | 新密码。 |

**示例**

``` JavaScript
function test() {
    let wksOne
    let wksPassword
    wksOne = Application.ActiveSheet

    // Establish a range that can allow edits on the protected worksheet.
    wksOne.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), "123")

    alert("Cells A1 to A4 can be edited on the protected worksheet.")

    // Change the password.
    wksOne.Protection.AllowEditRanges.Item(1).ChangePassword("456")

    alert("The password for these cells has been changed.")
}
```

### AllowEditRange.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 AllowEditRange 对象的变量。

### AllowEditRange.Unprotect

取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。

**语法**

**express.Unprotect ( *Password* )**

\*express是一个代表 AllowEditRange 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                             |
|------------|-----------|----------|------------------------------------------------------------------------------------------------------------------|
| *Password* | 可选      | Variant  | 一个字符串，它指定用于解除单元格区域保护的密码，此密码是区分大小写的。如果单元格区域不设密码保护，则忽略此参数。 |

**说明**

***如果您忘记了密码，将不能取消工作表或工作簿的保护。建议将密码和对应文档名妥善保存。***

## 成员属性

### AllowEditRange.Range

返回一个 Range 代表，它代表在受保护工作表上可编辑的区域的子集。

**语法**

**express.Range**

\*express是一个代表 AllowEditRange 对象的变量。

### AllowEditRange.Title

返回或设置受保护的工作上可编辑的单元格区域的标题。 **String** 型，可读写。

**语法**

**express.Title**

\*express是一个代表 AllowEditRange 对象的变量。

### AllowEditRange.Users

返回工作表上受保护区域的一个 **UserAccessList** 对象。

**语法**

**express.Users**

\*express是一个代表 AllowEditRange 对象的变量。

**示例**

``` JavaScript
function test() {
    let wksSheet = Application.ActiveSheet

    // Display name of user with access to protected range.
    alert(wksSheet.Protection.AllowEditRanges.Item(1).Users.Item(1).Name)
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/AllowEditRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
