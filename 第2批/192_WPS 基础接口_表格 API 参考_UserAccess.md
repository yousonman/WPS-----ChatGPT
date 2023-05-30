# UserAccess 对象

## 简介

代表对受保护区域的用户访问。

## 说明

使用 UserAccessList 集合的 Add 方法或 Item 属性可返回一个 UserAccess 对象。

一旦返回了 UserAccess 对象，您就可以使用 AllowEdit 属性来确定是否允许访问工作表中某个特定区域。下例添加一个在受保护的工作表上可编辑的区域，并通知用户该区域的标题。

``` JavaScript
function test(){
    let wksSheet = Application.ActiveSheet
    
    //Add a range that can be edited on the protected worksheet.
    wksSheet.Protection.AllowEditRanges.Add("Test", Range("A1"))
    
    //Notify the user the title of the range that can be edited.
    MsgBox(wksSheet.Protection.AllowEditRanges.Item(1).Title)
}
```

## 方法

| 序号 | 名称                         | 说明       |
|:----:|:-----------------------------|:-----------|
|  1   | [Delete](#UserAccess.Delete) | 删除对象。 |

## 属性

| 序号 | 名称                               | 说明                                                                      |
|:----:|:-----------------------------------|:--------------------------------------------------------------------------|
|  1   | [AllowEdit](#UserAccess.AllowEdit) | 返回或设置一个 Boolean 值，它指明是否允许用户访问受保护工作表的指定区域。 |
|  2   | [Name](#UserAccess.Name)           | 返回或设置一个 String 值，它代表对象的名称。                              |

## 成员方法

### UserAccess.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 UserAccess 对象的变量。

## 成员属性

### UserAccess.AllowEdit

返回或设置一个 **Boolean** 值，它指明是否允许用户访问受保护工作表的指定区域。

**语法**

**express.AllowEdit**

\*express是一个代表 UserAccess 对象的变量。

### UserAccess.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 UserAccess 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/UserAccess](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
