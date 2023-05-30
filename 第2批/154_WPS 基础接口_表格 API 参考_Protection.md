# Protection 对象

## 简介

代表工作表可使用的各种保护选项类型。

## 说明

使用 Worksheet 对象的 Protection 属性可返回一个 Protection 对象。

返回一个 Protection 对象后，就可用该对象的下列属性来设置或返回保护选项。

- AllowDeletingColumns
- AllowDeletingRows
- AllowFiltering
- AllowFormattingCells
- AllowFormattingColumns
- AllowFormattingRows
- AllowInsertingColumns
- AllowInsertingHyperlinks
- AllowInsertingRows
- AllowSorting
- AllowUsingPivotTables

下例通过在最上面的行中放三个成员并保留该工作表说明了如何使用 Protection 对象的 AllowInserting.Columns 属性。

然后，此示例检查插入列的保护设置是否是 False ，如果必要，则将其设置为 True 。最后，通知用户插入一个列。

``` JavaScript
function SetProtection(){
    Range("A1").Formula = "1"
    Range("B1").Formula = "3"
    Range("C1").Formula = "4" Application.ActiveSheet.Protect()
    
    // Check the protection setting of the worksheet and act accordingly.
    if(Application.ActiveSheet.Protection.AllowInsertingColumns == false){ Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, true)
       alert("Insert a column between 1 and 3")
    }
    else{
       alert("Insert a column between 1 and 3")
    }
}
```

## 属性

| 序号 | 名称                                                             | 说明                                                                                  |
|:----:|:-----------------------------------------------------------------|:--------------------------------------------------------------------------------------|
|  1   | [AllowDeletingColumns](#Protection.AllowDeletingColumns)         | 如果允许删除受保护工作表上的列，则返回 True 。 Boolean 类型，只读。                   |
|  2   | [AllowDeletingRows](#Protection.AllowDeletingRows)               | 如果允许删除受保护工作表上的行，则返回 True 。 Boolean 类型，只读。                   |
|  3   | [AllowEditRanges](#Protection.AllowEditRanges)                   | 返回 AllowEditRanges 对象。                                                           |
|  4   | [AllowFiltering](#Protection.AllowFiltering)                     | 如果允许用户使用工作表受保护之前设置的“自动筛选”，则返回 True 。 Boolean 类型，只读。 |
|  5   | [AllowFormattingCells](#Protection.AllowFormattingCells)         | 如果允许对受保护的工作表上的单元格进行格式设置，则返回 True 。 Boolean 类型，只读。   |
|  6   | [AllowFormattingColumns](#Protection.AllowFormattingColumns)     | 如果允许对受保护的工作表上的列进行格式设置，则返回 True 。 Boolean 类型，只读。       |
|  7   | [AllowFormattingRows](#Protection.AllowFormattingRows)           | 如果允许对受保护的工作表上的行进行格式设置，则返回 True 。 Boolean 类型，只读。       |
|  8   | [AllowInsertingColumns](#Protection.AllowInsertingColumns)       | 如果允许在受保护的工作表上插入列，则返回 True 。 Boolean 类型，只读。                 |
|  9   | [AllowInsertingHyperlinks](#Protection.AllowInsertingHyperlinks) | 如果允许在受保护的工作表上插入超链接，则返回 True 。 Boolean 类型，只读。             |
|  10  | [AllowInsertingRows](#Protection.AllowInsertingRows)             | 如果允许用户在受保护的工作表上插入行，则返回 True 。 Boolean 类型，只读。             |
|  11  | [AllowSorting](#Protection.AllowSorting)                         | 如果允许在受保护的工作表上使用排序选项，则返回 True 。 Boolean 类型，只读。           |
|  12  | [AllowUsingPivotTables](#Protection.AllowUsingPivotTables)       | 如果允许用户在受保护的工作表上处理数据透视表，则返回 True 。 Boolean 类型，只读。     |

## 成员属性

### Protection.AllowDeletingColumns

如果允许删除受保护工作表上的列，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowDeletingColumns**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowDeletingColumns** 属性。

对于受保护的工作表，必须取消对包含要删除的单元格的列的锁定。

**示例**

``` JavaScript
function ProtectionOptions(){
    /*本示例取消对受保护的工作表上的列 A 的锁定，然后允许用户删除列 A 并通知用户。*/
    Application.ActiveSheet.Unprotect()
    
    // Unlock column A.
    Columns.Item("A:A").Locked = false
    
    // Allow column A to be deleted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowDeletingColumns == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("Column A can be deleted on this protected worksheet.")
}
```

### Protection.AllowDeletingRows

如果允许删除受保护工作表上的行，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowDeletingRows**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowDeletingRows** 属性。

对于受保护的工作表，必须取消对包含要删除的单元格的行的锁定。

**示例**

``` JavaScript
function ProtectionOptions(){
    /*本示例取消对受保护的工作表上的第 1 行的锁定，然后允许用户删除第 1 行并通知用户。*/
    Application.ActiveSheet.Unprotect()
    
    // Unlock row 1.
    Rows.Item("1:1").Locked = false
    
    // Allow row 1 to be deleted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowDeletingRows == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("Row 1 can be deleted on this protected worksheet.")
}
```

### Protection.AllowEditRanges

返回 **AllowEditRanges** 对象。

**语法**

**express.AllowEditRanges**

\*express是一个代表 Protection 对象的变量。

**示例**

``` JavaScript
function UseAllowEditRanges(){
    /*在本示例中，ET 允许用户编辑活动工作表上的区域 A1:A4，并将指定区域的标题和地址通知用户。*/
    let wksOne = Application.ActiveSheet

    // Unprotect worksheet.
    wksOne.Unprotect()

    // Establish a range that can allow edits on the protected worksheet.
    wksOne.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), "123")

    //Notify the user the title and address of the range.
    let WKSone = wksOne.Protection.AllowEditRanges.Item(1)

    MsgBox("Title of range: " + WKSone.Title)
    MsgBox("Address of range: " + WKSone.Range.Address)
}
```

### Protection.AllowFiltering

如果允许用户使用工作表受保护之前设置的“自动筛选”，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowFiltering**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowFiltering** 属性。

**AllowFiltering** 属性允许用户更改已有的“自动筛选”上的筛选条件。用户不能创建或删除受保护的工作表上的“自动筛选”。

若要筛选受保护的工作表上的单元格，则必须先取消对这些单元格的锁定。

**示例**

``` JavaScript
function ProtectionOptions(){
    /*本示例允许用户筛选受保护的工作表上的第 1 行并通知用户。*/
    Application.ActiveSheet.Unprotect()
        
    // Unlock row 1.
    Rows.Item("1:1").Locked = false
        
    // Allow row 1 to be filtered on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowFiltering == false){
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("Row 1 can be filtered on this protected worksheet.")
}
```

### Protection.AllowFormattingCells

如果允许对受保护的工作表上的单元格进行格式设置，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowFormattingCells**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowFormattingCells** 属性。

使用该属性可禁用“保护”选项卡，从而允许用户更改所有格式，但不能取消对区域的锁定或隐藏。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户对受保护的工作表上的单元格进行格式设置，并通知用户。*/
    Application.ActiveSheet.Unprotect()
    
    // Allow cells to be formatted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowFormattingCells == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, true)
    }
    alert("Cells can be formatted on this protected worksheet.")
}
```

### Protection.AllowFormattingColumns

如果允许对受保护的工作表上的列进行格式设置，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowFormattingColumns**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowFormattingColumns** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户对受保护的工作表上的列进行格式设置，并通知用户。*/
    Application.ActiveSheet.Unprotect()

    // Allow columns to be formatted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowFormattingColumns == false ){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, true)
    }
    alert("Columns can be formatted on this protected worksheet.")
}
```

### Protection.AllowFormattingRows

如果允许对受保护的工作表上的行进行格式设置，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowFormattingRows**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowFormattingRows** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户对受保护的工作表上的行进行格式设置，并通知用户。*/
    Application.ActiveSheet.Unprotect()

    // Allow rows to be formatted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowFormattingRows == false){
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, true)
    }
    alert("Rows can be formatted on this protected worksheet.")
}
```

### Protection.AllowInsertingColumns

如果允许在受保护的工作表上插入列，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowInsertingColumns**

\*express是一个代表 Protection 对象的变量。

**说明**

默认情况下，插入的列继承了其左边的列的格式设置，这意味着该列可能包含锁定的单元格。也就是说，用户可能不能删除插入的列。

可以使用 **Protect** 方法参数设置 **AllowInsertingColumns** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户在受保护的工作表上插入列，并通知用户。*/
    Application.ActiveSheet.Unprotect()
        
    // Allow columns to be inserted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowInsertingColumns == false){
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, true)
    }
    alert("Columns can be inserted on this protected worksheet.")
}
```

### Protection.AllowInsertingHyperlinks

如果允许在受保护的工作表上插入超链接，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowInsertingHyperlinks**

\*express是一个代表 Protection 对象的变量。

**说明**

在受保护的工作表上，只能将超链接插入未锁定或未保护的单元格中。

可以使用 **Protect** 方法参数设置 **AllowInsertingHyperlinks** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户在受保护的工作表上的单元格 A1 中插入超链接，并通知用户。*/
    Application.ActiveSheet.Unprotect()
        
    // Unlock cell A1.
    Range("A1").Locked = false
        
    // Allow hyperlinks to be inserted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowInsertingHyperlinks == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("Hyperlinks can be inserted on this protected worksheet.")
}
```

### Protection.AllowInsertingRows

如果允许用户在受保护的工作表上插入行，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowInsertingRows**

\*express是一个代表 Protection 对象的变量。

**说明**

可以使用 **Protect** 方法参数设置 **AllowInsertingRows** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户在受保护的工作表上插入行，并通知用户。*/
    Application.ActiveSheet.Unprotect()

    // Allow rows to be inserted on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowInsertingRows == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, true)
    }
    alert("Rows can be inserted on this protected worksheet.")
}
```

### Protection.AllowSorting

如果允许在受保护的工作表上使用排序选项，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowSorting**

\*express是一个代表 Protection 对象的变量。

**说明**

在受保护的工作表中，只能对未锁定或未保护的单元格进行排序。

可以使用 Protect 方法参数设置 **AllowSorting** 属性。

**示例**

``` JavaScript
function test(){
    /*本示例允许用户对受保护的工作表上未锁定或未保护的单元格进行排序，并通知用户。*/
    Application.ActiveSheet.Unprotect()

    // Unlock cells A1 through B5.
    Range("A1:B5").Locked = false

    // Allow sorting to be performed on the protected worksheet.
    if(Application.ActiveSheet.Protection.AllowSorting == false){
       Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("For cells A1 through B5, sorting can be performed on the protected worksheet.")
}
```

### Protection.AllowUsingPivotTables

如果允许用户在受保护的工作表上处理数据透视表，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AllowUsingPivotTables**

\*express是一个代表 Protection 对象的变量。

**说明**

**AllowUsingPivotTables** 属性应用于 <span class="glossary" "appendpopup(this,'idh_xldefnonolap_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none;">非 OLAP 源数据</span> 。

可以使用 **Protect** 方法参数设置 **AllowUsingPivotTables** 属性。

**示例**

``` JavaScript
function test() {
    /*本示例允许用户访问数据透视表并通知用户。本示例假定非 OLAP 数据透视表位于活动的工作表上。*/
    Application.ActiveSheet.Unprotect()

    // Allow pivot tables to be manipulated on a protected worksheet.
    if(Application.ActiveSheet.Protection.AllowUsingPivotTables == false) {
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, true)
    }
    alert("Pivot tables can be manipulated on the protected worksheet.")
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Protection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
