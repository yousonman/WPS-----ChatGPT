# PickerProperty 对象

## 简介

代表一个对象，以便传递自定义属性。

## 说明

``` JavaScript
//下面的代码设置选取器对话框的属性，然后显示选取器对话框。
function test()
{
    // Configure the Picker Dialog properties.
    let objPickerDialog = Application.PickerDialog
    objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
    objPickerDialog.Title = "Sample Picker Dialog"
    let objPickerProperties = objPickerDialog.Properties
    let objPickerProperty = objPickerProperties.Add("SiteUrl", "http://my", Application.Enum.msoPickerFieldText)
    let objPickerExistingResults = objPickerDialog.CreatePickerResults()
    let objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")
     
    // Show the Picker Dialog and get the results.
    let objPickerResults = objPickerDialog.Show(true, objPickerExistingResult)
}
```

## 属性

| 序号 | 名称                                       | 说明                                                                           |
|:----:|:-------------------------------------------|:-------------------------------------------------------------------------------|
|  1   | [Application](#PickerProperty.Application) | 获取一个 Application 对象，该对象代表 PickerProperty 对象的容器应用程序。只读  |
|  2   | [Creator](#PickerProperty.Creator)         | 获取一个 32 位整数，该整数指示在其中创建了 PickerProperty 对象的应用程序。只读 |
|  3   | [Id](#PickerProperty.Id)                   | 检索关联的 PickerProperty 对象的唯一 Id。只读                                  |
|  4   | [Type](#PickerProperty.Type)               | 检索选取器属性的类型。只读                                                     |
|  5   | [Value](#PickerProperty.Value)             | 检索选取器属性的值。只读                                                       |

## 成员属性

### PickerProperty.Application

获取一个 **Application** 对象，该对象代表 **PickerProperty** 对象的容器应用程序。只读

**语法**

**express.Application**

\*express是一个代表 PickerProperty 对象的变量。

### PickerProperty.Creator

获取一个 32 位整数，该整数指示在其中创建了 **PickerProperty** 对象的应用程序。只读

**语法**

**express.Creator**

\*express是一个代表 PickerProperty 对象的变量。

### PickerProperty.Id

检索关联的 **PickerProperty** 对象的唯一 Id。只读

**语法**

**express.Id**

\*express是一个代表 PickerProperty 对象的变量。

### PickerProperty.Type

检索选取器属性的类型。只读

**语法**

**express.Type**

\*express是一个代表 PickerProperty 对象的变量。

### PickerProperty.Value

检索选取器属性的值。只读

**语法**

**express.Value**

\*express是一个代表 PickerProperty 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/PickerProperty](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
