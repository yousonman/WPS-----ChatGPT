# SignatureSetup 对象

## 简介

代表用于设置签名数据包的信息。

## 说明

``` JavaScript
/*以下示例将设置签名数据包的 SignatureSetup 对象的各种属性。*/
function test(){
    var objSigSetup = Application.ActiveWorkbook.Signatures
    objSigSetup.AllowComments = true 
    objSigSetup.ShowSignDate = true 
    objSigSetup.SigningInstructions = "Please sign this document." 
    objSigSetup.SuggestedSignerEmail = "jdow@example.com"
}
```

## 属性

| 序号 | 名称                                                         | 说明                                                                                     |
|:----:|:-------------------------------------------------------------|:-----------------------------------------------------------------------------------------|
|  1   | [AdditionalXml](#SignatureSetup.AdditionalXml)               | 获取或设置在设置过程中添加到签名的任何附加 XML 信息。可读/写。                           |
|  2   | [AllowComments](#SignatureSetup.AllowComments)               | 获取或设置一个 Boolean 类型的值，指定签署人是否可以在 “签名” 对话框中输入注释。可读/写。 |
|  3   | [Application](#SignatureSetup.Application)                   | 获取一个代表 SignatureSetup 对象的容器应用程序的 Application 对象。只读。                |
|  4   | [Creator](#SignatureSetup.Creator)                           | 获取一个 32 位整数，指示创建 SignatureSetup 对象时所使用的应用程序。只读。               |
|  5   | [Id](#SignatureSetup.Id)                                     | 获取文档的签名提供程序的 ID。只读。                                                      |
|  6   | [ReadOnly](#SignatureSetup.ReadOnly)                         | 获取一个 Boolean 类型的值，指示 SignatureSetup 对象是否为只读。只读。                    |
|  7   | [ShowSignDate](#SignatureSetup.ShowSignDate)                 | 获取或设置一个 Boolean 类型的值，指示是否应显示文档的签署日期。可读/写。                 |
|  8   | [SignatureProvider](#SignatureSetup.SignatureProvider)       | 获取一个标识所安装的签名提供程序加载项的值。只读。                                       |
|  9   | [SigningInstructions](#SignatureSetup.SigningInstructions)   | 获取或设置用于签署文档的指令。可读/写。                                                  |
|  10  | [SuggestedSigner](#SignatureSetup.SuggestedSigner)           | 获取或设置文档主要签署人的名称。可读/写。                                                |
|  11  | [SuggestedSignerEmail](#SignatureSetup.SuggestedSignerEmail) | 获取或设置文档签署人的电子邮件地址。可读/写。                                            |
|  12  | [SuggestedSignerLine2](#SignatureSetup.SuggestedSignerLine2) | 获取或设置建议签署人信息的第二行（例如，职务）。可读/写。                                |

## 成员属性

### SignatureSetup.AdditionalXml

获取或设置在设置过程中添加到签名的任何附加 XML 信息。可读/写。

**语法**

**express.AdditionalXml**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.AllowComments

获取或设置一个 **Boolean** 类型的值，指定签署人是否可以在 **“签名”** 对话框中输入注释。可读/写。

**语法**

**express.AllowComments**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.Application

获取一个代表 **SignatureSetup** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.Creator

获取一个 32 位整数，指示创建 **SignatureSetup** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.Id

获取文档的签名提供程序的 ID。只读。

**语法**

**express.Id**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.ReadOnly

获取一个 **Boolean** 类型的值，指示 **SignatureSetup** 对象是否为只读。只读。

**语法**

**express.ReadOnly**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.ShowSignDate

获取或设置一个 **Boolean** 类型的值，指示是否应显示文档的签署日期。可读/写。

**语法**

**express.ShowSignDate**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.SignatureProvider

获取一个标识所安装的签名提供程序加载项的值。只读。

**语法**

**express.SignatureProvider**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.SigningInstructions

获取或设置用于签署文档的指令。可读/写。

**语法**

**express.SigningInstructions**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.SuggestedSigner

获取或设置文档主要签署人的名称。可读/写。

**语法**

**express.SuggestedSigner**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.SuggestedSignerEmail

获取或设置文档签署人的电子邮件地址。可读/写。

**语法**

**express.SuggestedSignerEmail**

\*express是一个代表 SignatureSetup 对象的变量。

### SignatureSetup.SuggestedSignerLine2

获取或设置建议签署人信息的第二行（例如，职务）。可读/写。

**语法**

**express.SuggestedSignerLine2**

\*express是一个代表 SignatureSetup 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SignatureSetup](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
