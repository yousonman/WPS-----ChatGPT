# Options 对象

## 简介

代表 WPP 中的应用程序选项。

## 属性

| 序号 | 名称                                                | 说明                                                                                                      |
|:----:|:----------------------------------------------------|:----------------------------------------------------------------------------------------------------------|
|  1   | [DisplayPasteOptions](#Options.DisplayPasteOptions) | WPP 的 MsoTrue 属性值用于显示 “粘贴选项” 按钮，该按钮显示在新粘贴文本的正下方。可读/写 MsoTriState 类型。 |

## 成员属性

### Options.DisplayPasteOptions

WPP 的 **MsoTrue** 属性值用于显示 **“粘贴选项”** 按钮，该按钮显示在新粘贴文本的正下方。可读/写 **MsoTriState** 类型。

**语法**

**express.DisplayPasteOptions**

\*express是一个代表 Options 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue                                       |

**示例**

如果 **“粘贴选项”** 按钮为禁用状态，本示例将启用该选项。

``` JavaScript
function ShowPasteOptionsButton(){
let opt = Application.Options
    if(opt.DisplayPasteOptions == false){
        opt.DisplayPasteOptions = true
    }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/Options](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
