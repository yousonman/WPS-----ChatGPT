# MultiThreadedCalculation 对象

## 简介

返回或设置并发计算模式。

## 属性

| 序号 | 名称                                                 | 说明                                                                                                                                                               |
|:----:|:-----------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#MultiThreadedCalculation.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#MultiThreadedCalculation.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Enabled](#MultiThreadedCalculation.Enabled)         | 使用 Enabled 属性可以在运行时启用或禁用 MultiThreadedCalculation 对象。可读/写。                                                                                   |
|  4   | [Parent](#MultiThreadedCalculation.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  5   | [ThreadCount](#MultiThreadedCalculation.ThreadCount) | 获取进程的线程总数，这些线程是指定的 MultiThreadedCalculation 对象的一部分。                                                                                       |
|  6   | [ThreadMode](#MultiThreadedCalculation.ThreadMode)   | 返回或设置指定的 MultiThreadedCalculation 对象的线程模式。可读/写 XlThreadMode 类型。                                                                              |

## 成员属性

### MultiThreadedCalculation.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

### MultiThreadedCalculation.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### MultiThreadedCalculation.Enabled

使用 **Enabled** 属性可以在运行时启用或禁用 **MultiThreadedCalculation** 对象。可读/写。

**语法**

**express.Enabled**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

### MultiThreadedCalculation.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

### MultiThreadedCalculation.ThreadCount

获取进程的线程总数，这些线程是指定的 **MultiThreadedCalculation** 对象的一部分。

**语法**

**express.ThreadCount**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

### MultiThreadedCalculation.ThreadMode

返回或设置指定的 **MultiThreadedCalculation** 对象的线程模式。可读/写 **XlThreadMode** 类型。

**语法**

**express.ThreadMode**

\*express是一个代表 MultiThreadedCalculation 对象的变量。

**说明**

线程模式可以设置为 **XlThreadModeAutomatic** 或 **XlThreadModeManual** 。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/MultiThreadedCalculation](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
