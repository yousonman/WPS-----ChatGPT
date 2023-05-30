# Env 对象

## 简介

Env 对象主要用于取系统环境基本信息，这个对象目前提供了取用户目录、临时目录等相关信息，后续可能会根据实际需求不断扩充。

## 说明

以下代码示例，取到系统临时目录并打印到控制台。

``` JavaScript
let tempPath = Application.Env.GetTempPath()
console.log(tempPath)  //如果是jde环境，则是Debug.Print(tempPath)
```

## 方法

| 序号 | 名称                                            | 说明                                                  |
|:----:|:------------------------------------------------|:------------------------------------------------------|
|  1   | [GetAppDataPath](#Env.GetAppDataPath)           | 取系统%AppData%/Roming目录的路径，仅Windows平台适用。 |
|  2   | [GetDesktopDpi](#Env.GetDesktopDpi)             | 取系统DPI。                                           |
|  3   | [GetHomePath](#Env.GetHomePath)                 | 取系统HOME目录的路径，代表当前用户主目录。            |
|  4   | [GetProgramDataPath](#Env.GetProgramDataPath)   | 取系统ProgramData目录的路径，仅windows平台适用。      |
|  5   | [GetProgramFilesPath](#Env.GetProgramFilesPath) | 取系统ProgramFiles目录的路径，仅Windows平台适用。     |
|  6   | [GetRootPath](#Env.GetRootPath)                 | 取系统根目录的路径。                                  |
|  7   | [GetTempPath](#Env.GetTempPath)                 | 取系统临时目录的路径。                                |

## 成员方法

### Env.GetAppDataPath

取系统%AppData%/Roming目录的路径，仅Windows平台适用。

**语法**

**express.GetAppDataPath ()**

\*express是一个代表 Env 对象的变量。

### Env.GetDesktopDpi

取系统DPI。

**语法**

**express.GetDesktopDpi ()**

\*express是一个代表 Env 对象的变量。

### Env.GetHomePath

取系统HOME目录的路径，代表当前用户主目录。

**语法**

**express.GetHomePath ()**

\*express是一个代表 Env 对象的变量。

### Env.GetProgramDataPath

取系统ProgramData目录的路径，仅windows平台适用。

**语法**

**express.GetProgramDataPath ()**

\*express是一个代表 Env 对象的变量。

### Env.GetProgramFilesPath

取系统ProgramFiles目录的路径，仅Windows平台适用。

**语法**

**express.GetProgramFilesPath ()**

\*express是一个代表 Env 对象的变量。

### Env.GetRootPath

取系统根目录的路径。

**语法**

**express.GetRootPath ()**

\*express是一个代表 Env 对象的变量。

### Env.GetTempPath

取系统临时目录的路径。

**语法**

**express.GetTempPath ()**

\*express是一个代表 Env 对象的变量。

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/系统环境/Env](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
