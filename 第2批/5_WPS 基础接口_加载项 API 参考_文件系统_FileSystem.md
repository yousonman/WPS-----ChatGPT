# FileSystem 对象

## 简介

FileSystem 对象包括对文件和文件夹的一些基本和常见的操作接口，后续会根据实际需求持续增加。

## 说明

以下代码示例用于判断在临时目录下是否有a.txt这个文件

``` JavaScript
let tempPath = Application.Env.GetTempPath()
if (Application.FileSystem.Exists(tempPath + 'a.txt'))
    alert("a.txt文件存在")
```

## 方法

| 序号 | 名称                                                   | 说明                                           |
|:----:|:-------------------------------------------------------|:-----------------------------------------------|
|  1   | [AppendFile](#FileSystem.AppendFile)                   | 往文件末尾添加数据                             |
|  2   | [copyFileSync](#FileSystem.copyFileSync)               | 生成文件副本                                   |
|  3   | [Exists](#FileSystem.Exists)                           | 判断一个文件和文件夹是否存在。                 |
|  4   | [existsSync](#FileSystem.existsSync)                   | 判断目录是否存在                               |
|  5   | [Mkdir](#FileSystem.Mkdir)                             | 根据给定的path创建一个文件夹。                 |
|  6   | [mkdirSync](#FileSystem.mkdirSync)                     | 创建目录                                       |
|  7   | [mkdtempSync](#FileSystem.mkdtempSync)                 | 创建临时目录                                   |
|  8   | [readAsBinaryString](#FileSystem.readAsBinaryString)   | 读取指定路径的文件，返回二进制字符串数据。     |
|  9   | [readdirSync](#FileSystem.readdirSync)                 | 获取目录下的子目录对象数组（包含目录相关属性） |
|  10  | [ReadFile](#FileSystem.ReadFile)                       | 获取文件的内容                                 |
|  11  | [readFileString](#FileSystem.readFileString)           | 读取指定路径的文件，返回字符串。               |
|  12  | [Remove](#FileSystem.Remove)                           | 删除指定的path所代表的文件或文件夹。           |
|  13  | [rmdirSync](#FileSystem.rmdirSync)                     | 删除目录                                       |
|  14  | [tmpdir](#FileSystem.tmpdir)                           | 获取系统的临时文件目录                         |
|  15  | [unlinkSync](#FileSystem.unlinkSync)                   | 删除文件                                       |
|  16  | [writeAsBinaryString](#FileSystem.writeAsBinaryString) | 写二进制字符串对象到指定文件。                 |
|  17  | [WriteFile](#FileSystem.WriteFile)                     | 创建文件                                       |
|  18  | [writeFileString](#FileSystem.writeFileString)         | 写字符串到指定路径的文件。                     |

## 属性

| 序号 | 名称                           | 说明 |
|:----:|:-------------------------------|:-----|
|  1   | [Reserve](#FileSystem.Reserve) |      |

## 成员方法

### FileSystem.AppendFile

往文件末尾添加数据

**语法**

**express.AppendFile ( *FilePath* , *Data* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明             |
|------------|-----------|----------|------------------|
| *FilePath* | 必选      | String   | 要写入的文件名   |
| *Data*     | 必选      | String   | 要写入文件的内容 |

### FileSystem.copyFileSync

生成文件副本

**语法**

**express.copyFileSync ( *SrcFilePath* , *DestFilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                     |
|----------------|-----------|----------|--------------------------|
| *SrcFilePath*  | 必选      | String   | 源文件的路径包括文件名称 |
| *DestFilePath* | 必选      | String   | 新创建的文件副本文件名称 |

### FileSystem.Exists

判断一个文件和文件夹是否存在。

**语法**

**express.Exists ( *path* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                             |
|--------|-----------|----------|--------------------------------------------------|
| *path* | 必选      | String   | 文件或文件夹的路径，这个路径一定要是一个全路径。 |

### FileSystem.existsSync

判断目录是否存在

**语法**

**express.existsSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明               |
|------------|-----------|----------|--------------------|
| *FilePath* | 必选      | String   | 要判断的文件的名称 |

### FileSystem.Mkdir

根据给定的path创建一个文件夹。

**语法**

**express.Mkdir ( *path* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                               |
|--------|-----------|----------|----------------------------------------------------|
| *path* | 必选      | String   | 要创建的文件夹的路径，这个路径一定要是一个全路径。 |

### FileSystem.mkdirSync

创建目录

**语法**

**express.mkdirSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明           |
|------------|-----------|----------|----------------|
| *FilePath* | 必选      | String   | 创建目录的路径 |

### FileSystem.mkdtempSync

创建临时目录

**语法**

**express.mkdtempSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明           |
|------------|-----------|----------|----------------|
| *FilePath* | 必选      | String   | 临时目录的路径 |

### FileSystem.readAsBinaryString

读取指定路径的文件，返回二进制字符串数据。

**语法**

**express.readAsBinaryString ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明     |
|------------|-----------|----------|----------|
| *FilePath* | 必选      | String   | 文件路径 |

### FileSystem.readdirSync

获取目录下的子目录对象数组（包含目录相关属性）

**语法**

**express.readdirSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明               |
|------------|-----------|----------|--------------------|
| *FilePath* | 必选      | String   | 要获取的目录的路径 |

### FileSystem.ReadFile

获取文件的内容

**语法**

**express.ReadFile ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明         |
|------------|-----------|----------|--------------|
| *FilePath* | 必选      | String   | 读取文件内容 |

### FileSystem.readFileString

读取指定路径的文件，返回字符串。

**语法**

**express.readFileString ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明     |
|------------|-----------|----------|----------|
| *FilePath* | 必选      | String   | 文件路径 |

### FileSystem.Remove

删除指定的path所代表的文件或文件夹。

**语法**

**express.Remove ( *path* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                                     |
|--------|-----------|----------|----------------------------------------------------------|
| *path* | 必选      | String   | 要删除的文件或文件夹的路径，这个路径一定要是一个全路径。 |

### FileSystem.rmdirSync

删除目录

**语法**

**express.rmdirSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                 |
|------------|-----------|----------|----------------------|
| *FilePath* | 必选      | String   | 要删除的文件目录路径 |

### FileSystem.tmpdir

获取系统的临时文件目录

**语法**

**express.tmpdir ()**

\*express是一个代表 FileSystem 对象的变量。

### FileSystem.unlinkSync

删除文件

**语法**

**express.unlinkSync ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明             |
|------------|-----------|----------|------------------|
| *FilePath* | 必选      | String   | 要删除的文件路径 |

### FileSystem.writeAsBinaryString

写二进制字符串对象到指定文件。

**语法**

**express.writeAsBinaryString ( *FilePath* , *binaryString* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明             |
|----------------|-----------|----------|------------------|
| *FilePath*     | 必选      | String   | 文件路径         |
| *binaryString* | 必选      | String   | 二进制字符串数据 |

### FileSystem.WriteFile

创建文件

**语法**

**express.WriteFile ( *FilePath* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明     |
|------------|-----------|----------|----------|
| *FilePath* | 必选      | String   | 文件路径 |

### FileSystem.writeFileString

写字符串到指定路径的文件。

**语法**

**express.writeFileString ( *FilePath* , *data* )**

\*express是一个代表 FileSystem 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明       |
|------------|-----------|----------|------------|
| *FilePath* | 必选      | String   | 文件路径   |
| *data*     | 必选      | String   | 字符串数据 |

## 成员属性

### FileSystem.Reserve

**语法**

**express.Reserve**

\*express是一个代表 FileSystem 对象的变量。

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/文件系统/FileSystem](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
