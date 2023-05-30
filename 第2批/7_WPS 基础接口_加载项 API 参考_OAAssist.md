# OAAssist 对象

## 方法

| 序号 | 名称                                               | 说明                                                                 |
|:----:|:---------------------------------------------------|:---------------------------------------------------------------------|
|  1   | [CoCreateInstance](#OAAssist.CoCreateInstance)     | 根据给定的程序标识符从注册表找出对应的类标识符，再创建对应类的实例。 |
|  2   | [DownloadFile](#OAAssist.DownloadFile)             | 从指定的url处下载文件，返回的字符串为保存地址。                      |
|  3   | [ExecuteInCOMAddIns](#OAAssist.ExecuteInCOMAddIns) | 执行COM加载项                                                        |
|  4   | [ShellExecute](#OAAssist.ShellExecute)             | 使用壳服务打开网页或调起进程。                                       |
|  5   | [UploadFile](#OAAssist.UploadFile)                 | 向指定链接上传文件。                                                 |
|  6   | [WebNotify](#OAAssist.WebNotify)                   | 向客户端发送推送。                                                   |

## 成员方法

### OAAssist.CoCreateInstance

根据给定的程序标识符从注册表找出对应的类标识符，再创建对应类的实例。

**语法**

**express.CoCreateInstance ( *ProgId* )**

\*express是一个代表 OAAssist 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明       |
|----------|-----------|----------|------------|
| *ProgId* | 必选      | String   | 程序标识符 |

### OAAssist.DownloadFile

从指定的url处下载文件，返回的字符串为保存地址。

**语法**

**express.DownloadFile ( *Url* , *Success* , *Fail* )**

\*express是一个代表 OAAssist 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                 |
|-----------|-----------|----------|----------------------|
| *Url*     | 必选      | String   | 下载链接             |
| *Success* | 可选      | Variant  | 下载成功时的回调函数 |
| *Fail*    | 可选      | Variant  | 下载失败时的回调函数 |

**返回值**

String

**说明**

第二、三个参数传入时为异步下载，不传入时为同步下载。

### OAAssist.ExecuteInCOMAddIns

执行COM加载项

**语法**

**express.ExecuteInCOMAddIns ( *Addin* , *OnAction* )**

\*express是一个代表 OAAssist 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明          |
|------------|-----------|----------|---------------|
| *Addin*    | 必选      | String   | COM加载项名字 |
| *OnAction* | 必选      | String   | 执行函数名字  |

### OAAssist.ShellExecute

使用壳服务打开网页或调起进程。

**语法**

**express.ShellExecute ( *Url* , *Params* )**

\*express是一个代表 OAAssist 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明             |
|----------|-----------|----------|------------------|
| *Url*    | 可选      | String   | 网页或进程的地址 |
| *Params* | 可选      | String   | 启动参数         |

**说明**

当不传入参数时，该函数立即返回。对于带有“http”或“https”字样的地址，会打开对应的网页。以下是启动wps进程的参数:

| 传入参数         | 启动进程  |
|------------------|-----------|
| ksowebstartupwps | WPS文字   |
| ksowebstartupwpp | WPS演示   |
| ksowebstartupet  | WPS表格   |
| ksowpscloudsvr   | WPS云服务 |

同样也可直接传入地址调起其余进程。

注意：当前进程的终止并不会导致被调起进程的终止。

### OAAssist.UploadFile

向指定链接上传文件。

**语法**

**express.UploadFile ( *FileName* , *FilePath* , *Url* , *Field* , *Success* , *Fail* )**

\*express是一个代表 OAAssist 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                           |
|------------|-----------|----------|--------------------------------|
| *FileName* | 必选      | String   | 上传文件名                     |
| *FilePath* | 必选      | String   | 文件路径                       |
| *Url*      | 必选      | String   | 上传链接                       |
| *Field*    | 可选      | String   | https Header中上传时的name属性 |
| *Success*  | 可选      | Variant  | 上传成功时的回调函数           |
| *Fail*     | 必选      | Variant  | 上传失败时的回调函数           |

**返回值**

Boolean

### OAAssist.WebNotify

向客户端发送推送。

**语法**

**express.WebNotify ( *notifyStr* , *Force* )**

\*express是一个代表 OAAssist 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明       |
|-------------|-----------|----------|------------|
| *notifyStr* | 必选      | String   | 推送的内容 |
| *Force*     | 必选      | Boolean  | 是否强制   |

**返回值**

Boolean

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/OAAssist](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
