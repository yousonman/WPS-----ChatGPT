# ApiEvent 对象

## 简介

ApiEvent 对象用于接收 WPS 应用程序各种事件通知。

## 说明

|                  |
|:-----------------|
| WPS 开发人员参考 |

ApiEvent 对象用于注册和反注册 WPS 事件回调函数，注册一个事件回调函数的接口为： ApiEvent.AddApiEventListener(eventName, func) ， 其中 eventName 表示要响应的事件类型，这些事件类型由wps内部规定， func 表示要回调的js函数，当wps有某个事件被触发时，这些注册的js函数会被wps执行。 例如需要在打开文档时，弹出一个打开文档后的消息框，示例代码如下如下：

``` JavaScript
Application.ApiEvent.AddApiEventListener("DocumentOpen", (doc)=>{alert("文档已打开，文档名是： " + doc.Name)});
```

同样，如果不再需要某个已注册的回调函数，可以把相关的回调函数进行反注册，wps内部也会在网页关闭时，把网页中所有已注册的回调函数自动反注册，反注册一个回调函数的方式为： ApiEvent.RemoveApiEventListener(eventName) 。其中 eventName 表示事件类型，反注册后当前网页不再有此事件的回调函数被执行。

``` JavaScript
Application.ApiEvent.RemoveApiEventListener("DocumentOpen")
```

api事件分为询问型事件和通知型事件，询问型事件是指wps在准备做某一个事情时发出事件回调，第三方代码有机会来控制这个事情， 通知型事件是指wps做了某个事件后通知第三方代码这个事情已经结束，这种情况下事情已然发生。以上面的几个事件类型为例， DocumentBeforeClose 和 DocumentBeforeSave 事件是询问型事件，其它几个是通知型事件。下面分别对两种事件各举一个例子

1.询问型事件： 示例：当用户想保存一个文件时，通过js事件回调禁止保存这个文件,代码如下：

``` JavaScript
 let func = (doc)=>{
    alert(doc.Name)
    Application.ApiEvent.Cancel = true
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeSave", func) 
```

上面代码的意思是在文件保存前 OnDocBeforeSave 这个js函数被调用到，此时js函数里边通过 ApiEvent . Cancel 这个属性设置了要取消掉这次流程， wps执行完 OnBeforeSave 这个js函数后，检测到 Cancel 属性被设置为true, 从而中止了保存流程。 对于询问型事件，一般是通过 Cancel 属性来继续或是中止流程的执行。

2.通知型事件： 示例：示例：当用户打开文件后弹出这个文件的磁盘路径，代码如下：

``` JavaScript
 let func = (doc)=>{
    alert(doc.Name)
}                                       
Application.ApiEvent.AddApiEventListener("DocumentOpen", func) 
```

上面代码的意思是在文件打开这个事件发生后，wps把事件发生后通知给了网页上注册的回调函数。

关于事件类型的定义，在wps/et/wpp中各有不同，事件类型也基本上保持与office标准的事件类型相对应，详细的事件类型如下表：

wps文字支持的事件类型如下：

| 名称                              | 级别        | 类型   | 触发时机                                              |
|-----------------------------------|-------------|--------|-------------------------------------------------------|
| DocumentOpen                      | Application | 通知型 | 文档打开时触发。                                      |
| DocumentBeforeClose               | Application | 询问型 | 任一打开的文档关闭之前触发此事件。                    |
| DocumentBeforeSave                | Application | 询问型 | 任一打开的文档保存之前触发此事件。                    |
| DocumentAfterClose                | Application | 通知型 | 任一打开的文档关闭之后触发此事件。                    |
| DocumentChange                    | Application | 通知型 | 任一打开的文档内容发生变化时触发此事件。              |
| DocumentBeforePrint               | Application | 询问型 | 任一打开的文档打印之前触发此事件。                    |
| DocumentNew                       | Application | 通知型 | 新建任一文档时触发此事件。                            |
| DocumentRightsInfo                | Application | 通知型 | 当操作文档权限时会触发此事件。                        |
| DocumentBeforeCopy                | Application | 询问型 | 在任一文档复制之前触发该事件。                        |
| DocumentBeforePaste               | Application | 询问型 | 在任一文档粘贴之前触发该事件。                        |
| DocumentBeforeNew                 | Application | 询问型 | 在任一文档新建之前触发该事件。                        |
| DocumentBeforeOpen                | Application | 询问型 | 在任一文档打开之前触发该事件。                        |
| DocumentAfterSave                 | Application | 通知型 | 在任一文档保存之后触发该事件。                        |
| DocumentViewFocusIn               | Application | 通知型 | 在任一文档获取到焦点之后触发该事件。                  |
| DocumentViewFocusOut              | Application | 通知型 | 在任一文档失去焦点之后触发该事件。                    |
| DocumentAfterPrint                | Application | 通知型 | 在任一文档打印之后触发该事件。                        |
| WindowActivate                    | Application | 通知型 | 当激活文档窗口时触发此事件。                          |
| WindowDeactivate                  | Application | 通知型 | 任一文档窗口被切换到非激活状态时触发此事件。          |
| WindowSelectionChange             | Application | 通知型 | 活动窗口所选内容更改时将触发此事件。                  |
| WindowBeforeRightClick            | Application | 询问型 | 当右击任一文档窗口之前触发此事件。                    |
| WindowBeforeDoubleClick           | Application | 询问型 | 当双击任一文档窗口之前触发此事件。                    |
| ApplicationQuit                   | Application | 通知型 | 当退出文字组件进程时触发此事件。                      |
| AfterLogin                        | Application | 通知型 | 在用户登陆之后触发此事件。                            |
| AfterLogout                       | Application | 通知型 | 在用户登出之后触发此事件。                            |
| AfterVIPInfoChanged               | Application | 通知型 | 在用户会员信息变化之后触发此事件。                    |
| AfterWebExtensionDataRangeChange  | Application | 通知型 | 在网络扩展数据范围变化之后触发此事件。                |
| AfterTaskPaneShow                 | Application | 通知型 | 当任务窗格显示之后触发此事件。                        |
| AfterTaskPaneHidden               | Application | 通知型 | 当任务窗格隐藏之后触发此事件。                        |
| QueryDocumentPrintCopies          | Application | 询问型 | 当打印复本排队时触发此事件。                          |
| ContentChange                     | Application | 通知型 | 当文档内容发生变化时触发此事件。                      |
| SelectionAfterStyleChange         | Application | 通知型 | 当选择的内容字体变化之后触发此事件。                  |
| FileAfterSave                     | Application | 通知型 | 在文件保存之后触发该事件。                            |
| ContentControlAfterAdd            | Application | 通知型 | 当内容控件被添加到文档之后触发此事件。                |
| ContentControlBeforeDelete        | Application | 通知型 | 当内容控件被从文档删除之前触发此事件。                |
| ContentControlOnExit              | Application | 通知型 | 当退出内容控件时触发此事件。                          |
| ContentControlOnEnter             | Application | 通知型 | 当进入内容控件时触发此事件。                          |
| ContentControlBeforeStoreUpdate   | Application | 通知型 | 当内容控件的内容去更新XML映射内容时触发此事件。       |
| ContentControlBeforeContentUpdate | Application | 通知型 | 当内容控件的内容发生来自XML映射内容更新时触发此事件。 |

wps表格支持的api事件为：

| 名称                           | 级别        | 类型   | 触发时机                                                         |
|--------------------------------|-------------|--------|------------------------------------------------------------------|
| WorkbookOpen                   | Application | 通知型 | 当打开一个工作簿时触发此事件。                                   |
| NewWorkbook                    | Application | 通知型 | 当新建工作簿时触发此事件。                                       |
| WorkbookBeforeSave             | Application | 询问型 | 任一工作簿被保存之前触发此事件。                                 |
| WorkbookAfterSave              | Application | 通知型 | 任一工作簿被保存之后触发此事件。                                 |
| WorkbookBeforeClose            | Application | 询问型 | 任一打开的工作簿关闭之前触发此事件。                             |
| WorkbookBeforePrint            | Application | 询问型 | 在打印任一打开的工作簿之前触发此事件。                           |
| SheetSelectionChange           | Application | 通知型 | 该工作簿任一工作表上的选定区域发生更改时，将触发此事件。         |
| SheetChange                    | Application | 通知型 | 当用户或外部链接更改了该工作簿任一工作表中的单元格时触发此事件。 |
| SheetActivate                  | Application | 通知型 | 当激活任一工作表时触发此事件。                                   |
| SheetDeactivate                | Application | 通知型 | 当该工作簿任一工作表被切换到非激活状态时触发此事件。             |
| WindowActivate                 | Application | 通知型 | 任一工作簿窗口被激活时，将触发此事件。                           |
| WindowDeactivate               | Application | 通知型 | 任一工作簿窗口被切换到非激活状态时触发此事件。                   |
| SheetBeforeDoubleClick         | Application | 询问型 | 当双击任一工作表之前触发此事件。                                 |
| SheetBeforeRightClick          | Application | 询问型 | 当右击任一工作表之前触发此事件 。                                |
| SheetCalculate                 | Application | 通知型 | 在任一工作表进行计算时触发此事件。                               |
| WorkbookActivate               | Application | 通知型 | 任一工作簿被激活时，将触发此事件。                               |
| WorkbookDeactivate             | Application | 通知型 | 任一工作簿被切换到非激活状态时触发此事件。                       |
| WorkbookNewSheet               | Application | 通知型 | 在任一打开的工作簿中创建新工作表时触发此事件。                   |
| WindowResize                   | Application | 通知型 | 任一工作簿窗口调整大小时将触发此事件。                           |
| SheetFollowHyperlink           | Application | 通知型 | 单击任一工作表的超链接时触 发此事件。                            |
| AfterCalculate                 | Application | 通知型 | 当计算完成后触发此事件。                                         |
| SheetBeforeDelete              | Application | 通知型 | 当删除任一工作表之前触发此事件。                                 |
| DocumentRightsInfo             | Application | 通知型 | 当操作文档权限时会触发此事件。                                   |
| AfterLogin                     | Application | 通知型 | 当用户登陆之后触发此事件。                                       |
| AfterLogout                    | Application | 通知型 | 当用户登出之后触发此事件。                                       |
| AfterVIPInfoChanged            | Application | 通知型 | 当用户的会员信息变化之后触发此事件。                             |
| AfterWebExtWatchingDataUpdated | Application | 通知型 | 当网页拓展监控的数据发生变化之后触发此事件                       |
| AfterTaskPaneShow              | Application | 通知型 | 当任务窗格显示之后触发此事件。                                   |
| AfterTaskPaneHidden            | Application | 通知型 | 当任务窗格隐藏之后触发此事件                                     |
| DocumentBeforeNew              | Application | 询问型 | 在任一文档新建之前触发该事件。                                   |
| DocumentBeforeOpen             | Application | 询问型 | 在任一文档打开之前触发该事件。                                   |
| DocumentBeforeCopy             | Application | 询问型 | 在任一文档复制之前触发该事件。                                   |
| DocumentBeforePaste            | Application | 询问型 | 在任一文档粘贴之前触发该事件。                                   |
| FileAfterSave                  | Application | 通知型 | 在任一文档保存之后触发该事件。                                   |
| LinkedDataTypeConvert          | Application | 通知型 | 在连接数据类型转变时触发该事件。                                 |
| LinkedDataTypeChange           | Application | 通知型 | 在连接数据类型改变时触发该事件。                                 |
| LinkedDataTypeRefresh          | Application | 通知型 | 在连接数据类型刷新时触发该事件。                                 |
| LinkedDataTypeCancel           | Application | 通知型 | 在连接数据类型取消时触发该事件。                                 |
| DocumentAfterPrint             | Application | 通知型 | 该文档打印之后触发该事件。                                       |

wps演示支持的api事件为：

| 名称                    | 级别        | 类型   | 触发时机                               |
|-------------------------|-------------|--------|----------------------------------------|
| AfterLogin              | Application | 通知型 | 用户登录之后触发该事件。               |
| AfterLogout             | Application | 通知型 | 用户登出之后触发该事件。               |
| AfterPresentationOpen   | Application | 通知型 | 任一演示文稿打开后触发该事件。         |
| AfterTaskPaneHidden     | Application | 通知型 | 当任务窗格隐藏之后触发此事件。         |
| AfterTaskPaneShow       | Application | 通知型 | 当任务窗格显示之后触发此事件。         |
| DocumentAfterPrint      | Application | 通知型 | 当文档打印之后触发此事件。             |
| DocumentBeforeCopy      | Application | 询问型 | 当文档复制之前触发此事件。             |
| DocumentBeforeNew       | Application | 询问型 | 当文档新建之前触发此事件。             |
| DocumentBeforeOpen      | Application | 询问型 | 当文档打开之前触发此事件。             |
| DocumentBeforePaste     | Application | 询问型 | 当文档粘贴之前触发此事件。             |
| DocumentRightsInfo      | Application | 通知型 | 当操作文档权限时会触发此事件。         |
| FileAfterSave           | Application | 通知型 | 当文件保存之后触发此事件。             |
| NewPresentation         | Application | 通知型 | 当新建演示文稿时触发此事件。           |
| PresentationBeforeClose | Application | 询问型 | 任一打开的演示文稿关闭之前触发此事件。 |
| PresentationBeforePrint | Application | 询问型 | 任一演示文稿打印之前触发此事件。       |
| PresentationBeforeSave  | Application | 询问型 | 任一演示文稿被保存之前触发此事件。     |
| PresentationClose       | Application | 通知型 | 任一演示文稿关闭时触发此事件。         |
| PresentationCloseFinal  | Application | 通知型 | 任一演示文稿关闭之后触发此事件。       |
| PresentationNewSlide    | Application | 通知型 | 任一演示文稿新建幻灯片时触发此事件。   |
| PresentationOpen        | Application | 通知型 | 任一演示文稿打开时触发此事件。         |
| PresentationPrint       | Application | 通知型 | 任一演示文稿打印时触发此事件。         |
| PresentationSave        | Application | 通知型 | 任一演示文稿保存时触发此事件。         |
| Quit                    | Application | 通知型 | 任一演示文稿退出时触发此事件。         |
| SlideSelectionChanged   | Application | 通知型 | 演示文稿幻灯片选择变化时触发此事件。   |
| SlideShowBegin          | Application | 通知型 | 演示文稿开始播放时触发此事件。         |
| SlideShowEnd            | Application | 通知型 | 演示文稿结束播放时触发此事件。         |
| SlideShowNextClick      | Application | 通知型 | 演示文稿播放时下一次点击触发此事件。   |
| SlideShowNextSlide      | Application | 通知型 | 演示文稿播放下一个幻灯片时触发此事件。 |
| SlideShowOnNext         | Application | 通知型 | 演示文稿播放点击下一页时触发此事件。   |
| SlideShowOnPrevious     | Application | 通知型 | 演示文稿播放点击上一页时触发此事件。   |
| WindowActivate          | Application | 通知型 | 当激活演示窗口时触发此事件。           |
| WindowSelectionChange   | Application | 通知型 | 活动窗口所选内容更改时将触发此事件。   |

## 方法

| 序号 | 名称                                                       | 说明                  |
|:----:|:-----------------------------------------------------------|:----------------------|
|  1   | [AddApiEventListener](#ApiEvent.AddApiEventListener)       | 注册wps事件的回调。   |
|  2   | [RemoveApiEventListener](#ApiEvent.RemoveApiEventListener) | 反注册wps事件的回调。 |

## 属性

| 序号 | 名称                       | 说明                                                       |
|:----:|:---------------------------|:-----------------------------------------------------------|
|  1   | [Cancel](#ApiEvent.Cancel) | 当事件是询问型事件时，此属性用于设置是否中止事件进行的状态 |

## 成员方法

### ApiEvent.AddApiEventListener

注册wps事件的回调。

**语法**

**express.AddApiEventListener ( *eventName* , *func* )**

\*express是一个代表 ApiEvent 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                           |
|-------------|-----------|----------|------------------------------------------------------------------------------------------------|
| *eventName* | 必选      | String   | 代表事件类型，事件类型由wps内部规定，有关wps支持的事件类型，可以参数ApiEvent对象说明相关文档。 |
| *func*      | 必选      | js函数   | 代表要注册的事件回调，当事件发生时，此函数被调用到。                                           |

### ApiEvent.RemoveApiEventListener

反注册wps事件的回调。

**语法**

**express.RemoveApiEventListener ( *eventName* )**

\*express是一个代表 ApiEvent 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                           |
|-------------|-----------|----------|------------------------------------------------------------------------------------------------|
| *eventName* | 必选      | String   | 代表事件类型，事件类型由wps内部规定，有关wps支持的事件类型，可以参数ApiEvent对象说明相关文档。 |

## 成员属性

### ApiEvent.Cancel

当事件是询问型事件时，此属性用于设置是否中止事件进行的状态

**语法**

**express.Cancel**

\*express是一个代表 ApiEvent 对象的变量。

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/事件/ApiEvent](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
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
# ToggleButton 对象

## 简介

创建一个切换开关的按钮。

## 说明

创建一个切换开关的按钮。

## 方法

| 序号 | 名称                       | 说明                       |
|:----:|:---------------------------|:---------------------------|
|  1   | [Move](#ToggleButton.Move) | 将对象移动到参数指定的位置 |

## 属性

| 序号 | 名称                                         | 说明                                                      |
|:----:|:---------------------------------------------|:----------------------------------------------------------|
|  1   | [AutoSize](#ToggleButton.AutoSize)           | 指定该对象是否根据内容自动设置它的大小                    |
|  2   | [BackColor](#ToggleButton.BackColor)         | 获取或设置指定对象的内部颜色                              |
|  3   | [BackStyle](#ToggleButton.BackStyle)         | 您可以使用该属性指定控件是否透明                          |
|  4   | [Caption](#ToggleButton.Caption)             | 获取或设置出现在控件中的文本                              |
|  5   | [ControlSource](#ToggleButton.ControlSource) | 您可以使用该属性指定控件中显示的数据                      |
|  6   | [Enabled](#ToggleButton.Enabled)             | 您可以通过该属性设置或返回是否启用该控件                  |
|  7   | [Font](#ToggleButton.Font)                   | 指定对象的字体，只读                                      |
|  8   | [ForeColor](#ToggleButton.ForeColor)         | 您可以通过该属性指定控件中文本的颜色                      |
|  9   | [Height](#ToggleButton.Height)               | 获取或设置指定对象的高度                                  |
|  10  | [HeightPolicy](#ToggleButton.HeightPolicy)   | 获取或设置指定对像的高度策略                              |
|  11  | [Left](#ToggleButton.Left)                   | 您可以通过该属性指定控件在对象或报表中的位置              |
|  12  | [Name](#ToggleButton.Name)                   | 您可以通过该属性指定或确定表示对象名称的字符串，只读      |
|  13  | [TabIndex](#ToggleButton.TabIndex)           | 您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置 |
|  14  | [TabStop](#ToggleButton.TabStop)             | 您可以通过该属性指定是否可以通过Tab键将焦点移到控件上     |
|  15  | [TextAlign](#ToggleButton.TextAlign)         | 该属性用于指定对象中文本的对齐方式                        |
|  16  | [Top](#ToggleButton.Top)                     | 您可以通过该属性指定控件在对象或报表中的位置              |
|  17  | [Value](#ToggleButton.Value)                 | 指定该按钮是否被按下。true是按下，false是弹起。           |
|  18  | [Visible](#ToggleButton.Visible)             | 返回或设置该对象是否可见                                  |
|  19  | [Width](#ToggleButton.Width)                 | 获取或设置指定对象的宽度                                  |
|  20  | [WidthPolic](#ToggleButton.WidthPolic)       | 获取或设置指定对象的宽度策略                              |
|  21  | [WordWrap](#ToggleButton.WordWrap)           | 指定控件的内容是否在行尾自动换行                          |

## 成员方法

### ToggleButton.Move

将对象移动到参数指定的位置

**语法**

**express.Move ( *Left* , *Top* , *Width* , *Height* )**

\*express是一个代表 ToggleButton 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                               |
|----------|-----------|----------|------------------------------------|
| *Left*   | 必选      | Number   | 对象移动后相对于窗口的左边缘的距离 |
| *Top*    | 可选      | Number   | 对象移动后相对于窗口的上边缘的距离 |
| *Width*  | 可选      | Number   | 对象移动后的预期宽度               |
| *Height* | 可选      | Number   | 对象移动后的预期长度               |

**说明**

将对象移动到参数指定的位置

**示例**

``` JavaScript
{#_examples}/*将ToggleButton1移动到窗体左上角，并将移动后的Label长宽设置为100*/
function func1()
{
    UserForm1.ToggleButton1.Move(100,100,100,100)
}
```

## 成员属性

### ToggleButton.AutoSize

指定该对象是否根据内容自动设置它的大小

**语法**

**express.AutoSize**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

指定该对象是否根据内容自动设置它的大小

**示例**

``` JavaScript
/*设置ToggleButton1控件为根据内容自动设置大小*/
function func1()
{
    UserForm1.ToggleButton1.AutoSize = true
}
```

### ToggleButton.BackColor

获取或设置指定对象的内部颜色

**语法**

**express.BackColor**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置指定对象的内部颜色

**示例**

``` JavaScript
/*将ToggleButton1设置为十六进制下665898表示的颜色（紫色）*/
function func1()
{
    UserForm1.ToggleButton1.BackColor = 0x665898
}
```

### ToggleButton.BackStyle

您可以使用该属性指定控件是否透明

**语法**

**express.BackStyle**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以使用该属性指定控件是否透明

**示例**

``` JavaScript
/*将ToggleButton1设置为透明*/
function func1()
{
    UserForm1.ToggleButton1.BackStyle = 0
}
```

### ToggleButton.Caption

获取或设置出现在控件中的文本

**语法**

**express.Caption**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置出现在控件中的文本

**示例**

``` JavaScript
/*将ToggleButton1中的文本设置为"ToggleButton1"*/
function func1()
{
    UserForm1.ToggleButton1.Caption = "ToggleButton1"
}
```

### ToggleButton.ControlSource

您可以使用该属性指定控件中显示的数据

**语法**

**express.ControlSource**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以使用该属性指定控件中显示的数据。您可以显示和编辑绑定到表的数据、查询的数据或SQL语句中的字段的数据。您还可以显示表达式的结果

### ToggleButton.Enabled

您可以通过该属性设置或返回是否启用该控件

**语法**

**express.Enabled**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性设置或返回是否启用该控件

**示例**

``` JavaScript
/*启用该控件*/
function func1()
{
    UserForm1.ToggleButton1.Enabled = true
}
```

### ToggleButton.Font

指定对象的字体，只读

**语法**

**express.Font**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

指定对象的字体，只读

### ToggleButton.ForeColor

您可以通过该属性指定控件中文本的颜色

**语法**

**express.ForeColor**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定控件中文本的颜色

**示例**

``` JavaScript
/*将ToggleButton1中的文本颜色设置为十六进制658978所表示的颜色*/
function func1()
{
    UserForm1.ToggleButton1.ForeColor = 0x658978
}
```

### ToggleButton.Height

获取或设置指定对象的高度

**语法**

**express.Height**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置指定对象的高度

**示例**

``` JavaScript
/*将ToggleButton1控件的高度设置为100*/
function func1()
{
    UserForm1.ToggleButton1.Height = 100
}
```

### ToggleButton.HeightPolicy

获取或设置指定对像的高度策略

**语法**

**express.HeightPolicy**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置指定对像的高度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变高度 |
| Fixed     | 2   | 固定高度 |

**示例**

``` JavaScript
/*将ToggleButton1的高度设置为固定*/
function func1()
{
    UserForm1.ToggleButton1.HeightPolicy = 2
}
```

### ToggleButton.Left

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Left**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ToggleButton1的位置设置到距离窗体左边缘100像素*/
function func1()
{
    UserForm1.ToggleButton1.Left = 100
}
```

### ToggleButton.Name

您可以通过该属性指定或确定表示对象名称的字符串，只读

**语法**

**express.Name**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定或确定表示对象名称的字符串，只读

### ToggleButton.TabIndex

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**语法**

**express.TabIndex**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定控件在窗体或报表上的Tab键顺序中的位置

**示例**

``` JavaScript
/*将ToggleButton1设置为Tab键顺序为的第2个位置*/
function func1()
{
    UserForm1.ToggleButton1.TabIndex = 2
}
```

### ToggleButton.TabStop

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**语法**

**express.TabStop**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定是否可以通过Tab键将焦点移到控件上

**示例**

``` JavaScript
/*将ToggleButton1设置为可以通过Tab键将焦点移到该控件上*/
function func1()
{
    UserForm1.ToggleButton1.TabStop = true
}
```

### ToggleButton.TextAlign

该属性用于指定对象中文本的对齐方式

**语法**

**express.TextAlign**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

该属性用于指定对象中文本的对齐方式，包含以下值：

| 设置   | 值  | 描述                     |
|--------|-----|--------------------------|
| Left   | 1   | 文本数字和日期都左对齐   |
| Center | 2   | 文本数字和日期都居中对齐 |
| Right  | 3   | 文本数字和日期都右对齐   |

**示例**

``` JavaScript
/*将ToggleButton1中的文本设置为居中对齐*/
function func1()
{
    UserForm1.ToggleButton1.TextAlign = 2
}
```

### ToggleButton.Top

您可以通过该属性指定控件在对象或报表中的位置

**语法**

**express.Top**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

您可以通过该属性指定控件在对象或报表中的位置

**示例**

``` JavaScript
/*将ToggleButton1的位置设置到距离窗体上边缘100像素*/
function func1()
{
    UserForm1.ToggleButton1.Top = 100
}
```

### ToggleButton.Value

指定该按钮是否被按下。true是按下，false是弹起。

**语法**

**express.Value**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

指定该按钮是否被按下。true是按下，false是弹起。

**示例**

``` JavaScript
/*将ToggleButton1设置为按下*/
function func1()
{
    UserForm1.ToggleButton1.Value = true
}
```

### ToggleButton.Visible

返回或设置该对象是否可见

**语法**

**express.Visible**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

返回或设置该对象是否可见

**示例**

``` JavaScript
/*将ToggleButton1设置为可见*/
function func1()
{
    UserForm1.ToggleButton1.Visible = true
}
```

### ToggleButton.Width

获取或设置指定对象的宽度

**语法**

**express.Width**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置指定对象的宽度

**示例**

``` JavaScript
/*将ToggleButton1的宽度设置为100*/
function func1()
{
    UserForm1.ToggleButton1.Width = 100
}
```

### ToggleButton.WidthPolic

获取或设置指定对象的宽度策略

**语法**

**express.WidthPolic**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

获取或设置指定对象的宽度策略，可设置为以下值：

| 设置      | 值  | 描述     |
|-----------|-----|----------|
| Expanding | 1   | 可变宽度 |
| Fixed     | 2   | 固定宽度 |

**示例**

``` JavaScript
/*将ToggleButton1的宽度设置为固定*/
function func1()
{
    UserForm1.ToggleButton1.WidthPolicy = 2
}
```

### ToggleButton.WordWrap

指定控件的内容是否在行尾自动换行

**语法**

**express.WordWrap**

\*express是一个代表 ToggleButton 对象的变量。

**说明**

指定控件的内容是否在行尾自动换行

**示例**

``` JavaScript
/*将ToggleButton1设置为控件的内容在行尾自动换行*/
function func1()
{
    UserForm1.ToggleButton1.WordWrap = true
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/ToggleButton](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CustomProperties 对象

## 简介

由代表附加信息的 CustomProperty 对象组成的集合，这些信息可用作 XML 的元数据。

## 说明

使用 Worksheet 对象的 CustomProperties 属性返回 CustomProperties 集合。

返回 CustomProperties 集合后，可根据选择向工作表和智能标记中添加元数据。

若要向工作表添加元数据，请在 Add 方法中使用 CustomProperties 属性。

下例演示了该功能。在此示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值。

``` JavaScript
function test() {
    let wksSheet1 = Application.ActiveSheet

    // Add metadata to worksheet.
    wksSheet1.CustomProperties.Add("Market", "Nasdaq")

    // Display metadata.
    let cusProperties = wksSheet1.CustomProperties.Item(1)
    alert(cusProperties.Name + "\t" + cusProperties.Value)
}
```

## 方法

| 序号 | 名称                           | 说明                                                                    |
|:----:|:-------------------------------|:------------------------------------------------------------------------|
|  1   | [Add](#CustomProperties.Add)   | 添加自定义属性信息，返回 一个代表自定义属性信息的 CustomProperty 对象。 |
|  2   | [Item](#CustomProperties.Item) | 从集合中返回一个对象。                                                  |

## 属性

| 序号 | 名称                                         | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#CustomProperties.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#CustomProperties.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#CustomProperties.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#CustomProperties.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### CustomProperties.Add

添加自定义属性信息，返回 一个代表自定义属性信息的 **CustomProperty** 对象。

**语法**

**express.Add ( *Name* , *Value* )**

\*express是一个代表 CustomProperties 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明               |
|---------|-----------|----------|--------------------|
| *Name*  | 必选      | String   | 自定义属性的名称。 |
| *Value* | 必选      | Variant  | 自定义属性的值。   |

**示例**

``` JavaScript
/* 此示例向活动工作表添加标识符信息，并将其名称和值返回给用户*/
function test() {

    let wksSheet1 = Application.ActiveSheet

    // Add metadata to worksheet.
    wksSheet1.CustomProperties.Add("Market", "Nasdaq")

    // Display metadata.
    let cusProperties = wksSheet1.CustomProperties.Item(1)
    alert(cusProperties.Name + "\t" + cusProperties.Value)

}
```

### CustomProperties.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 CustomProperties 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**示例**

``` JavaScript
/* 下例演示了该功能。在此示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值*/
function test() {

    let wksSheet1 = Application.ActiveSheet

    // Add metadata to worksheet.
    wksSheet1.CustomProperties.Add("Market", "Nasdaq")

    // Display metadata.
    let cusProperties = wksSheet1.CustomProperties.Item(1)
        alert(cusProperties.Name + "\t" + cusProperties.Value)

}
```

## 成员属性

### CustomProperties.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 CustomProperties 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
function test(){
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### CustomProperties.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 CustomProperties 对象的变量。

### CustomProperties.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 CustomProperties 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### CustomProperties.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 CustomProperties 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/CustomProperties](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Errors 对象

## 简介

------------------------------------------------------------------------

代表区域内的各种电子表格错误。

## 说明

使用 Range 对象的 Errors 属性可返回 Errors 对象。

返回 Errors 对象后，可使用 Error 对象的 Value 属性检查特定的错误检查条件。

``` JavaScript
/*下例将一个数字作为文本放在单元格 A1 中，然后当单元格 A1 的值包含文本格式的数字时通知用户。*/
function test(){
 //Place a number written as text in cell A1.
    Range("A1").Formula = "'1"

    if(Range("A1").Errors.Item(xlNumberAsText).Value == true) {
        alert("Cell A1 has a number as text.")
    }
    else {
        alert("Cell A1 is a number.")
    }

}
```

## 属性

| 序号 | 名称                               | 说明                                                                                                                                                                                                                            |
|:----:|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Errors.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#Errors.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Item](#Errors.Item)               | 返回 Error 对象的单个成员。                                                                                                                                                                                                     |
|  4   | [Parent](#Errors.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员属性

### Errors.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Errors 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### Errors.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Errors 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Errors.Item

返回 **Error** 对象的单个成员。

**语法**

**express.Item**

\*express是一个代表 Errors 对象的变量。

**说明**

*Index* 可以是以下常量之一。

|                                                                   |
|-------------------------------------------------------------------|
| **xlEvaluateToError** ：单元格计算为错误值。                      |
| **xlTextDate** ：单元格包含用 2 位数表示年份的文本日期。          |
| **xlNumberAsText** ：单元格包含以文本形式存储的数字。             |
| **xlInconsistentFormula** ：单元格包含一个区域中不一致的公式。    |
| **xlOmittedCells** ：单元格包含一个忽略了区域中某个单元格的公式。 |
| **xlUnlockedFormulaCells** ：取消锁定的单元格包含一个公式。       |
| **xlEmptyCellReferences** ：单元格包含一个引用空单元格的公式。    |

### Errors.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Errors 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Errors](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FormatCondition 对象

## 简介

代表条件格式。

## 说明

FormatCondition 对象是 FormatConditions 集合的成员。对于给定区域， FormatConditions 集合中包含的条件格式不能超过三个。

使用 A dd 方法可新建条件格式。如果区域内存在多种格式，则可使用 Modify 方法更改其中一种格式，或使用 Del ete 方法删除一种格式，然后使用 Add 方法创建一种新格式。

使用 FormatCondition 对象的 Font 、 Borders 和 Interior 属性可控制已设置格式的单元格的外观。条件格式对象模型不支持这些对象的某些属性。下表列出所有可与条件格式一起使用的属性。

<table style="font-size: 14px; width: 100%;">
<colgroup>
<col style="width: 50%" />
<col style="width: 50%" />
</colgroup>
<thead>
<tr class="header" style="vertical-align:middle;">
<th style="font-size: 14px">对象</th>
<th style="font-size: 14px">属性</th>
</tr>
</thead>
<tbody>
<tr class="odd" style="vertical-align:middle;">
<td style="padding-right: 8px; padding-left: 8px">Font</td>
<td style="padding-right: 8px; padding-left: 8px">Bold
<p>Color</p>
<p>ColorIndex</p>
<p>FontStyle</p>
<p>Italic</p>
<p>Strikethrough</p>
<p>Underline</p>
<p>无法使用会计用下划线样式。</p></td>
</tr>
<tr class="even" style="vertical-align:middle;">
<td style="padding-right: 8px; padding-left: 8px">Border</td>
<td style="padding-right: 8px; padding-left: 8px">Bottom
<p>Color</p>
<p>Left</p>
<p>Right</p>
<p>Style</p>
<p>可使用下列边框样式（其他均不可用）： xlNone 、 xlSolid 、 xlDash 、 xlDot 、 xlDashDot 、 xlDashDotDot 、 xlGray50 、 xlGray75 和 xlGray25 。</p>
<p>Top</p>
<p>Weight</p>
<p>可使用下列边框粗细（其他均不可用）： xlWeightHairline 和 xlWeightThin 。</p></td>
</tr>
<tr class="odd" style="vertical-align:middle;">
<td style="padding-right: 8px; padding-left: 8px">Interior</td>
<td style="padding-right: 8px; padding-left: 8px">Color
<p>ColorIndex</p>
<p>Pattern</p>
<p>PatternColorIndex</p></td>
</tr>
</tbody>
</table>

使用 FormatConditions ( *index* )（其中 *index* 为条件格式的索引号）可返回 FormatCondition 对象。下例为 E1:E10 单元格的现有条件格式设置格式属性。

``` JavaScript
function test() {
    let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1)
    let range3 = range2.Borders
    range3.LineStyle = xlContinuous
    range3.Weight = xlThin
    range3.ColorIndex = 6
    let range4 = range2.Font
    range4.Bold = true
    range4.ColorIndex = 3
}
```

## 方法

| 序号 | 名称                                                          | 说明                                                                              |
|:----:|:--------------------------------------------------------------|:----------------------------------------------------------------------------------|
|  1   | [Delete](#FormatCondition.Delete)                             | 删除对象。                                                                        |
|  2   | [Modify](#FormatCondition.Modify)                             | 更改现有条件格式。                                                                |
|  3   | [ModifyAppliesToRange](#FormatCondition.ModifyAppliesToRange) | 设置此格式规则所应用于的单元格区域。                                              |
|  4   | [SetFirstPriority](#FormatCondition.SetFirstPriority)         | 将此条件格式规则的优先级值设置为“1”，以便在工作表上的所有其他规则之前计算此规则。 |
|  5   | [SetLastPriority](#FormatCondition.SetLastPriority)           | 为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则。        |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FormatCondition.Application)   | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [AppliesTo](#FormatCondition.AppliesTo)       | 返回一个 Range 对象，指定格式规则将应用于的单元格区域。                                                                                                                                                                         |
|  3   | [Borders](#FormatCondition.Borders)           | 返回一个 Borders 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。                                                                                                                                         |
|  4   | [Creator](#FormatCondition.Creator)           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  5   | [DateOperator](#FormatCondition.DateOperator) | 指定格式条件中所用的日期运算符。可读/写。                                                                                                                                                                                       |
|  6   | [Font](#FormatCondition.Font)                 | 返回一个 Font 对象，它代表指定对象的字体。                                                                                                                                                                                      |
|  7   | [Formula1](#FormatCondition.Formula1)         | 返回与条件格式或者数据有效性相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。 String 型，只读。                                                                                                                      |
|  8   | [Formula2](#FormatCondition.Formula2)         | 返回与条件格式或数据有效性验证第二部分相关联的值或表达式。仅用于数据有效性条件格式 Operator 属性为 xlBetween 或 xlNotBetween 的情况。可为常量值、字符串值、单元格引用或公式。 String 类型，只读。                               |
|  9   | [Interior](#FormatCondition.Interior)         | 返回一个 Interior 对象，它代表指定对象的内部。                                                                                                                                                                                  |
|  10  | [NumberFormat](#FormatCondition.NumberFormat) | 在条件格式规则的计算结果为 True 时返回或设置应用于单元格的数字格式。 Variant 型，可读写。                                                                                                                                       |
|  11  | [Operator](#FormatCondition.Operator)         | 返回一个 Long 值，它代表条件格式的操作符。                                                                                                                                                                                      |
|  12  | [PTCondition](#FormatCondition.PTCondition)   | 返回一个 Boolean 值，指明是否将条件格式应用于数据透视表。只读。                                                                                                                                                                 |
|  13  | [Parent](#FormatCondition.Parent)             | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  14  | [Priority](#FormatCondition.Priority)         | 返回或设置条件格式规则的优先级值。当工作表中存在多个条件格式规则时，优先级确定求值的顺序。                                                                                                                                      |
|  15  | [ScopeType](#FormatCondition.ScopeType)       | 返回或设置 XlPivotConditionScope 枚举的常量之一，该常量确定条件格式在应用于数据透视表图表时的范围。                                                                                                                             |
|  16  | [StopIfTrue](#FormatCondition.StopIfTrue)     | 返回或设置一个 Boolean 值，该值确定在当前规则的计算结果为 True 时是否应计算单元格上的其他格式规则。                                                                                                                             |
|  17  | [Text](#FormatCondition.Text)                 | 返回或设置一个 String 类型的值，该值指定条件格式规则所用的文本字符串。                                                                                                                                                          |
|  18  | [TextOperator](#FormatCondition.TextOperator) | 返回或设置一个 XlContainsOperator 枚举常量，该常量指定条件格式规则执行的文本搜索。                                                                                                                                              |
|  19  | [Type](#FormatCondition.Type)                 | 返回一个包含 xlFormatConditionType 值的 Long 值，它代表对象类型。                                                                                                                                                               |

## 成员方法

### FormatCondition.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 FormatCondition 对象的变量。

### FormatCondition.Modify

更改现有条件格式。

**语法**

**express.Modify ( *Type* , *Operator* , *Formula1* , *Formula2* )**

\*express是一个代表 FormatCondition 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型              | 说明                                                                                                 |
|------------|-----------|-----------------------|------------------------------------------------------------------------------------------------------|
| *Type*     | 必选      | XlFormatConditionType | 指定条件格式是基于单元格值还是基于表达式。                                                           |
| *Operator* | 可选      | Variant               | 一个 XlFormatConditionOperator 值，它代表条件格式运算符。如果 Type 设为 xlExpression，则忽略该参数。 |
| *Formula1* | 可选      | Variant               | 与条件格式相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。                               |
| *Formula2* | 可选      | Variant               | 与条件格式相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。                               |

**示例**

``` JavaScript
/* 本示例修改单元格区域 E1:E10 的现有条件格式*/
Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1).Modify(xlCellValue, xlLess, "=$a$1")
```

### FormatCondition.ModifyAppliesToRange

设置此格式规则所应用于的单元格区域。

**语法**

**express.ModifyAppliesToRange ( *Range* )**

\*express是一个代表 FormatCondition 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                       |
|---------|-----------|----------|----------------------------|
| *Range* | 必选      | Range    | 此格式规则将应用于的区域。 |

**说明**

该区域必须采用 A1 引用样式，并全部包含在 FormatConditions 集合的父工作表内。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可以使用货币符号，但会被忽略。

您也可以在区域的任意部分中使用局部定义名称，但该名称必须使用宏语言。

### FormatCondition.SetFirstPriority

将此条件格式规则的优先级值设置为“1”，以便在工作表上的所有其他规则之前计算此规则。

**语法**

**express.SetFirstPriority ()**

\*express是一个代表 FormatCondition 对象的变量。

**说明**

当工作表中有多个条件格式规则时，如果之前未将此规则设置为优先级“1”，此方法将导致工作表上的所有其他现有规则的优先级都增加 1。

| 注释                                     |
|------------------------------------------|
| 将在工作表级别应用条件格式规则的优先级。 |

### FormatCondition.SetLastPriority

为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则。

**语法**

**express.SetLastPriority ()**

\*express是一个代表 FormatCondition 对象的变量。

**说明**

优先级的实际值将等于工作表上条件格式规则的总数。当工作表中有多个条件格式规则时，此方法将导致优先级值大于此规则的规则的优先级增加 1。

| 注释                                     |
|------------------------------------------|
| 将在工作表级别应用条件格式规则的优先级。 |

## 成员属性

### FormatCondition.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 FormatCondition 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息
let myObject = Application.ActiveWorkbook
if(myObject.Application.Value == "ET") {
    alert("This is an ET Application object.")
}
else {
    alert("This is not an ET Application object.")
}
```

### FormatCondition.AppliesTo

返回一个 **Range** 对象，指定格式规则将应用于的单元格区域。

**语法**

**express.AppliesTo**

\*express是一个代表 FormatCondition 对象的变量。

### FormatCondition.Borders

返回一个 **Borders** 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。

**语法**

**express.Borders**

\*express是一个代表 FormatCondition 对象的变量。

**示例**

``` JavaScript
/* 此示例将 Sheet1 中单元格 B2 的底部边框颜色设置为红色细边框*/
function SetRangeBorder() {
    
    let range2 = Application.Worksheets.Item("Sheet1").Range("B2").Borders.Item(xlEdgeBottom)
    range2.LineStyle = xlContinuous
    range2.Weight = xlThin
    range2.ColorIndex = 3

}
```

### FormatCondition.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FormatCondition 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### FormatCondition.DateOperator

指定格式条件中所用的日期运算符。可读/写。

**语法**

**express.DateOperator**

\*express是一个代表 FormatCondition 对象的变量。

### FormatCondition.Font

返回一个 **Font** 对象，它代表指定对象的字体。

**语法**

**express.Font**

\*express是一个代表 FormatCondition 对象的变量。

### FormatCondition.Formula1

返回与条件格式或者数据有效性相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。 **String** 型，只读。

**语法**

**express.Formula1**

\*express是一个代表 FormatCondition 对象的变量。

**示例**

``` JavaScript
/* 如果单元格区域 E1:E10 的第一个条件格式的公式指定的是“小于 5”，则此示例将对其进行更改*/
let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1)
    if(range2.Operator == xlLess && range2.Formula1 == "5") {
        range2.Modify(xlCellValue, xlLess, "10")
    }
```

### FormatCondition.Formula2

返回与条件格式或数据有效性验证第二部分相关联的值或表达式。仅用于数据有效性条件格式 **Operator** 属性为 **xlBetween** 或 **xlNotBetween** 的情况。可为常量值、字符串值、单元格引用或公式。 **String** 类型，只读。

**语法**

**express.Formula2**

\*express是一个代表 FormatCondition 对象的变量。

**示例**

``` JavaScript
/* 如果单元格区域 E1:E10 的第一个条件格式的公式指定的是“介于 5 和 10 之间”，则本示例将对其进行更改*/
let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1)
    if(range2.Operator == xlBetween && range2.Formula1 == "5" && range2.Formula2 == "10") {
        range2.Modify(xlCellValue, xlLess, "10")
    }
```

### FormatCondition.Interior

返回一个 **Interior** 对象，它代表指定对象的内部。

**语法**

**express.Interior**

\*express是一个代表 FormatCondition 对象的变量。

### FormatCondition.NumberFormat

在条件格式规则的计算结果为 **True** 时返回或设置应用于单元格的数字格式。 **Variant** 型，可读写。

**语法**

**express.NumberFormat**

\*express是一个代表 FormatCondition 对象的变量。

**说明**

数字格式是使用 **“单元格格式”** 对话框的 **“数字”** 选项卡上显示的相同格式代码指定的。您可以使用内置的数字格式，例如 ` "General" ` 或者 <span style="color:#000000;text-decoration-line:none;font-family:&quot;white-space:normal;background-color:#FFFFFF;">创建自定义数字格式</span> 。

### FormatCondition.Operator

返回一个 **Long** 值，它代表条件格式的操作符。

**语法**

**express.Operator**

\*express是一个代表 FormatCondition 对象的变量。

**示例**

``` JavaScript
/* 如果单元格区域 E1:E10 的条件格式一的公式指定的是“小于 5”，则此示例将对其进行更改*/
let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1)
    if(range2.Operator == xlLess && range2.Formula1 == "5") {
        range2.Modify(xlCellValue, xlBetween, "5", "15")
    }
```

### FormatCondition.PTCondition

返回一个 **Boolean** 值，指明是否将条件格式应用于数据透视表。只读。

**语法**

**express.PTCondition**

\*express是一个代表 FormatCondition 对象的变量。

### FormatCondition.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FormatCondition 对象的变量。

### FormatCondition.Priority

返回或设置条件格式规则的优先级值。当工作表中存在多个条件格式规则时，优先级确定求值的顺序。

**语法**

**express.Priority**

\*express是一个代表 FormatCondition 对象的变量。

**说明**

设置优先级时，值必须为介于 1 和工作表上条件格式规则总数之间的正整数。对于工作表上的所有规则，优先级必须为唯一值，因此，如果更改指定条件格式规则的优先级，将可能会导致工作表上其他规则的优先级值移位。

### FormatCondition.ScopeType

返回或设置 **XlPivotConditionScope** 枚举的常量之一，该常量确定条件格式在应用于数据透视表图表时的范围。

**语法**

**express.ScopeType**

\*express是一个代表 FormatCondition 对象的变量。

**说明**

默认值为 **xlSelectionScope** ，它使用 **AppliesTo** 属性设置范围。

### FormatCondition.StopIfTrue

返回或设置一个 **Boolean** 值，该值确定在当前规则的计算结果为 **True** 时是否应计算单元格上的其他格式规则。

**语法**

**express.StopIfTrue**

\*express是一个代表 FormatCondition 对象的变量。

**说明**

为了支持向后兼容性，此属性的默认值为 **True** ，而这与在用户界面中创建的格式规则的行为相反。

### FormatCondition.Text

返回或设置一个 **String** 类型的值，该值指定条件格式规则所用的文本字符串。

**语法**

**express.Text**

\*express是一个代表 FormatCondition 对象的变量。

**说明**

如果 **Type** 属性不是设置为 **xlTextString** ，则忽略此属性。

### FormatCondition.TextOperator

返回或设置一个 **XlContainsOperator** 枚举常量，该常量指定条件格式规则执行的文本搜索。

**语法**

**express.TextOperator**

\*express是一个代表 FormatCondition 对象的变量。

### FormatCondition.Type

返回一个包含 **xlFormatConditionType** 值的 **Long** 值，它代表对象类型。

**语法**

**express.Type**

\*express是一个代表 FormatCondition 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/FormatCondition](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# HeaderFooter 对象

## 简介

代表单一页眉或页脚。 HeaderFooter 对象是 HeadersFooters 集合的一个成员。 HeadersFooters 集合包含指定工作簿节中的所有页眉和页脚。

## 说明

您也可以使用 HeaderFooter 属性和 Selection 对象来返回单个 HeaderFooter 对象。

| 注释                                                 |
|------------------------------------------------------|
| 不能将 HeaderFooter 对象添加到 HeadersFooters 集合。 |

使用 PageSetup 对象的 DifferentFirstPageHeaderFooter 属性可指定不同的首页。

使用 PageNumbers 对象的 Add 方法可在页眉或页脚中添加页码。以下示例在活动工作簿第一节的主页脚中添加页码。

``` JavaScript
function test() {
    let sheet1 = Application.ActiveSheet.PageSetup
    sheet1.CenterHeader = "&D&T"
    sheet1.OddAndEvenPagesHeaderFooter = false
    sheet1.DifferentFirstPageHeaderFooter = false
    sheet1.ScaleWithDocHeaderFooter = true
    sheet1.AlignMarginsHeaderFooter = true
} 
```

## 属性

| 序号 | 名称                             | 说明                                                                      |
|:----:|:---------------------------------|:--------------------------------------------------------------------------|
|  1   | [Picture](#HeaderFooter.Picture) | 返回一个 Picture 对象，该对象代表指定页眉或页脚中包含的图片字段。只读。   |
|  2   | [Text](#HeaderFooter.Text)       | 返回或设置一个 Text 对象，该对象代表指定页眉或页脚中包含的文本。可读/写。 |

## 成员属性

### HeaderFooter.Picture

返回一个 **Picture** 对象，该对象代表指定页眉或页脚中包含的图片字段。只读。

**语法**

**express.Picture**

\*express是一个代表 HeaderFooter 对象的变量。

### HeaderFooter.Text

返回或设置一个 **Text** 对象，该对象代表指定页眉或页脚中包含的文本。可读/写。

**语法**

**express.Text**

\*express是一个代表 HeaderFooter 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/HeaderFooter](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Legend 对象

## 简介

代表图标中的图例。每个图表只能有一个图例。

## 说明

Legend 对象包含一个或多个 LegendEntry 对象；每个 LegendEntry 对象都包含一个 LegendKey 对象。

除非 HasLegend 属性是 True ，否则无法看到图表的图例。如果该属性为 False ， Legend 对象的属性和方法将会失败。

``` JavaScript
/*使用 Legend 属性可返回 Legend 对象。下例将工作表一上嵌入图表一的图例字形设置为加粗。*/
function test(){
Application.Worksheets.Item(1).ChartObjects(1).Chart.Legend.Font.Bold = true
}
```

## 方法

| 序号 | 名称                                   | 说明                                                                                       |
|:----:|:---------------------------------------|:-------------------------------------------------------------------------------------------|
|  1   | [Clear](#Legend.Clear)                 | 清除整个对象。                                                                             |
|  2   | [Delete](#Legend.Delete)               | 删除对象。                                                                                 |
|  3   | [LegendEntries](#Legend.LegendEntries) | 返回表示图例中的单个图例项（ LegendEntry 对象）或图例项集合（ LegendEntries 对象）的对象。 |
|  4   | [Select](#Legend.Select)               | 选择对象。                                                                                 |

## 属性

| 序号 | 名称                                       | 说明                                                                                                                                                                                                                            |
|:----:|:-------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Legend.Application)         | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#Legend.Creator)                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Format](#Legend.Format)                   | 返回 ChartFormat 对象。只读。                                                                                                                                                                                                   |
|  4   | [Heigh](#Legend.Heigh)                     | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |
|  5   | [IncludeInLayout](#Legend.IncludeInLayout) | 如果在确定图表布局时图例将占用图表布局空间，则为 True 。默认值是 True 。 Boolean 类型，可读写。                                                                                                                                 |
|  6   | [Left](#Legend.Left)                       | 返回或设置 Double 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。                                                                                                                       |
|  7   | [Name](#Legend.Name)                       | 返回一个 String 值，它代表对象的名称。                                                                                                                                                                                          |
|  8   | [Parent](#Legend.Parent)                   | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  9   | [Position](#Legend.Position)               | 返回或设置一个 XlLegendPosition 值，它代表图表上图例的位置。                                                                                                                                                                    |
|  10  | [Shadow](#Legend.Shadow)                   | 返回或设置一个 Boolean 值，它确定对象是否有阴影。                                                                                                                                                                               |
|  11  | [Top](#Legend.Top)                         | 返回或设置一个 Double 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。                                                                                                                      |
|  12  | [Width](#Legend.Width)                     | 返回或设置一个 Double 值，它代表对象的宽度（以磅为单位）。                                                                                                                                                                      |

## 成员方法

### Legend.Clear

清除整个对象。

**语法**

**express.Clear ()**

\*express是一个代表 Legend 对象的变量。

**返回值**

Variant

### Legend.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 Legend 对象的变量。

**返回值**

Variant

### Legend.LegendEntries

返回表示图例中的单个图例项（ **LegendEntry** 对象）或图例项集合（ **LegendEntries** 对象）的对象。

**语法**

**express.LegendEntries ( *Index* )**

\*express是一个代表 Legend 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明               |
|---------|-----------|----------|--------------------|
| *Index* | 可选      | Variant  | 图例输入项的编号。 |

**示例**

``` JavaScript
/*本示例设置 Chart1 上图例输入项一的字体。*/
Application.Charts.Item("Chart1").Legend.LegendEntries(1).Font.Name = "Arial"
```

### Legend.Select

选择对象。

**语法**

**express.Select ()**

\*express是一个代表 Legend 对象的变量。

**返回值**

Variant

## 成员属性

### Legend.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Legend 对象的变量。

**示例**

``` JavaScript
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### Legend.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Legend 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Legend.Format

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

\*express是一个代表 Legend 对象的变量。

### Legend.Heigh

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Heigh**

\*express是一个代表 Legend 对象的变量。

### Legend.IncludeInLayout

如果在确定图表布局时图例将占用图表布局空间，则为 **True** 。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.IncludeInLayout**

\*express是一个代表 Legend 对象的变量。

**说明**

此属性对于图表是否处于自动版式模式无影响。如果用户使用 **“图表上方”** 命令添加标题，则图表将变得较小。如果用户随后删除标题，或者选择一个覆盖标题选项，则图表将变得较大，就好像标题不在图表上一样。

### Legend.Left

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 Legend 对象的变量。

### Legend.Name

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 Legend 对象的变量。

### Legend.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Legend 对象的变量。

### Legend.Position

返回或设置一个 **XlLegendPosition** 值，它代表图表上图例的位置。

**语法**

**express.Position**

\*express是一个代表 Legend 对象的变量。

**示例**

``` JavaScript
/*本示例将图表中的图例移到图表的底部。*/
Application.Charts.Item(1).Legend.Position = xlLegendPositionBottom
```

### Legend.Shadow

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

\*express是一个代表 Legend 对象的变量。

### Legend.Top

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 Legend 对象的变量。

### Legend.Width

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 Legend 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Legend](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# LegendEntries 对象

## 简介

指定的图表图例中所有 LegendEntry 对象的集合。

## 说明

每个图例项都有两部分：一部分是该项的文本，它是与该图例项相关联的数据系列或趋势线的名称；另一部分是项标志，它在图表中以直观的方式将图例项以及与之相关联的数据系列或趋势线链接起来。项标志及其相关系列或趋势线的格式属性包含在 LegendKey 对象中。

``` JavaScript
/*使用 LegendEntries 方法可返回 LegendEntries 集合。下例在嵌入式图表一中的图例项集合中循环，并更改这些图例项的字体颜色。*/
function test(){
let lgd = Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.Legend
for(let i = 
 1;  i < = lgd.LegendEntries.Count; i++){
    lgd.LegendEntries(i).Font.ColorIndex = 5
}
}
```

使用 LegendEntries ( *index* )（其中 *index* 是图例项索引号）可返回一个 LegendEntry 对象。不能按名称返回图例项。

图例项的编号代表图例项在图例中的位置。 ` LegendEntries(1) ` 位于图例的顶部； ` LegendEntries(LegendEntries.Count) ` 位于图例的底部。

``` JavaScript
/*下例将嵌入的第一个图表中图例顶部的图例项（这通常是第一个数据系列的图例）的文字字体设置为斜体。*/
Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.Legend.LegendEntries(1).Font.Italic = true
```

## 方法

| 序号 | 名称                        | 说明                   |
|:----:|:----------------------------|:-----------------------|
|  1   | [Item](#LegendEntries.Item) | 从集合中返回一个对象。 |

## 属性

| 序号 | 名称                                      | 说明                                                                                                                                                                                                                            |
|:----:|:------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#LegendEntries.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#LegendEntries.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#LegendEntries.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#LegendEntries.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### LegendEntries.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 LegendEntries 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *Index* | 必选      | Variant  | 对象的索引号。 |

**返回值**

包含在集合中的一个 LegendEntry 对象。

**示例**

``` JavaScript
/*此示例对 Sheet1 中第一张嵌入式图表的图例（通常是第一个数据系列的图例）顶部的图例项的文本字体进行更改。*/
function test(){
Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.Legend.LegendEntries.(1).Font.Italic = true
}
```

## 成员属性

### LegendEntries.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 LegendEntries 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### LegendEntries.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 LegendEntries 对象的变量。

### LegendEntries.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 LegendEntries 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### LegendEntries.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 LegendEntries 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/LegendEntries](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# LinkFormat 对象

## 简介

包含链接 OLE 对象的属性。

## 说明

如果指定 Shape 对象不代表一个链接对象，则 LinkFormat 属性会失败。

``` JavaScript
/*使用 LinkFormat 属性可返回 LinkFormat 对象。下例更新 Shapes 集合中的一个 OLE 对象。*/
Application.Worksheets.Item(1).Shapes.Item(1).LinkFormat.Update()
```

## 方法

| 序号 | 名称                         | 说明       |
|:----:|:-----------------------------|:-----------|
|  1   | [Update](#LinkFormat.Update) | 更新链接。 |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                                                                                            |
|:----:|:---------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#LinkFormat.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [AutoUpdate](#LinkFormat.AutoUpdate)   | 如果源更改时 LinkFormat 对象自动进行更新，则此属性为 True 。 Boolean 类型，只读。                                                                                                                                               |
|  3   | [Creator](#LinkFormat.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Locked](#LinkFormat.Locked)           | 返回或设置一个 Boolean 值，它指明对象是否已被锁定。                                                                                                                                                                             |
|  5   | [Parent](#LinkFormat.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### LinkFormat.Update

更新链接。

**语法**

**express.Update ()**

\*express是一个代表 LinkFormat 对象的变量。

## 成员属性

### LinkFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 LinkFormat 对象的变量。

**说明**

``` JavaScript
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### LinkFormat.AutoUpdate

如果源更改时 **LinkFormat** 对象自动进行更新，则此属性为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.AutoUpdate**

\*express是一个代表 LinkFormat 对象的变量。

### LinkFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 LinkFormat 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### LinkFormat.Locked

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

\*express是一个代表 LinkFormat 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True** ；如果在工作表处于受保护状态时仍能修改对象，则返回 **False** 。

### LinkFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 LinkFormat 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/LinkFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# OLEFormat 对象

## 简介

包含 OLE 对象的属性。

## 说明

如果 Shape 对象不代表某个链接的或嵌入的对象，则 OLEFormat 属性失败。

## 方法

| 序号 | 名称                            | 说明                              |
|:----:|:--------------------------------|:----------------------------------|
|  1   | [Activate](#OLEFormat.Activate) | 使当前图表成为活动图表。          |
|  2   | [Verb](#OLEFormat.Verb)         | 向指定的 OLE 对象服务器发送动词。 |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#OLEFormat.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Creator](#OLEFormat.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  3   | [Object](#OLEFormat.Object)           | 返回与此 OLE 对象相联系的 OLE 自动化对象。 Object 型，只读。                                                                                                                                                                    |
|  4   | [Parent](#OLEFormat.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  5   | [progID](#OLEFormat.progID)           | 返回对象的程序标识符。 String 型，只读。                                                                                                                                                                                        |

## 成员方法

### OLEFormat.Activate

使当前图表成为活动图表。

**语法**

**express.Activate ()**

\*express是一个代表 OLEFormat 对象的变量。

### OLEFormat.Verb

向指定的 OLE 对象服务器发送动词。

**语法**

**express.Verb ( *Verb* )**

\*express是一个代表 OLEFormat 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                                                                                                                                                                     |
|--------|-----------|-----------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Verb* | 可选      | XlOLEVerb | OLE 对象服务器将执行其操作的动词。如果省略此参数，则发送默认动词。对象的源应用程序决定哪些动词可用。OLE 对象的典型动词为 Open 和 Primary（由 XlOLEVerb 常量 xlOpen 和 xlPrimary 表示）。 |

## 成员属性

### OLEFormat.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
　　　　let myObject = ActiveWorkbook
    if (myObject.Application.Value == "ET"){
　　　　    MsgBox ("This is an ET Application object.")
　　　　}
　　　　else{}
　　　　    MsgBox ("This is not an ET Application object.")
　　　　}}
```

### OLEFormat.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 OLEFormat 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### OLEFormat.Object

返回与此 OLE 对象相联系的 OLE 自动化对象。 **Object** 型，只读。

**语法**

**express.Object**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

``` JavaScript
/*此示例在 sheet1 的内嵌 WPS 文档对象的开始处插入文字。注意，在 With 控制结构内的三条语句为 WordBasic 语句。*/
function test(){
　　　　let wordObj = Worksheets.Item("Sheet1").OLEObjects(1)
　　　　wordObj.Activate()
　　　　let wbc = wordObj.Object.Application.WordBasic
　　　　wbc.StartOfDocument
　　　　wbc.Insert ("This is the beginning")
　　　　wbc.InsertPara
}
```

### OLEFormat.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

### OLEFormat.progID

返回对象的程序标识符。 **String** 型，只读。

**语法**

**express.progID**

\*express是一个代表 OLEFormat 对象的变量。

**示例**

``` JavaScript
/*此示例为第一张工作表中所有 OLE 对象创建程序标识符列表。*/
function test(){
　　　　let rw = 0
　　　　for (let o = 1; o <= Worksheets.Item(1).OLEObjects.Count; o++){
  　　　　  rw = rw + 1
 　　　　   Worksheets.Item(2).cells.Item(rw, 1).Value2 = Worksheets.Item(1).OLEObjects(o).ProgId
　　　　}
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/OLEFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Outline 对象

## 简介

代表工作表上的分级显示。

## 方法

| 序号 | 名称                              | 说明                                      |
|:----:|:----------------------------------|:------------------------------------------|
|  1   | [ShowLevels](#Outline.ShowLevels) | 显示指定行号和/或列号在分级显示中的层次。 |

## 属性

| 序号 | 名称                                        | 说明                                                                                                                                                                                                                            |
|:----:|:--------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Outline.Application)         | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [AutomaticStyles](#Outline.AutomaticStyles) | 如果分级显示使用自动样式，则该属性值为 True 。 Boolean 类型，可读写。                                                                                                                                                           |
|  3   | [Creator](#Outline.Creator)                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#Outline.Parent)                   | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |
|  5   | [SummaryColumn](#Outline.SummaryColumn)     | 返回或设置分级显示中的汇总列的位置。 XlSummaryColumn 类型，可读写。                                                                                                                                                             |
|  6   | [SummaryRow](#Outline.SummaryRow)           | 返回或设置分级显示中汇总行的位置。 XlSummaryRow 类型，可读写。                                                                                                                                                                  |

## 成员方法

### Outline.ShowLevels

显示指定行号和/或列号在分级显示中的层次。

**语法**

**express.ShowLevels ( *RowLevels* , *ColumnLevels* )**

\*express是一个代表 Outline 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                                                                                           |
|----------------|-----------|----------|------------------------------------------------------------------------------------------------------------------------------------------------|
| *RowLevels*    | 可选      | Variant  | 指定分级显示中的行层次。如果该分级显示包含的层次数少于指定层次，则 ET 显示所有层次。如果该参数为 0（零）或者省略该参数，则不对行采取任何操作。 |
| *ColumnLevels* | 可选      | Variant  | 指定分级显示中的列层次。如果该分级显示包含的层次数少于指定层次，则 ET 显示所有层次。如果该参数为 0（零）或者省略该参数，则不对列采取任何操作。 |

**返回值**

Variant

**说明**

必须至少指定一个参数。

**示例**

本示例对工作表 Sheet1 的第一行到第三行的行层次和第一列的列层次进行分级显示。

``` JavaScript
Worksheets.Item("Sheet1").Outline.ShowLevels (3, 1)
```

## 成员属性

### Outline.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Outline 对象的变量。

**示例**

``` JavaScript
/*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test(){
　　　　let myObject = ActiveWorkbook
       if (myObject.Application.Value == "ET"){
　　　　    MsgBox ("This is an ET Application object.")
　　　　}
　　　　else{
　　　　    MsgBox ("This is not an ET Application object.")
　　　　}
}
```

### Outline.AutomaticStyles

如果分级显示使用自动样式，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.AutomaticStyles**

\*express是一个代表 Outline 对象的变量。

**示例**

``` JavaScript
/*本示例将 Sheet1 的分级显示设置为使用自动样式。*/
function test(){
　　　　Worksheets.Item("Sheet1").Outline.AutomaticStyles = true
}
```

### Outline.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Outline 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Outline.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Outline 对象的变量。

### Outline.SummaryColumn

返回或设置分级显示中的汇总列的位置。 **XlSummaryColumn** 类型，可读写。

**语法**

**express.SummaryColumn**

\*express是一个代表 Outline 对象的变量。

**说明**

| XlSummaryColumn 可为以下 常量之一。                       |
|-----------------------------------------------------------|
| xlSummaryOnRight 分级显示中的汇总列位于明细数据列的右侧。 |
| xlSummaryOnLeft 分级显示中的汇总列位于明细数据列的左侧。  |

**示例**

``` JavaScript
/*本示例创建一个使用自动样式的分级显示，汇总行位于明细数据行的上方，汇总列位于明细数据列的右侧。*/
function test(){
　　　　Worksheets.Item("Sheet1").Activate()
　　　　Selection.AutoOutline()
　　　　let ole = ActiveSheet.Outline
　　　　ole.SummaryRow = xlAbove
　　　　ole.SummaryColumn = xlRight
　　　　ole.AutomaticStyles = true
}
```

### Outline.SummaryRow

返回或设置分级显示中汇总行的位置。 **XlSummaryRow** 类型，可读写。

**语法**

**express.SummaryRow**

\*express是一个代表 Outline 对象的变量。

**说明**

| XlSummaryRow 可为以下常量之一。                         |
|---------------------------------------------------------|
| xlSummaryBelow 分级显示中的汇总行位于明细数据行的下方。 |
| xlSummaryAbove 分级显示中的汇总行位于明细数据行的上方。 |

若要进行 WPS 风格的分级显示，请将 **XlSummaryRow** 设置为 **xlAbove** ，使分类标题位于明细信息的上方。若要进行会计风格的分级显示，请将 **XlSummaryRow** 设置为 **xlBelow** ，使求和位于明细信息的下方。

**示例**

``` JavaScript
/*本示例创建一个使用自动样式的分级显示，汇总行位于明细数据行的上方，汇总列位于明细数据列的右侧。*/
function test(){
　　　　Worksheets.Item("Sheet1").Activate()
　　　　Selection.AutoOutline()
　　　　let ole = ActiveSheet.Outline
　　　　ole.SummaryRow = xlAbove
　　　　ole.SummaryColumn = xlRight
　　　　ole.AutomaticStyles = true
}
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Outline](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotField 对象

## 简介

代表数据透视表中的一个字段。

## 说明

PivotField 对象是 PivotFields 集合的成员。 PivotFields 集合包含数据透视表中的所有字段，包括隐藏字段。

下列属性返回数据透视表字段的子集，在某些情况下，使用这些属性会更为方便：

- ColumnFields 属性
- DataFields 属性
- HiddenFields 属性
- PageFields 属性
- RowFields 属性
- VisibleFields 属性

使用 PivotFields ( *index* )（其中 *index* 是字段名称或索引号）可返回一个 PivotField 对象。下例使 Sheet3 上第一张数据透视表中的字段“Year”成为行字段。

``` JavaScript
Application.Worksheets.Item("sheet3").PivotTables(1).PivotFields("year").Orientation = xlRowField
```

## 方法

| 序号 | 名称                                               | 说明                                                                                                                                                                                           |
|:----:|:---------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddPageItem](#PivotField.AddPageItem)             | 向具有多个项的页面字段添加其他项。                                                                                                                                                             |
|  2   | [AutoShow](#PivotField.AutoShow)                   | 显示指定数据透视表中行字段、页字段或列字段顶部或底部数据项的个数。                                                                                                                             |
|  3   | [AutoSort](#PivotField.AutoSort)                   | 为数据透视表建立自动字段排序规则。                                                                                                                                                             |
|  4   | [CalculatedItems](#PivotField.CalculatedItems)     | 返回一个 CalculatedItems 集合，该集合表示指定数据透视表中的所有计算项。只读。                                                                                                                  |
|  5   | [ClearAllFilters](#PivotField.ClearAllFilters)     | 调用此方法将删除当前应用于透视字段的所有筛选。包括从透视字段的 PivotFilters 集合中删除所有筛选以及删除应用于透视字段的所有手动筛选。如果透视字段位于报表筛选区域中，则所选项将被设置为默认项。 |
|  6   | [ClearLabelFilters](#PivotField.ClearLabelFilters) | 此方法将删除透视字段的 PivotFilters 集合中的所有标签筛选或所有日期筛选。                                                                                                                       |
|  7   | [ClearManualFilter](#PivotField.ClearManualFilter) | 提供一个简便方法，可将数据透视表中透视字段的所有项的 Visible 属性设置为 True ，并清空 OLAP 数据透视表中的 HiddenItemsList 和 VisibleItemsList 集合。                                           |
|  8   | [ClearValueFilters](#PivotField.ClearValueFilters) | 调用此方法将删除透视字段的 PivotFilters 集合中的所有值筛选。                                                                                                                                   |
|  9   | [Delete](#PivotField.Delete)                       | 删除对象。                                                                                                                                                                                     |
|  10  | [DrillTo](#PivotField.DrillTo)                     | DrillTo 方法支持从另一个透视字段深化到指定的透视字段。                                                                                                                                         |
|  11  | [PivotItems](#PivotField.PivotItems)               | 返回一个对象，该对象表示指定字段中的单个数据透视表项（ PivotItem 对象）或所有可见项和隐藏项的集合（ PivotItems 对象）。只读。                                                                  |
|  12  | [RepeatAllLabels](#PivotField.RepeatAllLabels)     | 指定是否要对指定数据透视表中的所有透视字段重复项目标签。                                                                                                                                       |

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.AllItemsVisible">AllItemsVisible</a></td>
<td style="text-align: left;" data-valian="middle"><p>用于检索指示是否对透视字段应用任何手动筛选的 Boolean 值。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 <span>Application</span> 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.AutoShowCount">AutoShowCount</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定数据透视表字段中自动显示的首项号或末项号。 Long 类型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.AutoShowField">AutoShowField</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回数据字段的名称，该字段用于判断在指定数据透视表字段中自动显示的是首项还是末项。 String 类型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.AutoShowRange">AutoShowRange</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定数据透视表字段中自动显示首项，则返回 xlTop ；如果自动显示末项，则返回 xlBottom 。 Long 类型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.AutoShowType">AutoShowType</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果启用指定数据透视表字段的 AutoShow ，则返回 xlAutomatic ；如果禁用 AutoShow ，则返回 xlManual 。 Long 类型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.AutoSortCustomSubtotal">AutoSortCustomSubtotal</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回用于对指定数据透视表字段进行自动排序的自定义分类汇总的名称。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.AutoSortField">AutoSortField</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回用于对指定数据透视表字段进行自动排序的数据字段的名称。 String 类型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.AutoSortOrder">AutoSortOrder</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回用于对指定数据透视表字段进行自动排序的次序。可为 <span>XlSortOrder</span> 常量之一。 Long 类型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.AutoSortPivotLine">AutoSortPivotLine</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回用于对指定数据透视表字段进行自动排序的 PivotLine 的名称。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.BaseField">BaseField</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置自定义计算的基准字段。本属性仅对数据字段有效。 Variant 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.BaseItem">BaseItem</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置自定义计算基准字段的数据项，仅对数据字段有效。 Variant 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Calculation">Calculation</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 <span>XlPivotFieldCalculation</span> 值，它代表指定字段执行的计算类型。此属性仅对数据字段有效。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Caption">Caption</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 String 值，它代表数据透视字段的标签文本。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.ChildField">ChildField</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>PivotField</span> 对象，该对象表示特定字段的子字段（如果该字段已分组且有子字段）。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">16</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.ChildItems">ChildItems</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个对象，它代表一个数据透视表数据项（ <span>PivotItem</span> 对象）或所有数据项的集合（ <span>PivotItems</span> 对象），该集合可以由指定字段中的子项组成，也可以由指定数据项中的子项组成。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">17</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">18</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.CubeField">CubeField</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回派生出指定数据透视表字段的 <span>CubeField</span> 对象。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">19</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.CurrentPage">CurrentPage</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置页字段的当前页显示（仅对页字段有效）。 <span>PivotItem</span> 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">20</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.CurrentPageList">CurrentPageList</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置对应于项目列表的字符串数组，该项目列表包含于数据透视表的多项目页字段中。 Variant 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">21</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.CurrentPageName">CurrentPageName</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定数据透视表上的当前显示页。该页名称将出现在页字段中。注意，只有当已存在当前显示页时，本属性才有效。 String 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">22</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DatabaseSort">DatabaseSort</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 True ，则允许手动更改数据透视表字段中项目的位置。如果该字段中没有手动定位的项目，则返回 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">23</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DataRange">DataRange</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>Range</span> 对象，如下表所示。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">24</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DataType">DataType</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>XlPivotFieldDataType</span> 值，它代表数据透视表字段中的数据类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">25</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DisplayAsCaption">DisplayAsCaption</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>此属性用于将透视字段的成员属性显示为标题。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">26</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DisplayAsTooltip">DisplayAsTooltip</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>此属性用于指定在工具提示中是否显示透视字段的特定成员属性。可读写。 Boolean 。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">27</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DisplayInReport">DisplayInReport</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>此属性用于指定在数据透视表中是否显示指定的成员属性 PivotField。可读/写 Boolean 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">28</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DragToColumn">DragToColumn</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定字段能被拖动到列位置上，则为 True 。默认值是 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">29</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DragToData">DragToData</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定字段可被拖动到数据位置上，则为 True 。默认值是 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">30</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DragToHide">DragToHide</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果通过将字段拖离数据透视表可隐藏该字段，则为 True 。默认值是 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">31</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DragToPage">DragToPage</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果字段可被拖动到页位置上，则为 True 。默认值是 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">32</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DragToRow">DragToRow</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果字段可被拖动到行位置上，则为 True 。默认值是 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">33</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.DrilledDown">DrilledDown</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定数据透视表字段或数据透视表项的标志设置为“drilled”（展开或可见），则为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">34</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.EnableItemSelection">EnableItemSelection</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果为 False ，则在用户界面中禁止使用下拉字段的功能。默认值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">35</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.EnableMultiplePageItems">EnableMultiplePageItems</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>用于指定对于页面区域中的字段是否在筛选器下拉列表中显示复选框。可读/写 Boolean 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">36</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Formula">Formula</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 String 值，它代表 A1 样式表示法和宏语言中的对象的公式。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">37</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Function">Function</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置对数据透视表字段汇总时所使用的函数（仅用于数据字段）。 <span>XlConsolidationFunction</span> 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">38</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.GroupLevel">GroupLevel</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一组字段中指定字段的位置（如果该字段是分组字段集合中的成员）。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">39</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Hidden">Hidden</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>此属性用于隐藏 OLAP 层次结构的各个级别。可读/写 Boolean 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">40</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.HiddenItems">HiddenItems</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个对象，该对象表示单个隐藏数据透视表项（ <span>PivotItem</span> 对象），或表示指定字段中所有隐藏项的集合（ <span>PivotItems</span> 对象）。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">41</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.HiddenItemsList">HiddenItemsList</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个字符串数组，该数组为数据透视表字段的隐藏项。 Variant 类型，可读写</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">42</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.IncludeNewItemsInFilter">IncludeNewItemsInFilter</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>在将手动筛选应用于透视字段时，此属性允许开发人员指定是应跟踪排除的项目还是应跟踪包含的项目。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">43</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.IsCalculated">IsCalculated</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果数据透视表字段为计算字段或计算项，则此属性为 True 。 Boolean 类型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">44</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.IsMemberProperty">IsMemberProperty</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果数据透视表字段包含成员属性，则返回 True 。 Boolean 类型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">45</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.LabelRange">LabelRange</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>Range</span> 对象，它代表包含字段标签的单元格。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">46</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.LayoutBlankLine">LayoutBlankLine</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果在数据透视表的指定行字段后插入了一个空行，则该值为 True 。默认值为 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">47</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.LayoutCompactRow">LayoutCompactRow</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>指定在选择行时是否压缩透视字段（在一列中显示多个透视字段的项目）。可读/写 Boolean 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">48</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.LayoutForm">LayoutForm</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定的数据透视表项出现的方式，即以表格格式还是以分级显示格式显示。 <span>XlLayoutFormType</span> 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">49</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.LayoutPageBreak">LayoutPageBreak</a></td>
<td style="text-align: left;" data-valian="middle"></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">50</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.LayoutSubtotalLocation">LayoutSubtotalLocation</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置与指定字段相关（在其上面或下面）的数据透视表字段分类汇总的位置。 <span>XlSubtototalLocationType</span> 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">51</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.MemberPropertyCaption">MemberPropertyCaption</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>设置 MemberPropertyCaption 属性可控制将哪个成员属性用作指定级别的标题。可读写 Boolean 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">52</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.MemoryUsed">MemoryUsed</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回对象当前使用的内存大小，以字节表示。 Long 型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">53</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Name">Name</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 String 值，它代表对象的名称。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">54</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.NumberFormat">NumberFormat</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 String 值，它代表对象的格式代码。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">55</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Orientation">Orientation</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 <span>XlPivotFieldOrientation</span> 值，它代表字段在指定的数据透视表中的位置。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">56</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">57</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.ParentField">ParentField</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>PivotField</span> 对象，该对象表示作为指定对象的组父级的数据透视表字段。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">58</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.ParentItems">ParentItems</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个对象，该对象表示指定字段中的单个数据透视表项（ <span>PivotItem</span> 对象），或者作为组父级的所有项的一个集合（ <span>PivotItems</span> 对象）。指定字段必须是另一个字段的组父级。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">59</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.PivotFilters">PivotFilters</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置指定 PivotField 对象的 PivotFilters。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">60</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Position">Position</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 Variant 值，他代表数据透视表上所有字段的位置，包括属于所有行、列、页以及数据方向上的字段（第一字段、第二字段、第三字段等等）。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">61</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.PropertyOrder">PropertyOrder</a></td>
<td style="text-align: left;" data-valian="middle"><p>只对属于成员属性字段的数据透视表字段有效。返回一个 Long 类型的数值，该数值表示成员属性在其所属的多维数据集字段内的显示位置。可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">62</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.PropertyParentField">PropertyParentField</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 PivotField 对象，该对象表示字段中属性所从属的字段。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">63</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.RepeatLabels">RepeatLabels</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置在数据透视表中是否对指定的透视字段重复项目标签。可读/写</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">64</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.ServerBased">ServerBased</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定数据透视表的数据源为外部数据源，并且只检索与选定页字段相匹配的数据项，则该属性值为 True 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">65</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.ShowAllItems">ShowAllItems</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果显示数据透视表中的所有项目（即使这些项目中不包含汇总数据），则该值为 True 。默认值为 False 。 Boolean 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">66</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.ShowDetail">ShowDetail</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>获取或设置指定的 <span>PivotField</span> 是否显示明细数据。可读/写 Boolean 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">67</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.ShowingInAxis">ShowingInAxis</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>指明透视字段当前在数据透视表中是否可见。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">68</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.SourceCaption">SourceCaption</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>SourceCaption 属性只适用于 OLAP 数据透视表，并从 OLAP 服务器中返回透视字段的原始标题。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">69</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.SourceName">SourceName</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 String 值，它代表指定的对象出现在指定的数据透视表的原始源数据中时的名称。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">70</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.StandardFormula">StandardFormula</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 String 值，该值指定使用标准英语（美国）格式的公式。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">71</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.SubtotalName">SubtotalName</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置显示在指定数据透视表的分类汇总列或行标题中的文本字符串标志。默认值为“Subtotal”。 String 类型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">72</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Subtotals">Subtotals</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置与指定字段同时显示的分类汇总。仅对非数据字段有效。 Variant 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">73</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.TotalLevels">TotalLevels</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回当前字段组中的字段总数。如果字段没有分组或数据源基于的是 <span>OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。）</span> ，则 TotalLevels 将返回值 1。 Long 类型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">74</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.UseMemberPropertyAsCaption">UseMemberPropertyAsCaption</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>此属性用于控制是否将成员属性标题用于透视字段的 PivotItem 标题。可读/写 Boolean 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">75</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.Value">Value</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回或设置一个 String 值，它代表数据透视表中指定的字段的名称。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">76</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.VisibleItems">VisibleItems</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个对象，该对象表示指定字段中的单个可见数据透视表项（ <span>PivotItem</span> 对象），或指定字段中的所有可见项的集合（ <span>PivotItems</span> 对象）。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">77</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotField.VisibleItemsList">VisibleItemsList</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Variant 类型的值，该值指定一个字符串数组，字符串代表应用于透视字段的手动筛选中的包含项。可读/写。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### PivotField.AddPageItem

向具有多个项的页面字段添加其他项。

**语法**

**express.AddPageItem ( *Item* , *ClearList* )**

\*express是一个代表 PivotField 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                           |
|-------------|-----------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Item*      | 必选      | String   | PivotItem 对象的源名称，该名称与特定的联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 成员的唯一名称相对应。 |
| *ClearList* | 可选      | Variant  | 如果为 False（默认值），则向现有列表中添加一个页面项。如果为 True，则删除所有当前项，然后添加 Item。                                                                                                           |

**说明**

为了避免运行时错误，数据源必须是 OLAP 源，选择的字段当前必须位于页面位置中，并且 **EnableMultiplePageItems** 属性必须设置为 **True** 。

**示例**

``` JavaScript
/*在本例中，ET 添加一个源名称为“[Product].[All Products].[Food].[Eggs]”的页面项。本示例假定 OLAP 数据透视表位于活动工作表上。*/
function test(){
 // The source is an OLAP database and you can manually reorder items.
    ActiveSheet.PivotTables(1).CubeFields.Item("[Product]"). EnableMultiplePageItems = true
                                
    // Add the page item titled "[Product].[All Products].[Food].[Eggs]".
    ActiveSheet.PivotTables(1).PivotFields("[Product]").AddPageItem ("[Product].[All Products].[Food].[Eggs]")
}
```

### PivotField.AutoShow

显示指定数据透视表中行字段、页字段或列字段顶部或底部数据项的个数。

**语法**

**express.AutoShow ( *Type* , *Range* , *Count* , *Field* )**

\*express是一个代表 PivotField 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                    |
|---------|-----------|----------|-----------------------------------------------------------------------------------------|
| *Type*  | 必选      | Long     | 使用 xlAutomatic 可使指定数据透视表显示与指定条件匹配的项。使用 xlManual 可禁用此功能。 |
| *Range* | 必选      | Long     | 所要显示的项的起始位置。可以是下列两个常量之一：xlTop 或 xlBottom。                     |
| *Count* | 必选      | Long     | 要显示的项数。                                                                          |
| *Field* | 必选      | String   | 基础数据字段的名称。必须指定唯一名称（由 SourceName 属性返回）而不是显示的名称。        |

**示例**

``` JavaScript
/*本示例按销售总额大小只显示销售总额最高的两个公司：*/
Application.ActiveSheet.PivotTables("Pivot1").PivotFields("Company").AutoShow(xlAutomatic, xlTop, 2, "Sum of Sales")
```

### PivotField.AutoSort

为数据透视表建立自动字段排序规则。

**语法**

**express.AutoSort ( *Order* , *Field* , *PivotLine* , *CustomSubtotal* )**

\*express是一个代表 PivotField 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                                                                                     |
|------------------|-----------|----------|------------------------------------------------------------------------------------------|
| *Order*          | 必选      | Long     | 指定排序顺序的 XlSortOrder 的常量之一。                                                  |
| *Field*          | 必选      | String   | 排序关键字段的名称。必须指定唯一名称（由 SourceName 属性返回的名称），而不是显示的名称。 |
| *PivotLine*      | 可选      | Variant  | 列上的一行或数据透视表中的一行。                                                         |
| *CustomSubtotal* | 可选      | Variant  | 自定义分类汇总字段。                                                                     |

**示例**

``` JavaScript
/*本例以销售总额为基准对“Company”字段按降序排序。*/
Application.ActiveSheet.PivotTables(1).PivotFields("Company").AutoSort(xlDescending, "Sum of Sales")
 
```

### PivotField.CalculatedItems

返回一个 **CalculatedItems** 集合，该集合表示指定数据透视表中的所有计算项。只读。

**语法**

**express.CalculatedItems ()**

\*express是一个代表 PivotField 对象的变量。

**返回值**

CalculatedItems

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该方法返回一个长度为零的集合。

**示例**

``` JavaScript
/*本示例创建计算项及其公式的列表。*/
function test(){
let pt = Application.Worksheets.Item(1).PivotTables(1)
let r = 0
for (let i = 1; i <= pt.PivotFields("Sales").CalculatedItems().Count; i++){
    r++
    Application.Worksheets.Item(3).Cells.Item(r, 1).Value2 = pt.PivotFields("Sales").CalculatedItems().Item(i).Name
    Application.Worksheets.Item(3).Cells.Item(r, 2).Value2 = pt.PivotFields("Sales").CalculatedItems().Item(i).Formula
}
}
```

### PivotField.ClearAllFilters

调用此方法将删除当前应用于透视字段的所有筛选。包括从透视字段的 **PivotFilters** 集合中删除所有筛选以及删除应用于透视字段的所有手动筛选。如果透视字段位于报表筛选区域中，则所选项将被设置为默认项。

**语法**

**express.ClearAllFilters ()**

\*express是一个代表 PivotField 对象的变量。

### PivotField.ClearLabelFilters

此方法将删除透视字段的 **PivotFilters** 集合中的所有标签筛选或所有日期筛选。

**语法**

**express.ClearLabelFilters ()**

\*express是一个代表 PivotField 对象的变量。

**说明**

下表列出了此方法将删除的各种标签筛选类型。

|                                     |
|-------------------------------------|
|                                     |
| **xlCaptionEquals**                 |
| **xlCaptionDoesNotEqual**           |
| **xlCaptionIsGreaterThan**          |
| **xlCaptionIsGreaterThanOrEqualTo** |
| **xlCaptionIsLessThan**             |
| **xlCaptionIsLessThanOrEqualTo**    |
| **xlCaptionBeginsWith**             |
| **xlCaptionDoesNotBeginWith**       |
| **xlCaptionEndsWith**               |
| **xlCaptionDoesNotEndWith**         |
| **xlCaptionContains**               |
| **xlCaptionDoesNotContain**         |
| **xlCaptionIsBetween**              |
| **xlCaptionIsNotBetween**           |
|                                     |
|                                     |

下表列出了此方法将删除的各种日期筛选类型。

|                                 |
|---------------------------------|
|                                 |
| **xlSpecificDate**              |
| **xlNotSpecificDate**           |
| **xlBefore**                    |
| **xlBeforeOrEqualTo**           |
| **xlAfter**                     |
| **xlAfterOrEqualTo**            |
| **xlDateBetween**               |
| **xlDateNotBetween**            |
| **xlDateToday**                 |
| **xlDateYesterday**             |
| **xlDateTomorrow**              |
| **xlDateThisWeek**              |
| **xlDateLastWeek**              |
| **xlDateNextWeek**              |
| **xlDateThisMonth**             |
| **xlDateLastMonth**             |
| **xlDateNextMonth**             |
| **xlDateThisQuarter**           |
| **xlDateLastQuarter**           |
| **xlDateNextQuarter**           |
| **xlDateThisYear**              |
| **xlDateLastYear**              |
| **xlDateNextYear**              |
| **xlYearToDate**                |
| **xlAllDatesInPeriodQuarter1**  |
| **xlAllDatesInPeriodQuarter2**  |
| **xlAllDatesInPeriodQuarter3**  |
| **xlAllDatesInPeriodQuarter4**  |
| **xlAllDatesInPeriodJanuary**   |
| **xlAllDatesInPeriodFebruary**  |
| **xlAllDatesInPeriodMarch**     |
| **xlAllDatesInPeriodApril**     |
| **xlAllDatesInPeriodMay**       |
| **xlAllDatesInPeriodJune**      |
| **xlAllDatesInPeriodJuly**      |
| **xlAllDatesInPeriodAugust**    |
| **xlAllDatesInPeriodSeptember** |
| **xlAllDatesInPeriodOctober**   |
| **xlAllDatesInPeriodNovember**  |
| **xlAllDatesInPeriodDecember**  |

### PivotField.ClearManualFilter

提供一个简便方法，可将数据透视表中透视字段的所有项的 **Visible** 属性设置为 **True** ，并清空 OLAP 数据透视表中的 **HiddenItemsList** 和 **VisibleItemsList** 集合。

**语法**

**express.ClearManualFilter ()**

\*express是一个代表 PivotField 对象的变量。

**说明**

此方法适用于数据透视表中的 **PivotField** 对象和 OLAP 数据透视表中的 **CubeField** 对象。如果对 OLAP 数据透视表中的透视字段调用此方法，将返回运行时错误。

调用此方法后，以下集合将为空： **HiddenItemsList** 、 **HiddenItems** 、 **VisibleItemsList** 和 **VisibleItems** 。

### PivotField.ClearValueFilters

调用此方法将删除透视字段的 **PivotFilters** 集合中的所有值筛选。

**语法**

**express.ClearValueFilters ()**

\*express是一个代表 PivotField 对象的变量。

**说明**

下表列出了此方法将删除的各种值筛选类型。

|                                   |
|-----------------------------------|
|                                   |
| **xlTopCount**                    |
| **xlBottomCount**                 |
| **xlTopPercent**                  |
| **xlBottomPercent**               |
| **xlTopSum**                      |
| **xlBottomSum**                   |
| **xlValueEquals**                 |
| **xlValueDoesNotEqual**           |
| **xlValueIsGreaterThan**          |
| **xlValueIsGreaterThanOrEqualTo** |
| **xlValueIsLessThan**             |
| **xlValueIsLessThanOrEqualTo**    |
| **xlValueIsBetween**              |
| **xlValueIsNotBetween**           |

### PivotField.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 PivotField 对象的变量。

### PivotField.DrillTo

**DrillTo** 方法支持从另一个透视字段深化到指定的透视字段。

**语法**

**express.DrillTo ( *PivotFieldName* )**

\*express是一个代表 PivotField 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型 | 说明                           |
|------------------|-----------|----------|--------------------------------|
| *PivotFieldName* | 必选      | String   | 要深化到的 PivotField 的名称。 |

**说明**

只能在实际位于数据透视表上的字段上执行此操作。因此，要么包含所请求的透视字段的层次结构必须在数据透视表的行或列中，要么属性/关系字段必须在数据透视表的行或列中，并且放在至少一个其他属性/关系字段的旁边。而且，要深化到的字段必须在同一个层次结构或属性链中。如果不符合这些条件，则会发生运行时错误。

### PivotField.PivotItems

返回一个对象，该对象表示指定字段中的单个数据透视表项（ **PivotItem** 对象）或所有可见项和隐藏项的集合（ **PivotItems** 对象）。只读。

**语法**

**express.PivotItems ( *Index* )**

\*express是一个代表 PivotField 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                     |
|---------|-----------|----------|--------------------------|
| *Index* | 可选      | Variant  | 要返回的项的名称或编号。 |

**返回值**

Variant

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该集合是根据唯一的名称（由 **SourceName** 属性返回的名称），而不是根据显示的名称建立索引。

**示例**

``` JavaScript
/*本示例将“product”字段中所有项的名称添加到新工作表的列表中。*/
function test(){
let nwSheet = Application.Worksheets.Add()
    nwSheet.Activate()
    let pvtTable = Worksheets.Item("Sheet2").Range("A3").PivotTable
    let rw = 0
    for(let i = 1; i <= pvtTable.PivotFields("product").PivotItems.Count; i++){
        rw++
        nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("product").PivotItems(i).Name
    }
}
```

### PivotField.RepeatAllLabels

指定是否要对指定数据透视表中的所有透视字段重复项目标签。

**语法**

**express.RepeatAllLabels ( *Repeat* )**

\*express是一个代表 PivotField 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型                 | 说明 |
|----------|-----------|--------------------------|------|
| *Repeat* | 必选      | XlPivotFieldRepeatLabels |      |

**说明**

使用 **RepeatAllLabels** 方法对应于 **“数据透视表工具设计”** 选项卡的 **“报表布局”** 下拉列表中的 **“重复所有项目标签”** 和 **“不重复项目标签”** 命令。

若要指定是否要对单个数据透视表重复项目标签，请使用 **RepeatLabels** 属性。

## 成员属性

### PivotField.AllItemsVisible

用于检索指示是否对透视字段应用任何手动筛选的 Boolean 值。只读。

**语法**

**express.AllItemsVisible**

\*express是一个代表 PivotField 对象的变量。

**说明**

此属性提供了一种简单方式，便于检查是否对透视字段或多维数据集字段应用手动筛选。

对于 OLAP 数据透视表，此属性仅适用于 **CubeField** 对象。如果尝试从 OLAP 数据透视表中的 **PivotField** 对象获取此属性或设置此属性，将返回运行时错误。

对于数据透视表，此属性适用于 **PivotField** 对象。

默认值为 **True** 。未应用手动筛选（无论 **IncludeNewItemsInFilter** 属性为 **True** 还是为 **False** ）时，此属性将自动设置为 **True** 。应用任何手动筛选（无论 **IncludeNewItemsInFilter** 属性为 **True** 还是为 **False** ）时，此属性将自动设置为 **False** 。

此属性直接反映透视字段或多维数据集字段的筛选下拉列表中 **“全选”** 复选框的状态。

### PivotField.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotField 对象的变量。

**示例**

``` JavaScript
*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
    let myObject = Application.ActiveWorkbook
    if (myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    } else {
        alert("This is not an ET Application object.")
    }
}
```

### PivotField.AutoShowCount

返回指定数据透视表字段中自动显示的首项号或末项号。 **Long** 类型，只读。

**语法**

**express.AutoShowCount**

\*express是一个代表 PivotField 对象的变量。

**示例**

``` JavaScript
/*本示例在消息框中显示“Salesman”字段的 AutoShow 参数值。*/
function test(){
let pi = Application.Worksheets.Item(1).PivotTables(1).PivotFields("salesman")
if(pi.AutoShowType == xlAutomatic){
    let rn
    let r = pi.AutoShowRange
    if(r == xlTop){
        rn = "top"
    }
    else{
        rn = "bottom"
    }
    alert("PivotTable report is showing " + rn + " " + pi.AutoShowCount + " items in " + pi.Name + " field by " + pi.AutoShowField)
}
else{
    MsgBox("PivotTable report is not using AutoShow for this field")
}
}
```

### PivotField.AutoShowField

返回数据字段的名称，该字段用于判断在指定数据透视表字段中自动显示的是首项还是末项。 **String** 类型，只读。

**语法**

**express.AutoShowField**

\*express是一个代表 PivotField 对象的变量。

**示例**

``` JavaScript
/*本示例在消息框中显示“Salesman”字段的 AutoShow 参数值。*/
function test(){
let pi = Worksheets.Item(1).PivotTables(1).PivotFields("salesman")                              
if(pi.AutoShowType == xlAutomatic){
    let rn
    let r = pi.AutoShowRange
    if(r == xlTop){
        rn = "top"
    }
    else{
        rn = "bottom"
    }
    alert("PivotTable report is showing " + rn + " " + pi.AutoShowCount + " items in " + pi.Name + " field by " + pi.AutoShowField)
}
else{
    alert("PivotTable report is not using AutoShow for this field")
}
}
```

### PivotField.AutoShowRange

如果指定数据透视表字段中自动显示首项，则返回 **xlTop** ；如果自动显示末项，则返回 **xlBottom** 。 **Long** 类型，只读。

**语法**

**express.AutoShowRange**

\*express是一个代表 PivotField 对象的变量。

**示例**

``` JavaScript
/*本示例在消息框中显示“Salesman”字段的 AutoShow 参数值。*/
function test(){
let pi = Worksheets.Item(1).PivotTables(1).PivotFields("salesman")
if(pi.AutoShowType == xlAutomatic){
    let rn
    let r = pi.AutoShowRange
    if(r == xlTop){
        rn = "top"
    }
    else{
        rn = "bottom"
    }
    MsgBox("PivotTable report is showing " + rn + " " + pi.AutoShowCount + " items in " + pi.Name + " field by " + pi.AutoShowField)
}
else{
    MsgBox("PivotTable report is not using AutoShow for this field")
}
}
```

### PivotField.AutoShowType

如果启用指定数据透视表字段的 **AutoShow** ，则返回 **xlAutomatic** ；如果禁用 **AutoShow** ，则返回 **xlManual** 。 **Long** 类型，只读。

**语法**

**express.AutoShowType**

\*express是一个代表 PivotField 对象的变量。

**示例**

``` JavaScript
/*本示例在消息框中显示“Salesman”字段的 AutoShow 参数值。*/
function test(){
let pi = Application.Worksheets.Item(1).PivotTables(1).PivotFields("salesman")
if(pi.AutoShowType == xlAutomatic){
    let rn
    let r = pi.AutoShowRange
    if(r == xlTop){
        rn = "top"
    }
    else{
        rn = "bottom"
    }
    alert("PivotTable report is showing " + rn + " " + pi.AutoShowCount + " items in " + pi.Name + " field by " + pi.AutoShowField)
}
else{
    alert("PivotTable report is not using AutoShow for this field")
}
}
```

### PivotField.AutoSortCustomSubtotal

返回用于对指定数据透视表字段进行自动排序的自定义分类汇总的名称。只读。

**语法**

**express.AutoSortCustomSubtotal**

\*express是一个代表 PivotField 对象的变量。

**说明**

默认值是 1（自动）。如果 **AutoSortCustomSubtotal** 属性设置是 1（自动），则数据按常规分类汇总进行排序。 **AutoSortCustomSubtotal** 属性可以具有下表中列出的其中一个索引值。

|     |              |
|-----|--------------|
| 1   | 自动         |
| 2   | 求和         |
| 3   | 计数         |
| 4   | 平均值       |
| 5   | 最大值       |
| 6   | 最小值       |
| 7   | 乘积         |
| 8   | 数值计数     |
| 9   | 标准偏差     |
| 10  | 总体标准偏差 |
| 11  | 方差         |
| 12  | 总体方差     |

只有实际显示在数据透视表中的自定义分类汇总才支持排序，因此如果尝试将 **AutoSortCustomSubtotal** 设置为某个值，该值代表一个不在数据透视表视图中的自定义分类汇总，则会返回一个运行时错误。

如果基于自定义分类汇总应用排序，且从数据透视表中删除该分类汇总，则 **AutoSortCustomSubtotal** 属性将自动设置为默认值 (1)。

### PivotField.AutoSortField

返回用于对指定数据透视表字段进行自动排序的数据字段的名称。 **String** 类型，只读。

**语法**

**express.AutoSortField**

\*express是一个代表 PivotField 对象的变量。

**示例**

``` JavaScript
本示例在消息框中显示“Product”字段的 AutoSort 参数值。

示例代码 
let pi = Worksheets.Item(1).PivotTables(1).PivotFields("product")
let aso
switch(pi.AutoSortOrder){
    case xlManual:
        aso = "manual"
        break;
    case xlAscending:
        aso = "ascending"
        break;
    case xlDescending:
        aso = "descending"
        break;
}
MsgBox(" sorted in " + aso + " order by " + pi.AutoSortField)
```

### PivotField.AutoSortOrder

返回用于对指定数据透视表字段进行自动排序的次序。可为 **XlSortOrder** 常量之一。 **Long** 类型，只读。

**语法**

**express.AutoSortOrder**

\*express是一个代表 PivotField 对象的变量。

**示例**

本示例在消息框中显示“Product”字段的 AutoSort 参数值。

``` JavaScript
let pi = Worksheets.Item(1).PivotTables(1).PivotFields("product")
let aso
switch(pi.AutoSortOrder){
    case xlManual:
        aso = "manual"
        break;
    case xlAscending:
        aso = "ascending"
        break;
    case xlDescending:
        aso = "descending"
        break;
}
MsgBox(" sorted in " + aso + " order by " + pi.AutoSortField)
```

### PivotField.AutoSortPivotLine

table { word-break:break-all; }

返回用于对指定数据透视表字段进行自动排序的 PivotLine 的名称。只读。

**语法**

**express.AutoSortPivotLine**

\*express是一个代表 PivotField 对象的变量。

### PivotField.BaseField

返回或设置自定义计算的基准字段。本属性仅对数据字段有效。 **Variant** 类型，可读写。

**语法**

**express.BaseField**

\*express是一个代表 PivotField 对象的变量。

**说明**

该属性对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

本示例将工作表 Sheet1 上数据透视表中的数据字段设置为计算与基准字段的差，设置基准字段为“ORDER_DATE”字段，并将基准数据项设为“89-5-16”。

``` JavaScript
let pi =  Worksheets("Sheet1").Range("A3").PivotField
    pi.Calculation = xlDifferenceFrom
    pi.BaseField = "ORDER_DATE"
    pi.BaseItem = "5/16/89"
```

### PivotField.BaseItem

返回或设置自定义计算基准字段的数据项，仅对数据字段有效。 **Variant** 类型，可读写。

**语法**

**express.BaseItem**

\*express是一个代表 PivotField 对象的变量。

**说明**

该属性对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

本示例将工作表 Sheet1 上数据透视表中的数据字段设置为计算与基准字段的差，设置基准字段为“ORDER_DATE”字段，并将基准数据项设为“89-5-16”。

``` JavaScript
let pi = Worksheets("Sheet1").Range("A3").PivotField
    pi.Calculation = xlDifferenceFrom
    pi.BaseField = "ORDER_DATE"
    pi.BaseItem = "5/16/89"
```

### PivotField.Calculation

返回或设置一个 **XlPivotFieldCalculation** 值，它代表指定字段执行的计算类型。此属性仅对数据字段有效。

**语法**

**express.Calculation**

\*express是一个代表 PivotField 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性仅能返回或被设置为 **xlNormal** 。

**示例**

本示例将工作表 Sheet1 上数据透视表中的数据字段设置为计算与基准字段的差，设置基准字段为“ORDER_DATE”字段，并将基准数据项设为“89-5-16”。

``` JavaScript
let pi = Worksheets("Sheet1").Range("A3").PivotField
    pi.Calculation = xlDifferenceFrom
    pi.BaseField = "ORDER_DATE"
    pi.BaseItem = "5/16/89"
```

### PivotField.Caption

table { word-break:break-all; }

返回一个 **String** 值，它代表数据透视字段的标签文本。

**语法**

**express.Caption**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

下表显示 **Caption** 属性及其相关属性的示例值，假设存在一个具有唯一名称为“\[Europe\].\[France\].\[Paris\]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，并且存在一个包含名为“Paris”项的非 OLAP 数据源。

| 属性           | 值（OLAP 数据源）                       | 值（非 OLAP 数据源）        |
|----------------|-----------------------------------------|-----------------------------|
| **Caption**    | Paris                                   | Paris                       |
| **Name**       | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |
| **SourceName** | \[Europe\].\[France\].\[Paris\]（只读） | （与 SQL 属性值相同；只读） |
| **Value**      | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |

在将一个索引值指定到 **PivotItems** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源）                                      | 语法（非 OLAP 数据源）         |
|----------------------------------------------------------|--------------------------------|
| expression.PivotItems("\[Europe\].\[France\].\[Paris\]") | expression.PivotItems("Paris") |

在使用 **Item** 属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源）             | 名称（非 OLAP 数据源） |
|---------------------------------|------------------------|
| \[Europe\].\[France\].\[Paris\] | Paris                  |

### PivotField.ChildField

返回一个 **PivotField** 对象，该对象表示特定字段的子字段（如果该字段已分组且有子字段）。只读。

**语法**

**express.ChildField**

\*express是一个代表 PivotField 对象的变量。

**说明**

如果指定字段没有子字段，则使用该属性会出现错误。

该属性对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

本示例显示“REGION2”字段的子字段的名称。

``` JavaScript
let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
MsgBox("The name of the child field is " + pvtTable.PivotFields("REGION2").ChildField.Name)
```

### PivotField.ChildItems

返回一个对象，它代表一个数据透视表数据项（ **PivotItem** 对象）或所有数据项的集合（ **PivotItems** 对象），该集合可以由指定字段中的子项组成，也可以由指定数据项中的子项组成。只读。

**语法**

**express.ChildItems**

\*express是一个代表 PivotField 对象的变量。

**说明**

此属性对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

此示例将“vegetables”数据项的所有子数据项添加到一张新工作表上的数据清单中。

``` JavaScript
let nwSheet = Worksheets.Add()
    nwSheet.Activate()
let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable
let rw = 0
let pvtItem = pvtTable.PivotFields("product").PivotItems("vegetables").ChildItems
for(let i = 1; i <= pvtItem.Count; i++){
    rw = rw + 1
    nwSheet.Cells.Item(rw, 1).Value2 = pvtItem(i).Name
}
```

### PivotField.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotField.CubeField

返回派生出指定数据透视表字段的 **CubeField** 对象。只读。

**语法**

**express.CubeField**

\*express是一个代表 PivotField 对象的变量。

**说明**

**示例**

本示例为第一张工作表中第一个基于 联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 的数据透视表中所有分层字段创建一个多维数据集域名称列表。本示例假定在第一张工作表中存在数据透视表。

``` JavaScript
function UseCubeField(){
    let objNewSheet = Worksheets.Add()
        objNewSheet.Activate()
    let intRow = 1
    let objPF = Worksheets.Item(1).PivotTables(1).PivotFields
    for(let i = 1; i <= objPF.Count; i++){
        if(objPF(i).CubeField.CubeFieldType == xlHierarchy){
            objNewSheet.Cells.Item(intRow, 1).Value2 = objPF(i).Name
            intRow++
        }
    }
}
```

### PivotField.CurrentPage

返回或设置页字段的当前页显示（仅对页字段有效）。 **PivotItem** 类型，可读写。

**语法**

**express.CurrentPage**

\*express是一个代表 PivotField 对象的变量。

**示例**

本示例将 Sheet1 上数据透视表中的当前页名称返回到字符串变量 ` strPgName ` 中。

``` JavaScript
let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
let strPgName = pvtTable.PivotFields("Country").CurrentPage.Name
```

### PivotField.CurrentPageList

返回或设置对应于项目列表的字符串数组，该项目列表包含于数据透视表的多项目页字段中。 **Variant** 类型，可读写。

**语法**

**express.CurrentPageList**

\*express是一个代表 PivotField 对象的变量。

**说明**

为了避免运行时错误，数据源必须是 OLAP 源，选择的字段当前必须位于页面位置中，并且 **EnableMultiplePageItems** 属性必须设置为 **True** 。

**示例**

本示例设置页字段以列出数据透视表的“Food”项目。本示例假定数据透视表位于活动工作表上。

``` JavaScript
function UseCurrentPageList(){
    let pvtTable = ActiveSheet.PivotTables(1)
    let pvtField = pvtTable.PivotFields("[Product]")
                                    
    // To avoid run-time errors set the following property to True.
    pvtTable.CubeFields("[Product]").EnableMultiplePageItems = true
                                    
    // Set the page list to "Food".
    pvtField.CurrentPageList = "[Product].[All Products].[Food]"
}
```

### PivotField.CurrentPageName

返回或设置指定数据透视表上的当前显示页。该页名称将出现在页字段中。注意，只有当已存在当前显示页时，本属性才有效。 **String** 类型，可读写。

**语法**

**express.CurrentPageName**

\*express是一个代表 PivotField 对象的变量。

**说明**

本属性应用于与 OLAP 数据源相连的数据透视表。如果用未与 OLAP 数据源相连的数据透视表返回或设置本属性，则将导致运行时错误。

**示例**

本示例将活动工作表上第一个数据透视表的当前显示页的名称设置为“USA”。

``` JavaScript
ActiveSheet.PivotTables("PivotTable1").PivotFields("[Customers]").CurrentPageName = "[Customers].[All Customers].[USA]"
```

### PivotField.DatabaseSort

如果为 **True** ，则允许手动更改数据透视表字段中项目的位置。如果该字段中没有手动定位的项目，则返回 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DatabaseSort**

\*express是一个代表 PivotField 对象的变量。

**说明**

如果数据源不是 联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，则 **DatabaseSort** 属性返回 **False** 。

如果数据源是 OLAP，并且字段中既没有应用自定义排序也没有应用自动排序，那么该属性返回 **True** 。

对于 OLAP 数据透视表，如果将 **DatabaseSort** 属性设置为 **True** ，则会删除应用于字段的所有自定义排序或自动排序（也就是说，建立连接时数据透视表恢复为默认的状态）。

如果没有应用自动排序，那么将 **DatabaseSort** 属性设置为 **False** 时，会使排序次序变为当前的项目次序。

将 **DatabaseSort** 属性设置为 **True** 或 **False** 都会引起更新。

对于非 OLAP 源或 OLAP 数据字段，如果将 **DatabaseSort** 属性设置为 **True** ，则会导致运行时错误。

**示例**

本示例判断数据源是否是 OLAP 数据源，并通知用户。本示例假定 OLAP 数据透视表位于活动工作表上。

``` JavaScript
function UseDatabaseSort(){
    let pvtTable = ActiveSheet.PivotTables(1)
    let pvtField = pvtTable.PivotFields("[Product].[Product Family]")
                                    
    // Determine source type for the PivotTable report.
    if(pvtField.DatabaseSort == true){
        MsgBox("The source is OLAP; you can manually reorder items.")
    }
    else{
        MsgBox("The data source might not be OLAP.")
    }
}
```

### PivotField.DataRange

返回一个 **Range** 对象，如下表所示。只读。

**语法**

**express.DataRange**

\*express是一个代表 PivotField 对象的变量。

**说明**

| 对象             | 数据区域               |
|------------------|------------------------|
| “数据”字段       | 字段中包含的数据       |
| 行、列或者页字段 | 字段中的数据项         |
| 数据项           | 数据项规范所限定的数据 |

**示例**

此示例选定名为“REGION”字段中的数据透视表数据项。

``` JavaScript
let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
Worksheets.Item("Sheet1").Activate()
pvtTable.PivotFields("REGION").DataRange.Select()
```

### PivotField.DataType

返回一个 **XlPivotFieldDataType** 值，它代表数据透视表字段中的数据类型。

**语法**

**express.DataType**

\*express是一个代表 PivotField 对象的变量。

**示例**

本示例显示“ORDER_DATE”字段的数据类型。

``` JavaScript
let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
switch(pvtTable.PivotFields("ORDER_DATE").DataType) {
    case xlText:
        MsgBox("The field contains text data")
        break
    case xlNumber:
        MsgBox("The field contains numeric data")
        break
    case xlDate:
        MsgBox("The field contains date data")
}
```

### PivotField.DisplayAsCaption

table { word-break:break-all; }

此属性用于将透视字段的成员属性显示为标题。只读。

**语法**

**express.DisplayAsCaption**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

当给定成员属性用作标题时，此属性返回 **True** ，当透视字段的成员属性未用作标题时，返回 **False** 。尝试将此属性用于并非成员属性的透视字段时，将返回运行时错误。

### PivotField.DisplayAsTooltip

table { word-break:break-all; }

此属性用于指定在工具提示中是否显示透视字段的特定成员属性。可读写。 **Boolean** 。

**语法**

**express.DisplayAsTooltip**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

如果尝试从透视字段的非成员属性获取或设置这些属性，将返回运行时错误。

当 **DisplayAsTooltip** 属性设置为 **True** 时，将在工具提示中显示给定成员属性，当此属性设置为 **False** 时，不会在工具提示中显示给定成员属性。默认值为 **True** 。

### PivotField.DisplayInReport

table { word-break:break-all; }

此属性用于指定在数据透视表中是否显示指定的成员属性 PivotField。可读/写 **Boolean** 类型。

**语法**

**express.DisplayInReport**

\*express是一个代表 PivotField 对象的变量。

### PivotField.DragToColumn

如果指定字段能被拖动到列位置上，则为 **True** 。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DragToColumn**

\*express是一个代表 PivotField 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的度量字段，该值为 **False** 。

**示例**

**示例**

此示例防止将第一张工作表上第一个数据透视表中的“Year”字段拖动到列位置上。

``` JavaScript
Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").DragToColumn = false
```

### PivotField.DragToData

如果指定字段可被拖动到数据位置上，则为 **True** 。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DragToData**

\*express是一个代表 PivotField 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的度量字段，该值为 **False** 。

**示例**

此示例防止将“Year”字段拖动到第一张工作表上第一个数据透视表中的数据位置上。

``` JavaScript
Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").DragToData = false
```

### PivotField.DragToHide

如果通过将字段拖离数据透视表可隐藏该字段，则为 **True** 。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DragToHide**

\*express是一个代表 PivotField 对象的变量。

**说明**

**示例**

此示例防止将第一张工作表上第一个数据透视表中的“Year”字段拖离该报表。

``` JavaScript
Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").DragToHide = false
```

### PivotField.DragToPage

如果字段可被拖动到页位置上，则为 **True** 。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DragToPage**

\*express是一个代表 PivotField 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的度量字段，该值为 **False** 。

**示例**

此示例防止将第一张工作表上第一个数据透视表中的“Year”字段拖动到页位置上。

``` JavaScript
Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").DragToPage = false
```

### PivotField.DragToRow

如果字段可被拖动到行位置上，则为 **True** 。默认值是 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DragToRow**

\*express是一个代表 PivotField 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的度量字段，该值为 **False** 。

**示例**

此示例防止第一张工作表上第一个数据透视表中的“Year”字段被拖动到行位置上。

``` JavaScript
Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Year").DragToRow = false
```

### PivotField.DrilledDown

如果指定数据透视表字段或数据透视表项的标志设置为“drilled”（展开或可见），则为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DrilledDown**

\*express是一个代表 PivotField 对象的变量。

**说明**

此属性仅对 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源有效。

如果字段或项被隐藏，则不可设置此属性。

**示例**

此示例将活动工作表上第三个数据透视表中状态字段的所有项的标记设置为“not drilled”。

``` JavaScript
ActiveSheet.PivotTables("PivotTable3").PivotFields("state").DrilledDown = false
```

### PivotField.EnableItemSelection

如果为 **False** ，则在用户界面中禁止使用下拉字段的功能。默认值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.EnableItemSelection**

\*express是一个代表 PivotField 对象的变量。

**说明**

如果 OLAP 数据透视表字段不是层次结构中的最高级别，则将出现运行时错误。

**示例**

**示例**

本示例确定使用下拉字段的选定项目的设置，并且在必要时启用该功能。本示例假定数据透视表位于活动工作表上。 示例代码 function UseEnableItemSelection() { let pvtTable = ActiveSheet.PivotTables(1) let pvtField = pvtTable.RowFields(1) // Determine setting for property and enable if necessary. if(pvtField.EnableItemSelection == false) { pvtField.EnableItemSelection = true MsgBox("Item selection enabled for fields.") } else { MsgBox("Item selection is already enabled for fields.") } }

### PivotField.EnableMultiplePageItems

table { word-break:break-all; }

用于指定对于页面区域中的字段是否在筛选器下拉列表中显示复选框。可读/写 **Boolean** 类型。

**语法**

**express.EnableMultiplePageItems**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

为 OLAP 保留现有属性值。

| 注释                                                                                                                                                                                                                                                                                                                                                                  |
|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 在 ET 2007 或更高版本中，如果创建先于 ET 2007 的 OLAP 数据透视表 (PivotTable.Version \< 3)，且 **PivotTable** 对象的 **SubtotalHiddenPageItems** 属性和 **PivotField** 对象的 **EnableMultiplePageItems** 属性均设置为 **True** ，则更改页面区域的筛选器下拉菜单中的复选框的状态将不会有效果。在这种情况下，筛选器将始终设置为 **All** ，并包含未选中（隐藏）的项目。 |

### PivotField.Formula

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

**语法**

**express.Formula**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

此属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回空字符串。如果单元格包含公式， **Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

### PivotField.Function

返回或设置对数据透视表字段汇总时所使用的函数（仅用于数据字段）。 **XlConsolidationFunction** 类型，可读写。

**语法**

**express.Function**

\*express是一个代表 PivotField 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，本属性是只读的，并且总是返回 **xlUnknown** 。对于其他的数据源，本属性不能被设置为 **xlUnknown** 。

**示例**

本示例对活动工作表中第一个数据透视表的“Sum of 1994”字段进行设置，使之使用“SUM”函数。

### PivotField.GroupLevel

返回一组字段中指定字段的位置（如果该字段是分组字段集合中的成员）。只读。

**语法**

**express.GroupLevel**

\*express是一个代表 PivotField 对象的变量。

**说明**

该属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

处于最高级的父字段（即最左侧的父字段）为第一级，其子字段为第二级，依此类推。

**示例**

本示例检查包含活动单元格的字段是否为最高一级的父字段，如果是则显示一个消息框。

``` JavaScript
Worksheets.Item("Sheet1").Activate()
if(ActiveCell.PivotField.GroupLevel == 1) {
    MsgBox("This is the highest-level parent field.")
}
```

### PivotField.Hidden

table { word-break:break-all; }

此属性用于隐藏 OLAP 层次结构的各个级别。可读/写 **Boolean** 类型。

**语法**

**express.Hidden**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

对于隐藏的级别，返回 **True** ；对于可见（非隐藏）的级别，返回 **False** 。

此属性只适用于 OLAP 层次结构的级别。

对于关系数据源和 OLAP 特性，此属性始终为 **False** 。如果试图为特性或关系字段将其设置为 **True** ，将会返回运行时错误。 如果此设置对于同一层次结构的所有其他级别已设置为 **True** ，则无法为某个级别将此属性设置为 **True** 。试图这样做将会返回运行时错误。必须为层次结构中的至少一个级别将此属性设置为 **False** 。

### PivotField.HiddenItems

返回一个对象，该对象表示单个隐藏数据透视表项（ **PivotItem** 对象），或表示指定字段中所有隐藏项的集合（ **PivotItems** 对象）。只读。

**语法**

**express.HiddenItems**

\*express是一个代表 PivotField 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该属性总返回一个空集合。

**示例**

本示例将“product”字段的所有隐藏数据项的名称的列表添加到一张新工作表中。

``` JavaScript
let nwSheet = Worksheets.Add()
nwSheet.Activate()
let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable
let rw = 0
for(let i=1;i <= pvtTable.PivotFields("product").HiddenItems.Count;i++) {
    rw++
    nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("product").HiddenItems.Item(i).Name
}
```

### PivotField.HiddenItemsList

返回或设置一个字符串数组，该数组为数据透视表字段的隐藏项。 **Variant** 类型，可读写

**语法**

**express.HiddenItemsList**

\*express是一个代表 PivotField 对象的变量。

**说明**

**HiddenItemsList** 属性只对 联机分析处理 (OLAP) ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源有效，如果将该属性用于非 OLAP 数据源，则将返回一个运行错误。

**示例**

本示例设置项目列表，从而只显示某些项目。本示例假定 OLAP 数据透视表位于活动工作表上。

``` JavaScript
function UseHiddenItemsList() {
    ActiveSheet.PivotTables(1).PivotFields(1).HiddenItemsList = ["[Product].[All Products].[Food]","[Product].[All Products].[Drink]"]
}
```

### PivotField.IncludeNewItemsInFilter

table { word-break:break-all; }

在将手动筛选应用于透视字段时，此属性允许开发人员指定是应跟踪排除的项目还是应跟踪包含的项目。 **Boolean** 类型，可读写。

**语法**

**express.IncludeNewItemsInFilter**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

此属性的默认值为 **False** 。

应用手动筛选后，开发人员可以将 **IncludeNewItemsInFilter** 属性设置为 **True** 来跟踪排除的项目，也可以将此属性设置为 **False** 来跟踪包含的项目。如果在 **IncludeNewItemsInFilter** 属性设置为 **True** 后向源数据中添加了新项目，则这些新项目只有在进行新的刷新操作后才会显示在数据透视表中，因为它们未处在 ET 正在跟踪的项目集合中。

对于 OLAP 层次结构，此设置在 **CubeField** 对象上设置，而试图在作为层次结构一部分的透视字段上设置它将会失败（并会产生运行时错误）。可以通过作为层次结构一部分的透视字段获取此设置，而且它将返回与对应的 **CubeField.IncludeNewItemsInFilter** 属性相同的结果。

切换此设置后，以下集合将为空： **HiddenItemsList** 、 **HiddenItems** 、 **VisibleItemsList** 和 **VisibleItems** 。 当 **IncludeNewItemsInFilter** 设置为 **False** 时， **HiddenItemsList** 和 **HiddenItems** 集合为空，并且无法向它们添加项目。试图添加项目时会返回运行时错误。 当 **IncludeNewItemsInFilter** 设置为 **True** 时， **VisibleItemsList** 和 **VisibleItems** 集合为空，并且无法向它们添加项目。试图添加项目时会返回运行时错误。

在数据透视表中，此设置在 **PivotField** 对象上设置。切换此设置将不会改变筛选器状态。

### PivotField.IsCalculated

如果数据透视表字段为计算字段或计算项，则此属性为 **True** 。 **Boolean** 类型，只读。

**语法**

**express.IsCalculated**

\*express是一个代表 PivotField 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性始终返回 **False** 。

**示例**

如果指定数据透视表包含任何计算字段，此示例将禁用 **“数据透视表字段”** 对话框。

``` JavaScript
let pt = Worksheets.Item(1).PivotTables("Pivot1")
for(let i=1;i <= pt.PivotFields().Count;i++) {
    if(pt.PivotFields(i).IsCalculated) {
        pt.EnableFieldDialog = false
    }
}
```

### PivotField.IsMemberProperty

如果数据透视表字段包含成员属性，则返回 **True** 。 **Boolean** 类型，只读。

**语法**

**express.IsMemberProperty**

\*express是一个代表 PivotField 对象的变量。

**说明**

如果没有使用 联机分析处理 (OLAP) ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，则该属性将返回一个运行时错误。

**示例**

本示例确定数据透视表字段是否包含成员属性，并通知用户。本示例假定数据透视表位于活动的工作表上，并与 OLAP 数据源相连。 示例代码 function CheckForMembers() { let pvtTable = ActiveSheet.PivotTables(1) let pvtField = pvtTable.PivotFields(1) // Determine if member properties exist and notify user. if(pvtField.IsMemberProperty == true) { MsgBox("The PivotField contains member properties.") } else { MsgBox("The PivotField does not contain member properties.") } }

### PivotField.LabelRange

返回一个 **Range** 对象，它代表包含字段标签的单元格。只读。

**语法**

**express.LabelRange**

\*express是一个代表 PivotField 对象的变量。

**说明**

**示例**

此示例选择“ORDER_DATE”字段的字段按钮。

``` JavaScript
let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
let pvtField = pvtTable.PivotFields("ORDER_DATE")
Worksheets.Item("Sheet1").Activate()
pvtField.LabelRange.Select()
```

### PivotField.LayoutBlankLine

如果在数据透视表的指定行字段后插入了一个空行，则该值为 **True** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.LayoutBlankLine**

\*express是一个代表 PivotField 对象的变量。

**说明**

可以对任意数据透视表字段设置此属性；然而，只有当指定字段是除最里层（最低 级别?（级别：OLAP 维的一部分。在维中，数据的明细级别会有高、低之分，如在时间维内分为年、季度、月和日等级别。） ）行字段之外的其他行字段时，该空行才会出现。对于非 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，当重新排列该字段或将其添加到数据透视表中时，此属性的值不会改变。

不能在数据透视表的空行中输入数据。

**示例**

本示例在活动工作表的第一个数据透视表的状态字段后添加一个空行。

``` JavaScript
ActiveSheet.PivotTables("PivotTable1").PivotFields("state").LayoutBlankLine = true
```

### PivotField.LayoutCompactRow

table { word-break:break-all; }

指定在选择行时是否压缩透视字段（在一列中显示多个透视字段的项目）。可读/写 **Boolean** 类型。

**语法**

**express.LayoutCompactRow**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

如果 **LayoutForm** 属性设置为 **xlTabular** ，则将此属性设置为 **True** （默认值）将不会立即在数据透视表中产生可看到的效果，但可以独立地设置此属性。

### PivotField.LayoutForm

返回或设置指定的数据透视表项出现的方式，即以表格格式还是以分级显示格式显示。 **XlLayoutFormType** 类型，可读写。

**语法**

**express.LayoutForm**

\*express是一个代表 PivotField 对象的变量。

**说明**

|                                                                                        |
|----------------------------------------------------------------------------------------|
| **XlLayoutFormType** 可为以下 **XlLayoutFormType** 常量之一。                          |
| **xlTabular** 。默认值。                                                               |
| **xlOutline** 。 **LayoutSubtotalLocation** 属性指定分类汇总在数据透视表中出现的位置。 |

可以对任意数据透视表字段设置此属性。然而，只有当指定字段是除最里层（最低 级别 ?（级别：OLAP 维的一部分。在维中，数据的明细级别会有高、低之分，如在时间维内分为年、季度、月和日等级别。） ）行字段之外的其他行字段时，该格式才会出现。对于非 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，当重新排列该字段或将其添加到数据透视表中或从数据透视表中删除时，此属性的值不会更改。

**示例**

本示例以分级显示格式显示当前活动工作表上第一个数据透视表中的状态字段，然后在该字段顶部显示分类汇总。

``` JavaScript
ActiveSheet.PivotTables("PivotTable1").PivotFields("state").LayoutForm = xlOutline
ActiveSheet.PivotTables("PivotTable1").PivotFields("state").LayoutSubtotalLocation = xlTop
```

### PivotField.LayoutPageBreak

**语法**

**express.LayoutPageBreak**

\*express是一个代表 PivotField 对象的变量。

**说明**

虽然可以对任意数据透视表字段设置此属性，但只有当指定字段是除最里层（最低 级别 ?（级别：OLAP 维的一部分。在维中，数据的明细级别会有高、低之分，如在时间维内分为年、季度、月和日等级别。） ）行字段之外的其他行字段时，打印选项才会出现。对于非 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，当重新排列该字段或将其添加到数据透视表中或从数据透视表中删除时，此属性的值不会更改。

**示例**

本示例在活动工作表的第一个数据透视表的状态字段之后添加一个分页符。

``` JavaScript
ActiveSheet.PivotTables("PivotTable1").PivotFields("state").LayoutPageBreak = true
```

### PivotField.LayoutSubtotalLocation

返回或设置与指定字段相关（在其上面或下面）的数据透视表字段分类汇总的位置。 **XlSubtototalLocationType** 类型，可读写。

**语法**

**express.LayoutSubtotalLocation**

\*express是一个代表 PivotField 对象的变量。

**说明**

可以对任意具有分级显示格式的数据透视表字段设置此属性；但是，只有当指定字段是除最里层（最低 级别?（级别：OLAP 维的一部分。在维中，数据的明细级别会有高、低之分，如在时间维内分为年、季度、月和日等级别。） ）行字段之外的其他行字段时，格式才会出现。对于非 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，当重新排列该字段或将其添加到数据透视表中或从数据透视表中删除时，此属性的值不会更改。

**LayoutForm** 属性确定报表是以表格格式还是以分级显示格式显示。

**示例**

本示例以分级显示格式显示当前活动工作表上第一个数据透视表中的状态字段，然后在该字段顶部显示分类汇总。

``` JavaScript
ActiveSheet.PivotTables("PivotTable1").PivotFields("state").LayoutForm = xlOutline
ActiveSheet.PivotTables("PivotTable1").PivotFields("state").LayoutSubtotalLocation = xlAtTop
```

### PivotField.MemberPropertyCaption

table { word-break:break-all; }

设置 **MemberPropertyCaption** 属性可控制将哪个成员属性用作指定级别的标题。可读写 **Boolean** 类型。

**语法**

**express.MemberPropertyCaption**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

| 注释                                                                                               |
|----------------------------------------------------------------------------------------------------|
| 只有在为透视字段将 **UseMemberPropertyAsCaption** 设置为 **True** 时，此设置才会产生可看到的效果。 |

设置 **MemberPropertyCaption** 后，在切换 **UseMemberPropertyAsCaption** 属性时会记住此设置。

### PivotField.MemoryUsed

返回对象当前使用的内存大小，以字节表示。 **Long** 型，只读。

**语法**

**express.MemoryUsed**

\*express是一个代表 PivotField 对象的变量。

**示例**

此示例弹出一个消息框，该消息框中显示了 ET 当前使用的内存字节数。

``` JavaScript
MsgBox("ET is currently using " + Application.MemoryUsed + " bytes")
```

### PivotField.Name

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 PivotField 对象的变量。

### PivotField.NumberFormat

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表对象的格式代码。

**语法**

**express.NumberFormat**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

只能对数据字段设置 **NumberFormat** 属性。

格式代码与 **“设置单元格格式”** 对话框中的 **“格式代码”** 选项是同一个字符串。 **Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

### PivotField.Orientation

返回或设置一个 **XlPivotFieldOrientation** 值，它代表字段在指定的数据透视表中的位置。

**语法**

**express.Orientation**

\*express是一个代表 PivotField 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，当设置某一层中的一个字段的此属性值时，也会同时设置同一层中所有其他字段的方向。维字段只能在数据透视表的行、列和页字段区域中进行定向。度量字段则只能在数据区域中进行定向。将某一层或数据字段设置为 **xlHidden** 时，会将该层或字段从数据透视表中移出。

**示例**

本示例显示“ORDER_DATE”字段的方向。

``` JavaScript
let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
let pvtField = pvtTable.PivotFields("ORDER_DATE")
switch(pvtField.Orientation) {
    case xlHidden:
        MsgBox("Hidden field")
        break
    case xlRowField:
        MsgBox("Row field")
        break
    case xlColumnField:
        MsgBox("Column field")
        break
    case xlPageField:
        MsgBox("Page field")
        break
    case xlDataField:
        MsgBox("Data field")
}
```

### PivotField.Parent

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotField 对象的变量。

### PivotField.ParentField

返回一个 **PivotField** 对象，该对象表示作为指定对象的组父级的数据透视表字段。只读。

**语法**

**express.ParentField**

\*express是一个代表 PivotField 对象的变量。

**说明**

**示例**

本示例显示包含活动单元格的字段的组父级字段名称。

``` JavaScript
Worksheets.Item("Sheet1").Activate()
MsgBox("The active field is a child of the field " + ActiveCell.PivotField.ParentField.Name)
```

### PivotField.ParentItems

返回一个对象，该对象表示指定字段中的单个数据透视表项（ **PivotItem** 对象），或者作为组父级的所有项的一个集合（ **PivotItems** 对象）。指定字段必须是另一个字段的组父级。只读。

**语法**

**express.ParentItems**

\*express是一个代表 PivotField 对象的变量。

**说明**

该属性不可用于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。

**示例**

本示例创建“product”字段中所有组父项的名称的列表。

``` JavaScript
let nwSheet = Worksheets.Add()
nwSheet.Activate()
let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable
let rw = 0
for(let i=1;i <= pvtTable.PivotFields("product").ParentItems.Count;i++) {
    rw++
    nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("product").ParentItems.Item(i).Name
}
```

### PivotField.PivotFilters

table { word-break:break-all; }

返回或设置指定 **PivotField** 对象的 PivotFilters。只读。

**语法**

**express.PivotFilters**

\*express是一个代表 PivotField 对象的变量。

### PivotField.Position

table { word-break:break-all; }

返回一个 Variant 值，他代表数据透视表上所有字段的位置，包括属于所有行、列、页以及数据方向上的字段（第一字段、第二字段、第三字段等等）。

**语法**

**express.Position**

\*express是一个代表 PivotField 对象的变量。

### PivotField.PropertyOrder

只对属于成员属性字段的数据透视表字段有效。返回一个 **Long** 类型的数值，该数值表示成员属性在其所属的多维数据集字段内的显示位置。可读写。

**语法**

**express.PropertyOrder**

\*express是一个代表 PivotField 对象的变量。

**说明**

设置该属性将重新安排该多维数据集字段的属性的顺序。该属性从一开始。允许的区域从一到为层次结构显示的成员属性字段的最大数量。

如果 **IsMemberProperty** 属性为 **False** ，则使用 **PropertyOrder** 属性将导致一个运行时错误。

**示例**

本示例判断第四个字段中是否存在成员属性，如果存在，则显示成员属性的位置。 ET 将该信息通知用户。本示例假定数据透视表位于活动工作表上，并且基于联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。 示例代码 function CheckPropertyOrder() { let pvtTable = ActiveSheet.PivotTables(1) let pvtField = pvtTable.PivotFields(4) // Check for member properties and notify user. if(pvtField.IsMemberProperty == false) { MsgBox("No member properties present.") } else { MsgBox("The property order of the members is: " + pvtField.PropertyOrder) } }

### PivotField.PropertyParentField

返回一个 **PivotField** 对象，该对象表示字段中属性所从属的字段。

**语法**

**express.PropertyParentField**

\*express是一个代表 PivotField 对象的变量。

**说明**

只对属于成员属性字段的字段有效。

如果 **IsMemberProperty** 属性为 **False** ，则使用 **PropertyParentField** 属性将返回一个运行时错误。

**示例**

本示例确定第四个字段中是否存在成员属性，如果存在，则确定属性所属的字段。 ET 将该信息通知用户。本示例假定数据透视表位于活动工作表上，并且基于联机分析处理 (OLAP)?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。 示例代码 function CheckParentField() { let pvtTable = ActiveSheet.PivotTables(1) let pvtField = pvtTable.PivotFields(4) // Check for member properties and notify user. if(pvtField.IsMemberProperty == false) { MsgBox("No member properties present.") } else { MsgBox("The parent field of the members is: " + pvtField.PropertyParentField.Name) } }

### PivotField.RepeatLabels

table { word-break:break-all; }

返回或设置在数据透视表中是否对指定的透视字段重复项目标签。可读/写

**语法**

**express.RepeatLabels**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

如果对指定的透视字段重复项目标签，则为 **True** ；否则为 **False** 。

**RepeatLabels** 属性的设置对应于数据透视表中的字段的 **“字段设置”** 对话框 **“布局和打印”** 选项卡上的 **“重复项目标签”** 复选框。

若要指定是否在单次操作中对数据透视表内的所有透视字段重复项目标签，请使用 **RepeatAllLabels** 方法。

### PivotField.ServerBased

如果指定数据透视表的数据源为外部数据源，并且只检索与选定页字段相匹配的数据项，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.ServerBased**

\*express是一个代表 PivotField 对象的变量。

**说明**

该属性不能应用于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，这时其值总是 **False** 。

如果本属性设为 **True** ，则指定数据库中只有与选定页字段项相匹配的记录可被检索到。以后每当用户更改选定页字段时，新的选定页字段项将作为参数传递给查询，并且高速缓存将得到刷新。

如果下列某个条件成立，则不能对该属性进行设置：

- 字段被分组。
- 使用的不是外部数据源。
- 高速缓存由两个或者多个数据透视表共享使用。
- 字段的数据类型是不能基于服务器的类型（即备注字段或 OLE 对象）。

**示例**

本示例列出所有基于服务器的页字段。

``` JavaScript
let r = 0
for(let i=1;i <= ActiveSheet.PivotTables(1).PageFields().Count;i++) {
    if(ActiveSheet.PivotTables(1).PageFields(i).ServerBased == true) {
        r++
        Worksheets.Item(2).Cells.Item(r, 1).Value2 = ActiveSheet.PivotTables(1).PageFields(i).Name
    }
}
```

### PivotField.ShowAllItems

如果显示数据透视表中的所有项目（即使这些项目中不包含汇总数据），则该值为 **True** 。默认值为 **False** 。 **Boolean** 类型，可读写。

**语法**

**express.ShowAllItems**

\*express是一个代表 PivotField 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该值总是 **False** 。

**示例**

本示例显示第一张工作表上第一张数据透视表中“Month”字段的所有行，包括没有数据的月份。

``` JavaScript
Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Month").ShowAllItems = true
```

### PivotField.ShowDetail

table { word-break:break-all; }

获取或设置指定的 **PivotField** 是否显示明细数据。可读/写 **Boolean** 类型。

**语法**

**express.ShowDetail**

\*express是一个代表 PivotField 对象的变量。

### PivotField.ShowingInAxis

table { word-break:break-all; }

指明透视字段当前在数据透视表中是否可见。只读。

**语法**

**express.ShowingInAxis**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

如果透视字段当前可见，则此属性返回 **True** 。如果透视字段当前不可见，则返回 **False** 。透视字段的可见性取决于透视字段标题是否在表格（非压缩）形式的数据透视表中可见。

### PivotField.SourceCaption

table { word-break:break-all; }

**SourceCaption** 属性只适用于 OLAP 数据透视表，并从 OLAP 服务器中返回透视字段的原始标题。只读。

**语法**

**express.SourceCaption**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

可以将 **PivotField** . **Name** 或 **PivotField** . **SourceName** 用于非 OLAP 数据透视表。

### PivotField.SourceName

table { word-break:break-all; }

返回一个 **String** 值，它代表指定的对象出现在指定的数据透视表的原始源数据中时的名称。

**语法**

**express.SourceName**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

如果用户在创建数据透视表之后重命名数据项，此属性的值可能会与当前数据项名称不同。

下表显示 **SourceName** 属性及其相关属性的示例值，假设存在唯一名称为“\[Europe\].\[France\].\[Paris\]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，并且存在包含名为“Paris”的项的非 OLAP 数据源。

| 属性           | 值（OLAP 数据源）                       | 值（非 OLAP 数据源）        |
|----------------|-----------------------------------------|-----------------------------|
| **Caption**    | Paris                                   | Paris                       |
| **Name**       | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |
| **SourceName** | \[Europe\].\[France\].\[Paris\]（只读） | （与 SQL 属性值相同，只读） |
| **Value**      | \[Europe\].\[France\].\[Paris\]（只读） | Paris                       |

在将一个索引值指定到 **PivotItems** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源）                                      | 语法（非 OLAP 数据源）         |
|----------------------------------------------------------|--------------------------------|
| expression.PivotItems("\[Europe\].\[France\].\[Paris\]") | expression.PivotItems("Paris") |

在使用 **Item** 属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源）             | 名称（非 OLAP 数据源） |
|---------------------------------|------------------------|
| \[Europe\].\[France\].\[Paris\] | Paris                  |

### PivotField.StandardFormula

返回或设置一个 **String** 值，该值指定使用标准英语（美国）格式的公式。可读写。

**语法**

**express.StandardFormula**

\*express是一个代表 PivotField 对象的变量。

**说明**

**StandardFormula** 属性主要影响具有日期或数字格式的项目名称。该属性提供了一种方法可指定或查询给定计算项的公式。

**StandardFormula** 属性是国际通用的， **Formula** 属性则不是。

**示例**

本示例向“Decimals”字段中添加 10，并将其显示为数据字段中的计算项。本示例假定数据透视表位于活动工作簿上，并且标题为“Decimals”的字段位于模拟运算表中。

``` JavaScript
function UseStandardFomula() {
    let pvtTable = ActiveSheet.PivotTables(1)
    // Change calculated field of decimals by adding '10'.
    pvtTable.CalculatedFields().Item(1).StandardFormula = "Decimals + 10"
}
```

### PivotField.SubtotalName

返回或设置显示在指定数据透视表的分类汇总列或行标题中的文本字符串标志。默认值为“Subtotal”。 **String** 类型，可读写。

**语法**

**express.SubtotalName**

\*express是一个代表 PivotField 对象的变量。

**说明**

**示例**

本示例将活动工作表上第二个数据透视表的状态字段中的分类汇总标志设置为“Regional Subtotal”（取代默认的“Subtotal”）。

``` JavaScript
ActiveSheet.PivotTables("PivotTable2").PivotFields("state").SubtotalName = "Regional Subtotal"
```

### PivotField.Subtotals

返回或设置与指定字段同时显示的分类汇总。仅对非数据字段有效。 **Variant** 类型，可读写。

**语法**

**express.Subtotals**

\*express是一个代表 PivotField 对象的变量。

**说明**

如果索引为 **True** ，则该字段显示分类汇总。如果索引 1（自动）为 **True** ，则其他所有值将设置为 **False** 。

| 索引 | 含义       |
|------|------------|
| 1    | 自动       |
| 2    | 总计       |
| 3    | 计数       |
| 4    | 平均       |
| 5    | 最大值     |
| 6    | 最小值     |
| 7    | 乘积       |
| 8    | Count Nums |
| 9    | StdDev     |
| 10   | StdDevp    |
| 11   | Var        |
| 12   | Varp       |

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源， *Index* 只能返回或设置为 1（自动）。返回的数组对第一个数组元素总是包含 **True** 或 **False** ，而对所有其他元素则包含 **False** 。数组中所有元素的值都为 **False** ，则表明没有分类汇总。

**示例**

本示例将包含活动单元格的字段设置为显示分类汇总求和。

``` JavaScript
Worksheets.Item("Sheet1").Activate()
ActiveCell.PivotField.Subtotals(2, true)
```

### PivotField.TotalLevels

返回当前字段组中的字段总数。如果字段没有分组或数据源基于的是 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） ，则 **TotalLevels** 将返回值 1。 **Long** 类型，只读。

**语法**

**express.TotalLevels**

\*express是一个代表 PivotField 对象的变量。

**说明**

同一分组内的所有字段具有相同的 **TotalLevels** 值。

**示例**

本示例显示包含活动单元格的字段组中字段的总数。

``` JavaScript
Worksheets.Item("Sheet1").Activate()
MsgBox("This group has " + ActiveCell.PivotField.TotalLevels + " levels")
```

### PivotField.UseMemberPropertyAsCaption

table { word-break:break-all; }

此属性用于控制是否将成员属性标题用于透视字段的 PivotItem 标题。可读/写 **Boolean** 类型。

**语法**

**express.UseMemberPropertyAsCaption**

\*express是一个代表 PivotField 对象的变量。

**说明**

table { word-break:break-all; }

如果将透视字段的 **UseMemberPropertyAsCaption** 设置为 **True** ，则 **MemberPropertyCaption** 将指定要显示的成员属性标题。如果未指定，则该透视字段的第一个成员属性（按数据源顺序）将显示为该透视字段的项目的标题。

如果将 **UseMemberPropertyAsCaption** 设置为 **False** ，则常规 PivotItem 标题将用于该透视字段。

如果尝试将没有成员属性的透视字段的 **UseMemberPropertyAsCaption** 设置为 **True** ，则会返回一个运行时错误。对于没有成员属性的透视字段，该属性将始终为 **False** 。

### PivotField.Value

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表数据透视表中指定的字段的名称。

**语法**

**express.Value**

\*express是一个代表 PivotField 对象的变量。

### PivotField.VisibleItems

返回一个对象，该对象表示指定字段中的单个可见数据透视表项（ **PivotItem** 对象），或指定字段中的所有可见项的集合（ **PivotItems** 对象）。只读。

**语法**

**express.VisibleItems**

\*express是一个代表 PivotField 对象的变量。

**说明**

对于 OLAP ?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，本属性是只读的，并且总返回 **True** 。没有隐藏项。

**示例**

本示例向新工作表上的列表中添加“Product”字段中所有可见的数据项名称。

``` JavaScript
let nwSheet = Worksheets.Add()
nwSheet.Activate()
let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable
let rw = 0
for(let i=1;i <= pvtTable.PivotFields("Product").VisibleItems.Count;i++) {
    rw++
    nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("Product").VisibleItems.Item(i).Name
}
```

### PivotField.VisibleItemsList

返回或设置一个 **Variant** 类型的值，该值指定一个字符串数组，字符串代表应用于透视字段的手动筛选中的包含项。可读/写。

**语法**

**express.VisibleItemsList**

\*express是一个代表 PivotField 对象的变量。

**说明**

此属性只适用于 OLAP 数据透视表。

**示例**

此示例显示 OLAP 数据透视表中包含在内的手动筛选。

``` JavaScript
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography].[Country]").VisibleItemsList = ["[Customer].[Customer Geography].[Country].[Australia]"]
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography].[State-Province]").VisibleItemsList = [""]
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography].[City]").VisibleItemsList = [""]
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography].[Postal Code]").VisibleItemsList = [""]
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography].[Full Name]").VisibleItemsList = [""]
```

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotField](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PivotLineCells 对象

## 简介

table { word-break:break-all; }

特定 PivotLine 的 PivotCell 对象的集合。

## 说明

table { word-break:break-all; }

使用 PivotLineCells ( *index* ) 方法可以返回或指定集合中特定 PivotCell 对象的位置。您也可以指定 PivotField 对象或 PivotField 的名称以返回单个 PivotCell 对象。

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLineCells.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 <span>Application</span> 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLineCells.Count">Count</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回 PivotLineCells 集合中的项数。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLineCells.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLineCells.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<table>
<tbody>
<tr class="odd">
<td></td>
</tr>
</tbody>
</table></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#PivotLineCells.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回指定 PivotLineCells 对象的父对象。只读。</p></td>
</tr>
</tbody>
</table>

## 成员属性

### PivotLineCells.Application

table { word-break:break-all; }

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 PivotLineCells 对象的变量。

**说明**

table { word-break:break-all; }

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### PivotLineCells.Count

table { word-break:break-all; }

返回 **PivotLineCells** 集合中的项数。只读。

**语法**

**express.Count**

\*express是一个代表 PivotLineCells 对象的变量。

### PivotLineCells.Creator

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 PivotLineCells 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### PivotLineCells.Item

table { word-break:break-all; }

|     |
|-----|
|     |

**语法**

**express.Item**

\*express是一个代表 PivotLineCells 对象的变量。

### PivotLineCells.Parent

table { word-break:break-all; }

返回指定 **PivotLineCells** 对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 PivotLineCells 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/PivotLineCells](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ShapeRange 对象

## 简介

代表形状区域，它是文档中的一组形状。

## 说明

形状区域可以只包含文档中的一个形状，或者也可包含所有形状。您可以在形状区域中包含所需的任意形状（在文档中的所有形状中选取，或从选定内容中的所有形状中选取）。例如，您可以构造一个 ShapeRange 集合，它包含文档中前三个形状、文档中所有选定的形状，或文档中所有的任意多边形的。

1\. 返回指定名称或索引号的一组形状

使用 ` Shapes.Range ` ( *index* )（其中 *index* 是形状的名称或索引号，或由形状的名称或索引号组成的数组）可返回代表文档中的一组形状的 ShapeRange 集合。您可以使用 Array 函数来构造名称或索引号的数组。下例设置 *myDocument* 上形状一和三的填充图案。

``` JavaScript
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick)
}
```

下例设置 *myDocument* 上名为“Oval 4”和“Rectangle 5”的形状的填充图案。

虽然可以使用 Range 属性来返回任意数量的形状或幻灯片，但如果您只想返回一个集合成员，则使用 Item 方法会更简单。例如， ` Shapes(1) ` 比 ` Shapes.Range(1) ` 简单。

``` JavaScript
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    let myRange = myDocument.Shapes.Range(["Oval 4", "Rectangle 5"])
    myRange.Fill.Patterned(msoPatternHorizontalBrick)
}
```

2\. 返回文档中全部或部分选定的形状

使用 Selection 对象的 ShapeRange 属性可返回选定对象中的所有形状。下例设置窗口一中选定对象的所有形状的填充前景色（假定所选内容中至少有一个形状）。

``` JavaScript
Application.Windows.Item(1).Selection.ShapeRange.Fill.ForeColor.RGB = (255, 0, 255)
```

使用 ` Selection.ShapeRange ` ( *index* )（其中 *index* 是形状的名称或索引号）返回某一选定的形状。下例设置了窗口一内选定形状的集合中形状二的前景填充色（假定该选定内容中至少有两个形状）。

``` JavaScript
Application.Windows.Item(1).Selection.ShapeRange.Item(2).Fill.ForeColor.RGB = (255, 0, 255)
```

## 方法

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Align">Align</a></td>
<td style="text-align: left;" data-valian="middle"><p>对齐指定形状区域中的形状。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Apply">Apply</a></td>
<td style="text-align: left;" data-valian="middle"><p>应用通过 PickUp 方法复制的指定形状格式。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Delete">Delete</a></td>
<td style="text-align: left;" data-valian="middle"><p>删除对象。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Distribute">Distribute</a></td>
<td style="text-align: left;" data-valian="middle"><p>水平或垂直地分布指定的形状区域中的各形状。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Duplicate">Duplicate</a></td>
<td style="text-align: left;" data-valian="middle"><p>复制对象，并返回对新复制对象的引用。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Flip">Flip</a></td>
<td style="text-align: left;" data-valian="middle"><p>绕指定形状的水平或垂直对称轴翻转该形状。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Group">Group</a></td>
<td style="text-align: left;" data-valian="middle"><p>将指定区域中的形状形成一组。</p>
<p>返回值： 一个代表组合形状的 Shape 对象。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.IncrementLeft">IncrementLeft</a></td>
<td style="text-align: left;" data-valian="middle"><p>以指定磅数水平移动指定形状。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.IncrementRotation">IncrementRotation</a></td>
<td style="text-align: left;" data-valian="middle"><p>使指定的形状按指定度数值绕 Z 轴旋转。使用 Rotation 属性可设置形状的绝对转角。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.IncrementTop">IncrementTop</a></td>
<td style="text-align: left;" data-valian="middle"><p>以指定磅数垂直移动指定形状。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>从集合中返回一个对象。</p>
<p>返回值： 包含在集合中的一个 Shape 对象。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.PickUp">PickUp</a></td>
<td style="text-align: left;" data-valian="middle"><p>复制指定形状的格式。使用 Apply 方法可将复制的格式应用到其他形状。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Regroup">Regroup</a></td>
<td style="text-align: left;" data-valian="middle"><p>将指定形状原先所属的形状区域重新组合。以单一的 Shape 对象返回重新组合的形状。</p>
<p>返回值类型： Shape</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.RerouteConnections">RerouteConnections</a></td>
<td style="text-align: left;" data-valian="middle"><p>此方法将重排连接在指定形状上的所有连接符；如果指定的形状是连接符，就重排该连接符。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ScaleHeight">ScaleHeight</a></td>
<td style="text-align: left;" data-valian="middle"><p>按指定的比例调整形状的高度。对于图片和 OLE 对象，可以指定是相对于原有尺寸还是相对于当前尺寸来调整该形状。对于不是图片和 OLE 对象的形状，总是相对于其当前大小来调整高度。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">16</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ScaleWidth">ScaleWidth</a></td>
<td style="text-align: left;" data-valian="middle"><p>按指定的比例调整形状的宽度。对于图片和 OLE 对象，可以指定是相对于原有尺寸还是相对于当前尺寸来调整该形状。对于不是图片和 OLE 对象的形状，总是相对于其当前大小来调整宽度。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">17</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Select">Select</a></td>
<td style="text-align: left;" data-valian="middle"><p>选择对象。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">18</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.SetShapesDefaultProperties">SetShapesDefaultProperties</a></td>
<td style="text-align: left;" data-valian="middle"><p>将指定形状的格式设置为形状的默认格式。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">19</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Ungroup">Ungroup</a></td>
<td style="text-align: left;" data-valian="middle"><p>取消指定形状或者形状区域中组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。</p>
<p>返回值： 一个 ShapeRange 对象，它代表取消组合的形状。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">20</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ZOrder">ZOrder</a></td>
<td style="text-align: left;" data-valian="middle"><p>将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-次序中的位置）。</p></td>
</tr>
</tbody>
</table>

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Adjustments">Adjustments</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Adjustmen ts 对象，它包括指定形状的所有调整操作的调整值。应用于任何代表自选图形、艺术字或连接符的 ShapeRange 对象。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.AlternativeText">AlternativeText</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个当 ShapeRange 对象保存为网页时，该对象的描述性（可选）文本字符串。 String 型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.AutoShapeType">AutoShapeType</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定的 Shape 或 ShapeRange 对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。 MsoAutoShapeType 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.BackgroundStyle">BackgroundStyle</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置背景样式。可读/写 MsoBackgroundStyleIndex 类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.BlackWhiteMode">BlackWhiteMode</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个值，该值指明以黑白模式查看演示文稿时指定形状出现的方式。 MsoBlackWhiteMode ，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Callout">Callout</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 CalloutFor mat 对象，它包含指定形状的标注格式属性。应用于代表线形标注的 ShapeRange 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Chart">Chart</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Char t 对象，该对象代表形状区域中包含的图表。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Child">Child</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定的形状是子形状，或者如果形状区域中的所有形状都是同一个父形状的子形状，则返回 msoTrue 。 MsoTriState 类型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ConnectionSiteCount">ConnectionSiteCount</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定形状中的连接部位的数量。 Long 型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Connector">Connector</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定的形状是连接符，则此属性为 True 。 MsoTriState 类型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ConnectorFormat">ConnectorFormat</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个包含连接符格式属性的 ConnectorFo rmat 对象。应用于代表连接符的 ShapeRange 对象。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Count">Count</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Long 值，它代表集合中对象的数量。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">14</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">15</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Fill">Fill</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定形状的 FillFormat 对象或指定图表的 ChartFillFormat 对象，这些对象包含形状或图表的填充格式属性。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">16</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Glow">Glow</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个包含形状区域发光格式属性的指定形状区域的 GlowFormat 对象。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">17</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.GroupItems">GroupItems</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 GroupSha pes 对象，它代表指定形状组中的单个形状。使用 GroupShapes 对象的 Ite m 方法可从形状组中返回单个形状。应用于代表成组形状的 ShapeRange 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">18</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.HasChart">HasChart</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回形状区域是否包含图表。 MsoTriState 类型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">19</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Height">Height</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Single 值，它代表对象的高度（以磅为单位）。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">20</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.HorizontalFlip">HorizontalFlip</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定的形状绕水平对称轴翻转，则为 True 。 MsoTriState ，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">21</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ID">ID</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Long 值，它代表指定对象的类型。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">22</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Left">Left</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 Single 值，它代表从对象左边缘到工作表的 A 列左边缘或到图表上的图表区左边缘的距离（以磅为单位）。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">23</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Line">Line</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 LineFormat 对象，它包含指定形状的线条格式属性。对于线条， LineFormat 对象代表线条本身；而对于带有边框的形状， LineFormat 对象代表边框。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">24</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.LockAspectRatio">LockAspectRatio</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定的形状在调整大小时其原始比例保持不变，则此属性为 True 。如果调整大小时可以分别更改形状的高度和宽度，则此属性为 False 。 MsoTriState 类型，可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">25</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Name">Name</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 String 值，它代表对象的名称。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">26</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Nodes">Nodes</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Shape Nodes 集合，它代表指定形状的几何描述。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">27</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定对象的父对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">28</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ParentGroup">ParentGroup</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Shape 对象，它代表子形状或子形状区域的共同父形状。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">29</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.PictureFormat">PictureFormat</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 PictureFo rmat 对象，它包含指定形状的图片格式属性。应用于代表图片或 OLE 对象的 ShapeRange 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">30</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Reflection">Reflection</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定形状区域的 ReflectionFormat 对象，该对象包含形状区域的映像格式属性。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">31</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Rotation">Rotation</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设形状的旋转角度（以度为单位）。 Single 型，可读写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">32</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Shadow">Shadow</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个只读的 ShadowF ormat 对象，它包含指定形状的阴影格式属性。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">33</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ShapeStyle">ShapeStyle</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置 MsoShapeStyleIndex 类型的值，该值代表形状区域的形状样式。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">34</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.SoftEdge">SoftEdge</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定形状区域的 SoftEdgeFormat 对象，该对象包含形状区域的柔化边缘格式设置属性。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">35</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.TextEffect">TextEffect</a></td>
<td style="text-align: left;" data-valian="middle"><p>该属性返回一个 TextEffec tFormat 对象，它包含指定形状的文本效果格式属性。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">36</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.TextFrame">TextFrame</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 TextFram e 对象，它包含指定形状的对齐方式和定位属性。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">37</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.TextFrame2">TextFrame2</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 TextFra me2 对象，该对象包含指定形状区域的文本格式。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">38</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ThreeD">ThreeD</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 ThreeDForm at 对象，它包含指定形状的三维效果格式属性。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">39</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Title">Title</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置与指定的形状区域关联的可选文本的标题。可读/写。</p>
<p>返回String值</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">40</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Top">Top</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Single 值，它代表形状范围内最上面的形状的上边缘到工作表上边缘的距离（以磅为单位）。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">41</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Type">Type</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个代表形状类型的 MsoShapeType 值。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">42</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.VerticalFlip">VerticalFlip</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果指定的形状绕垂直坐标轴翻转，则此属性为 True 。 MsoTriState 类型，只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">43</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Vertices">Vertices</a></td>
<td style="text-align: left;" data-valian="middle"><p>将指定任意多边形形状的顶点（及贝塞尔曲线的控制点）坐标作为一系列 <span class="glossary" data-"appendpopup(this,'ofcoordinatepair_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">坐标对</span> 返回。可将此属性返回的数组用作 AddCu rve 方法或 AddP olyLine 方法的参数。 Variant 型，只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">44</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Visible">Visible</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 MsoTriState 值，它确定对象是否可见。可读写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">45</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.Width">Width</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置一个 Single 值，它代表对象的宽度（以磅为单位）。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">46</td>
<td style="text-align: left;" data-valian="middle"><a href="#ShapeRange.ZOrderPosition">ZOrderPosition</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定形状在 z-顺序中的位置。 Long 型，只读。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### ShapeRange.Align

对齐指定形状区域中的形状。

**语法**

**express.Align ( *AlignCmd* , *RelativeTo* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型    | 说明                                   |
|--------------|-----------|-------------|----------------------------------------|
| *AlignCmd*   | 必选      | MsoAlignCmd | 指定某个指定形状范围中形状的对齐方式。 |
| *RelativeTo* | 必选      | MsoTriState | 不在 ET 中使用。必须为 False。         |

**示例**

``` JavaScript
/*本示例将 myDocument 中指定范围的所有形状的左边与该范围内最左边的形状的左边对齐。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.SelectAll()
    Selection.ShapeRange.Align(msoAlignLefts, false)
}
```

### ShapeRange.Apply

应用通过 **PickUp** 方法复制的指定形状格式。

**语法**

**express.Apply ()**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*此示例复制 myDocument 上第一个形状的格式，然后将复制的格式应用到第二个形状上。*/
  let myDocument = Application.Worksheets.Item(1)
  myDocument.Shapes.Item(1).PickUp()
  myDocument.Shapes.Item(2).Apply()
}
```

### ShapeRange.Delete

删除对象。

**语法**

**express.Delete ()**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Distribute

水平或垂直地分布指定的形状区域中的各形状。

**语法**

**express.Distribute ( *DistributeCmd* , *RelativeTo* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称            | 必选/可选 | 数据类型         | 说明                                       |
|-----------------|-----------|------------------|--------------------------------------------|
| *DistributeCmd* | 必选      | MsoDistributeCmd | 指定该范围内的形状是水平分布还是垂直分布。 |
| *RelativeTo*    | 必选      | MsoTriState      | 不在 ET 中使用。必须为 False。             |

**示例**

``` JavaScript
function test(){
  /*本示例在 myDocument 上定义了一个包含所有自选形状对象的形状子集，然后水平地分布该子集中的形状。最左边的形状将保留在原位。*/
  let myDocument = Application.Worksheets.Item(1)
  let myShapes = myDocument.Shapes
  let numAutoShapes
  let autoShpArray = []
  let numShapes = myShapes.Count
  
  if(numShapes > 1) {
      numAutoShapes = 0
      for(i = 1 ;i <= numShapes; i++){
          if(myShapes.Item(i).Type == msoAutoShape) {
              numAutoShapes++
              autoShpArray[numAutoShapes] = myShapes.Item(i).Name
          }
      }
      if(numAutoShapes > 1) {
          let asRange = myShapes.Range(autoShpArray)
          asRange.Distribute(msoDistributeHorizontally, false)
      }
  }
}
```

### ShapeRange.Duplicate

复制对象，并返回对新复制对象的引用。

**语法**

**express.Duplicate ()**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Flip

绕指定形状的水平或垂直对称轴翻转该形状。

**语法**

**express.Flip ( *FlipCmd* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型   | 说明                             |
|-----------|-----------|------------|----------------------------------|
| *FlipCmd* | 必选      | MsoFlipCmd | 指定形状是水平翻转还是垂直翻转。 |

**示例**

``` JavaScript
function test(){
  /*此示例向 myDocument 中添加三角形，复制该三角形，然后垂直翻转所复制的三角形，并将其设置为红色。*/
  let myDocument = Application.Worksheets.Item(1)
  let shape = myDocument.Shapes.AddShape(msoShapeRightTriangle, 10, 10, 50, 50).Duplicate()
  shape.Fill.ForeColor.RGB = (255, 0, 0)
  shape.Flip(msoFlipVertical)
}
```

### ShapeRange.Group

将指定区域中的形状形成一组。

返回值： 一个代表组合形状的 **Shape** 对象。

**语法**

**express.Group ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

由于一组形 状作为单个形状处理，所以创建和分解形状组将更改 Shapes 集合中的项目数，并更改集合中受影响的项目之后的项目索引号。

Range 对象必须是数据透视表字段的数据区域中的单个单元格。如果试图对多个单元格应用该方法，将会失败（不显示错误消息）。

### ShapeRange.IncrementLeft

以指定磅数水平移动指定形状。

**语法**

**express.IncrementLeft ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                       |
|-------------|-----------|----------|----------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 以磅为单位指定形状水平移动的距离。正值使形状向右移动，负值使形状向左移动。 |

**示例**

``` JavaScript
function test(){
  /*此示例复制 myDocument 上的第一个形状，并为复制的形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/
  let myDocument = Application.Worksheets.Item(1)
  let duplicate = myDocument.Shapes.Item(1).Duplicate()
  duplicate.Fill.PresetTextured(msoTextureGranite)
  duplicate.IncrementLeft(70)
  duplicate.IncrementTop(-50)
  duplicate.IncrementRotation(30)
}
```

### ShapeRange.IncrementRotation

使指定的形状按指定度数值绕 Z 轴旋转。使用 **Rotation** 属性可设置形状的绝对转角。

**语法**

**express.IncrementRotation ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                         |
|-------------|-----------|----------|------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定形状的水平旋转量，以度为单位。正值为顺时针旋转形状，负值逆时针旋转形状。 |

**说明**

如果要将三维形状绕 X 轴或 Y 轴旋转，请分别使用 **I ncrementRotationX** 方法或 **IncrementRotationY** 方法。

**示例**

``` JavaScript
function test(){
  /*此示例复制 myDocument 上的第一个形状，并为复制的形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/
  let myDocument = Application.Worksheets.Item(1)
  let duplicate = myDocument.Shapes.Item(1).Duplicate()
  duplicate.Fill.PresetTextured(msoTextureGranite)
  duplicate.IncrementLeft(70)
  duplicate.IncrementTop(-50)
  duplicate.IncrementRotation(30)
}
```

### ShapeRange.IncrementTop

以指定磅数垂直移动指定形状。

**语法**

**express.IncrementTop ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                     |
|-------------|-----------|----------|--------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定形状对象垂直移动的距离，以磅为单位。正值将形状下移，负值将形状上移。 |

**示例**

``` JavaScript
function test(){
  /*此示例复制 myDocument 上的第一个形状，并为复制的形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角。*/
  let myDocument = Worksheets.Item(1)
  let duplicate = myDocument.Shapes.Item(1).Duplicate()
  duplicate.Fill.PresetTextured(msoTextureGranite)
  duplicate.IncrementLeft(70)
  duplicate.IncrementTop(-50)
  duplicate.IncrementRotation(30)
}
```

### ShapeRange.Item

从集合中返回一个对象。

返回值： 包含在集合中的一个 **Shape** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 对象的名称或索引号。 |

**示例**

``` JavaScript
function test()
{
    /*本示例对某个形状区域中第二个形状的 OnAction 属性进行设置。如果变量 sr 不代表 ShapeRange 对象，则本示例无效。*/
    let sr = Application.Worksheets.Item(1).Shapes.Range([1, 3])
    sr.Item(2).OnAction = "ShapeAction"
}
```

### ShapeRange.PickUp

复制指定形状的格式。使用 **Apply** 方法可将复制的格式应用到其他形状。

**语法**

**express.PickUp ()**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*此示例复制 myDocument 上第一个形状的格式，然后将复制的格式应用到第二个形状上。*/
  let myDocument = Application.Worksheets.Item(1)
  myDocument.Shapes.Item(1).PickUp()
  myDocument.Shapes.Item(2).Apply()
}
```

### ShapeRange.Regroup

将指定形状原先所属的形状区域重新组合。以单一的 **Shape** 对象返回重新组合的形状。

返回值类型： Shape

**语法**

**express.Regroup ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

**Regroup** 方法在指定的 **ShapeRange** 集合中查找原先曾被组合的形状，并仅还原所找到的第一个形状原来所属的形状组。因此，如果指定的形状子集所包含的各个图形原来曾属于不同的图形组，就只会恢复其中的一个图形组。

请注意，由于一组形状作为单个形状处理，因此将形状组合及取消组合都会改变 Shapes 集合中的项目数，而且还会改变集合中受影响的项目后面的项目索引号。

**示例**

``` JavaScript
/*本示例对当前活动窗口中选定的形状进行重新分组。如果这些形状并未进行过组合和取消组合操作，则本示例将失效。*/
Application.ActiveWindow.Selection.ShapeRange.Regroup()
```

### ShapeRange.RerouteConnections

此方法将重排连接在指定形状上的所有连接符；如果指定的形状是连接符，就重排该连接符。

**语法**

**express.RerouteConnections ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

重排连接符使其以最短的路径连接形状。重排时， **RerouteConnections** 方法可能会断开连接符的两端并将其重新连接到形状的其他位置。

如果该方法应用于一个连接符，则只重排该连接符；如果该方法应用于一个已连接的形状，则重排该形状上所有的连接符。

**示例**

``` JavaScript
function test(){
  /*本示例将两个矩形添加到 myDocument，用曲线连接符连接两个矩形，然后重排连接符使两个矩形间采用最短的路径。请注意，RerouteConnections 方法用于调整连接符的大小和位置并决定要连接到哪个连接结点，因此，最初为 BeginConnect 和 EndConnect 方法的 ConnectionSite 参数指定的值是不相关的。*/
  let myDocument = ApplicationWorksheets.Item(1)
  let s = myDocument.Shapes
  let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
  let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
  let newConnector = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100)
  newConnector.ConnectorFormat.BeginConnect(firstRect, 1)
  newConnector.ConnectorFormat.EndConnect(secondRect, 1)
  newConnector.RerouteConnections()
}
```

### ShapeRange.ScaleHeight

按指定的比例调整形状的高度。对于图片和 OLE 对象，可以指定是相对于原有尺寸还是相对于当前尺寸来调整该形状。对于不是图片和 OLE 对象的形状，总是相对于其当前大小来调整高度。

**语法**

**express.ScaleHeight ( *Factor* , *RelativeToOriginalSize* , *Scale* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型    | 说明                                                                                                                                                                     |
|--------------------------|-----------|-------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Factor*                 | 必选      | Single      | 指定形状调整后的高度与当前或原始高度的比例。例如，若要将一个矩形放大百分之五十，请将此参数设为 1.5。                                                                     |
| *RelativeToOriginalSize* | 必选      | MsoTriState | 如果为 msoTrue，则相对于形状的原有尺寸来调整高度。如果该值为 msoFalse，则相对于形状的当前尺寸来调整高度。仅当指定的形状是图片或 OLE 对象时，才能将此参数指定为 msoTrue。 |
| *Scale*                  | 可选      | Variant     | MsoScaleFrom 的常量之一，它指定调整形状大小时，该形状哪一部分的位置将保持不变。                                                                                          |

**说明**

|                                                       |
|-------------------------------------------------------|
| **MsoTriState** 可以是下列 MsoTriState 常量之一。     |
| **msoCTrue** 。不应用于此属性。                       |
| **msoFalse** 。相对于形状的当前尺寸来调整形状的大小。 |
| **msoTriStateMixed** 。不应用于此属性。               |
| **msoTriStateToggle** 。不应用于此属性。              |
| **msoTrue** 。相对于形状的初始尺寸来调整形状的大小。  |

**示例**

``` JavaScript
function test(){
  /*此示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/
  let myDocument = Application.Worksheets.Item(1)
  let s = myDocument.Shapes     
  for(let i = 1; i <= s.Count; i++) {
      switch(s.Item(i).Type) {
      case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture:
          s.Item(i).ScaleHeight(1.75, msoTrue)
          s.Item(i).ScaleWidth(1.75, msoTrue)
          break
      default:
          s.Item(i).ScaleHeight(1.75, msoFalse)
          s.Item(i).ScaleWidth(1.75, msoFalse)
          break
    }
  }
}
```

### ShapeRange.ScaleWidth

按指定的比例调整形状的宽度。对于图片和 OLE 对象，可以指定是相对于原有尺寸还是相对于当前尺寸来调整该形状。对于不是图片和 OLE 对象的形状，总是相对于其当前大小来调整宽度。

**语法**

**express.ScaleWidth ( *Factor* , *RelativeToOriginalSize* , *Scale* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型    | 说明                                                                                                         |
|--------------------------|-----------|-------------|--------------------------------------------------------------------------------------------------------------|
| *Factor*                 | 必选      | Single      | 指定形状调整后的宽度与当前或原始宽度的比例。例如，若要将一个矩形放大百分之五十，请将此参数设为 1.5。         |
| *RelativeToOriginalSize* | 必选      | MsoTriState | 如果为 False，则相对于形状的原有尺寸来调整宽度。仅当指定的形状是图片或 OLE 对象时，才能将此参数指定为 True。 |
| *Scale*                  | 可选      | Variant     | MsoScaleFrom 的常量之一，它指定调整形状大小时，该形状哪一部分的位置将保持不变。                              |

**说明**

|                                                                   |
|-------------------------------------------------------------------|
| **MsoTriState** 可为以下 MsoTriState 常量之一。                   |
| **msoCTrue** 。不应用于此属性。                                   |
| **msoFalse** 。相对于形状的当前尺寸来调整形状的大小。             |
| **msoTriStateMixed** 。不应用于此属性。                           |
| **msoTriStateToggle** 。不应用于此属性。                          |
| **msoTrue** 。仅当指定的形状是图片或 OLE 对象时，才能使用此参数。 |

**示例**

``` JavaScript
function test(){
  /*本示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%。*/
  let myDocument = Application.Worksheets.Item(1)
  let s = myDocument.Shapes     
  for(let i = 1; i <= s.Count; i++) {
      switch(s.Item(i).Type) {
      case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoLinkedPicture, msoPicture:
          s.Item(i).ScaleHeight(1.75, msoTrue)
          s.Item(i).ScaleWidth(1.75, msoTrue)
          break
      default:
          s.Item(i).ScaleHeight(1.75, msoFalse)
          s.Item(i).ScaleWidth(1.75, msoFalse)
          break
    }
  }
}
```

### ShapeRange.Select

选择对象。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                                                                              |
|-----------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------|
| *Replace* | 可选      | Variant  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

### ShapeRange.SetShapesDefaultProperties

将指定形状的格式设置为形状的默认格式。

**语法**

**express.SetShapesDefaultProperties ()**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*此示例向 myDocument 中添加矩形，设置该矩形的填充样式，将矩形的格式设置为该矩形的默认格式，然后在文档中添加另一个较小的矩形。第二个矩形的填充格式与第一个矩形一样。*/
  let myDocument = Application.Worksheets.Item(1)
  let s = myDocument.Shapes
  let a =s.AddShape(msoShapeRectangle, 5, 5, 80, 60)
  let f = a.Fill
  f.ForeColor.RGB = (0, 0, 255)
  f.BackColor.RGB = (0, 204, 255)
  f.Patterned(msoPatternHorizontalBrick)
  
  // Set formatting as default formatting
  a.SetShapesDefaultProperties()
  
  // Create new shape with default formatting
  s.AddShape(msoShapeRectangle, 90, 90, 40, 30)
}
```

### ShapeRange.Ungroup

取消指定形状或者形状区域中组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。

返回值： 一个 **ShapeRange** 对象，它代表取消组合的形状。

**语法**

**express.Ungroup ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

由于一组形状被视为单个对象，因此对形状进行组合或取消组合操作时，将会更改 **Shapes** 集合中的项目数，并更改集合中受影响的项目之后的项目索引号。

**示例**

``` JavaScript
/*此示例取消 myDocument 中的所有形状组合，并分解所有的图片和 OLE 对象。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    let s = myDocument.Shapes
    for(let i = 1; i <= s.Count; i++) {
        s.Item(i).Ungroup()
    }
}

/*此示例取消 myDocument 中的所有形状组合，但并不取消文档中图片和 OLE 对象的组合。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    let s = myDocument.Shapes
    for(let i = 1; i <= s.Count; i++) {
        if(s.Item(i).Type == msoGroup) {
            s.Item(i).Ungroup()
        }
    }
}
```

### ShapeRange.ZOrder

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-次序中的位置）。

**语法**

**express.ZOrder ( *ZOrderCmd* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型     | 说明                                       |
|-------------|-----------|--------------|--------------------------------------------|
| *ZOrderCmd* | 必选      | MsoZOrderCmd | 指定根据其他形状将指定的形状移到什么位置。 |

**说明**

|                                                   |
|---------------------------------------------------|
| **MsoZOrderCmd** 可为以下 MsoZOrderCmd 常量之一。 |
| **msoBringForward**                               |
| **msoBringInFrontOfText** 。 仅用于 WPS。         |
| **msoBringToFront**                               |
| **msoSendBackward**                               |
| **msoSendBehindText** 。仅用于 WPS。              |
| **msoSendToBack**                                 |

使用 **ZOrderPosition** 属性可确定形状在 z-次序中的当前位置。

**示例**

## 成员属性

### ShapeRange.Adjustments

返回一个 **Adjustmen ts** 对象，它包括指定形状的所有调整操作的调整值。应用于任何代表自选图形、艺术字或连接符的 **ShapeRange** 对象。

**语法**

**express.Adjustments**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.AlternativeText

返回或设置一个当 **ShapeRange** 对象保存为网页时，该对象的描述性（可选）文本字符串。 **String** 型，可读写。

**语法**

**express.AlternativeText**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

可选文字可以显示在 Web 浏览器中某个形状的图像位置上，或者当鼠标指针停留在图像上时，这些文本将直接显示在图像之上（在支持这些功能的浏览器中）。

### ShapeRange.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test()
{
    /*本示例显示一条有关创建 myObject 的应用程序的消息。*/
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### ShapeRange.AutoShapeType

返回或设置指定的 **Shape** 或 **ShapeRange** 对象的形状类型，该对象必须代表自选图形，而不能代表直线、任意多边形图形或连接符。 **MsoAutoShapeType** 类型，可读写。

**语法**

**express.AutoShapeType**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

改变一个形状的类型时，该形状保留其大小、颜色和其他属性。

使用 **Conn ectorFormat** 对象的 **Typ e** 属性设置或返回连接符类型。

**示例**

``` JavaScript
/*本示例将 myDocument 中的所有 16 角星都替换为 32 角星。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    for(let i=1;i <= myDocument.Shapes.Count;i++) {
        if(myDocument.Shapes.Item(i).AutoShapeType == msoShape16pointStar) {
            myDocument.Shapes.Item(i).AutoShapeType = msoShape32pointStar
        }
    }
}
```

### ShapeRange.BackgroundStyle

返回或设置背景样式。可读/写 **MsoBackgroundStyleIndex** 类型。

**语法**

**express.BackgroundStyle**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.BlackWhiteMode

返回或设置一个值，该值指明以黑白模式查看演示文稿时指定形状出现的方式。 **MsoBlackWhiteMode** ，可读写。

**语法**

**express.BlackWhiteMode**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*此示例使 wksOne 上的第一个形状以黑白模式显示。当用户以黑白模式查看演示文稿时，第一个形状为黑色，而不管它在颜色模式中为何种颜色。*/
function UseBlackWhiteMode() {
    let wksOne = Application.Worksheets.Item(1)
    wksOne.Shapes.Item(1).BlackWhiteMode = msoBlackWhiteGrayOutline
}
```

### ShapeRange.Callout

返回一个 **CalloutFor mat** 对象，它包含指定形状的标注格式属性。应用于代表线形标注的 **ShapeRange** 对象。只读。

**语法**

**express.Callout**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加椭圆和指向该椭圆的标注。该标注的文字没有边框，但用垂直的强调线分开标注文字和标注线。*/
function test(){
  let myDocument = Application.Worksheets.Item(1)
  let myShape = myDocument.Shapes
  myShape.AddShape(msoShapeOval, 180, 200, 280, 130)
  let Add = myShape.AddCallout(msoCalloutTwo, 420, 170, 170, 40)
  Add.TextFrame.Characters().Text = "My oval"
  let Call = Add.Callout
  Call.Accent = true
  Call.Border = false
}
```

### ShapeRange.Chart

返回 **Char t** 对象，该对象代表形状区域中包含的图表。只读。

**语法**

**express.Chart**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Child

如果指定的形状是子形状，或者如果形状区域中的所有形状都是同一个父形状的子形状，则返回 **msoTrue** 。 **MsoTriState** 类型，只读。

**语法**

**express.Child**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

|                                                              |
|--------------------------------------------------------------|
| **MsoTriState** 可为以下 **MsoTriState** 常量之一:           |
| **msoCTrue** 。不应用于此属性。                              |
| 如果选择的形状不是子形状，则为 **msoFalse** 。               |
| 如果只有一些选择的形状是子形状，则为 **msoTriStateMixed** 。 |
| **msoTriStateToggle** 。不应用于此属性。                     |
| 如果选择的形状是子形状，则为 **msoTrue** 。                  |

**示例**

``` JavaScript
/*本示例选择画布中的第一个形状，并且如果选择的形状是子形状，则用指定的颜色填充该形状。本示例假定活动工作表上的一个绘图画布中包含多个形状。*/
function FillChildShape() {
    //Select the first shape in the drawing canvas.
    Application.ActiveSheet.Shapes.Item(1).CanvasItems(1).Select()
    //Fill selected shape if it is a child shape.
    if(Selection.ShapeRange.Child == msoTrue) {
        Selection.ShapeRange.Fill.ForeColor.RGB = (100, 0, 200)
    }
    else {
        MsgBox("This shape is not a child shape.")
    }
}
```

### ShapeRange.ConnectionSiteCount

返回指定形状中的连接部位的数量。 **Long** 型，只读。

**语法**

**express.ConnectionSiteCount**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*此示例向 myDocument 中添加两个矩形，并且用两个连接符连接这两个矩形。两个连接符的起始点都连到第一个矩形上的第一个连接部位；连接符的终止点分别连到第二个矩形上的第一个和最后一个连接部位上。*/
  let myDocument = Application.Worksheets.Item(1)
  let s = myDocument.Shapes
  let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
  let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
  let lastsite = secondRect.ConnectionSiteCount
  let Con1 = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100).ConnectorFormat
  Con1.BeginConnect(firstRect, 1)
  Con1.EndConnect(secondRect, 1)
  let Con2 = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100).ConnectorFormat
  Con2.BeginConnect(firstRect, 1)
  Con2.EndConnect(secondRect, lastsite)
}
```

### ShapeRange.Connector

如果指定的形状是连接符，则此属性为 **True** 。 **MsoTriState** 类型，只读。

**语法**

**express.Connector**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*本示例删除了 myDocument 中所有的连接符。*/
  let myDocument = Application.Worksheets.Item(1)
  let myShape = myDocument.Shapes
  for(let i=myShape.Count;i >= 1;i--) { 
      if(myShape.Item(i).Connector) {
          myShape.Item(i).Delete()
      }
  }
}
```

### ShapeRange.ConnectorFormat

返回一个包含连接符格式属性的 **ConnectorFo rmat** 对象。应用于代表连接符的 **ShapeRange** 对象。只读。

**语法**

**express.ConnectorFormat**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*本示例向 myDocument 中添加两个矩形，用一个连接符连接这两个矩形，并自动将连接符路径修改为最短路径，然后断开矩形间的连接符。*/
  let myDocument = Application.Worksheets.Item(1)
  let s = myDocument.Shapes
  let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
  let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
  let c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0)
  let Con = c.ConnectorFormat
  Con.BeginConnect(firstRect, 1)
  Con.EndConnect(secondRect, 1)
  c.RerouteConnections()
  Con.BeginDisconnect()
  Con.EndDisconnect()
}
```

### ShapeRange.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### ShapeRange.Fill

返回指定形状的 **FillFormat** 对象或指定图表的 **ChartFillFormat** 对象，这些对象包含形状或图表的填充格式属性。只读。

**语法**

**express.Fill**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*本示例向 myDocument 中添加矩形，然后设置该矩形的填充格式的前景色、背景色和渐变。*/
function test()
{
    let myDocument = Application.Worksheets.Item(1)
    let myShape = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill
    myShape.ForeColor.RGB = (128, 0, 0)
    myShape.BackColor.RGB = (170, 170, 170)
    myShape.TwoColorGradient(msoGradientHorizontal, 1)
}
```

### ShapeRange.Glow

返回一个包含形状区域发光格式属性的指定形状区域的 **GlowFormat** 对象。只读。

**语法**

**express.Glow**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

发光效果为图形添加了色彩鲜明的彩色边缘。

### ShapeRange.GroupItems

返回一个 **GroupSha pes** 对象，它代表指定形状组中的单个形状。使用 **GroupShapes** 对象的 **Ite m** 方法可从形状组中返回单个形状。应用于代表成组形状的 **ShapeRange** 对象。只读。

**语法**

**express.GroupItems**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*此示例向 myDocument 中添加三个三角形，将它们组合，再设置整个组的颜色，然后只更改第二个三角形的颜色。*/
  let myDocument = Application.Worksheets.Item(1)
  let myShape = myDocument.Shapes
  myShape.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"
  myShape.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"
  myShape.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"
  let myGroup = myShape.Range(["shpOne", "shpTwo", "shpThree"]).Group()
  myGroup.Fill.PresetTextured(msoTextureBlueTissuePaper)
  myGroup.GroupItems.Item(2).Fill.PresetTextured(msoTextureGreenMarble)
}
```

### ShapeRange.HasChart

返回形状区域是否包含图表。 **MsoTriState** 类型，只读。

**语法**

**express.HasChart**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Height

返回或设置一个 **Single** 值，它代表对象的高度（以磅为单位）。

**语法**

**express.Height**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.HorizontalFlip

如果指定的形状绕水平对称轴翻转，则为 **True** 。 **MsoTriState** ，只读。

**语法**

**express.HorizontalFlip**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
funcion test(){
  /*如果 myDocument 中的形状进行过水平或垂直翻转，则此示例将恢复其原始状态。*/
  let myDocument = Application.Worksheets.Item(1)
  for(let i=1;i <= myDocument.Shapes.Count;i++) {
      if(myDocument.Shapes.Item(i).HorizontalFlip) {
          myDocument.Shapes.Item(i).Flip(msoFlipHorizontal)
      }
      if(myDocument.Shapes.Item(i).VerticalFlip) {
          myDocument.Shapes.Item(i).Flip(msoFlipVertical)
      }
  }
}
```

### ShapeRange.ID

返回一个 Long 值，它代表指定对象的类型。

**语法**

**express.ID**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

可将 ID 标签用作其他 HTML 文档或相同网页上的超链接引用。

### ShapeRange.Left

返回或设置 **Single** 值，它代表从对象左边缘到工作表的 A 列左边缘或到图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Line

返回一个 **LineFormat** 对象，它包含指定形状的线条格式属性。对于线条， **LineFormat** 对象代表线条本身；而对于带有边框的形状， **LineFormat** 对象代表边框。只读。

**语法**

**express.Line**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*此示例将一条蓝色虚线添至 myDocument。*/
  let myDocument = Application.Worksheets.Item(1)
  let myLine = myDocument.Shapes.AddLine(10, 10, 250, 250).Line
  myLine.DashStyle = msoLineDashDotDot
  myLine.ForeColor.RGB = (50, 0, 128)
  
  /*此示例向 myDocument 添加一个十字形，然后将其边框设为红色、8 磅宽。*/
  let myDocument = Application.Worksheets.Item(1)
  let myLine = myDocument.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line
  myLine.Weight = 8
  myLine.ForeColor.RGB = 255, 0, 0
  
}
```

### ShapeRange.LockAspectRatio

如果指定的形状在调整大小时其原始比例保持不变，则此属性为 **True** 。如果调整大小时可以分别更改形状的高度和宽度，则此属性为 **False** 。 **MsoTriState** 类型，可读写。

**语法**

**express.LockAspectRatio**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

|                                                          |
|----------------------------------------------------------|
| **MsoTriState** 可为以下 **MsoTriState** 常量之一:       |
| **msoCTrue**                                             |
| **msoFalse** 。调整大小时可以分别更改形状的高度和宽度。  |
| **msoTriStateMixed**                                     |
| **msoTriStateToggle**                                    |
| **msoTrue** 。指定的形状在调整大小时其原始比例保持不变。 |

**示例**

``` JavaScript
function test(){
  /*本示例向 myDocument 中添加一个立方体。可移动并调整该立方体，但不能重新设置其比例。*/
  let myDocument = Application.Worksheets.Item(1)
  myDocument.Shapes.AddShape(msoShapeCube, 50, 50, 100, 200).LockAspectRatio = msoTrue
}
```

### ShapeRange.Name

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Nodes

返回一个 **Shape Nodes** 集合，它代表指定形状的几何描述。

**语法**

**express.Nodes**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

此属性适用于代表任意多边形的 **Shape** 或 **ShapeRange** 对象。

**示例**

``` JavaScript
function test(){
  /*本示例在 myDocument 中第三个形状的第四个结点后添加一个带有一段曲线的平滑结点。第三个形状必须是至少有四个结点的任意多边形。*/
  let myDocument = Application.Worksheets.Item(1)
  let myNode = myDocument.Shapes.Item(3).Nodes
  myNode.Insert(4, msoSegmentCurve, msoEditingSmooth, 210, 100)
}
```

### ShapeRange.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.ParentGroup

返回 **Shape** 对象，它代表子形状或子形状区域的共同父形状。

**语法**

**express.ParentGroup**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*在此示例中，ET 向活动工作表添加两个形状，然后通过删除组合的父形状来删除这两个形状。*/
function ParentGroup() {
    let myShape = Application.ActiveSheet.Shapes
    myShape.AddShape(1, 10, 10, 100, 100)
    myShape.AddShape(2, 110, 120, 100, 100)
    myShape.Range([1, 2]).Group()
    // Using the child shape in the group get the Parent shape.
    let pgShape = ActiveSheet.Shapes.Item(1).GroupItems.Item(1).ParentGroup
    alert("The two shapes will now be deleted.")
    // Delete the parent shape.
    pgShape.Delete()
}
```

### ShapeRange.PictureFormat

返回一个 **PictureFo rmat** 对象，它包含指定形状的图片格式属性。应用于代表图片或 OLE 对象的 **ShapeRange** 对象。只读。

**语法**

**express.PictureFormat**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*本示例设置了 myDocument 中第一个形状的亮度和对比度。第一个形状必须是图片或 OLE 对象。*/
  let myDocument = Application.Worksheets.Item(1)
  let Pic = myDocument.Shapes.Item(1).PictureFormat
  Pic.Brightness = 0.3
  Pic.Contrast = .75
}
```

### ShapeRange.Reflection

返回指定形状区域的 **ReflectionFormat** 对象，该对象包含形状区域的映像格式属性。只读。

**语法**

**express.Reflection**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Rotation

返回或设形状的旋转角度（以度为单位）。 **Single** 型，可读写。

**语法**

**express.Rotation**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

旋转量始终四舍五入到最接近的整数。

### ShapeRange.Shadow

返回一个只读的 **ShadowF ormat** 对象，它包含指定形状的阴影格式属性。

**语法**

**express.Shadow**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.ShapeStyle

返回或设置 **MsoShapeStyleIndex** 类型的值，该值代表形状区域的形状样式。可读/写。

**语法**

**express.ShapeStyle**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.SoftEdge

返回指定形状区域的 **SoftEdgeFormat** 对象，该对象包含形状区域的柔化边缘格式设置属性。只读。

**语法**

**express.SoftEdge**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.TextEffect

该属性返回一个 **TextEffec tFormat** 对象，它包含指定形状的文本效果格式属性。

**语法**

**express.TextEffect**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*此示例中，如果 myDocument 的第三个形状为艺术字，则将它的字体样式设置为加粗。*/
  let myDocument = Application.Worksheets.Item(1)
  let myShape = myDocument.Shapes.Item(3)
  if(myShape.Type == msoTextEffect) {
      myShape.TextEffect.FontBold = true
  }
}
```

### ShapeRange.TextFrame

返回一个 **TextFram e** 对象，它包含指定形状的对齐方式和定位属性。

**语法**

**express.TextFrame**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*此示例将形状一中文本框内的文本设置为两端对齐。如果形状一中不包含文本框，则此示例无效。*/
Application.Worksheets.Item(1).Shapes.Item(1).TextFrame.HorizontalAlignment = xlHAlignJustify
```

### ShapeRange.TextFrame2

返回一个 **TextFra me2** 对象，该对象包含指定形状区域的文本格式。只读。

**语法**

**express.TextFrame2**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.ThreeD

返回一个 **ThreeDForm at** 对象，它包含指定形状的三维效果格式属性。只读。

**语法**

**express.ThreeD**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*本例设置应用于 myDocument 中第一个形状的三维效果：深度、延伸色、延伸方向以及照明方向。*/
  let myDocument = Application.Worksheets.Item(1)
  let myThree = myDocument.Shapes.Item(1).ThreeD
  myThree.Visible = true
  myThree.Depth = 50
  myThree.ExtrusionColor.RGB = (255, 100, 255)
  // RGB value for purple
  myThree.SetExtrusionDirection(msoExtrusionTop)
  myThree.PresetLightingDirection = msoLightingLeft
}
```

### ShapeRange.Title

返回或设置与指定的形状区域关联的可选文本的标题。可读/写。

返回String值

**语法**

**express.Title**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

若要设置形状区域的可选文本的描述性文本字符串，请使用 **Alternativ eText** 属性。

### ShapeRange.Top

返回或设置一个 **Single** 值，它代表形状范围内最上面的形状的上边缘到工作表上边缘的距离（以磅为单位）。

**语法**

**express.Top**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Type

返回一个代表形状类型的 **MsoShapeType** 值。

**语法**

**express.Type**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.VerticalFlip

如果指定的形状绕垂直坐标轴翻转，则此属性为 **True** 。 **MsoTriState** 类型，只读。

**语法**

**express.VerticalFlip**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
function test(){
  /*如果 myDocument 中的形状进行过水平或垂直翻转，则本示例将恢复其原始状态。*/
  let myDocument = Application.Worksheets.Item(1)
  for(let i=1;i <= myDocument.Shapes.Count;i++) {
      if(myDocument.Shapes.Item(i).HorizontalFlip) {
          myDocument.Shapes.Item(i).Flip(msoFlipHorizontal)
      }
      if(myDocument.Shapes.Item(i).VerticalFlip) {
          myDocument.Shapes.Item(i).Flip(msoFlipVertical)
      }
  }
}
```

### ShapeRange.Vertices

将指定任意多边形形状的顶点（及贝塞尔曲线的控制点）坐标作为一系列 <span class="glossary" "appendpopup(this,'ofcoordinatepair_1_1')"="" style="color: rgb(0, 0, 0); text-decoration-line: none; font-family: &quot;Microsoft YaHei&quot;, Tahoma, Arial, Helvetica, sans-serif; white-space: normal;">坐标对</span> 返回。可将此属性返回的数组用作 **AddCu rve** 方法或 **AddP olyLine** 方法的参数。 **Variant** 型，只读。

**语法**

**express.Vertices**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

下表显示 **Vertices** 属性数组 ` vertArray() ` 中的值如何与三角形的 vertices 坐标相关联。

| vertArray 元素      | 内容                                   |
|---------------------|----------------------------------------|
| ` vertArray(1, 1) ` | 第一个顶点与文档的左边界之间的水平距离 |
| ` vertArray(1, 2) ` | 第一个顶点与文档的顶端之间的垂直距离   |
| ` vertArray(2, 1) ` | 第二个顶点与文档的左边界之间的水平距离 |
| ` vertArray(2, 2) ` | 第二个顶点与文档的顶端之间的垂直距离   |
| ` vertArray(3, 1) ` | 第三个顶点与文档的左边界之间的水平距离 |
| ` vertArray(3, 2) ` | 第三个顶点与文档的顶端之间的垂直距离   |

**示例**

``` JavaScript
function test(){
  /*此示例将 myDocument 上第一个形状的顶点坐标分配给数组变量 vertArray()，并显示第一个顶点的坐标。*/
  let myDocument = Application.Worksheets.Item(1)
  let myShape = myDocument.Shapes.Item(1)
  let vertArray = [myShape.Vertices]
  x1 = vertArray[1, 1]
  y1 = vertArray[1, 2]
  alert("First vertex coordinates: " + x1 + ", " + y1)
  
}
function test(){
  /*此示例创建一个曲线，该曲线的几何说明与myDocument中第一个形状相同。此示例要求第一个形状必须包含3n+1个顶点。*/
  let myDocument = Worksheets.Item(1)
  let myShape = myDocument.Shapes
  myShape.AddCurve(myShape.Item(1).Vertices)
  
}
```

### ShapeRange.Visible

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Width

返回或设置一个 **Single** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.ZOrderPosition

返回指定形状在 z-顺序中的位置。 **Long** 型，只读。

**语法**

**express.ZOrderPosition**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

若要设置形状在 z-顺序中的位置，请用 **ZOrder** 方法。

形状在 z-顺序中的位置对应于在 Shapes 集合中的索引号。例如，如果在 ` myDocument ` 中有四个形状，则 ` myDocument.Shapes(1) ` 表达式返回 z-顺序中最后一个形状，表达式 ` myDocument.Shapes(4) ` 返回 z-顺序中第一个形状。

每次向集合中添加一个新的形状时，默认情况下它会被添加到 z-次序的最前端。

**示例**

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ShapeRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Sort 对象

## 简介

代表数据区域的排序方式。

## 说明

以下示例在活动工作表中生成一个数据区域并对这些数据排序。

``` JavaScript
function test(){
    //Building data to sort on the active sheet.
    Range("A1").Value2 = "Name"
    Range("A2").Value2 = "Bill"
    Range("A3").Value2 = "Rod"
    Range("A4").Value2 = "John"
    Range("A5").Value2 = "Paddy"
    Range("A6").Value2 = "Kelly"
    Range("A7").Value2 = "William"
    Range("A8").Value2 = "Janet"
    Range("A9").Value2 = "Florence"
    Range("A10").Value2 = "Albert"
    Range("A11").Value2 = "Mary"
    alert("The list is out of order.  Hit Ok to continue..." + jsInformation)
                                    
    //Selecting a cell within the range.
    Range("A2").Select()
                                    
    //Applying sort.
    let sor = Application.ActiveWorkbook.Worksheets.Item(ActiveSheet.Name).Sort
    sor.SortFields.Clear()
    sor.SortFields.Add(Range("A2:A11"), xlSortOnValues, xlAscending, xlSortNormal)
    sor.SetRange(Range("A1:A11"))
    sor.Header = xlYes
    sor.MatchCase = false
    sor.Orientation = xlTopToBottom
    sor.SortMethod = xlPinYin
    sor.Apply()
    alert("Sort complete." + jsInformation)                              
}
```

## 方法

| 序号 | 名称                       | 说明                                       |
|:----:|:---------------------------|:-------------------------------------------|
|  1   | [Apply](#Sort.Apply)       | 根据当前应用的排序状态对区域进行排序。     |
|  2   | [SetRange](#Sort.SetRange) | 设置 Sort 对象的起始字符和结束字符的位置。 |

## 属性

| 序号 | 名称                             | 说明                                                                                                                                                               |
|:----:|:---------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#Sort.Application) | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 Application 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 Application 对象。只读。 |
|  2   | [Creator](#Sort.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                         |
|  3   | [Header](#Sort.Header)           | 指定第一行是否包含标题信息。可读/写 XlYesNoGuess 类型。                                                                                                            |
|  4   | [MatchCase](#Sort.MatchCase)     | 如果设置为 True ，则执行区分大小写的排序，如果设置为 False ，则执行不区分大小写的排序。可读/写。                                                                   |
|  5   | [Orientation](#Sort.Orientation) | 指定排序方向。可读/写 XlSortOrientation 类型。                                                                                                                     |
|  6   | [Parent](#Sort.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                       |
|  7   | [Rng](#Sort.Rng)                 | 返回要执行排序的值的区域。只读。                                                                                                                                   |
|  8   | [SortFields](#Sort.SortFields)   | 可用来在工作簿、列表和自动筛选上存储排序状态。只读 SortFields 类型。                                                                                               |
|  9   | [SortMethod](#Sort.SortMethod)   | 指定中文排序方法。可读/写 XlSortMethod 类型。                                                                                                                      |

## 成员方法

### Sort.Apply

根据当前应用的排序状态对区域进行排序。

**语法**

**express.Apply ()**

\*express是一个代表 Sort 对象的变量。

### Sort.SetRange

设置 **Sort** 对象的起始字符和结束字符的位置。

**语法**

**express.SetRange ( *Rng* )**

\*express是一个代表 Sort 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明                   |
|-------|-----------|----------|------------------------|
| *Rng* | 必选      | Range    | 指定 Sort 集合的区域。 |

**说明**

仅当对工作表区域应用排序时才能使用 **SetRange** ，当区域在表内时不能使用它。

## 成员属性

### Sort.Application

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 Sort 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

### Sort.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Sort 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### Sort.Header

指定第一行是否包含标题信息。可读/写 **XlYesNoGuess** 类型。

**语法**

**express.Header**

\*express是一个代表 Sort 对象的变量。

**说明**

默认值为 **xlNo** 。如果希望 ET 确定标题，可以指定 **xlGuess** 。

### Sort.MatchCase

如果设置为 **True** ，则执行区分大小写的排序，如果设置为 **False** ，则执行不区分大小写的排序。可读/写。

**语法**

**express.MatchCase**

\*express是一个代表 Sort 对象的变量。

### Sort.Orientation

指定排序方向。可读/写 **XlSortOrientation** 类型。

**语法**

**express.Orientation**

\*express是一个代表 Sort 对象的变量。

**说明**

|                                                              |
|--------------------------------------------------------------|
| XlSortOrientation 可以是下列 **XlSortOrientation** 常量之一: |
| **xlSortColumns**                                            |
| **xlSortRows**                                               |

### Sort.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 Sort 对象的变量。

### Sort.Rng

返回要执行排序的值的区域。只读。

**语法**

**express.Rng**

\*express是一个代表 Sort 对象的变量。

### Sort.SortFields

可用来在工作簿、列表和自动筛选上存储排序状态。只读 **SortFields** 类型。

**语法**

**express.SortFields**

\*express是一个代表 Sort 对象的变量。

### Sort.SortMethod

指定中文排序方法。可读/写 **XlSortMethod** 类型。

**语法**

**express.SortMethod**

\*express是一个代表 Sort 对象的变量。

**说明**

|                                                  |
|--------------------------------------------------|
| XlSortMethod 可以是以下 **SortMethod** 常量之一: |
| **xlStroke**                                     |
| **xlPinYin**                                     |

> WPS 开发文档： [WPS 基础接口/表格 API 参考/Sort](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# UserAccessList 对象

## 简介

代表受保护区域用户访问权限的 UserAccess 对象的集合。

## 说明

使用受保护的 Range 对象的 Users 属性可返回一个 UserAccessList 集合。

一旦返回了 UserAccessList 集合，您就可以使用 Count 属性来确定能够访问受保护区域的用户的数量。在下例中，ET 通知用户能够访问第一个受保护区域的用户的数量。本示例假定活动工作表中存在受保护区域。

``` JavaScript
function test(){
    let wksSheet = Application.ActiveSheet
                                    
    //Notify the user the number of users that can access the protected range.
    alert(wksSheet.Protection.AllowEditRanges.Item(1).Users.Count)
}
```

## 方法

| 序号 | 名称                                   | 说明                                         |
|:----:|:---------------------------------------|:---------------------------------------------|
|  1   | [Add](#UserAccessList.Add)             |                                              |
|  2   | [DeleteAll](#UserAccessList.DeleteAll) | 删除有权访问工作表上受保护的区域的所有用户。 |

## 属性

| 序号 | 名称                           | 说明                                       |
|:----:|:-------------------------------|:-------------------------------------------|
|  1   | [Count](#UserAccessList.Count) | 返回一个 Long 值，它代表集合中对象的数量。 |
|  2   | [Item](#UserAccessList.Item)   | 从集合中返回一个对象。                     |

## 成员方法

### UserAccessList.Add

**语法**

**express.Add ( *Name* , *AllowEdit* )**

\*express是一个代表 UserAccessList 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                      |
|-------------|-----------|----------|---------------------------------------------------------------------------|
| *Name*      | 必选      | String   | 用户访问列表的名称。                                                      |
| *AllowEdit* | 必选      | Boolean  | 如果为 True，则允许访问列表中的用户编辑受保护工作表上的可编辑单元格区域。 |

**返回值**

一个代表新的用户访问列表的 UserAccess 对象。

**说明**

添加用户访问列表。

### UserAccessList.DeleteAll

删除有权访问工作表上受保护的区域的所有用户。

**语法**

**express.DeleteAll ()**

\*express是一个代表 UserAccessList 对象的变量。

**示例**

``` JavaScript
//在本示例中，ET 将删除有权访问活动工作表中第一个受保护的区域的所有用户。本示例假定工作表不受保护。
function test()
{
    let wksSheet = Application.ActiveSheet
                                    
    //Remove all users with access to the first protected range.
    wksSheet.Protection.AllowEditRanges.Item(1).Users.DeleteAll()
}
```

## 成员属性

### UserAccessList.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 UserAccessList 对象的变量。

### UserAccessList.Item

从集合中返回一个对象。

**语法**

**express.Item**

\*express是一个代表 UserAccessList 对象的变量。

**说明**

有关返回集合中单个成员的详细信息，请参阅 返回集合中的对象 。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/UserAccessList](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# VPageBreaks 对象

## 简介

打印区域中垂直分页符的集合。

## 说明

每一个垂直分页符都由一个 VPageBreak 对象代表。

当 Application 属性、 Count 属性、 Creator 属性、 Item 属性、 Parent 属性或 Add 方法与 VPageBreaks 属性一起使用时：

- 对于自动打印区域， VPageBreaks 属性只应用于打印区域内的分页符。
- 对于同一区域中的用户定义的打印区域， VPageBreaks 属性应用于所有的分页符。

使用 VPageBreaks 属性可返回 VPageBreaks 集合。使用 Add 方法可添加一个垂直分页符。

如果添加的分页符不和打印区域交叠，则新添加的 VPageBreak 对象将不会出现在打印区域的 VPageBreaks 集合中。如果重新调整打印区域或重新定义打印区域，将会改变集合的内容。

下例给活动单元格左边添加一个垂直分页符。

``` JavaScript
Application.ActiveSheet.VPageBreaks.Add(ActiveCell)
```

## 方法

| 序号 | 名称                      | 说明                                                          |
|:----:|:--------------------------|:--------------------------------------------------------------|
|  1   | [Add](#VPageBreaks.Add)   | 添加垂直分页符。返回 一个代表新垂直分页符的 VPageBreak 对象。 |
|  2   | [Item](#VPageBreaks.Item) | 从集合中返回一个对象。                                        |

## 属性

| 序号 | 名称                                    | 说明                                                                                                                                                                                                                            |
|:----:|:----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#VPageBreaks.Application) | 如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
|  2   | [Count](#VPageBreaks.Count)             | 返回一个 Long 值，它代表集合中对象的数量。                                                                                                                                                                                      |
|  3   | [Creator](#VPageBreaks.Creator)         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 Long 类型。                                                                                                                                                      |
|  4   | [Parent](#VPageBreaks.Parent)           | 返回指定对象的父对象。只读。                                                                                                                                                                                                    |

## 成员方法

### VPageBreaks.Add

添加垂直分页符。返回 一个代表新垂直分页符的 **VPageBreak** 对象。

**语法**

**express.Add ( *Before* )**

\*express是一个代表 VPageBreaks 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                            |
|----------|-----------|----------|-------------------------------------------------|
| *Before* | 必选      | Object   | 一个 Range 对象。新的分页符将添加到该区域左侧。 |

**示例**

``` JavaScript
/* 本示例在单元格 F25 上方添加水平分页符，在其左侧添加垂直分页符*/
function test() {
    Application.Worksheets.Item(1).HPageBreaks.Add(Worksheets.Item(1).Range("F25"))
    Application.Worksheets.Item(1).VPageBreaks.Add(Worksheets.Item(1).Range("F25"))
}
```

### VPageBreaks.Item

从集合中返回一个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 VPageBreaks 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明           |
|---------|-----------|----------|----------------|
| *Index* | 必选      | Long     | 对象的索引号。 |

**示例**

``` JavaScript
/* 此示例更改第一个垂直分页符的位置*/
Application.Worksheets.Item(1).VPageBreaks.Item(1).Location = Worksheets.Item(1).Range("e5")
```

## 成员属性

### VPageBreaks.Application

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

\*express是一个代表 VPageBreaks 对象的变量。

**示例**

``` JavaScript
/* 本示例显示一条有关创建 myObject 的应用程序的消息*/
function test() {
    let myObject = Application.ActiveWorkbook
    if(myObject.Application.Value == "ET") {
        alert("This is an ET Application object.")
    }
    else {
        alert("This is not an ET Application object.")
    }
}
```

### VPageBreaks.Count

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

\*express是一个代表 VPageBreaks 对象的变量。

### VPageBreaks.Creator

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 VPageBreaks 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。 **Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

### VPageBreaks.Parent

返回指定对象的父对象。只读。

**语法**

**express.Parent**

\*express是一个代表 VPageBreaks 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/VPageBreaks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# AddIn 对象

## 简介

代表已经安装或尚未安装的单个加载项。 AddIn 对象是 AddIns 集合的成员。 AddIns 集合包含可供 WPS 使用的所有加载项，不论它们当前是否加载。 AddIns 集合中包括在 “工具” 菜单的 “模板和加载项” 对话框中显示的共用模板或 WPS 加载项库 (WLL)。

## 说明

可以使用 AddIns ( *index* ) 返回单个 AddIn 对象（其中 *index* 是加载项的名称或索引编号）。名称的拼写必须和 “模板和加载项” 对话框中显示的完全一样（但大小写可以不一样）。以下示例将 Letter.dot 模板作为共用模板加载。

``` JavaScript
Application.AddIns.Item("Letter.dot").Installed = true
```

索引编号代表加载项在 “模板和加载项” 对话框中的加载项列表上的位置。下列指令显示第一个可用加载项的路径。

``` JavaScript
function test() {
    if (Application.AddIns.Count >= 1) {
        alert(Application.AddIns.Item(1).Path)
    }
}
```

以下示例在活动文档开始处创建一个加载项列表。该列表包含了每一个有效加载项的名称、路径和加载状态。

``` JavaScript
function test() {
    let rng = Application.ActiveDocument.Range(0, 0)
    rng.InsertAfter("Name" + "\t" + Path + "\t" + Installed)
    rng.InsertParagraphAfter()
    for (let i = 1; i <= Application.AddIns.Count; i++) {
        rng.InsertAfter(Application.AddIns.Item(i).Name + "\t" + Application.AddIns.Item(i).Path + "\t" + Application.AddIns.Item(i).Installed)
        rng.InsertParagraphAfter()
    }
    rng.ConvertToTable()
}
```

使用 Add 方法将一个加载项添加到可用加载项列表上并且使用 *Install* 参数安装它（可选）。

``` JavaScript
Application.AddIns.Add("C:\\Templates\\Other\\Letter.dot", true)
```

要安装可用加载项列表中显示的加载项，请使用 Installed 属性。

``` JavaScript
Application.AddIns.Item("Letter.dot").Installed = true
```

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td>使用 <span>Compiled</span> 属性判断一个 AddIn 对象是模板还是 WLL。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

## 方法

| 序号 | 名称                    | 说明               |
|:----:|:------------------------|:-------------------|
|  1   | [Delete](#AddIn.Delete) | 删除指定的加载项。 |

## 属性

| 序号 | 名称                              | 说明                                                                                                                                           |
|:----:|:----------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#AddIn.Application) | 返回一个 Application 对象，该对象代表 WPS 应用程序。                                                                                           |
|  2   | [Autoload](#AddIn.Autoload)       | 如果该属性为 True ，则在 WPS 启动时，自动加载指定的加载项。位于 WPS 程序文件夹的 Startup 子文件夹中的加载项可以自动加载。 Boolean 类型，只读。 |
|  3   | [Compiled](#AddIn.Compiled)       | 如果为 True ，则指定的加载项是一个 WPS 加载项库 (WLL). 如果为 False ，则加载项为一个模板。 Boolean 类型，只读。                                |
|  4   | [Creator](#AddIn.Creator)         | 返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。 Long 类型，只读。                                                                 |
|  5   | [Index](#AddIn.Index)             | 返回一个 Long 类型的值，该值代表项在集合中的位置。只读。                                                                                       |
|  6   | [Installed](#AddIn.Installed)     | 如果为 True ，则指定的加载项已安装（即加载）。可在 “工具” 菜单的 “模板和加载项” 对话框中选择需要加载的加载项。可读/写 Boolean 类型。           |
|  7   | [Name](#AddIn.Name)               | 返回加载项的名称。 String 类型，只读。                                                                                                         |
|  8   | [Parent](#AddIn.Parent)           | 返回一个 Object 类型值，该值代表指定 AddIn 对象的父对象。                                                                                      |
|  9   | [Path](#AddIn.Path)               | 返回已安装加载项的位置。 String 类型，只读。                                                                                                   |

## 成员方法

### AddIn.Delete

删除指定的加载项。

**语法**

**express.Delete ()**

\*express是一个代表 AddIn 对象的变量。

## 成员属性

### AddIn.Application

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

\*express是一个代表 AddIn 对象的变量。

### AddIn.Autoload

如果该属性为 **True** ，则在 WPS 启动时，自动加载指定的加载项。位于 WPS 程序文件夹的 Startup 子文件夹中的加载项可以自动加载。 **Boolean** 类型，只读。

**语法**

**express.Autoload**

\*express是一个代表 AddIn 对象的变量。

**示例**

``` JavaScript
/*此示例显示 WPS 启动时自动加载的每一加载项的名称。*/
function test() {    let blnFound = false    for (let i = 1; i <= Application.AddIns.Count; i++) {        if (Application.AddIns.Item(i).Autoload == true) {            alert(Application.AddIns.Item(i).Name)            blnFound = true        }    }    if (blnFound !== true) {        alert("No add-ins were loaded automatically.")    }}
```

``` JavaScript
/*此示例判断名为“Gallery.dot”的加载项是否自动加载。*/
function test() {
    let addinLoop

    for (let i = 1; i <= Application.AddIns.Count; i++) {
        addinLoop = Application.AddIns.Item(i)
        if (addinLoop.Name.toLowerCase().indexOf("gallery.dot") != -1) {
            if (addinLoop.Autoload == true) {
                alert("Autoload")
            }
        }
    }
}
```

### AddIn.Compiled

如果为 **True** ，则指定的加载项是一个 WPS 加载项库 (WLL). 如果为 **False** ，则加载项为一个模板。 **Boolean** 类型，只读。

**语法**

**express.Compiled**

\*express是一个代表 AddIn 对象的变量。

**示例**

``` JavaScript
/*此示例判断当前加载的 WLL 的数目。*/
function test() {
    let count = 0
    for (let i = 1; i <= Application.AddIns.Count; i++) {
        if (Application.AddIns.Item(i).Compiled == true && Application.AddIns.Item(i).Installed == true) {
            count++
        }
    }
    alert(count + " WLL's are loaded")
}
```

``` JavaScript
/*如果第一个加载项是模板，本示例卸载该模板并打开它。*/
function test() {
    if (Application.AddIns.Item(1).Compiled == false) {
        Application.AddIns.Item(1).Installed = false
        Documents.Open(Application.AddIns.Item(1).Path + Application.PathSeparator + Application.AddIns.Item(1).Name)
    }
}
```

### AddIn.Creator

返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。 **Long** 类型，只读。

**语法**

**express.Creator**

\*express是一个代表 AddIn 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td>该值也可用常量 <strong>wdCreatorCode</strong> 表示。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

### AddIn.Index

返回一个 **Long** 类型的值，该值代表项在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 AddIn 对象的变量。

### AddIn.Installed

如果为 **True** ，则指定的加载项已安装（即加载）。可在 **“工具”** 菜单的 **“模板和加载项”** 对话框中选择需要加载的加载项。可读/写 **Boolean** 类型。

**语法**

**express.Installed**

\*express是一个代表 AddIn 对象的变量。

**说明**

已卸载的加载项包含在 **AddIns** 集合中。若要从 AddIns 集合中删除模板或 WLL，可对 **AddIn** 对象应用 **Delete** 方法（加载项名称从 **“模板和加载项”** 对话框中删除）。若要卸载所有的模板和 WLL，可对 **AddIns** 集合应用 **Unload** 方法。

**示例**

``` JavaScript
/*此示例卸载名为“Gallery.dot”的全局模板。*/
Application.AddIns.Item("Gallery.dot").Installed = false
```

``` JavaScript
/*此示例加载 FindAll.wll。*/
Application.AddIns.Item("C:\\Templates\\FindAll.wll").Installed = true
```

``` JavaScript
/*此示例加载 Custom.dot。*/
Application.AddIns.Item("C:\\Program Files\\Microsoft Office\\Templates\\Custom.dot").Installed = true
```

``` JavaScript
/*如果 Dot1.dot 作为全局模板加载，本示例则在状态栏上显示一条消息。*/
function test() {
    if (Application.AddIns.Item("Dot1.dot").Installed == true) {
        StatusBar = "Dot1.dot is loaded"
    }
}
```

### AddIn.Name

返回加载项的名称。 **String** 类型，只读。

**语法**

**express.Name**

\*express是一个代表 AddIn 对象的变量。

### AddIn.Parent

返回一个 **Object** 类型值，该值代表指定 **AddIn** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 AddIn 对象的变量。

### AddIn.Path

返回已安装加载项的位置。 **String** 类型，只读。

**语法**

**express.Path**

\*express是一个代表 AddIn 对象的变量。

**说明**

此路径不包括尾随字符，例如，“C:\WPSOffice”或“http://MyServer”。使用 **PathSeparator** 属性可添加分隔文件夹与驱动器号的字符。使用 **Name** 属性可返回不带路径的文件名，而使用 **FullName** 属性可同时返回文件名和路径。

<table data-border="0" data-cellpadding="0" data-cellspacing="0" width="100%">
<colgroup>
<col style="width: 100%" />
</colgroup>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td><table>
<tbody>
<tr class="odd">
<td>可以使用 <strong>PathSeparator</strong> 属性建立 Web 地址，即使此地址中包含正斜杠 (/)，而 <strong>PathSeparator</strong> 属性默认使用反斜杠 (\)。</td>
</tr>
</tbody>
</table></td>
</tr>
</tbody>
</table>

**示例**

``` JavaScript
/*本示例显示 AddIns 集合中第一个加载项的路径。*/
function test() {
    if (Application.AddIns.Count >= 1) {
        alert(Application.AddIns.Item(1).Path)
    }
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/AddIn](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# DocumentFields 对象

## 简介

代表一个公文域集合。 DocumentField 对象是 DocumentFields 集合的成员。 DocumentFields 集合中包含当前在 WPS 中打开的所有 DocumentField 对象。

## 方法

| 序号 | 名称                             | 说明                                                |
|:----:|:---------------------------------|:----------------------------------------------------|
|  1   | [Add](#DocumentFields.Add)       |                                                     |
|  2   | [Exists](#DocumentFields.Exists) | 判断是否存在该名称的公文域，如果存在范围True。      |
|  3   | [Item](#DocumentFields.Item)     | 返回一个 DocumentField 对象，该对象通过索引来确定。 |

## 属性

| 序号 | 名称                                             | 说明                                                                               |
|:----:|:-------------------------------------------------|:-----------------------------------------------------------------------------------|
|  1   | [Count](#DocumentFields.Count)                   | 返回一个 Long 类型的值，代表该对象中 DocumentField 对象的数量，只读。              |
|  2   | [DefaultSorting](#DocumentFields.DefaultSorting) |  设置DocumentFields集合的排序方式，可以通过枚举值 WpsBookmarkSortBy 指定。 可读/写 |
|  3   | [InsertionMode](#DocumentFields.InsertionMode)   | 通过Boolean类型变量设置公文域对象的插入模式， 可读/写。                            |

## 成员方法

### DocumentFields.Add

**语法**

**express.Add ( *Name* , *Range* , *Hidden* , *PrintOut* , *ReadOnly* )**

\*express是一个代表 DocumentFields 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                      |
|------------|-----------|----------|-------------------------------------------|
| *Name*     | 必选      | String   | 要添加的公文域名称                        |
| *Range*    | 可选      | Variant  | 代表一个Range对象，表示该公文域添加的位置 |
| *Hidden*   | 可选      | Variant  | 该公文域添加后是否显示                    |
| *PrintOut* | 可选      | Variant  | 在文档打印时是否打印该公文域              |
| *ReadOnly* | 可选      | Variant  | 该公文域是否为只读类型                    |

### DocumentFields.Exists

判断是否存在该名称的公文域，如果存在范围True。

**语法**

**express.Exists ( *Name* )**

\*express是一个代表 DocumentFields 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明           |
|--------|-----------|----------|----------------|
| *Name* | 必选      | String   | 公文域得名称。 |

### DocumentFields.Item

返回一个 **DocumentField** 对象，该对象通过索引来确定。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 DocumentFields 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                 |
|---------|-----------|----------|----------------------|
| *Index* | 必选      | Variant  | 要获取的公文域索引值 |

## 成员属性

### DocumentFields.Count

返回一个 **Long** 类型的值，代表该对象中 **DocumentField** 对象的数量，只读。

**语法**

**express.Count**

\*express是一个代表 DocumentFields 对象的变量。

### DocumentFields.DefaultSorting

设置DocumentFields集合的排序方式，可以通过枚举值 WpsBookmarkSortBy 指定。 可读/写

**语法**

**express.DefaultSorting**

\*express是一个代表 DocumentFields 对象的变量。

### DocumentFields.InsertionMode

通过Boolean类型变量设置公文域对象的插入模式， 可读/写。

**语法**

**express.InsertionMode**

\*express是一个代表 DocumentFields 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/DocumentFields](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FormField 对象

## 简介

代表单个窗体域。 FormField 对象是 FormFields 集合的一个成员。

## 说明

使用 FormFields (index) 可返回一个 FormField 对象，其中 index 是书签名称或索引号。以下示例将 Text1 窗体域的结果设置为“Don Funk”。

``` JavaScript
Application.ActiveDocument.FormFields.Item("Text1").Result = "Don Funk"
```

索引号代表窗体域在所选内容、范围或文档中的位置。以下示例显示所选内容中第一个窗体域的名称。

``` JavaScript
function test() {
    if(Application.Selection.FormFields.Count >= 1) {
        alert(Application.Selection.FormFields.Item(1).Name)
    }
}
```

使用 Add 方法和 FormFields 对象可添加一个窗体域。以下示例在活动文档的开始处添加一个复选框，然后选中该复选框。

``` JavaScript
function test() {
    let ffield = Application.ActiveDocument.FormFields.Add(ActiveDocument.Range(0,0),wdFieldFormCheckBox)
    ffield.CheckBox.Value = true
}
```

使用 FormField 对象和 CheckBox 、 DropDown 和 TextInput 属性可返回 CheckDown 、 DropDown 和 TextInput 对象。以下示例选中名为“Check1”的复选框。

``` JavaScript
Application.ActiveDocument.FormFields.Item("Check1").CheckBox.Value = true
```

## 方法

| 序号 | 名称                        | 说明                               |
|:----:|:----------------------------|:-----------------------------------|
|  1   | [Copy](#FormField.Copy)     | 将指定窗体域复制到剪贴板。         |
|  2   | [Cut](#FormField.Cut)       | 将指定窗体域从文档中移到剪贴板上。 |
|  3   | [Delete](#FormField.Delete) | 删除指定窗体域。                   |
|  4   | [Select](#FormField.Select) | 选择指定的对象。                   |

## 属性

| 序号 | 名称                                          | 说明                                                                                          |
|:----:|:----------------------------------------------|:----------------------------------------------------------------------------------------------|
|  1   | [Application](#FormField.Application)         | 返回一个代表 WPS 应用程序的 Application 对象。                                                |
|  2   | [CalculateOnExit](#FormField.CalculateOnExit) | 如果退出指定的窗体域时自动更新对该域的引用，则该属性值为 True 。 Boolean 类型，可读写。       |
|  3   | [CheckBox](#FormField.CheckBox)               | 返回一个 CheckBox 对象，该对象代表复选框型窗体域。只读。                                      |
|  4   | [Creator](#FormField.Creator)                 | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                  |
|  5   | [DropDown](#FormField.DropDown)               | 返回一个 DropDown 对象，该对象代表下拉型窗体域。只读。                                        |
|  6   | [Enabled](#FormField.Enabled)                 | 如果启用窗体域，则该属性值为 True 。 Boolean 类型，可读写。                                   |
|  7   | [EntryMacro](#FormField.EntryMacro)           | 返回或设置指定窗体域（CheckBox、DropDown 或 TextInput）的输入宏的名称。 String 类型，可读写。 |
|  8   | [ExitMacro](#FormField.ExitMacro)             | 返回或设置指定窗体域（CheckBox、DropDown 或 TextInput）的退出宏的名称。 String 类型，可读写。 |
|  9   | [HelpText](#FormField.HelpText)               | 返回或设置当窗体域获得焦点并且用户按 F1 时在消息框中显示的文字。 String 类型，可读写。        |
|  10  | [Name](#FormField.Name)                       | 返回或设置指定对象的名称。 String 类型，可读写。                                              |
|  11  | [Next](#FormField.Next)                       | 返回集合中的下一个窗体域。只读。                                                              |
|  12  | [OwnHelp](#FormField.OwnHelp)                 | 指定当窗体域具有焦点并且用户按 F1 时在消息框中显示的文本的源。 Boolean 类型，可读写。         |
|  13  | [Range](#FormField.Range)                     | 返回一个 Range 对象，该对象代表窗体域中包含的文档部分。                                       |
|  14  | [Result](#FormField.Result)                   | 返回一个代表指定的窗体域的结果的 String 。可读写。                                            |
|  15  | [Type](#FormField.Type)                       | 返回域的类型。只读 WdFieldType 类型。                                                         |

## 成员方法

### FormField.Copy

将指定窗体域复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 FormField 对象的变量。

### FormField.Cut

将指定窗体域从文档中移到剪贴板上。

**语法**

**express.Cut ()**

\*express是一个代表 FormField 对象的变量。

**示例**

``` JavaScript
/*本示例删除活动文档的第一个域，并将该域粘贴到插入点。*/
function test() {
  if(Application.ActiveDocument.Fields.Count >= 1) {
      Application.ActiveDocument.Fields.Item(1).Cut()
      Application.Selection.Collapse(wdCollapseEnd)
      Application.Selection.Paste()
  }
}

/*本示例删除第一段的第一个单词，并将该单词粘贴到该段的末尾。*/
function test() {
  let ran= Application.ActiveDocument.Paragraphs.Item(1).Range
  ran.Words.Item(1).Cut()
  ran.Collapse(wdCollapseEnd)
  ran.Move(wdCharacter,-1)
  ran.Paste()
}

/*本示例删除所选内容，并将其粘贴到新文档中。*/
function test() {
  if(Application.Selection.Type == wdSelectionNormal) {
      Application.Selection.Cut()
      Application.Documents.Add().Content.Paste()
  }
}
```

### FormField.Delete

删除指定窗体域。

**语法**

**express.Delete ()**

\*express是一个代表 FormField 对象的变量。

### FormField.Select

选择指定的对象。

**语法**

**express.Select ()**

\*express是一个代表 FormField 对象的变量。

**说明**

使用此方法后，可使用 Selection 对象处理所选项目。有关详细信息，请参阅 处理 Selection 对象。

## 成员属性

### FormField.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 FormField 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### FormField.CalculateOnExit

如果退出指定的窗体域时自动更新对该域的引用，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.CalculateOnExit**

\*express是一个代表 FormField 对象的变量。

**说明**

一个 REF 域可用来引用窗体域中的内容。例如，{REF SubTotal} 引用由 SubTotal 书签标识的窗体域。

**示例**

``` JavaScript
/*本示例在退出 Form.doc 中的窗体域时可以自动更新对该域的引用。*/
function test()
{
for(let i=1;i<=Documents.Item("Form.doc").FormFields.Count;i++) {
    Documents.Item("Form.doc").FormFields.Item(i).CalculateOnExit = false
}
}
```

``` JavaScript
/*本示例在新文档中添加一个文字型窗体域和一个 REF 域。每次键入文字时并退出 Text1 域时，都将自动更新 REF 域。*/
function test()
{
let abc= Documents.Add()
abc.FormFields.Add(Selection.Range,wdFieldFormTextInput)
abc.Fields.Add(Selection.Range,wdFieldRef,"Text1")
abc.FormFields.Item("Text1").CalculateOnExit = true
abc.Protect(wdAllowOnlyFormFields)
}
```

### FormField.CheckBox

返回一个 **CheckBox** 对象，该对象代表复选框型窗体域。只读。

**语法**

**express.CheckBox**

\*express是一个代表 FormField 对象的变量。

**说明**

如果将 **CheckBox** 属性应用于非复选框型窗体域的 **FormField** 对象，则该属性有效，但是返回的对象的 **Valid** 属性将为 **False** 。

**示例**

``` JavaScript
/*本示例清除名为“Blue”的复选框。*/
ActiveDocument.FormFields.Item("Blue").CheckBox.Value = false
```

``` JavaScript
/*本示例将当前值与名为“Check1”的复选框的默认值进行比较。如果两值相等，则 blnSame 变量会设置为 True。*/
function test()
{
let ffTemp = ActiveDocument.FormFields.Item("Check1").CheckBox
let blnSame
if(ffTemp.Default == ffTemp.Value) {
    blnSame = true
}
else {
    blnSame = false
}
}
```

### FormField.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 FormField 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### FormField.DropDown

返回一个 **DropDown** 对象，该对象代表下拉型窗体域。只读。

**语法**

**express.DropDown**

\*express是一个代表 FormField 对象的变量。

**说明**

如果将 **DropDown** 属性应用于非下拉型窗体域的 **FormField** 对象，则该属性有效，但是返回的对象的 **Valid** 属性将为 **False** 。

**示例**

``` JavaScript
/*本示例显示名为“Colors”的下拉型窗体域中所选项的文本。*/
function test()
{
let ffDrop = ActiveDocument.FormFields.Item("Colors").DropDown
MsgBox(ffDrop.ListEntries.Item(ffDrop.Value).Name)
}
```

``` JavaScript
/*本示例向 Form.doc 中名为“Places”的下拉型窗体域中添加“Seattle”。*/
function test()
{
let trie=Documents.Item("Form.doc").FormFields.Item("Places").DropDown.ListEntries
trie.Add("Seattle")
}
```

### FormField.Enabled

如果启用窗体域，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.Enabled**

\*express是一个代表 FormField 对象的变量。

**说明**

如果启用了一个窗体域，则在填充该窗体时，可以更改其内容。

**示例**

``` JavaScript
/*如果活动文档中的第一个窗体域是一个已启用的复选框，则以下示例选中该复选框。*/
function test()
{
let ffFirst = ActiveDocument.FormFields.Item(1)
if(ffFirst.Enabled == true &&ffFirst.Type == wdFieldFormCheckBox) {
    ffFirst.CheckBox.Value = true
}
}
```

### FormField.EntryMacro

返回或设置指定窗体域（CheckBox、DropDown 或 TextInput）的输入宏的名称。 **String** 类型，可读写。

**语法**

**express.EntryMacro**

\*express是一个代表 FormField 对象的变量。

**说明**

当窗体域获得焦点时，输入宏开始运行。

**示例**

``` JavaScript
/*本示例为“Form.doc”中的第一个窗体域指定名为“Blue”的宏。*/
Documents.Item("Form.doc").FormFields.Item(1).EntryMacro = "Blue"
```

``` JavaScript
/*本示例为活动文档中名为“Text1”的窗体域指定一个名为“Breadth”的宏。*/
ActiveDocument.FormFields.Item("Text1").EntryMacro = "Breadth"
```

### FormField.ExitMacro

返回或设置指定窗体域（CheckBox、DropDown 或 TextInput）的退出宏的名称。 **String** 类型，可读写。

**语法**

**express.ExitMacro**

\*express是一个代表 FormField 对象的变量。

**说明**

当窗体域失去焦点时，退出宏开始运行。

**示例**

``` JavaScript
/*本示例为选定内容的第一个窗体域指定名为“Reformat”的宏。*/
function test()
{
if(Selection.FormFields.Count > 0 ) {
    Selection.FormFields.Item(1).ExitMacro = "Reformat"
}
}
```

``` JavaScript
/*本示例为“Form.doc”中最后一个窗体域指定名为“Blue”的宏。*/
function test()
{
let intMax = Documents.Item("Form.doc").FormFields.Count
Documents.Item("Form.doc").FormFields.Item(intMax).ExitMacro = "Blue"
}
```

### FormField.HelpText

返回或设置当窗体域获得焦点并且用户按 F1 时在消息框中显示的文字。 **String** 类型，可读写。

**语法**

**express.HelpText**

\*express是一个代表 FormField 对象的变量。

**说明**

如果将 **OwnHelp** 属性设置为 **True** ，则 **HelpText** 会指定该文字字符串的值。如果将 **OwnHelp** 设置为 **False** ，则 **HelpText** 会指定包含窗体域的帮助文字的“自动图文集”词条的名称。

**示例**

``` JavaScript
/*本示例为名为“Name”的窗体域设置帮助文字。*/
function test()
{
ActiveDocument.FormFields.Item("Name").OwnHelp = true
ActiveDocument.FormFields.Item("Name").HelpText = "Type your full legal name."
}
```

### FormField.Name

返回或设置指定对象的名称。 **String** 类型，可读写。

**语法**

**express.Name**

\*express是一个代表 FormField 对象的变量。

### FormField.Next

返回集合中的下一个窗体域。只读。

**语法**

**express.Next**

\*express是一个代表 FormField 对象的变量。

### FormField.OwnHelp

指定当窗体域具有焦点并且用户按 F1 时在消息框中显示的文本的源。 **Boolean** 类型，可读写。

**语法**

**express.OwnHelp**

\*express是一个代表 FormField 对象的变量。

**说明**

如果为 **True** ，则显示 **HelpText** 属性指定的文本。如果为 **False** ，则显示 **HelpText** 属性指定的“自动图文集”词条中的文本。

**示例**

``` JavaScript
/*本示例将当前节的第一个窗体域的帮助文字设置为名为“Sample”的“自动图文集”词条的内容。*/
function test()
{
let help=Selection.Sections.Item(1).Range.FormFields.Item(1)
help.OwnHelp = false
help.HelpText = "Sample"
}
```

### FormField.Range

返回一个 **Range** 对象，该对象代表窗体域中包含的文档部分。

**语法**

**express.Range**

\*express是一个代表 FormField 对象的变量。

### FormField.Result

返回一个代表指定的窗体域的结果的 **String** 。可读写。

**语法**

**express.Result**

\*express是一个代表 FormField 对象的变量。

**示例**

``` JavaScript
/*以下示例显示活动文档中每个窗体域的结果。*/
function test() {
  for(let i=1;i<=Application.ActiveDocument.FormFields.Count;i++) {
      alert(Application.ActiveDocument.FormFields.Item(i).Result)
  }
}
```

### FormField.Type

返回域的类型。只读 WdFieldType 类型。

**语法**

**express.Type**

\*express是一个代表 FormField 对象的变量。

**说明**

**安全性:** 动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/FormField](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# HeadersFooters 对象

## 方法

| 序号 | 名称                         | 说明                                                                 |
|:----:|:-----------------------------|:---------------------------------------------------------------------|
|  1   | [Item](#HeadersFooters.Item) | 返回一个 HeaderFooter 对象，该对象代表一个区域或部分中的页眉或页脚。 |

## 属性

| 序号 | 名称                                       | 说明                                                                         |
|:----:|:-------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#HeadersFooters.Application) | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#HeadersFooters.Count)             | 返回一个 Long 类型的值，该值代表集合中页眉和/或页脚的数目。只读。            |
|  3   | [Creator](#HeadersFooters.Creator)         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [Parent](#HeadersFooters.Parent)           | 返回一个 Object 类型值，该值代表指定 HeadersFooters 对象的父对象。           |

## 成员方法

### HeadersFooters.Item

返回一个 **HeaderFooter** 对象，该对象代表一个区域或部分中的页眉或页脚。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 HeadersFooters 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型            | 说明                               |
|---------|-----------|---------------------|------------------------------------|
| *Index* | 必选      | WdHeaderFooterIndex | 指定区域或部分中页眉或页脚的常量。 |

**示例**

``` JavaScript
/*本示例在活动文档中创建首页页眉并设置其格式。*/
function test(){
    Application.ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = true

    let rng = Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterFirstPage).Range
        rng.InsertBefore("Sales Report")
        rng.Font.Bold = true
        rng.Font.Size = "15"
        rng.Font.Color = wdColorBlue
        rng.Paragraphs.Alignment = wdAlignParagraphCenter
}
```

## 成员属性

### HeadersFooters.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 HeadersFooters 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### HeadersFooters.Count

返回一个 **Long** 类型的值，该值代表集合中页眉和/或页脚的数目。只读。

**语法**

**express.Count**

\*express是一个代表 HeadersFooters 对象的变量。

### HeadersFooters.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 HeadersFooters 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### HeadersFooters.Parent

返回一个 **Object** 类型值，该值代表指定 **HeadersFooters** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 HeadersFooters 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/HeadersFooters](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# InlineShape 对象

## 简介

代表文档的文字层中的对象。内嵌形状只能是图片、OLE 对象或 ActiveX 控件。 InlineShape 对象是 InlineShapes 集合的一个成员。 InlineShapes 集合包含文档、范围或所选内容中的所有内嵌形状。

## 说明

InlineShape 对象被视为字符，并作为字符置于文本行中。

使用 InlineShapes ( *Index* ) 可返回一个 InlineShape 对象，其中 *Index* 是索引号。内嵌形状没有名称。以下示例激活活动文档中的第一个内嵌形状。

``` JavaScript
Application.ActiveDocument.InlineShapes.Item(1).Activate()
```

Shape 对象锁定到某一文本范围，但可以自由浮动，并且可以放置在页面上的任何位置。可以使用 ConvertToInlineShape 方法和 ConvertToShape 方法来转换形状的类型。只能将图片、OLE 对象和 ActiveX 控件转换为内嵌形状。使用 Type 属性可返回内嵌形状的类型：图片、链接图片、嵌入的 OLE 对象、链接的 OLE 对象或者 ActiveX 控件。

## 方法

| 序号 | 名称                                          | 说明                                                                        |
|:----:|:----------------------------------------------|:----------------------------------------------------------------------------|
|  1   | [ConvertToShape](#InlineShape.ConvertToShape) | 将嵌入式图形转换为可自由浮动的图形。返回一个 Shape 对象，该对象代表新图形。 |
|  2   | [Delete](#InlineShape.Delete)                 | 删除指定的嵌入式图形。                                                      |
|  3   | [Reset](#InlineShape.Reset)                   | 删除对内嵌形状所做的更改。                                                  |
|  4   | [Select](#InlineShape.Select)                 | 选择指定的内嵌形状。                                                        |

## 属性

| 序号 | 名称                                                      | 说明                                                                                                                                                                                                                                |
|:----:|:----------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AlternativeText](#InlineShape.AlternativeText)           | 返回或设置一个 String 类型的值，该值代表与网页中图形相关联的 可选文字 ?（可选文字：由 Web 浏览器用于在图像下载过程中显示文本，其对象是那些关闭图形的用户，以及那些依赖屏幕阅读软件将屏幕上的图形转换为可读语言的用户。） 。可读写。 |
|  2   | [Application](#InlineShape.Application)                   | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                                                                                      |
|  3   | [Borders](#InlineShape.Borders)                           | 返回一个 Borders 集合，该集合代表指定图形的所有边框。                                                                                                                                                                               |
|  4   | [Chart](#InlineShape.Chart)                               | 返回一个 Chart 对象，该对象代表文档中内嵌形状集合内的图表。只读。                                                                                                                                                                   |
|  5   | [Creator](#InlineShape.Creator)                           | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                                                                                        |
|  6   | [Field](#InlineShape.Field)                               | 返回一个 Field 对象，该对象代表与指定内嵌形状相关联的域。只读。                                                                                                                                                                     |
|  7   | [Fill](#InlineShape.Fill)                                 | 返回一个 FillFormat 对象，该对象包含指定图形的填充格式属性。只读。                                                                                                                                                                  |
|  8   | [Glow](#InlineShape.Glow)                                 | 返回一个 GlowFormat 对象，该对象代表发光效果的格式属性。只读。                                                                                                                                                                      |
|  9   | [GroupItems](#InlineShape.GroupItems)                     | 返回一个 GroupShapes 集合，该集合代表针对内嵌形状分组到一起的形状。只读。                                                                                                                                                           |
|  10  | [HasChart](#InlineShape.HasChart)                         | 如果指定的形状是图表，则为 True 。只读。                                                                                                                                                                                            |
|  11  | [HasSmartArt](#InlineShape.HasSmartArt)                   | 如果形状上显示 SmartArt 图表则返回 True 。只读。                                                                                                                                                                                    |
|  12  | [Height](#InlineShape.Height)                             | 返回或设置内嵌形状的高度。 Single 类型，可读写。                                                                                                                                                                                    |
|  13  | [HorizontalLineFormat](#InlineShape.HorizontalLineFormat) | 返回一个 HorizontalLineFormat 对象，该对象包含指定的 InlineShape 对象的水平行的格式。只读。                                                                                                                                         |
|  14  | [Hyperlink](#InlineShape.Hyperlink)                       | 返回一个 Hyperlink 对象，该对象代表与指定内嵌形状相关联的超链接。只读。                                                                                                                                                             |
|  15  | [IsPictureBullet](#InlineShape.IsPictureBullet)           | 如果为 True ，则表明 InlineShape 对象是图片项目符号。 Boolean 类型，只读。                                                                                                                                                          |
|  16  | [Line](#InlineShape.Line)                                 | 返回一个 LineFormat 对象，该对象包含指定形状的线条格式属性。只读。                                                                                                                                                                  |
|  17  | [LinkFormat](#InlineShape.LinkFormat)                     | 返回一个 LinkFormat 对象，该对象代表链接到文件的指定内嵌形状的链接选项。只读。                                                                                                                                                      |
|  18  | [LockAspectRatio](#InlineShape.LockAspectRatio)           | 如果在调整指定形状的大小时保留其最初比例，则该属性值为 MsoTrue ；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 MsoFalse 。可读写 MsoTriState 类型。                                                                      |
|  19  | [OLEFormat](#InlineShape.OLEFormat)                       | 返回一个 OLEFormat 对象，该对象代表指定的内嵌形状的 OLE 特征（链接除外）。只读                                                                                                                                                      |
|  20  | [Parent](#InlineShape.Parent)                             | 返回一个 Object 类型值，该值代表指定 InlineShape 对象的父对象。                                                                                                                                                                     |
|  21  | [PictureFormat](#InlineShape.PictureFormat)               | 返回一个 PictureFormat 对象，该对象包含内嵌形状的图片格式属性。只读。                                                                                                                                                               |
|  22  | [Range](#InlineShape.Range)                               | 返回一个 Range 对象，该对象代表内嵌形状中包含的文档部分。                                                                                                                                                                           |
|  23  | [Reflection](#InlineShape.Reflection)                     | 返回一个 ReflectionFormat 对象，该对象代表形状的反射格式。只读。                                                                                                                                                                    |
|  24  | [ScaleHeight](#InlineShape.ScaleHeight)                   | 参照指定内嵌形状的原始尺寸缩放其高度。可读写 Single 类型。                                                                                                                                                                          |
|  25  | [ScaleWidth](#InlineShape.ScaleWidth)                     | 参照指定内嵌形状的原始尺寸缩放其宽度。可读写 Single 类型。                                                                                                                                                                          |
|  26  | [Script](#InlineShape.Script)                             | 返回一个 Script 对象，该对象代表与指定网页中的图像相关联的脚本或代码块。                                                                                                                                                            |
|  27  | [Shadow](#InlineShape.Shadow)                             | 返回一个 ShadowFormat 对象，该对象代表指定形状的阴影格式。只读。                                                                                                                                                                    |
|  28  | [SmartArt](#InlineShape.SmartArt)                         | 返回 SmartArt 对象，该对象提供一种处理与指定内嵌形状相关联的 SmartArt 的方式。只读。                                                                                                                                                |
|  29  | [SoftEdge](#InlineShape.SoftEdge)                         | 返回一个 SoftEdgeFormat 对象，该对象代表形状的软边缘格式。只读。                                                                                                                                                                    |
|  30  | [TextEffect](#InlineShape.TextEffect)                     | 返回一个 TextEffectFormat 对象，该对象包含指定内嵌形状的文本效果格式属性。只读。                                                                                                                                                    |
|  31  | [Title](#InlineShape.Title)                               | 返回或设置包含指定内嵌形状的标题的 String 类型值。可读/写。                                                                                                                                                                         |
|  32  | [Type](#InlineShape.Type)                                 | 返回内嵌形状的类型。只读 WdInlineShapeType 类型。                                                                                                                                                                                   |
|  33  | [Width](#InlineShape.Width)                               | 返回或设置指定内嵌形状的宽度（以磅为单位）。可读写 Long 类型。                                                                                                                                                                      |

## 成员方法

### InlineShape.ConvertToShape

将嵌入式图形转换为可自由浮动的图形。返回一个 **Shape** 对象，该对象代表新图形。

**语法**

**express.ConvertToShape ()**

\*express是一个代表 InlineShape 对象的变量。

**说明**

使用 **ConvertToShape** 方法前，最少应该已经对 **FreeformBuilder** 对象用过一次 **AddNodes** 方法。

**示例**

``` JavaScript
/*本示例将活动文档的第一个嵌入式图形转化为浮动的图形。*/
ActiveDocument.InlineShapes.Item(1).ConvertToShape()
```

### InlineShape.Delete

删除指定的嵌入式图形。

**语法**

**express.Delete ()**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.Reset

删除对内嵌形状所做的更改。

**语法**

**express.Reset ()**

\*express是一个代表 InlineShape 对象的变量。

**示例**

``` JavaScript
/*本示例首先将一幅图片作为内嵌形状插入，并更改其亮度，然后再将其亮度重新设置为初始的亮度。*/
function test()
{
let aInLine = ActiveDocument.InlineShapes.AddPicture("C:\\Windows\\Bubbles.bmp", Selection.Range)
aInLine.PictureFormat.Brightness = 0.5
MsgBox("Changing brightness back")
aInLine.Reset()
}
```

### InlineShape.Select

选择指定的内嵌形状。

**语法**

**express.Select ()**

\*express是一个代表 InlineShape 对象的变量。

**说明**

使用此方法之后，可用 **Selection** 属性处理所选项目。有关详细信息，请参阅 处理 Selection 对象。

## 成员属性

### InlineShape.AlternativeText

返回或设置一个 **String** 类型的值，该值代表与网页中图形相关联的 可选文字 ?（可选文字：由 Web 浏览器用于在图像下载过程中显示文本，其对象是那些关闭图形的用户，以及那些依赖屏幕阅读软件将屏幕上的图形转换为可读语言的用户。） 。可读写。

**语法**

**express.AlternativeText**

\*express是一个代表 InlineShape 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：为活动窗口中选定的图形设定可选文字。选定图形是一幅野鸭的图片。*/
ActiveWindow.Selection.InlineShapes.Item(1).AlternativeText = "This is a mallard duck."
```

### InlineShape.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 InlineShape 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### InlineShape.Borders

返回一个 **Borders** 集合，该集合代表指定图形的所有边框。

**语法**

**express.Borders**

\*express是一个代表 InlineShape 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### InlineShape.Chart

返回一个 **Chart** 对象，该对象代表文档中内嵌形状集合内的图表。只读。

**语法**

**express.Chart**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 InlineShape 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### InlineShape.Field

返回一个 **Field** 对象，该对象代表与指定内嵌形状相关联的域。只读。

**语法**

**express.Field**

\*express是一个代表 InlineShape 对象的变量。

**说明**

使用 **Fields** 属性可返回 **Fields** 集合。

**示例**

``` JavaScript
/*以下示例将一个图形作为内嵌形状插入（使用 INCLUDEPICTURE 域），然后显示该形状的域代码。*/
function test()
{
let iShapeNew =ActiveDocument.InlineShapes.AddPicture("C:\\Windows\\Tiles.bmp",true,false,Selection.Range)
MsgBox(ActiveDocument.InlineShapes.Item(1).Field.Code.Text)
}
```

### InlineShape.Fill

返回一个 **FillFormat** 对象，该对象包含指定图形的填充格式属性。只读。

**语法**

**express.Fill**

\*express是一个代表 InlineShape 对象的变量。

**示例**

``` JavaScript
/*本示例首先向 myDocument 添加一个矩形，然后为其设置前景色、背景色和矩形填充的渐变类型。*/
function test()
{
let myDocument = Documents.Item(1).Shapes.AddShape(msoShapeRectangle,90, 90, 90, 50).Fill
    myDocument.ForeColor.RGB = (128, 0, 0)
    myDocument.BackColor.RGB = (170, 170, 170)
    myDocument.TwoColorGradient(msoGradientHorizontal, 1)
}
```

### InlineShape.Glow

返回一个 **GlowFormat** 对象，该对象代表发光效果的格式属性。只读。

**语法**

**express.Glow**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.GroupItems

返回一个 **GroupShapes** 集合，该集合代表针对内嵌形状分组到一起的形状。只读。

**语法**

**express.GroupItems**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.HasChart

如果指定的形状是图表，则为 **True** 。只读。

**语法**

**express.HasChart**

\*express是一个代表 InlineShape 对象的变量。

**说明**

对于 OLE 图表，此属性总是返回 False。此时请使用 ` InlineShape.OLEFormat.ProgID ` ，并且检查是否存在下列可能的值：“Excel.Chart.8”、“MSGraph.Chart.8”、“Excel.Sheet.8”、“Excel.Chart.5”、“MSGraph.Chart.5”或“Excel.Sheet.5”。

### InlineShape.HasSmartArt

如果形状上显示 SmartArt 图表则返回 **True** 。只读。

**语法**

**express.HasSmartArt**

\*express是一个代表 InlineShape 对象的变量。

**示例**

``` JavaScript
/*下面的代码示例显示活动文档中的第一个内嵌形状是否包含 SmartArt。*/
function test()
{
let myInlineShape = ActiveDocument.InlineShapes.Item(1)

if(myInlineShape.HasSmartArt){
   MsgBox("The first shape contains SmartArt.")
}
else{
   MsgBox("The first shape contains no SmartArt.")
}
}
```

### InlineShape.Height

返回或设置内嵌形状的高度。 **Single** 类型，可读写。

**语法**

**express.Height**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.HorizontalLineFormat

返回一个 **HorizontalLineFormat** 对象，该对象包含指定的 **InlineShape** 对象的水平行的格式。只读。

**语法**

**express.HorizontalLineFormat**

\*express是一个代表 InlineShape 对象的变量。

**示例**

``` JavaScript
/*本示例将指定水平行的长度设定为窗口宽度的 50%。*/
function test()
{
ActiveDocument.InlineShapes.AddHorizontalLineStandard()
ActiveDocument.InlineShapes.Item(1).HorizontalLineFormat.PercentWidth = 50
}
```

### InlineShape.Hyperlink

返回一个 **Hyperlink** 对象，该对象代表与指定内嵌形状相关联的超链接。只读。

**语法**

**express.Hyperlink**

\*express是一个代表 InlineShape 对象的变量。

**说明**

如果不存在与指定形状相关联的超链接，则会发生错误。在此情况下，请将 **Add** 方法用于 **Hyperlinks** 集合来为指定形状添加超链接。以下示例说明了具体的操作方法。

``` JavaScript
ActiveDocument.Hyperlinks.Add(Selection.InlineShapes.Item(1), "http://www.wps.cn")
 
```

**示例**

``` JavaScript
/*以下示例显示活动文档中第一个形状的超链接地址。*/
MsgBox(ActiveDocument.Shapes.Item(1).Hyperlink.Address)
```

### InlineShape.IsPictureBullet

如果为 **True** ，则表明 **InlineShape** 对象是图片项目符号。 **Boolean** 类型，只读。

**语法**

**express.IsPictureBullet**

\*express是一个代表 InlineShape 对象的变量。

**说明**

虽然图片项目符号被看作嵌入式图形，但是搜索文档的 **InlineShapes** 集合并不返回图片项目符号。

**示例**

``` JavaScript
/*如果选定列表的格式被设为图片项目符号，本示例设置该列表的格式，否则，显示一条消息。*/
function IsSelectionAPictureBullet(shp){
    try{
        if(shp.IsPictureBullet == true){
          shp.Width = InchesToPoints(0.5)
          shp.Height = InchesToPoints(0.05)
        }
    }
    catch(exception){
          MsgBox("The selection is not a list or " + "does not contain picture bullets.")
    }
}
```

``` JavaScript
/*使用下列代码调用以上程序。*/
function CallPic(){
    IsSelectionAPictureBullet(Selection.Range.ListFormat.ListPictureBullet)
}
```

### InlineShape.Line

返回一个 **LineFormat** 对象，该对象包含指定形状的线条格式属性。只读。

**语法**

**express.Line**

\*express是一个代表 InlineShape 对象的变量。

**说明**

对于有边框的内嵌形状， **LineFormat** 对象代表边框。

### InlineShape.LinkFormat

返回一个 **LinkFormat** 对象，该对象代表链接到文件的指定内嵌形状的链接选项。只读。

**语法**

**express.LinkFormat**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.LockAspectRatio

如果在调整指定形状的大小时保留其最初比例，则该属性值为 **MsoTrue** ；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 **MsoFalse** 。可读写 **MsoTriState** 类型。

**语法**

**express.LockAspectRatio**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.OLEFormat

返回一个 **OLEFormat** 对象，该对象代表指定的内嵌形状的 OLE 特征（链接除外）。只读

**语法**

**express.OLEFormat**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.Parent

返回一个 **Object** 类型值，该值代表指定 **InlineShape** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.PictureFormat

返回一个 **PictureFormat** 对象，该对象包含内嵌形状的图片格式属性。只读。

**语法**

**express.PictureFormat**

\*express是一个代表 InlineShape 对象的变量。

**说明**

适用于代表图片或 OLE 对象的 **InlineShape** 对象。

### InlineShape.Range

返回一个 **Range** 对象，该对象代表内嵌形状中包含的文档部分。

**语法**

**express.Range**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.Reflection

返回一个 **ReflectionFormat** 对象，该对象代表形状的反射格式。只读。

**语法**

**express.Reflection**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.ScaleHeight

参照指定内嵌形状的原始尺寸缩放其高度。可读写 **Single** 类型。

**语法**

**express.ScaleHeight**

\*express是一个代表 InlineShape 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中第一个内嵌形状的高度和宽度设为其原高度和宽度的 150%。*/
function test()
{
ActiveDocument.InlineShapes.Item(1).ScaleHeight = 150
ActiveDocument.InlineShapes.Item(1).ScaleWidth = 150
}
```

### InlineShape.ScaleWidth

参照指定内嵌形状的原始尺寸缩放其宽度。可读写 **Single** 类型。

**语法**

**express.ScaleWidth**

\*express是一个代表 InlineShape 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中第一个内嵌形状的高度和宽度设为其原高度和宽度的 150%。*/
function test()
{
ActiveDocument.InlineShapes.Item(1).ScaleHeight = 150
ActiveDocument.InlineShapes.Item(1).ScaleWidth = 150
}
```

### InlineShape.Script

返回一个 **Script** 对象，该对象代表与指定网页中的图像相关联的脚本或代码块。

**语法**

**express.Script**

\*express是一个代表 InlineShape 对象的变量。

**说明**

如果该网页中不包含任何脚本，则不返回任何内容。

### InlineShape.Shadow

返回一个 **ShadowFormat** 对象，该对象代表指定形状的阴影格式。只读。

**语法**

**express.Shadow**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.SmartArt

返回 SmartArt 对象，该对象提供一种处理与指定内嵌形状相关联的 SmartArt 的方式。只读。

**语法**

**express.SmartArt**

\*express是一个代表 InlineShape 对象的变量。

**说明**

**SmartArt** 属性为与内嵌形状相关联的 SmartArt 图形进行交互提供了一个入口点。

**示例**

``` JavaScript
/*下面的代码示例向活动文档中添加 SmartArt 图形。*/
function test()
{
let myDoc = ActiveDocument
let myInlineShape = myDoc.InlineShapes.AddSmartArt(Application.SmartArtLayouts.Item(2), null, null, null, null, myDoc.Paragraphs.Item(2).Range)
let mySmartArt = myInlineShape.SmartArt
}
```

### InlineShape.SoftEdge

返回一个 **SoftEdgeFormat** 对象，该对象代表形状的软边缘格式。只读。

**语法**

**express.SoftEdge**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.TextEffect

返回一个 **TextEffectFormat** 对象，该对象包含指定内嵌形状的文本效果格式属性。只读。

**语法**

**express.TextEffect**

\*express是一个代表 InlineShape 对象的变量。

**示例**

``` JavaScript
/*以下示例中，如果 myDocument 的第三个形状为艺术字，则将它的字形设为加粗。*/
function test()
{
let myDocument = ActiveDocument
let myShapes = myDocument.Shapes.Item(3)

if(myShapes.Type == msoTextEffect){
    myShapes.TextEffect.FontBold = true
}
}
```

### InlineShape.Title

返回或设置包含指定内嵌形状的标题的 **String** 类型值。可读/写。

**语法**

**express.Title**

\*express是一个代表 InlineShape 对象的变量。

**说明**

使用 **Title** 属性可为内嵌形状提供可选文字标题。该属性将标题文本添加到 WPS 2015 中 **“设置形状格式”** 对话框的 **“可选文字”** 窗格上的 **“标题”** 文本框中。

| 注释                                                                                                                       |
|----------------------------------------------------------------------------------------------------------------------------|
| Web 浏览器在加载表格的过程中或表格丢失时显示可选文字。Web 搜索引擎利用可选文字帮助查找网页。可选文字也可用来帮助残障人士。 |

**示例**

``` JavaScript
/*下面的代码示例向活动文档的第一个内嵌形状中添加可选文字标题。 */
ActiveDocument.InlineShapes.Item(1).Title = "Desktop screenshot."
```

### InlineShape.Type

返回内嵌形状的类型。只读 WdInlineShapeType 类型。

**语法**

**express.Type**

\*express是一个代表 InlineShape 对象的变量。

### InlineShape.Width

返回或设置指定内嵌形状的宽度（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Width**

\*express是一个代表 InlineShape 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/InlineShape](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Pane 对象

## 简介

代表一个窗格。 Pane 对象是 Panes 集合的一个成员。 Panes 集合包含单个窗口的所有窗格。

## 说明

使用 Panes ( *Index* ) 可返回一个 Pane 对象，其中 *Index* 为索引号。以下示例关闭活动窗格。

``` JavaScript
function test()
{
if(ActiveDocument.ActiveWindow.Panes.Count >= 2) {
    ActiveDocument.ActiveWindow.ActivePane.Close()
}
}
```

使用 Add 方法或 Split 属性可添加一个窗格。以下示例按当前窗口大小的 20% 拆分活动窗口。

``` JavaScript
ActiveDocument.ActiveWindow.Panes.Add(20)
```

以下示例将活动窗口一分为二。

``` JavaScript
ActiveDocument.ActiveWindow.Split = true
```

您可以使用 SplitSpecial 属性显示单独窗格中的批注、脚注或尾注。

当窗口被拆分或在非页面视图下显示脚注、批注等信息时，窗口具有多个窗格。以下示例在普通视图中显示批注窗格，然后提示关闭该窗格。

``` JavaScript
/**/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdNormalView
if(ActiveDocument.Comments.Count >= 1) {
    ActiveDocument.ActiveWindow.View.SplitSpecial = wdPaneComments
    let response = MsgBox("Do you want to close the comments pane?",jsYesNo)
    if(response == jsResultYes) {
        ActiveDocument.ActiveWindow.ActivePane.Close()
    }
}
}
```

## 方法

| 序号 | 名称                                 | 说明                                                       |
|:----:|:-------------------------------------|:-----------------------------------------------------------|
|  1   | [Activate](#Pane.Activate)           | 激活指定的窗格。                                           |
|  2   | [AutoScroll](#Pane.AutoScroll)       | 在指定的窗格中自动滚动。                                   |
|  3   | [Close](#Pane.Close)                 | 关闭指定的邮件合并数据源、窗格或任务。                     |
|  4   | [LargeScroll](#Pane.LargeScroll)     | 将窗口或窗格滚动指定的屏数。                               |
|  5   | [NewFrameset](#Pane.NewFrameset)     | 基于指定的窗格新建一个框架页。                             |
|  6   | [PageScroll](#Pane.PageScroll)       | 将指定的窗口按页滚动。                                     |
|  7   | [SmallScroll](#Pane.SmallScroll)     | 将窗口滚动指定的行数。                                     |
|  8   | [TOCInFrameset](#Pane.TOCInFrameset) | 基于指定文档创建一个目录，并将其放入框架页左边的新框架中。 |

## 属性

| 序号 | 名称                                                         | 说明                                                                                  |
|:----:|:-------------------------------------------------------------|:--------------------------------------------------------------------------------------|
|  1   | [Application](#Pane.Application)                             | 返回一个代表 WPS 应用程序的 Application 对象。                                        |
|  2   | [BrowseWidth](#Pane.BrowseWidth)                             | 返回指定窗格中正文换行区域的宽度（以磅为单位）。 Long 类型，只读。                    |
|  3   | [Creator](#Pane.Creator)                                     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。          |
|  4   | [DisplayRulers](#Pane.DisplayRulers)                         | 如果在指定窗格中显示标尺，则该属性值为 True 。 Boolean 类型，可读写。                 |
|  5   | [DisplayVerticalRuler](#Pane.DisplayVerticalRuler)           | 如果在指定窗格中显示垂直标尺，则该属性值为 True 。 Boolean 类型，可读写。             |
|  6   | [Document](#Pane.Document)                                   | 返回一个 Document 对象，该对象与指定窗格相关。只读。                                  |
|  7   | [Frameset](#Pane.Frameset)                                   | 返回一个 Frameset 对象，该对象代表一个整体框架页或框架页的单个框架。只读。            |
|  8   | [HorizontalPercentScrolled](#Pane.HorizontalPercentScrolled) | 返回或设置水平滚动条位置（以文档宽度的百分比表示）。 Long 类型，可读写。              |
|  9   | [Index](#Pane.Index)                                         | 返回一个 Long 类型的值，该值代表项目在集合中的位置。只读。                            |
|  10  | [MinimumFontSize](#Pane.MinimumFontSize)                     | 返回或设置在指定窗格中显示的最小字号（以磅为单位）。 Long 类型，可读写。              |
|  11  | [Next](#Pane.Next)                                           | 返回一个 Pane 对象，该对象代表集合中的下一个文档窗格。只读。                          |
|  12  | [Pages](#Pane.Pages)                                         | 返回一个 Pages 集合，该集合代表文档中的页面。                                         |
|  13  | [Parent](#Pane.Parent)                                       | 返回一个 Object 类型值，该值代表指定 Pane 对象的父对象。                              |
|  14  | [Previous](#Pane.Previous)                                   | 返回一个 Pane 对象，该对象代表集合中的上一个文档窗格。只读。                          |
|  15  | [Selection](#Pane.Selection)                                 | 返回 Selection 对象，该对象代表文档窗格中的所选内容或插入点。只读。                   |
|  16  | [VerticalPercentScrolled](#Pane.VerticalPercentScrolled)     | 返回或设置垂直滚动条的位置（以文档长度的百分比表示）。可读写 Long 类型。              |
|  17  | [View](#Pane.View)                                           | 返回一个 View 对象，该对象代表指定窗格的视图。                                        |
|  18  | [Zooms](#Pane.Zooms)                                         | 返回一个 Zooms 集合，该集合代表每个视图（如普通视图、大纲视图和页面视图）的缩放选项。 |

## 成员方法

### Pane.Activate

激活指定的窗格。

**语法**

**express.Activate ()**

\*express是一个代表 Pane 对象的变量。

**示例**

``` JavaScript
/*本示例拆分活动窗口，然后激活第一个窗格。*/
function SplitWindow() {
    let activewindow = ActiveDocument.ActiveWindow
    activewindow.SplitVertical = 50
    activewindow.Panes.Item(1).Activate()
}
```

### Pane.AutoScroll

在指定的窗格中自动滚动。

**语法**

**express.AutoScroll ( *Velocity* )**

\*express是一个代表 Pane 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明                                                                                                    |
|------------|-----------|----------|---------------------------------------------------------------------------------------------------------|
| *Velocity* | 必选      | Long     | 滚动的速度。取值范围从 -100 到 100。使用 -100 表示以最快速度向后滚动，使用 100 表示以最快速度向前滚动。 |

**说明**

该方法将连续执行，直到通过按键或单击鼠标的方式手动将其停止。

**示例**

``` JavaScript
/*本示例使文档在活动窗格中慢速向后滚动。*/
ActiveDocument.ActiveWindow.ActivePane.AutoScroll(-20)
```

``` JavaScript
/*本示例使文档在活动窗格中以最快速度向前滚动。*/
ActiveDocument.ActiveWindow.ActivePane.AutoScroll(100)
```

### Pane.Close

关闭指定的邮件合并数据源、窗格或任务。

**语法**

**express.Close ()**

\*express是一个代表 Pane 对象的变量。

**示例**

``` JavaScript
/*本示例在拆分活动窗口时关闭活动窗格。*/
function test()
{
if(ActiveDocument.ActiveWindow.Panes.Count >= 2) {
    ActiveDocument.ActiveWindow.ActivePane.Close()
}
}
```

### Pane.LargeScroll

将窗口或窗格滚动指定的屏数。

**语法**

**express.LargeScroll ( *Down* , *Up* , *ToRight* , *ToLeft* )**

\*express是一个代表 Pane 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                 |
|-----------|-----------|----------|----------------------|
| *Down*    | 可选      | Variant  | 向下滚动窗口的屏数。 |
| *Up*      | 可选      | Variant  | 向上滚动窗口的屏数。 |
| *ToRight* | 可选      | Variant  | 向右滚动窗口的屏数。 |
| *ToLeft*  | 可选      | Variant  | 向左滚动窗口的屏数。 |

**说明**

此方法等效于在水平和垂直滚动条上单击滚动块以外的位置。

如果指定了 *Down* 和 *Up* 两个参数，则窗口滚动屏数为这两个参数的差。例如，如果 *Down* 为 2， *Up* 为 4，则窗口向上滚动两屏。同样，如果指定了 *ToLeft* 和 *ToRight* 两个参数，则窗口滚动屏数也为这两个参数的差。

这四个参数都可以取负值。如果没有指定参数，窗口将向下滚动一屏。

### Pane.NewFrameset

基于指定的窗格新建一个框架页。

**语法**

**express.NewFrameset ()**

\*express是一个代表 Pane 对象的变量。

**说明**

有关创建框架页的详细信息，请参阅 创建框架页。

**示例**

``` JavaScript
/*本示例将打开一个名为“Temp.doc”的文件，然后创建一个新的框架页，该框架页中唯一的一个框架中包含“Temp.doc”。*/
function test()
{
Documents.Open("C:\\Documents\\Temp.doc")
ActiveDocument.ActiveWindow.ActivePane.NewFrameset()
}
```

### Pane.PageScroll

将指定的窗口按页滚动。

**语法**

**express.PageScroll ( *Down* , *Up* )**

\*express是一个代表 Pane 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                               |
|--------|-----------|----------|----------------------------------------------------|
| *Down* | 可选      | Variant  | 向下滚动的页数。如果省略此参数，则参数值设置为 1。 |
| *Up*   | 可选      | Variant  | Variant 向上滚动的页数。                           |

**说明**

**PageScroll** 方法仅用于页面视图或 Web 版式视图。此方法不影响插入点的位置。

如果指定了 *Down* 和 *Up* 两个参数，则窗口滚动页数为这两个参数的差。例如，如果 *Down* 为 2， *Up* 为 4，则窗口向上滚动两页。

**示例**

``` JavaScript
/*以下示例使活动窗格向上滚动一页。*/
function test()
{
ActiveDocument.ActiveWindow.View.Type = wdPrintView
ActiveDocument.ActiveWindow.ActivePane.PageScroll(null,1)
}
```

### Pane.SmallScroll

将窗口滚动指定的行数。

**语法**

**express.SmallScroll ( *Down* , *Up* , *ToRight* , *ToLeft* )**

\*express是一个代表 Pane 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                                                             |
|-----------|-----------|----------|----------------------------------------------------------------------------------|
| *Down*    | 可选      | Variant  | 向下滚动窗口的行数。一“行”对应于单击一次垂直滚动条上的向下滚动箭头所滚动的距离。 |
| *Up*      | 可选      | Variant  | 向上滚动窗口的行数。一“行”对应于单击一次垂直滚动条上的向上滚动箭头所滚动的距离。 |
| *ToRight* | 可选      | Variant  | 向右滚动窗口的行数。一“行”对应于单击一次水平滚动条上的向右滚动箭头所滚动的距离。 |
| *ToLeft*  | 可选      | Variant  | 向左滚动窗口的行数。一“行”对应于单击一次水平滚动条上的向左滚动箭头所滚动的距离。 |

**说明**

此方法等效于单击水平和垂直滚动条上的滚动箭头。

如果指定了 *Down* 和 *Up* 两个参数，则窗口滚动行数为这两个参数的差。例如，如果 *Down* 为 3， *Up* 为 6，则窗口将向上滚动三行。同样，如果指定了 *ToLeft* 和 *ToRight* 两个参数，则窗口滚动行数也为这两个参数的差。

这四个参数都可以取负值。如果没有指定参数，窗口将下滚一行。

### Pane.TOCInFrameset

基于指定文档创建一个目录，并将其放入框架页左边的新框架中。

**语法**

**express.TOCInFrameset ()**

\*express是一个代表 Pane 对象的变量。

**说明**

有关创建框架页的详细信息，请参阅 创建框架页。

**示例**

``` JavaScript
/*本示例实现的功能是：打开一个名为“Proposal.doc”的文件，基于该文件创建一个框架页，并添加一个含有该文件目录的框架（位于页面的左侧）。*/
function test()
{
Documents.Open("C:\\Documents\\Proposal.doc")
ActiveDocument.ActiveWindow.ActivePane.NewFrameset()
ActiveDocument.ActiveWindow.ActivePane.TOCInFrameset()
}
```

## 成员属性

### Pane.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Pane 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Pane.BrowseWidth

返回指定窗格中正文换行区域的宽度（以磅为单位）。 **Long** 类型，只读。

**语法**

**express.BrowseWidth**

\*express是一个代表 Pane 对象的变量。

**说明**

此属性仅用于 Web 版式视图。

### Pane.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Pane 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Pane.DisplayRulers

如果在指定窗格中显示标尺，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayRulers**

\*express是一个代表 Pane 对象的变量。

**说明**

**DisplayRulers** 属性相当于 **“视图”** 菜单上的 **“标尺”** 命令。如果 **DisplayRulers** 为 **False** ，则无论 **DisplayVerticalRuler** 属性处于何状态，都不显示水平和垂直标尺。

**示例**

``` JavaScript
/*本示例将活动窗格切换到页面视图并显示水平和垂直标尺。*/
function test()
{
let activepane = ActiveDocument.ActiveWindow.ActivePane
activepane.View.Type = wdPrintView
activepane.DisplayRulers = true
activepane.DisplayVerticalRuler = true
}
```

### Pane.DisplayVerticalRuler

如果在指定窗格中显示垂直标尺，则该属性值为 **True** 。 **Boolean** 类型，可读写。

**语法**

**express.DisplayVerticalRuler**

\*express是一个代表 Pane 对象的变量。

**示例**

``` JavaScript
/*本示例将活动窗格切换到页面视图并显示水平和垂直标尺。*/
function test()
{
let activepane = ActiveDocument.ActiveWindow.ActivePane
activepane.View.Type = wdPrintView
activepane.DisplayRulers = true
activepane.DisplayVerticalRuler = true
}
```

### Pane.Document

返回一个 **Document** 对象，该对象与指定窗格相关。只读。

**语法**

**express.Document**

\*express是一个代表 Pane 对象的变量。

### Pane.Frameset

返回一个 **Frameset** 对象，该对象代表一个整体框架页或框架页的单个框架。只读。

**语法**

**express.Frameset**

\*express是一个代表 Pane 对象的变量。

**说明**

有关创建框架页的详细信息，请参阅 创建框架页。

**示例**

``` JavaScript
/*本示例紧邻指定框架的右侧添加一个新框架。*/
ActiveDocument.ActiveWindow.ActivePane.Frameset.AddNewFrame(wdFramesetNewRight)
 
```

### Pane.HorizontalPercentScrolled

返回或设置水平滚动条位置（以文档宽度的百分比表示）。 **Long** 类型，可读写。

**语法**

**express.HorizontalPercentScrolled**

\*express是一个代表 Pane 对象的变量。

**示例**

``` JavaScript
/*以下示例将 Document1 窗口的活动窗格水平滚动至最左边。*/
function test()
{
Windows.Item("Document1").Activate()
Windows.Item("Document1").ActivePane.HorizontalPercentScrolled = 0
}
```

### Pane.Index

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Pane 对象的变量。

### Pane.MinimumFontSize

返回或设置在指定窗格中显示的最小字号（以磅为单位）。 **Long** 类型，可读写。

**语法**

**express.MinimumFontSize**

\*express是一个代表 Pane 对象的变量。

**说明**

该属性将只影响 Web 版式视图中显示的文本，而 **“格式”** 工具栏上显示的磅值和用于打印的磅值则不会改变。

**示例**

``` JavaScript
/*本示例将活动窗口设置为大纲视图，并将活动窗格中的最小字号设置为 12 磅。*/
function test()
{
let activewindow = ActiveDocument.ActiveWindow
activewindow.View.Type = wdWebView
activewindow.ActivePane.MinimumFontSize = 12
}
```

``` JavaScript
/*本示例返回活动窗格的最小字号。*/
MsgBox(ActiveDocument.ActiveWindow.ActivePane.MinimumFontSize)
```

### Pane.Next

返回一个 **Pane** 对象，该对象代表集合中的下一个文档窗格。只读。

**语法**

**express.Next**

\*express是一个代表 Pane 对象的变量。

### Pane.Pages

返回一个 **Pages** 集合，该集合代表文档中的页面。

**语法**

**express.Pages**

\*express是一个代表 Pane 对象的变量。

**示例**

``` JavaScript
/*下列示例从活动文档左上角 0.5 英寸处穿过页面向页面右下角（距页面右边缘和下边缘 0.5 英寸处）创建一条线。*/
function test()
{
let objPage = ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1)

//Add new line to document
ActiveDocument.Shapes.AddLine(InchesToPoints(0.5), InchesToPoints(0.5), objPage.Width - InchesToPoints(0.5), objPage.Height - InchesToPoints(0.5))
 


}
```

### Pane.Parent

返回一个 **Object** 类型值，该值代表指定 **Pane** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Pane 对象的变量。

### Pane.Previous

返回一个 **Pane** 对象，该对象代表集合中的上一个文档窗格。只读。

**语法**

**express.Previous**

\*express是一个代表 Pane 对象的变量。

### Pane.Selection

返回 **Selection** 对象，该对象代表文档窗格中的所选内容或插入点。只读。

**语法**

**express.Selection**

\*express是一个代表 Pane 对象的变量。

### Pane.VerticalPercentScrolled

返回或设置垂直滚动条的位置（以文档长度的百分比表示）。可读写 **Long** 类型。

**语法**

**express.VerticalPercentScrolled**

\*express是一个代表 Pane 对象的变量。

**示例**

``` JavaScript
/*以下示例将 Document1 窗口的活动窗格垂直滚动至结尾。*/
function test()
{
let document = Windows.Item("Document1")
document.Activate()
document.ActivePane.VerticalPercentScrolled = 100
}
```

### Pane.View

返回一个 **View** 对象，该对象代表指定窗格的视图。

**语法**

**express.View**

\*express是一个代表 Pane 对象的变量。

**示例**

``` JavaScript
/*以下示例显示与 Windows 集合中的第一个窗口相关联的窗格的所有非打印字符。*/
function test()
{
for(let i = 1; i <= Windows.Item(1).Panes.Count; i++) {
    Windows.Item(1).Panes.Item(i).View.ShowAll = true
}
}
```

### Pane.Zooms

返回一个 **Zooms** 集合，该集合代表每个视图（如普通视图、大纲视图和页面视图）的缩放选项。

**语法**

**express.Zooms**

\*express是一个代表 Pane 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*本示例将在页面视图中打开的所有窗口的显示比例设为 100%。*/
function test()
{
for(let i = 1; i <= Windows.Count; i++) {
    Windows.Item(i).ActivePane.Zooms.Item(wdPrintView).Percentage = 100
}
}
```

``` JavaScript
/*本示例设置页面视图的显示比例，使整个页面可见。*/
ActiveDocument.ActiveWindow.Panes.Item(1).Zooms.Item(wdPrintView).PageFit = wdPageFitFullPage
 
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Pane](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Paragraphs 对象

## 方法

| 序号 | 名称                                                             | 说明                                                            |
|:----:|:-----------------------------------------------------------------|:----------------------------------------------------------------|
|  1   | [Add](#Paragraphs.Add)                                           | 返回一个 Paragraph 对象，该对象代表添加到文档中的新的空白段落。 |
|  2   | [Indent](#Paragraphs.Indent)                                     | 为一个或多个段落增加一个级别的缩进。                            |
|  3   | [IndentFirstLineCharWidth](#Paragraphs.IndentFirstLineCharWidth) | 将一个或多个段落的首行缩进指定的字符数。                        |

## 属性

| 序号 | 名称                                                                           | 说明                                                                                                                                                                                            |
|:----:|:-------------------------------------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AddSpaceBetweenFarEastAndAlpha](#Paragraphs.AddSpaceBetweenFarEastAndAlpha)   | 如果 WPS 将自动在指定段落的日文和拉丁文文字之间添加空格，则该属性值为 True 。如果仅对于某些指定段落将该属性设置为 True ，则该属性会返回 wdUndefined 。 Long 类型，可读写。                      |
|  2   | [AddSpaceBetweenFarEastAndDigit](#Paragraphs.AddSpaceBetweenFarEastAndDigit)   | 如果 WPS 将自动在指定段落的日文文字和数字之间添加空格，则该属性值为 True 。如果仅对于某些指定段落将该属性设置为 True ，则该属性会返回 wdUndefined 。 Long 类型，可读写。                        |
|  3   | [Alignment](#Paragraphs.Alignment)                                             | 返回或设置一个 WdParagraphAlignment 常量，该常量代表指定段落的对齐方式，可读写。                                                                                                                |
|  4   | [Application](#Paragraphs.Application)                                         | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                                                  |
|  5   | [AutoAdjustRightIndent](#Paragraphs.AutoAdjustRightIndent)                     | 如果 WPS 会根据您指定的每行字符数自动调整指定段落的右缩进，则该属性值为 True 。如果只将某些指定段落的 AutoAdjustRightIndent 属性设置为 True ，则该属性会返回 wdUndefined 。 Long 类型，可读写。 |
|  6   | [BaseLineAlignment](#Paragraphs.BaseLineAlignment)                             | 返回或设置一个 WdBaselineAlignment 常量，该常量代表行中字体的垂直位置，可读写。                                                                                                                 |
|  7   | [Borders](#Paragraphs.Borders)                                                 | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                                                                                                                           |
|  8   | [CharacterUnitFirstLineIndent](#Paragraphs.CharacterUnitFirstLineIndent)       | 返回或设置首行或悬挂缩进的值（以字符为单位）。用正值设置首行缩进，用负值设置悬挂缩进。 Single 类型，可读写。                                                                                    |
|  9   | [CharacterUnitLeftIndent](#Paragraphs.CharacterUnitLeftIndent)                 | 该属性返回或设置指定段落的左缩进量（以字符为单位）。 Single 类型，可读写。                                                                                                                      |
|  10  | [CharacterUnitRightIndent](#Paragraphs.CharacterUnitRightIndent)               | 该属性返回或设置指定段落的右缩进量（以字符为单位）。 Single 类型，可读写。                                                                                                                      |
|  11  | [Count](#Paragraphs.Count)                                                     | 返回一个 Long 类型的值，该值代表集合中的段数。只读。                                                                                                                                            |
|  12  | [Creator](#Paragraphs.Creator)                                                 | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                                                    |
|  13  | [DisableLineHeightGrid](#Paragraphs.DisableLineHeightGrid)                     | 如果该属性的值为 True ，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 DisableLineHeightGrid 属性设置为 True ，则返回 wdUndefined 。 Long 类型，可读写。    |
|  14  | [FarEastLineBreakControl](#Paragraphs.FarEastLineBreakControl)                 | 如果为 True ，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 FarEastLineBreakControl 属性设定为 True ，则返回 wdUndefined 。 Long 类型，可读写。                     |
|  15  | [First](#Paragraphs.First)                                                     | 返回一个 Paragraph 对象，该对象代表在 Paragraphs 集合中的第一个项目。                                                                                                                           |
|  16  | [FirstLineIndent](#Paragraphs.FirstLineIndent)                                 | 返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。 Single 类型，可读写。                                                                    |
|  17  | [Format](#Paragraphs.Format)                                                   | 返回或设置一个 ParagraphFormat 对象，该对象代表指定的一个或多个段落的格式。                                                                                                                     |
|  18  | [HalfWidthPunctuationOnTopOfLine](#Paragraphs.HalfWidthPunctuationOnTopOfLine) | 如果为 True ，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 True ，则此属性将返回 wdUndefined 。 Long 类型，可读写。                                        |
|  19  | [HangingPunctuation](#Paragraphs.HangingPunctuation)                           | 如果为 True ，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 True ，则返回 wdUndefined 。 Long 类型，可读写。                                                             |
|  20  | [Hyphenation](#Paragraphs.Hyphenation)                                         | 如果指定的段落进行自动断字，则该属性值为 True 。如果指定的段落不进行自动断字，则该属性值为 False 。可读写 Long 类型。                                                                           |
|  21  | [KeepTogether](#Paragraphs.KeepTogether)                                       | 在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 True 。可读写 Long 类型。                                                                                           |
|  22  | [KeepWithNext](#Paragraphs.KeepWithNext)                                       | 在 WPS 对文档重新分页时，如果指定段落与其下一段位于同一页上，则该属性值为 True 。可读写 Long 类型。                                                                                             |
|  23  | [Last](#Paragraphs.Last)                                                       | 返回一个 Paragraph 对象，该对象代表段落集合中的最后一个项目。                                                                                                                                   |
|  24  | [LeftIndent](#Paragraphs.LeftIndent)                                           | 返回或设置一个 Single 类型的值，该值代表指定段落的左缩进值（以磅为单位）。可读写。                                                                                                              |
|  25  | [LineSpacing](#Paragraphs.LineSpacing)                                         | 返回或设置指定段落的行距（以磅为单位）。 Single 类型，可读写。                                                                                                                                  |
|  26  | [LineSpacingRule](#Paragraphs.LineSpacingRule)                                 | 返回或设置指定段落的行距。可读写 WdLineSpacing 类型。                                                                                                                                           |
|  27  | [LineUnitAfter](#Paragraphs.LineUnitAfter)                                     | 返回或设置指定段落的段后间距（以网格线为单位）。可读写 Single 类型。                                                                                                                            |
|  28  | [LineUnitBefore](#Paragraphs.LineUnitBefore)                                   | 返回或设置指定段落的段前间距（以网格线为单位）。可读写 Single 类型。                                                                                                                            |
|  29  | [NoLineNumber](#Paragraphs.NoLineNumber)                                       | 如果取消指定段的行号，则该属性值为 True 。该属性值可以是 True 、 False 或 wdUndefined 。可读写 Long 类型。                                                                                      |
|  30  | [OutlineLevel](#Paragraphs.OutlineLevel)                                       | 返回或设置指定段落的大纲级别。可读写 WdOutlineLevel 类型。                                                                                                                                      |
|  31  | [PageBreakBefore](#Paragraphs.PageBreakBefore)                                 | 如果在指定段落前插入了分页符，则该属性值为 True 。该属性值可以是 True 、 False 或 wdUndefined 。可读写 Long 类型。                                                                              |
|  32  | [Parent](#Paragraphs.Parent)                                                   | 返回一个 Object 类型值，该值代表指定 Paragraphs 对象的父对象。                                                                                                                                  |
|  33  | [ReadingOrder](#Paragraphs.ReadingOrder)                                       | 返回或设置指定段落的读取次序而不改变其对齐方式。可读写 WdReadingOrder 类型。                                                                                                                    |
|  34  | [RightIndent](#Paragraphs.RightIndent)                                         | 返回或设置指定段落的右缩进量（以磅为单位）。可读写 Single 类型。                                                                                                                                |
|  35  | [Shading](#Paragraphs.Shading)                                                 | 返回一个 Shading 对象，该对象代表指定段落的底纹格式。                                                                                                                                           |
|  36  | [SpaceAfter](#Paragraphs.SpaceAfter)                                           | 返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 Single 类型。                                                                                                                       |
|  37  | [SpaceAfterAuto](#Paragraphs.SpaceAfterAuto)                                   | 如果 WPS 自动设置指定段落的段后间距，则该属性为 True 。可读/写 Long 类型。                                                                                                                      |
|  38  | [SpaceBefore](#Paragraphs.SpaceBefore)                                         | 返回或设置指定段落的段前间距（以磅为单位）。可读/写 Single 类型。                                                                                                                               |
|  39  | [SpaceBeforeAuto](#Paragraphs.SpaceBeforeAuto)                                 | 如果 WPS 自动设置指定段落的段前间距，则该属性为 True 。可读/写 Long 类型。                                                                                                                      |
|  40  | [Style](#Paragraphs.Style)                                                     | 返回或设置指定段落的样式。可读写 Variant 类型。                                                                                                                                                 |
|  41  | [TabStops](#Paragraphs.TabStops)                                               | 返回或设置一个 TabStops 集合，该集合代表指定段落中的所有自定义制表位。可读写。                                                                                                                  |
|  42  | [WidowControl](#Paragraphs.WidowControl)                                       | 在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 True 。该属性值可以是 True 、 False 或 wdUndefined 。可读写 Long 类型。                                     |
|  43  | [WordWrap](#Paragraphs.WordWrap)                                               | 如果 WPS 在指定段落中的西文单词中间断字换行，则该属性值为 True 。可读写 Long 类型。                                                                                                             |

## 成员方法

### Paragraphs.Add

返回一个 **Paragraph** 对象，该对象代表添加到文档中的新的空白段落。

**语法**

**express.Add ( *Range* )**

\*express是一个代表 Paragraphs 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                             |
|---------|-----------|----------|--------------------------------------------------|
| *Range* | 可选      | Variant  | 要在其前添加新段落的区域。新的段落不替换该区域。 |

**说明**

如果不指定 *Range* ，则将新段落添加到选定内容或区域之后，或者添加到文档最后，具体情况取决于 expression 的设置。

**示例**

``` JavaScript
/*本示例在选定内容之后添加一个段落。*/
Application.Selection.Paragraphs.Add()

/*本示例在选定内容中第一段之前添加一个段落标记。*/
Application.Selection.Paragraphs.Add(Selection.Paragraphs.Item(1).Range)

/*本示例在活动文档第二段之前添加一个段落标记。*/
Application.ActiveDocument.Paragraphs.Add(ActiveDocument.Paragraphs.Item(2).Range)

/*本示例在活动文档的末尾添加一个新的段落标记。*/
Application.ActiveDocument.Paragraphs.Add()
```

### Paragraphs.Indent

为一个或多个段落增加一个级别的缩进。

**语法**

**express.Indent ()**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

此方法等效于单击 **“格式”** 工具栏上的 **“增加缩进量”** 按钮。

**示例**

``` JavaScript
/*以下示例为活动文档中的所有段落增加两个级别的缩进，然后为第一段删除一个级别的缩进。*/
function test() {
  Application.ActiveDocument.Paragraphs.Indent()
  Application.ActiveDocument.Paragraphs.Item(1).Outdent()
}
```

### Paragraphs.IndentFirstLineCharWidth

将一个或多个段落的首行缩进指定的字符数。

**语法**

**express.IndentFirstLineCharWidth ( *Count* )**

\*express是一个代表 Paragraphs 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                               |
|---------|-----------|----------|------------------------------------|
| *Count* | 必选      | Integer  | 每个指定段落的首行要缩进的字符数。 |

**示例**

``` JavaScript
/*以下示例将活动文档中所有段落的首行缩进 10 个字符。*/
Application.ActiveDocument.Paragraphs.IndentFirstLineCharWidth(10)
```

## 成员属性

### Paragraphs.AddSpaceBetweenFarEastAndAlpha

如果 WPS 将自动在指定段落的日文和拉丁文文字之间添加空格，则该属性值为 **True** 。如果仅对于某些指定段落将该属性设置为 **True** ，则该属性会返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndAlpha**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：设置 WPS 在活动文档第一段的日文文字和拉丁文字之间自动添加空格。*/
ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndAlpha = true
```

### Paragraphs.AddSpaceBetweenFarEastAndDigit

如果 WPS 将自动在指定段落的日文文字和数字之间添加空格，则该属性值为 **True** 。如果仅对于某些指定段落将该属性设置为 **True** ，则该属性会返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndDigit**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例可实现的功能是：设定 WPS 自动在活动文档第一段的中文文字和数字之间添加空格。*/
ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndDigit = true
```

### Paragraphs.Alignment

返回或设置一个 **WdParagraphAlignment** 常量，该常量代表指定段落的对齐方式，可读写。

**语法**

**express.Alignment**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

对于所选择或安装的不同语言支持（例如美国英语），某些 **WdParagraphAlignment** 常量可能不可用。

### Paragraphs.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Paragraphs.AutoAdjustRightIndent

如果 WPS 会根据您指定的每行字符数自动调整指定段落的右缩进，则该属性值为 **True** 。如果只将某些指定段落的 **AutoAdjustRightIndent** 属性设置为 **True** ，则该属性会返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.AutoAdjustRightIndent**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 根据指定的每行字符数，自动调整所选段落的右缩进。*/
Selection.ParagraphFormat.AutoAdjustRightIndent = true
```

### Paragraphs.BaseLineAlignment

返回或设置一个 **WdBaselineAlignment** 常量，该常量代表行中字体的垂直位置，可读写。

**语法**

**express.BaseLineAlignment**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 自动调整活动文档中的基线字体对齐方式。*/
ActiveDocument.BaseLineAlignment = wdBaselineAlignAuto
```

### Paragraphs.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Paragraphs.CharacterUnitFirstLineIndent

返回或设置首行或悬挂缩进的值（以字符为单位）。用正值设置首行缩进，用负值设置悬挂缩进。 **Single** 类型，可读写。

**语法**

**express.CharacterUnitFirstLineIndent**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中第一段的首行缩进设为一个字符。*/
ActiveDocument.Paragraphs.Item(1).CharacterUnitFirstLineIndent = 1
```

``` JavaScript
/*本示例将活动文档中第二段的悬挂缩进设为 1.5 个字符。*/
ActiveDocument.Paragraphs.Item(2).CharacterUnitFirstLineIndent = -1.5
```

### Paragraphs.CharacterUnitLeftIndent

该属性返回或设置指定段落的左缩进量（以字符为单位）。 **Single** 类型，可读写。

**语法**

**express.CharacterUnitLeftIndent**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中第一段的左缩进设为从左边距缩进一个字符。*/
ActiveDocument.Paragraphs.Item(1).CharacterUnitLeftIndent = 1
 
```

### Paragraphs.CharacterUnitRightIndent

该属性返回或设置指定段落的右缩进量（以字符为单位）。 **Single** 类型，可读写。

**语法**

**express.CharacterUnitRightIndent**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中的所有段落的右缩进设为从右边距缩进一个字符。*/
ActiveDocument.Paragraphs.CharacterUnitRightIndent = 1
```

### Paragraphs.Count

返回一个 **Long** 类型的值，该值代表集合中的段数。只读。

**语法**

**express.Count**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例显示活动文档的段数。*/
alert("The active document contains " + ActiveDocument.Paragraphs.Count + " paragraphs.")
```

### Paragraphs.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Paragraphs.DisableLineHeightGrid

如果该属性的值为 **True** ，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 **DisableLineHeightGrid** 属性设置为 **True** ，则返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.DisableLineHeightGrid**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 在已指定每页行数的情况下，将选定段落中的字符与行网格对齐。*/
Selection.ParagraphFormat.DisableLineHeightGrid = true
```

### Paragraphs.FarEastLineBreakControl

如果为 **True** ，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 **FarEastLineBreakControl** 属性设定为 **True** ，则返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.FarEastLineBreakControl**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例设定 WPS 将东亚语言文字的换行规则应用于活动文档的第一段。*/
ActiveDocument.Paragraphs.Item(1).FarEastLineBreakControl = true
```

### Paragraphs.First

返回一个 **Paragraph** 对象，该对象代表在 **Paragraphs** 集合中的第一个项目。

**语法**

**express.First**

\*express是一个代表 Paragraphs 对象的变量。

### Paragraphs.FirstLineIndent

返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。 **Single** 类型，可读写。

**语法**

**express.FirstLineIndent**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例为活动文档的首段设置 1 英寸的首行缩进。*/
Application.ActiveDocument.Paragraphs.FirstLineIndent = InchesToPoints(1)

/*本示例为活动文档的第二段设置 0.5 英寸的悬挂缩进。InchesToPoints 方法用来将英寸转化为磅值。*/
Application.ActiveDocument.Paragraphs.FirstLineIndent = InchesToPoints(-0.5)
```

### Paragraphs.Format

返回或设置一个 **ParagraphFormat** 对象，该对象代表指定的一个或多个段落的格式。

**语法**

**express.Format**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档的所有段落靠左对齐。*/
ActiveDocument.Paragraphs.Format.Alignment = wdAlignParagraphLeft
```

### Paragraphs.HalfWidthPunctuationOnTopOfLine

如果为 **True** ，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 **True** ，则此属性将返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.HalfWidthPunctuationOnTopOfLine**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例设置 WPS 将活动文档第一段的行首标点符号改为半角字符。*/
ActiveDocument.Paragraphs.HalfWidthPunctuationOnTopOfLine = true
```

### Paragraphs.HangingPunctuation

如果为 **True** ，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 **True** ，则返回 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.HangingPunctuation**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例实现的功能是：允许活动文档第一段的标点可以溢出边界。*/
ActiveDocument.Paragraphs.HangingPunctuation = true
```

### Paragraphs.Hyphenation

如果指定的段落进行自动断字，则该属性值为 **True** 。如果指定的段落不进行自动断字，则该属性值为 **False** 。可读写 **Long** 类型。

**语法**

**express.Hyphenation**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例关闭活动文档中所有段落的自动断字功能。*/
ActiveDocument.Paragraphs.Hyphenation = false
```

### Paragraphs.KeepTogether

在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.KeepTogether**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例为活动文档中的各段落设置格式，以使在 WPS 对文档重新分页时同一段中的各行都位于同一页上。*/
ActiveDocument.Paragraphs.KeepTogether = true
```

### Paragraphs.KeepWithNext

在 WPS 对文档重新分页时，如果指定段落与其下一段位于同一页上，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.KeepWithNext**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

此属性可以是 **True** 、 **False** 或 **wdUndefined** 。

**示例**

``` JavaScript
/*以下示例将当前所选内容中的所有段落设置为位于同一页上。*/
Selection.Paragraphs.KeepWithNext = true
```

### Paragraphs.Last

返回一个 **Paragraph** 对象，该对象代表段落集合中的最后一个项目。

**语法**

**express.Last**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中的最后一个段落设置为右对齐。*/
ActiveDocument.Paragraphs.Last.Alignment = wdAlignParagraphRight
```

### Paragraphs.LeftIndent

返回或设置一个 **Single** 类型的值，该值代表指定段落的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中所有段落的左缩进都设置为 1 英寸。InchesToPoints 方法用来将英寸转换为磅。*/
ActiveDocument.Paragraphs.LeftIndent = InchesToPoints(1)
```

### Paragraphs.LineSpacing

返回或设置指定段落的行距（以磅为单位）。 **Single** 类型，可读写。

**语法**

**express.LineSpacing**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

使用 **LinesToPoints** 方法可将行数转换为相应的值（以磅为单位）。例如， ` LinesToPoints(2) ` 返回值 24。

可以在 **LineSpacingRule** 属性已设置为下列值之后对 **LineSpacing** 属性进行设置：

- **wdLineSpaceAtLeast** 行距可以大于或等于指定的 **LineSpacing** 值，但绝不能小于该值。
- **wdLineSpaceExactly** 指定的 **LineSpacing** 行距值绝不能更改，即使段落中使用了较大的字体。
- **wdLineSpaceMultiple** 必须指定 **LineSpacing** 属性值，以磅为单位。

**示例**

``` JavaScript
/*以下示例为所选段落设置 3 倍行距。*/
function test()
{
Selection.Paragraphs.LineSpacingRule = wdLineSpaceMultiple
Selection.Paragraphs.LineSpacing = LinesToPoints(3)
}
```

### Paragraphs.LineSpacingRule

返回或设置指定段落的行距。可读写 **WdLineSpacing** 类型。

**语法**

**express.LineSpacingRule**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

使用 **wdLineSpaceSingle** 、 **wdLineSpace1pt5** 或 **wdLineSpaceDouble** 可将行距设置为这些值之一。要将行距设置为固定磅值或多倍行距，还必须设置 **LineSpacing** 属性。

**示例**

``` JavaScript
/*以下示例为活动文档的所有段落设置 2 倍行距。*/
ActiveDocument.Paragraphs.LineSpacingRule = wdLineSpaceDouble
```

### Paragraphs.LineUnitAfter

返回或设置指定段落的段后间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitAfter**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中所有段落的段后间距设置为 1 个网格线。*/
ActiveDocument.Paragraphs.LineUnitAfter = 1
```

### Paragraphs.LineUnitBefore

返回或设置指定段落的段前间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitBefore**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中所有段落的段前间距设置为 1 个网格线。*/
ActiveDocument.Paragraphs.LineUnitBefore = 1
```

### Paragraphs.NoLineNumber

如果取消指定段的行号，则该属性值为 **True** 。该属性值可以是 **True** 、 **False** 或 **wdUndefined** 。可读写 **Long** 类型。

**语法**

**express.NoLineNumber**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

使用 **PageSetup** 对象的 **LineNumbering** 属性可设置行号。

**示例**

``` JavaScript
/*以下示例为活动文档设置行号。起始编号设置为 1，并且在文档中的所有节都连续编号。然后取消第二段的行号。*/
function test()
{
let linenum = ActiveDocument.PageSetup.LineNumbering
    linenum.Active = true
    linenum.StartingNumber = 1
    linenum.CountBy = 1
    linenum.RestartMode = wdRestartContinuous
ActiveDocument.Paragraphs.Item(2).NoLineNumber = true
}
```

### Paragraphs.OutlineLevel

返回或设置指定段落的大纲级别。可读写 **WdOutlineLevel** 类型。

**语法**

**express.OutlineLevel**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

如果段落应用了一种标题样式（从“标题 1”到“标题 9”），则大纲级别与该标题样式相同，并且不能更改。大纲级别仅在大纲视图或文档结构图窗格中可见。

**示例**

``` JavaScript
/*以下示例返回活动文档中所有段落的大纲级别。*/
let temp = ActiveDocument.Paragraphs.OutlineLevel
```

### Paragraphs.PageBreakBefore

如果在指定段落前插入了分页符，则该属性值为 **True** 。该属性值可以是 **True** 、 **False** 或 **wdUndefined** 。可读写 **Long** 类型。

**语法**

**express.PageBreakBefore**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*以下示例在所选内容的第一段前插入分页符。*/
Selection.Paragraphs.PageBreakBefore = true
```

### Paragraphs.Parent

返回一个 **Object** 类型值，该值代表指定 **Paragraphs** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Paragraphs 对象的变量。

### Paragraphs.ReadingOrder

返回或设置指定段落的读取次序而不改变其对齐方式。可读写 **WdReadingOrder** 类型。

**语法**

**express.ReadingOrder**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

使用 **Selection** 对象的 **LtrPara** 、 **LtrRun** 、 **RtlPara** 和 **RtlRun** 方法可更改段落的对齐方式和读取次序。

**示例**

``` JavaScript
/*以下示例将活动文档中所有段落的读取次序设置为从右向左。*/
ActiveDocument.Paragraphs.ReadingOrder = wdReadingOrderRtl
```

### Paragraphs.RightIndent

返回或设置指定段落的右缩进量（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.RightIndent**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*以下示例将活动文档中所有段落的右缩进都设置为距右边距 1 英寸。InchesToPoints 方法用来将英寸转换为磅。*/
ActiveDocument.Paragraphs.RightIndent = InchesToPoints(1)
```

### Paragraphs.Shading

返回一个 **Shading** 对象，该对象代表指定段落的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*以下示例对所选内容的所有段落应用黄色底纹。*/
function test()
{
let shadingParagraph = Selection.Paragraphs.Shading
    shadingParagraph.Texture = wdTexture12Pt5Percent
    shadingParagraph.BackgroundPatternColorIndex = wdYellow
    shadingParagraph.ForegroundPatternColorIndex = wdBlack
}
```

### Paragraphs.SpaceAfter

返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceAfter**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中所有段落的段后间距设置为 12 磅。*/
ActiveDocument.Paragraphs.SpaceAfter = 12
```

### Paragraphs.SpaceAfterAuto

如果 WPS 自动设置指定段落的段后间距，则该属性为 **True** 。可读/写 **Long** 类型。

**语法**

**express.SpaceAfterAuto**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceAfterAuto** 属性设置为 **True** ，则该属性返回 **wdUndefined** 。该属性可被设置为 **True** 或 **False** 。

如果将 **SpaceAfterAuto** 设置为 **True** ，则 **SpaceAfter** 属性将被忽略。

**示例**

``` JavaScript
/*本示例显示关于活动文档的 SpaceAfterAuto 属性设置的报告。*/
function test()
{
let mySpaceAfterAuto = ActiveDocument.Paragraphs.SpaceAfterAuto
switch(mySpaceAfterAuto){
    case 0:
        let x = "Spacing after paragraphs is handled " + "automatically for all paragraphs."
        break
    case 1:
        let x = "Spacing after paragraphs is handled " + "manually for all paragraphs."
        break
    case undefined:
        let x = "Spacing after paragraphs is handled " + "automatically for some paragraphs, " + "manually for some paragraphs."
}
}
```

### Paragraphs.SpaceBefore

返回或设置指定段落的段前间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceBefore**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*本示例将活动文档中所有段落的段前间距设置为 12 磅。*/
ActiveDocument.Paragraphs.SpaceBefore = 12
```

### Paragraphs.SpaceBeforeAuto

如果 WPS 自动设置指定段落的段前间距，则该属性为 **True** 。可读/写 **Long** 类型。

**语法**

**express.SpaceBeforeAuto**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceBeforeAuto** 属性设置为 **True** ，则该属性返回 **wdUndefined** 。该属性可被设置为 **True** 或 **False** 。

如果将 **SpaceBeforeAuto** 设置为 **True** ，则 **SpaceBefore** 属性将被忽略。

**示例**

``` JavaScript
/*本示例显示关于活动文档的 SpaceBeforeAuto 属性设置的报告。*/
function test()
{
let mySpaceBeforeAuto = ActiveDocument.Paragraphs.SpaceBeforeAuto
switch(mySpaceBeforeAuto){
    case 0:
        let x = "Spacing before paragraphs is handled " + "automatically for all paragraphs."
        break
    case 1:
        let x = "Spacing before paragraphs is handled " + "manually for all paragraphs."
        break
    case undefined:
        let x = "Spacing before paragraphs is handled " + "automatically for some paragraphs, " + "manually for some paragraphs."
}
}
```

### Paragraphs.Style

返回或设置指定段落的样式。可读写 **Variant** 类型。

**语法**

**express.Style**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

要设置该属性，请指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个代表样式的对象。

如果返回包含多个样式的范围的样式，则只返回第一个字符样式或段落样式。

### Paragraphs.TabStops

返回或设置一个 **TabStops** 集合，该集合代表指定段落中的所有自定义制表位。可读写。

**语法**

**express.TabStops**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

``` JavaScript
/*以下示例在活动文档各段的 2 英寸处添加居中的制表位。InchesToPoints 方法用来将英寸转换为磅。*/
ActiveDocument.Paragraphs.TabStops.Add(InchesToPoints(2),wdAlignTabCenter,null)
```

``` JavaScript
/*以下示例将文档中各段的制表位设置为与第一段的制表位一致。*/
function test()
{
let para1Tabs = ActiveDocument.Paragraphs.Item(1).TabStops
ActiveDocument.Paragraphs.TabStops = para1Tabs
}
```

### Paragraphs.WidowControl

在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 **True** 。该属性值可以是 **True** 、 **False** 或 **wdUndefined** 。可读写 **Long** 类型。

**语法**

**express.WidowControl**

\*express是一个代表 Paragraphs 对象的变量。

**示例**

``` JavaScript
/*以下示例给活动文档中的各段落设置格式，从而使段中的首行或末行有可能单独位于上页的页尾或下页的页首。*/
ActiveDocument.Paragraphs.WidowControl = false
```

### Paragraphs.WordWrap

如果 WPS 在指定段落中的西文单词中间断字换行，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.WordWrap**

\*express是一个代表 Paragraphs 对象的变量。

**说明**

如果仅将某些指定段落的此属性设置为 **True** ，则此属性返回 **wdUndefined** 。根据您所选择或安装的语言支持（例如，美国英语），该属性可能不可用。

**示例**

``` JavaScript
/*以下示例设置 WPS 在活动文档的第一段中的西文单词中间断字换行。*/
ActiveDocument.Paragraphs.Item(1).WordWrap = true
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Paragraphs](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Row 对象

## 简介

代表表格中的一行。 Row 对象是 Rows 集合的一个成员

## 说明

代表表格中的一行。 Row 对象是 Rows 集合的一个成员。 Rows 集合包含指定的所选内容、范围或表格中的所有行。

## 方法

| 序号 | 名称                                | 说明                 |
|:----:|:------------------------------------|:---------------------|
|  1   | [ConvertToText](#Row.ConvertToText) |                      |
|  2   | [Delete](#Row.Delete)               | 删除指定的表格行     |
|  3   | [Select](#Row.Select)               | 选择指定的表格行。   |
|  4   | [SetHeight](#Row.SetHeight)         | 设置表格行的高度。   |
|  5   | [SetLeftIndent](#Row.SetLeftIndent) | 设置表格中行的缩进。 |

## 属性

| 序号 | 名称                                              | 说明                                                                                                                                                                 |
|:----:|:--------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Alignment](#Row.Alignment)                       | 返回或设置一个 WdRowAlignment 常量，该常量代表指定行的对齐方式。可读写。                                                                                             |
|  2   | [AllowBreakAcrossPage](#Row.AllowBreakAcrossPage) | 如果允许分页符拆分表格中一行或多行中的文本，则该属性值为 True 。可读写 Long 类型。                                                                                   |
|  3   | [Application](#Row.Application)                   | 返回一个代表 WPS 应用程序的 Application 对象。                                                                                                                       |
|  4   | [Borders](#Row.Borders)                           | 返回一个 Borders 集合，该集合代表指定对象的所有边框。                                                                                                                |
|  5   | [Cells](#Row.Cells)                               | 返回一个 Cells 集合，该集合代表在某一列、行、选定内容或区域中的表格单元格。只读。                                                                                    |
|  6   | [Creator](#Row.Creator)                           | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。                                                                                         |
|  7   | [HeadingFormat](#Row.HeadingFormat)               | 如果为 True ，则将指定一行或数行的格式设置为表格标题。当表格不止一页时，可将多行的格式重复设置为表格标题。可以是 True 、 False 或 wdUndefined 。 Long 类型，可读写。 |
|  8   | [Height](#Row.Height)                             | 返回或设置表格中指定行的高度。Single 类型，可读写。                                                                                                                  |
|  9   | [HeightRule](#Row.HeightRule)                     | 返回或设置确定指定单元格或行高度的规则。 WdRowHeightRule 类型，可读写。                                                                                              |
|  10  | [ID](#Row.ID)                                     | 在文档另存为网页时返回或设置指定表格行的标识标签。可读写 String 类型。                                                                                               |
|  11  | [Index](#Row.Index)                               | 返回 Long 值，该值表示项目在集合中的位置。只读。                                                                                                                     |
|  12  | [IsFirst](#Row.IsFirst)                           | 如果指定的行是表格中的首行，则该属性值为 True 。否则为False                                                                                                          |
|  13  | [IsLast](#Row.IsLast)                             | 如果指定的行是表格中的首行，则该属性值为 True 。否则为False                                                                                                          |
|  14  | [LeftIndent](#Row.LeftIndent)                     | 返回或设置一个 Single 类型的值，该值代表指定表格行的左缩进值（以磅为单位）。可读写。                                                                                 |
|  15  | [NestingLevel](#Row.NestingLevel)                 | 返回指定表格行的嵌套层。只读 Long 类型。                                                                                                                             |
|  16  | [Next](#Row.Next)                                 | 返回一个 Row 对象，该对象代表表格中行的集合中的下一个表格行。只读。                                                                                                  |
|  17  | [Parent](#Row.Parent)                             | 返回一个 Object 类型值，该值代表指定 Row 对象的父对象。                                                                                                              |
|  18  | [Previous](#Row.Previous)                         | 返回一个 Row 对象，该对象代表指定行的前一个表格行。只读。                                                                                                            |
|  19  | [Range](#Row.Range)                               | 返回一个 Range 对象，该对象代表指定表格行内包含的文档部分                                                                                                            |
|  20  | [Shading](#Row.Shading)                           | 返回一个 Shading 对象，该对象代表指定对象的底纹格式。                                                                                                                |
|  21  | [SpaceBetweenColumns](#Row.SpaceBetweenColumns)   | 返回或设置指定的一行或多行中相邻列的文本之间的距离（以磅为单位）。可读写 Single 类型。                                                                               |

## 成员方法

### Row.ConvertToText

**语法**

**express.ConvertToText ( *Separator* , *NestedTables* )**

\*express是一个代表 Row 对象的变量。

**参数**

| 名称           | 必选/可选 | 数据类型 | 说明                                                                                                                                                                                                      |
|----------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Separator*    | 可选      | Variant  | 用以分隔被转换列的字符（被转换行由段落标记分隔）。可以是下列任何 WdTableFieldSeparator 常量：wdSeparateByCommas、wdSeparateByDefaultListSeparator、wdSeparateByParagraphs 或 wdSeparateByTabs（默认值）。 |
| *NestedTables* | 可选      | Variant  | 如果要将嵌套的表格转换为文本，则为 True。如果 Separator 不是 wdSeparateByParagraphs，则将忽略此参数。默认值为 True。                                                                                      |

**说明**

将表格转换为文本并返回一个 **Range** 对象，该对象代表带分隔符的文本。

### Row.Delete

删除指定的表格行

**语法**

**express.Delete ()**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
//删除选中区域的表格的第一行
Application.Selection.Rows.Item(1).Delete()
```

### Row.Select

选择指定的表格行。

**语法**

**express.Select ()**

\*express是一个代表 Row 对象的变量。

**说明**

使用此方法后，请使用 **Selection** 对象来处理所选行。有关详细信息，请参阅 处理 Selection 对象。

**示例**

``` JavaScript
/*以下示例选择 Report.doc 的第一张表格中的第一行。*/
Documents.Item("Report.doc").Tables.Item(1).Rows.Item(1).Select()
```

### Row.SetHeight

设置表格行的高度。

**语法**

**express.SetHeight ( *RowHeight* , *HeightRule* )**

\*express是一个代表 Row 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型        | 说明                         |
|--------------|-----------|-----------------|------------------------------|
| *RowHeight*  | 必选      | Single          | 行的高度，以磅为单位。       |
| *HeightRule* | 必选      | WdRowHeightRule | 用于确定指定行的高度的规则。 |

**示例**

``` JavaScript
/*以下示例创建一张表格，然后将首行的固定行高设置为 0.5 英寸（36 磅）。*/
function test()
{
let newDoc = Documents.Add()
let aTable = newDoc.Tables.Add(Selection.Range, 3, 3)
aTable.Rows.Item(1).SetHeight(InchesToPoints(0.5), wdRowHeightExactly)
}
```

### Row.SetLeftIndent

设置表格中行的缩进。

**语法**

**express.SetLeftIndent ( *LeftIndent* , *RulerStyle* )**

\*express是一个代表 Row 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型     | 说明                                                               |
|--------------|-----------|--------------|--------------------------------------------------------------------|
| *LeftIndent* | 必选      | Single       | 指定的一行或多行的当前左边缘与所需左边缘之间的距离（以磅为单位）。 |
| *RulerStyle* | 必选      | WdRulerStyle | 改变左缩进时，用来控制 WPS 调整表格的方式。                        |

**说明**

上述 **WdRulerStyle** 行为适用于左对齐的表格。居中和右对齐的表格的 **WdRulerStyle** 行为不可预测；在这些情况下，请谨慎使用 **SetLeftIndent** 方法。

**示例**

``` JavaScript
/*以下示例在新文档中创建一张表格，第一行缩进 0.5 英寸（36 磅）。改变左缩进时，单元格宽度将调整以保持表格的右边缘。*/
function test()
{
let docNew = Documents.Add()
let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)

tableNew.Rows.Item(1).SetLeftIndent(InchesToPoints(0.5),
    wdAdjustSameWidth)
}
```

## 成员属性

### Row.Alignment

返回或设置一个 **WdRowAlignment** 常量，该常量代表指定行的对齐方式。可读写。

**语法**

**express.Alignment**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*本示例使活动文档的第一个表格中第一行的所有单元格居中。*/
function CenterRows() {
    ActiveDocument.Tables.Item(1).Rows.Item(1).Alignment = wdAlignRowCenter
}
```

### Row.AllowBreakAcrossPage

如果允许分页符拆分表格中一行或多行中的文本，则该属性值为 **True** 。可读写 **Long** 类型。

**语法**

**express.AllowBreakAcrossPage**

\*express是一个代表 Row 对象的变量。

**说明**

该属性可以是 **True** 、 **False** 或 **wdUndefined** （仅允许拆分部分指定文本）。

**示例**

``` JavaScript
/*以下示例新建一篇包含 5x5 表格的文档，并防止在分页时拆分表格的第三行。*/
function test()
{
let docNew = Documents.Add()
let tableNew = docNew.Tables.Add(Selection.Range, 5, 5)
tableNew.Rows.Item(3).AllowBreakAcrossPages = false
}
```

``` JavaScript
/*本示例确定分页时是否可以拆分当前表格中的行。如果插入点不在表格中，则显示一个消息框。*/
function test()
{
Selection.Collapse(wdCollapseStart)
if(!Selection.Tables.Count) {
    MsgBox("The insertion point is not in a table.")
}
else {
    let lngAllowBreak = Selection.Rows.AllowBreakAcrossPages
}
}
```

### Row.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Row 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Row.Borders

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

\*express是一个代表 Row 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

### Row.Cells

返回一个 **Cells** 集合，该集合代表在某一列、行、选定内容或区域中的表格单元格。只读。

**语法**

**express.Cells**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
//返回对应的选区的第一行的单元格数
function test()
{
  let tcels = Application.Selection.Rows.Item(1).Cells
  tcels.Count
}
```

### Row.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Row 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Row.HeadingFormat

如果为 **True** ，则将指定一行或数行的格式设置为表格标题。当表格不止一页时，可将多行的格式重复设置为表格标题。可以是 **True** 、 **False** 或 **wdUndefined** 。 **Long** 类型，可读写。

**语法**

**express.HeadingFormat**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*本示例在活动文档的开始处创建一个 5x5 表格，然后将表格第一行的格式设置为表格标题。*/
function test()
{
let rngTemp = Selection.Range
let tableNew = ActiveDocument.Tables.Add(rngTemp, 5, 5)

tableNew.Rows.Item(1).HeadingFormat = true
}
```

``` JavaScript
/*本示例判定是否将插入点所在行的格式设置为表格标题。*/
function test()
{
if(Selection.Information(wdWithInTable)) {
    if(Selection.Rows.Item(1).HeadingFormat) { 
        MsgBox("The current row is a table heading")
    }
}
else {
    MsgBox("The insertion point is not in a table.")
}
}
```

### Row.Height

返回或设置表格中指定行的高度。Single 类型，可读写。

**语法**

**express.Height**

\*express是一个代表 Row 对象的变量。

**说明**

如果指定行的 **HeightRule** 属性为 **wdRowHeightAuto** ，则 **Height** 返回 **wdUndefined** ；设置 **Height** 属性会将 **HeightRule** 设置为 **wdRowHeightAtLeast** 。

**示例**

``` JavaScript
/*本示例将活动文档第一张表格的行高度设置为至少 20 磅。*/
ActiveDocument.Tables.Item(1).Rows.Height = 20
```

### Row.HeightRule

返回或设置确定指定单元格或行高度的规则。 **WdRowHeightRule** 类型，可读写。

**语法**

**express.HeightRule**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*本示例在新文档中创建一张 3x3 表格，并设置第二行的最小行高为 24 磅。*/
function test()
{
let newDoc = Documents.Add()
let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)
    myTable.Rows.Item(2).Height = 24
    myTable.Rows.Item(2).HeightRule = wdRowHeightAtLeast
}
```

### Row.ID

在文档另存为网页时返回或设置指定表格行的标识标签。可读写 **String** 类型。

**语法**

**express.ID**

\*express是一个代表 Row 对象的变量。

**说明**

可以将标签作为引用其他网页或当前文档中内容的超链接。

### Row.Index

返回 **Long** 值，该值表示项目在集合中的位置。只读。

**语法**

**express.Index**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
//获取当前激活文档的第一个表格的第二行的索引号
Application.ActiveDocument.Tables.Item(1).Rows.Item(2).Index

//获取选中区域的第二行所在的索引号
Application.Selection.Rows.Item(2).Index
```

### Row.IsFirst

如果指定的行是表格中的首行，则该属性值为 **True** 。否则为False

**语法**

**express.IsFirst**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
//定所选内容的首行是否为表格中的首行。
Application.Selection.Rows.Item(1).IsFirst
```

### Row.IsLast

如果指定的行是表格中的首行，则该属性值为 True 。否则为False

**语法**

**express.IsLast**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
Application.Selection.Rows.Item(1).IsLast
```

### Row.LeftIndent

返回或设置一个 **Single** 类型的值，该值代表指定表格行的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*以下示例设置活动文档中第一个表格的第一行的左缩进。*/
ActiveDocument.Tables.Item(1).Rows.Item(1).LeftIndent = InchesToPoints(1)
```

### Row.NestingLevel

返回指定表格行的嵌套层。只读 **Long** 类型。

**语法**

**express.NestingLevel**

\*express是一个代表 Row 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

### Row.Next

返回一个 **Row** 对象，该对象代表表格中行的集合中的下一个表格行。只读。

**语法**

**express.Next**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*如果所选内容位于表格中，则以下示例选择下一个表格行的内容。*/
function test()
{
if(Selection.Information(wdWithInTable)) {
    Selection.Rows.Item(1).Next.Select()
}
}
```

### Row.Parent

返回一个 **Object** 类型值，该值代表指定 **Row** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Row 对象的变量。

### Row.Previous

返回一个 **Row** 对象，该对象代表指定行的前一个表格行。只读。

**语法**

**express.Previous**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*如果所选内容位于表格中，则以下示例选择前一行的内容。*/
function test()
{
if(Selection.Information(wdWithInTable)) {
    Selection.Rows.Item(1).Previous.Select()
}
}
```

### Row.Range

返回一个 **Range** 对象，该对象代表指定表格行内包含的文档部分

**语法**

**express.Range**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
//获取表格 1 中的第一行内容并打印

function test()
{
    let rgText
    if(Application.ActiveDocument.Tables.Count >= 1)
        rgText = Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Range.Text
        alert(rgText)
}
```

### Row.Shading

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*以下示例将横线纹理应用于表格 1 的第一行。*/
function test()
{
if(ActiveDocument.Tables.Count >= 1) {
    let myShading = ActiveDocument.Tables.Item(1).Rows.Item(1).Shading
        myShading.Texture = wdTextureHorizontal
}
}
```

### Row.SpaceBetweenColumns

返回或设置指定的一行或多行中相邻列的文本之间的距离（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.SpaceBetweenColumns**

\*express是一个代表 Row 对象的变量。

**示例**

``` JavaScript
/*本示例在新文档中创建一个 3x3 表格，然后将第一行中的列间距设置为 0.5 英寸。*/
function test()
{
let newDoc = Documents.Add()
let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)
myTable.Rows.Item(1).SpaceBetweenColumns = InchesToPoints(0.5)
}
```

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Row](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# Tables 对象

## 简介

Table 对象的集合，这些对象代表所选内容、范围或文档中的表格。

## 方法

| 序号 | 名称                 | 说明                                                        |
|:----:|:---------------------|:------------------------------------------------------------|
|  1   | [Add](#Tables.Add)   | 返回一个 Table 对象，该对象代表添加到文档中的新的空白表格。 |
|  2   | [Item](#Tables.Item) | 返回集合中的单个 Table 对象。                               |

## 属性

| 序号 | 名称                                 | 说明                                                                         |
|:----:|:-------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [Application](#Tables.Application)   | 返回一个代表 WPS 应用程序的 Application 对象。                               |
|  2   | [Count](#Tables.Count)               | 返回一个 Long 类型的值，该值代表集合中表格的数量。只读。                     |
|  3   | [Creator](#Tables.Creator)           | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 Long 类型。 |
|  4   | [NestingLevel](#Tables.NestingLevel) | 返回指定表格的嵌套层。只读 Long 类型。                                       |
|  5   | [Parent](#Tables.Parent)             | 返回一个 Object 类型值，该值代表指定 Tables 对象的父对象。                   |

## 成员方法

### Tables.Add

返回一个 **Table** 对象，该对象代表添加到文档中的新的空白表格。

**语法**

**express.Add ( *Range* , *NumRows* , *NumColumns* , *DefaultTableBehavior* , *AutoFitBehavior* )**

\*express是一个代表 Tables 对象的变量。

**参数**

| 名称                   | 必选/可选 | 数据类型     | 说明                                                                                                                                                                                                                                      |
|------------------------|-----------|--------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Range*                | 必选      | Range object | 表格出现的区域。如果该区域未折叠，表格将替换该区域。                                                                                                                                                                                      |
| *NumRows*              | 必选      | Long         | 要在表格中包括的行数。                                                                                                                                                                                                                    |
| *NumColumns*           | 必选      | Long         | 要在表格中包括的列数。                                                                                                                                                                                                                    |
| *DefaultTableBehavior* | 可选      | Variant      | 设置一个值来指定 WPS 是否要根据单元格的内容自动调整表格单元格的大小（“自动调整”功能）。可以是下列常量之一： wdWord8TableBehavior （禁用“自动调整”功能）或 wdWord9TableBehavior （启用“自动调整”功能）。默认常量为 wdWord8TableBehavior 。 |
| *AutoFitBehavior*      | 可选      | Variant      | 用于设置 WPS 调整表格大小的“自动调整”规则。可以为一个 WdAutoFitBehavior 常量。                                                                                                                                                            |

**示例**

``` JavaScript
/*本示例在活动文档的开头添加一个 3 行 4 列的空表格*/
let myRange = Application.ActiveDocument.Range(0, 0)
Application.ActiveDocument.Tables.Add(myRange, 3, 4)

/*本示例在活动文档的结尾添加一个 6 行 10 列的空白新表格*/
let myRange = Application.ActiveDocument.Content
myRange.Collapse(wdCollapseEnd)
ActiveDocument.Tables.Add(myRange, 6, 10)

/*本示例向新文档中添加一个 3 行 5 列的表格，然后在表格的每个单元格中插入数据*/
function NewTable(){
    let docNew = Application.Documents.Add()
    let tblNew = docNew.Tables.Add(Selection.Range, 3, 5)
    for(let i = 1; i <= 3; i++){
        for(let j=1; j <=5; j++){
            tblNew.Cell(i, j).Range.InsertAfter("Cell: R" + i + ", C" + j)
        }
    }
    tblNew.Columns.AutoFit()
}
```

### Tables.Item

返回集合中的单个 **Table** 对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 Tables 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                         |
|---------|-----------|----------|--------------------------------------------------------------|
| *Index* | 必选      | Long     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

## 成员属性

### Tables.Application

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

\*express是一个代表 Tables 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

### Tables.Count

返回一个 **Long** 类型的值，该值代表集合中表格的数量。只读。

**语法**

**express.Count**

\*express是一个代表 Tables 对象的变量。

### Tables.Creator

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

\*express是一个代表 Tables 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

### Tables.NestingLevel

返回指定表格的嵌套层。只读 **Long** 类型。

**语法**

**express.NestingLevel**

\*express是一个代表 Tables 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

### Tables.Parent

返回一个 **Object** 类型值，该值代表指定 **Tables** 对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 Tables 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/Tables](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TaskPanes 对象

## 方法

| 序号 | 名称                    | 说明                                   |
|:----:|:------------------------|:---------------------------------------|
|  1   | [Item](#TaskPanes.Item) | 将指定任务窗格作为 TaskPane 对象返回。 |

## 属性

| 序号 | 名称                      | 说明                                                         |
|:----:|:--------------------------|:-------------------------------------------------------------|
|  1   | [Count](#TaskPanes.Count) | 返回一个 Long 类型的值，该值代表集合中任务窗格的数目。只读。 |

## 成员方法

### TaskPanes.Item

将指定任务窗格作为 **TaskPane** 对象返回。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 TaskPanes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型    | 说明             |
|---------|-----------|-------------|------------------|
| *Index* | 必选      | WdTaskPanes | 指定的任务窗格。 |

## 成员属性

### TaskPanes.Count

返回一个 **Long** 类型的值，该值代表集合中任务窗格的数目。只读。

**语法**

**express.Count**

\*express是一个代表 TaskPanes 对象的变量。

> WPS 开发文档： [WPS 基础接口/文字 API 参考/TaskPanes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ColorFormat 对象

## 简介

代表单色对象的颜色、带有渐变或图案填充的对象的前景色或背景色，或者指针的颜色。可以将颜色设为明确的红-绿-蓝值（使用 RGB 属性）或设为配色方案中的一种颜色（使用 SchemeColor 属性）。

## 说明

使用下表中列出的属性之一返回 ColorFormat 对象。

| 使用此属性     | 对此对象          | 返回一个代表以下颜色的 ColorFormat 对象      |
|----------------|-------------------|----------------------------------------------|
| DimColor       | AnimationSettings | 变暗对象使用的颜色                           |
| BackColor      | FillFormat        | 背景填充颜色（在底纹或带图案的填充中使用）   |
| ForeColor      | FillFormat        | 前景填充颜色（对于纯色填充，则代表填充颜色） |
| Color          | Font              | 项目符号或字符颜色                           |
| BackColor      | LineFormat        | 背景线条颜色（在带图案线条中使用）           |
| ForeColor      | LineFormat        | 前景线条颜色（或只是实线的线条颜色）         |
| ForeColor      | ShadowFormat      | 阴影颜色                                     |
| PointerColor   | SlideShowSettings | 演示文稿的默认指针颜色                       |
| PointerColor   | SlideShowView     | 幻灯片放映视图中的临时指针颜色               |
| ExtrusionColor | ThreeDFormat      | 延伸对象的侧面颜色                           |

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 <span>Application</span> 对象，该对象表示指定对象的创建者。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.ObjectThemeColor">ObjectThemeColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定 ColorFormat 对象的主题颜色。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回指定对象的父对象。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.RGB">RGB</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置指定颜色的 红-绿-蓝(RGB)值（RGB 值：RGB 函数的返回值；以红、绿、蓝的数值组合来指定颜色，其数值范围是由 0（零）到 255 的整数。） Long 类型，可读/写 。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.SchemeColor">SchemeColor</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回或设置已经应用的配色方案中的颜色，该配色方案与指定对象相关联。可读/写 PpColorSchemeIndex 类型。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.TintAndShade">TintAndShade</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 Single 类型的值，该值代表使指定形状的颜色是变亮，还是变暗。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#ColorFormat.Type">Type</a></td>
<td style="text-align: left;" data-valian="middle"><p>table { word-break:break-all; }</p>
<p>返回一个 MsoColorType 常量，该常量代表颜色的类型。只读。</p></td>
</tr>
</tbody>
</table>

## 成员属性

### ColorFormat.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 ColorFormat 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### ColorFormat.Creator

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

``` JavaScript
/*以下示例显示关于 myObject 创建者的消息。*/
function test(){
　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
　　　　if(myObject.Creator == parseInt(50575054,16)) {
　　　　    alert("This is a WPP object")
　　　　}
　　　　else {
　　　　    alert("This is not a WPP object")
　　　　}
}
```

### ColorFormat.ObjectThemeColor

返回或设置指定 ColorFormat 对象的主题颜色。可读/写。

**语法**

**express.ObjectThemeColor**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

**ObjectThemeColor** 属性的值可以是下列 常量之一。

**示例**

下例显示了如何使用 **ObjectThemeColor** 属性获取活动演示文稿中第一张幻灯片上第一个形状的前景填充的主题颜色。

``` JavaScript
function ObjectThemeColor_Example() {
    Debug.Print(ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.ForeColor.ObjectThemeColor)  
}
```

### ColorFormat.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 ColorFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test() {
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
    let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
    addShp.TextRange.Text = "Test text"
    addShp.Parent.Rotation = 45
}
```

### ColorFormat.RGB

返回或设置指定颜色的 红-绿-蓝(RGB)值（RGB 值：RGB 函数的返回值；以红、绿、蓝的数值组合来指定颜色，其数值范围是由 0（零）到 255 的整数。） **Long** 类型，可读/写 。

**语法**

**express.RGB**

\*express是一个代表 ColorFormat 对象的变量。

**示例**

``` JavaScript
/*本示例为活动演示文稿的第三个配色方案设定背景色，并将配色方案应用于基于幻灯片母版的演示文稿中所有的幻灯片。*/
function test(){
　　　　let a = Application.ActivePresentation
　　　　let cs1 = a.ColorSchemes.Item(3)
　　　　cs1.Colors(ppBackground).RGB = 128, 128, 0
　　　　a.SlideMaster.ColorScheme = cs1
}
```

### ColorFormat.SchemeColor

返回或设置已经应用的配色方案中的颜色，该配色方案与指定对象相关联。可读/写 **PpColorSchemeIndex** 类型。

**语法**

**express.SchemeColor**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

| PpColorSchemeIndex 可以是下列 PpColorSchemeIndex 常量之一。 |
|-------------------------------------------------------------|
| ppAccent1                                                   |
| ppAccent2                                                   |
| ppAccent3                                                   |
| ppBackground                                                |
| ppFill                                                      |
| ppForeground                                                |
| ppNotSchemeColor                                            |
| ppSchemeColorMixed                                          |
| ppShadow                                                    |
| ppTitle                                                     |

**示例**

``` JavaScript
/*本示例切换活动演示文稿中第一张幻灯片的两种背景色，一种是明确的红-绿-蓝值所定义的颜色，另一种是配色方案的背景色。*/
function test(){
　　　　let slides = Application.ActivePresentation.Slides.Item(1)
　　　　slides.FollowMasterBackground = false
　　　　let color = slides.Background.Fill.ForeColor
    if(color.Type == msoColorTypeScheme) {
　　　　    color.RGB = 0, 128, 128
　　　　}
　　　　else {
　　　　    color.SchemeColor = ppBackground
　　　　}
}
```

### ColorFormat.TintAndShade

返回一个 **Single** 类型的值，该值代表使指定形状的颜色是变亮，还是变暗。可读/写。

**语法**

**express.TintAndShade**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

**TintAndShade** 属性值的输入范围介于 -1（最暗）到 1（最亮）之间，0（零）为适中。

**示例**

本示例在活动文档中新建一个形状，并设置填充颜色，然后使彩色底纹变亮。

``` JavaScript
function PrinterPlate() {
    let shpHeart = Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShapeHeart, 150, 150, 250, 250)
    let color = shpHeart.Fill.ForeColor
    color.CMYK = 16111872
    color.TintAndShade = 0.3
    color.OverPrint = msoTrue
}
```

### ColorFormat.Type

table { word-break:break-all; }

返回一个 **MsoColorType** 常量，该常量代表颜色的类型。只读。

**语法**

**express.Type**

\*express是一个代表 ColorFormat 对象的变量。

**说明**

| MsoColorType 可以是下列 MsoColorType 常量之一。 |
|-------------------------------------------------|
| msoColorTypeCMS                                 |
| msoColorTypeCMYK                                |
| msoColorTypeInk                                 |
| msoColorTypeMixed                               |
| msoColorTypeRGB                                 |
| msoColorTypeScheme                              |

> WPS 开发文档： [WPS 基础接口/演示 API 参考/ColorFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PictureFormat 对象

## 简介

包含应用于图片和 OLE 对象的属性和方法。

## 说明

LinkFormat 对象包含仅应用于链接 OLE 对象的属性和方法。 OLEFormat 对象包含应用于链接和非链接 OLE 对象的属性和方法。

## 方法

| 序号 | 名称                                                      | 说明                                                                     |
|:----:|:----------------------------------------------------------|:-------------------------------------------------------------------------|
|  1   | [IncrementBrightness](#PictureFormat.IncrementBrightness) | 按指定数值更改图片的亮度。可以使用 Brightness 属性设置图片的绝对亮度。   |
|  2   | [IncrementContrast](#PictureFormat.IncrementContrast)     | 按指定数值更改图片的对比度。可以使用 Contrast 属性设置图片的绝对对比度。 |

## 属性

| 序号 | 名称                                                          | 说明                                                                                                                                                                                                                                             |
|:----:|:--------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#PictureFormat.Application)                     | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                          |
|  2   | [Brightness](#PictureFormat.Brightness)                       | 返回或设置指定图片或 OLE 对象的亮度。该属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数。可读/写 Single 类型。                                                                                                                                   |
|  3   | [ColorType](#PictureFormat.ColorType)                         | 返回或设置应用于指定图片或 OLE 对象的颜色转换类型。可读/写 MsoPictureColorType 类型。                                                                                                                                                            |
|  4   | [Contrast](#PictureFormat.Contrast)                           | 返回或设置指定图片或 OLE 对象的对比度。该属性的值必须是 0.0（最低对比度）到 1.0（最高对比度）的数。可读/写 Single 类型。                                                                                                                         |
|  5   | [Creator](#PictureFormat.Creator)                             | 返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。                                                                                    |
|  6   | [CropBottom](#PictureFormat.CropBottom)                       | 返回或设置从指定图片或 OLE 对象的底部裁剪下的磅值。可读/写 Single 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始高度为 100 磅，将其高度调整为 200 磅，再将 CropBottom 属性设置为 50，则会从图片底部裁剪 100 磅（而不是 50 镑）。 |
|  7   | [CropLeft](#PictureFormat.CropLeft)                           | 返回或设置从指定图片或 OLE 对象的左侧裁剪下的磅值。可读/写 Single 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始宽度为 100 磅，将其宽度调整为 200 磅，再将 CropLeft 属性设置为 50，则会从图片左侧裁剪 100 磅（而不是 50 磅）。   |
|  8   | [CropRight](#PictureFormat.CropRight)                         | 返回或设置从指定图片或 OLE 对象的右侧裁剪下的磅值。可读/写 Single 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始宽度为 100 磅，将其宽度调整为 200 磅，再将 CropRight 属性设置为 50，则会从图片右侧裁剪 100 磅（而不是 50 磅）。  |
|  9   | [CropTop](#PictureFormat.CropTop)                             | 返回或设置从指定图片或 OLE 对象的顶部裁剪下的磅值。可读/写 Single 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始高度为 100 磅，将其高度调整为 200 磅，再将 CropTop 属性设置为 50，则会从图片顶部裁剪 100 磅（而不是 50 镑）。    |
|  10  | [Parent](#PictureFormat.Parent)                               | 返回指定对象的父对象。                                                                                                                                                                                                                           |
|  11  | [TransparencyColor](#PictureFormat.TransparencyColor)         | 以红-绿-蓝 (RGB) 值返回或设置指定图片的透明色。要使该属性生效，必须将 TransparentBackground 属性设为 True 。仅适用于位图。可读/写 Long 类型。                                                                                                    |
|  12  | [TransparentBackground](#PictureFormat.TransparentBackground) | 决定图片中定义为透明色的部分是否变为透明。可读/写 MsoTriState 类型。仅适用于位图。                                                                                                                                                               |

## 成员方法

### PictureFormat.IncrementBrightness

按指定数值更改图片的亮度。可以使用 **Brightness** 属性设置图片的绝对亮度。

**语法**

**express.IncrementBrightness ( *Increment* )**

\*express是一个代表 PictureFormat 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                     |
|-------------|-----------|----------|--------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定图片 Brightness 属性值的改变量。正值增加图片亮度；负值降低图片亮度。 |

**说明**

图片的亮度调整不能超出 **Brightness** 属性的上下限。例如，如果将 **Brightness** 属性初始设为 0.9，并指定 **Increment** 参数为 0.3，则最后的亮度不是 1.2，而是 **Brightness** 属性的上限，即 1.0。

**示例**

``` JavaScript
/*本示例创建 myDocument 上第一个形状的副本，再移动该副本并使之变暗。要执行本示例，第一个形状必须是图片或 OLE 对象。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let mypane = myDocument.Shapes.Item(1).Duplicate()
 　　　　   mypane.PictureFormat.IncrementBrightness(-0.2)
  　　　　  mypane.IncrementLeft(50)
 　　　　   mypane.IncrementTop(50)
}
```

### PictureFormat.IncrementContrast

按指定数值更改图片的对比度。可以使用 **Contrast** 属性设置图片的绝对对比度。

**语法**

**express.IncrementContrast ( *Increment* )**

\*express是一个代表 PictureFormat 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                         |
|-------------|-----------|----------|------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 指定图片的 Contrast 属性值的改变量。正值表示对比度增大，负值表示对比度减小。 |

**说明**

图片的对比度调整不能超出 **Contrast** 属性的上下限。例如，如果将 **Contrast** 属性初始设为 0.9，并指定 **Increment** 参数为 0.3，则最后的对比度不是 1.2，而是 **Contrast** 属性的上限，即 1.0。

**示例**

``` JavaScript
/*本示例对 myDocument 上还没有设置为最大对比度的所有图片增加对比度。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　for(let i = 1; i <= myDocument.Shapes.Count; i++){
   　　　　 if(myDocument.Shapes.Item(i).Type == msoPicture){
     　　　　  myDocument.Shapes.Item(i).PictureFormat.IncrementContrast(0.1)
 　　　　   }
　　　　}
}
```

## 成员属性

### PictureFormat.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 PictureFormat 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres){
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### PictureFormat.Brightness

返回或设置指定图片或 OLE 对象的亮度。该属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数。可读/写 **Single** 类型。

**语法**

**express.Brightness**

\*express是一个代表 PictureFormat 对象的变量。

**示例**

``` JavaScript
/*本示例设置了 myDocument 中第一个形状的亮度。第一个形状必须为图片或 OLE 对象。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
   　　　　 myDocument.Shapes.Item(1).PictureFormat.Brightness = 0.3
}
```

### PictureFormat.ColorType

返回或设置应用于指定图片或 OLE 对象的颜色转换类型。可读/写 **MsoPictureColorType** 类型。

**语法**

**express.ColorType**

\*express是一个代表 PictureFormat 对象的变量。

**说明**

| MsoPictureColorType 可以是下列 MsoPictureColorType 常量之一。 |
|---------------------------------------------------------------|
| msoPictureAutomatic                                           |
| msoPictureBlackAndWhite                                       |
| msoPictureGrayscale                                           |
| msoPictureMixed                                               |
| msoPictureWatermark                                           |

**示例**

``` JavaScript
/*本示例将 myDocument 上的第一个形状的颜色变换方式设置为灰度级。第一个形状必须为图片或 OLE 对象。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　myDocument.Shapes.Item(1).PictureFormat.ColorType = msoPictureGrayscale
}
```

### PictureFormat.Contrast

返回或设置指定图片或 OLE 对象的对比度。该属性的值必须是 0.0（最低对比度）到 1.0（最高对比度）的数。可读/写 **Single** 类型。

**语法**

**express.Contrast**

\*express是一个代表 PictureFormat 对象的变量。

**示例**

``` JavaScript
/*本示例设置了 myDocument 中第一个形状的对比度。第一个形状必须为图片或 OLE 对象。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　myDocument.Shapes.Item(1).PictureFormat.Contrast = 0.8
}
```

### PictureFormat.Creator

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

\*express是一个代表 PictureFormat 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

``` JavaScript
/*以下示例显示关于 myObject 创建者的消息。*/
function test(){
　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)

　　　　if(myObject.Creator == parseInt(50575054, 16)){
　　　　   alert("This is aWPP object")
　　　　}
　　　　else{
　　　　   alert("This is not aWPP object")
　　　　}
}
```

### PictureFormat.CropBottom

返回或设置从指定图片或 OLE 对象的底部裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始高度为 100 磅，将其高度调整为 200 磅，再将 **CropBottom** 属性设置为 50，则会从图片底部裁剪 100 磅（而不是 50 镑）。

**语法**

**express.CropBottom**

\*express是一个代表 PictureFormat 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中第三个形状的底部裁剪了 20 磅。要使本示例执行，第三个形状必须是图片或 OLE 对象。*/
function test(){
　　　　let myDocument = ActivePresentation.Slides.Item(1)
　　　　myDocument.Shapes.Item(3).PictureFormat.CropBottom = 20
}
```

``` JavaScript
/*本示例将选定形状的底部裁剪用户指定的百分比，而不考虑该形状是否调整过。要使本示例执行，选定的形状必须是图片或 OLE 对象。*/
function test(){
　　　　let percentToCrop = Application.prompt("What percentage do you want to crop off the bottom of this picture?")
　　　　let shapeToCrop = Application.ActiveWindow.Selection.ShapeRange.Item(1)

　　　　let myPictureFormat = shapeToCrop.Duplicate()
　　　　    myPictureFormat.ScaleHeight(1, true)

　　　　let origHeight = myPictureFormat.Height
　　　　myPictureFormat.Delete()

　　　　let cropPoints = origHeight * percentToCrop / 100
　　　　shapeToCrop.PictureFormat.CropBottom = cropPoints
}
```

### PictureFormat.CropLeft

返回或设置从指定图片或 OLE 对象的左侧裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始宽度为 100 磅，将其宽度调整为 200 磅，再将 **CropLeft** 属性设置为 50，则会从图片左侧裁剪 100 磅（而不是 50 磅）。

**语法**

**express.CropLeft**

\*express是一个代表 PictureFormat 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中第三个形状的左侧裁剪了 20 磅。要使本示例执行，第三个形状必须是图片或 OLE 对象。*/
function test(){
　　　　let myDocument = ActivePresentation.Slides.Item(1)
　　　　myDocument.Shapes.Item(3).PictureFormat.CropLeft = 20
}
```

### PictureFormat.CropRight

返回或设置从指定图片或 OLE 对象的右侧裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始宽度为 100 磅，将其宽度调整为 200 磅，再将 **CropRight** 属性设置为 50，则会从图片右侧裁剪 100 磅（而不是 50 磅）。

**语法**

**express.CropRight**

\*express是一个代表 PictureFormat 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中第三个形状的右侧裁剪了 20 磅。要使本示例执行，第三个形状必须是图片或 OLE 对象。*/
function test(){
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　myDocument.Shapes.Item(3).PictureFormat.CropRight = 20
 }
```

### PictureFormat.CropTop

返回或设置从指定图片或 OLE 对象的顶部裁剪下的磅值。可读/写 **Single** 类型。裁剪以相对于图片的原始尺寸计算。例如，如果插入的图片原始高度为 100 磅，将其高度调整为 200 磅，再将 **CropTop** 属性设置为 50，则会从图片顶部裁剪 100 磅（而不是 50 镑）。

**语法**

**express.CropTop**

\*express是一个代表 PictureFormat 对象的变量。

**示例**

``` JavaScript
/*本示例在 myDocument 中第三个形状的顶部裁剪了 20 磅。要使本示例执行，第三个形状必须是图片或 OLE 对象。*/
function test(){
let myDocument = Application.ActivePresentation.Slides.Item(1)
myDocument.Shapes.Item(3).PictureFormat.CropTop = 20
}
```

### PictureFormat.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 PictureFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
　　　　let myPicture = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
   　　　　 myPicture.TextRange.Text = "Test text"
    　　　　myPicture.Parent.Rotation = 45
}
```

### PictureFormat.TransparencyColor

以红-绿-蓝 (RGB) 值返回或设置指定图片的透明色。要使该属性生效，必须将 **TransparentBackground** 属性设为 **True** 。仅适用于位图。可读/写 **Long** 类型。

**语法**

**express.TransparencyColor**

\*express是一个代表 PictureFormat 对象的变量。

**说明**

如果想透过图片的透明部分始终看到图片后面的对象，必须将图片的 **FillFormat** 对象的 **Visible** 属性设为 **False** 。如果图片有透明色且其 **FillFormat** 对象的 **Visible** 属性被设为 **True** ，则可以透过透明色看到图片的填充，但图片后面的对象将难以看到。

**示例**

``` JavaScript
/*本示例用函数 RGB(0, 0, 255) 返回 RGB 值，并将该 RGB 值对应的颜色设置为 myDocument 上第一个形状的透明色。要使本示例正常运行，第一个形状必须是位图。*/
function test(){
　　　　let blueScreen = (0, 0, 255)
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let myPicture = myDocument.Shapes.Item(1).PictureFormat
 　　　　   myPicture.TransparentBackground = true
 　　　　   myPicture.TransparencyColor = blueScreen
　　　　myDocument.Shapes.Item(1).Fill.Visible = false
}
```

### PictureFormat.TransparentBackground

决定图片中定义为透明色的部分是否变为透明。可读/写 **MsoTriState** 类型。仅适用于位图。

**语法**

**express.TransparentBackground**

\*express是一个代表 PictureFormat 对象的变量。

**说明**

使用 **TransparencyColor** 属性可设置透明色。

如果想透过图片的透明部分始终看到图片后面的对象，必须将图片的 **FillFormat** 对象的 **Visible** 属性设为 **msoFalse** 。如果图片有透明色且其 **FillFormat** 对象的 **Visible** 属性被设为 **msoTrue** ，则图片的填充可以透过透明色看到，但图片后面的对象将难以看到。

| MsoTriState 可以是下列 MsoTriState 常量之一。  |
|------------------------------------------------|
| msoCTrue                                       |
| msoFalse                                       |
| msoTriStateMixed                               |
| msoTriStateToggle                              |
| msoTrue 图片中颜色定义为透明色的部分变为透明。 |

**示例**

``` JavaScript
/*本示例用函数 RGB(0, 24, 240) 返回 RGB 值，并将该 RGB 值对应的颜色设置为 myDocument 上第一个形状的透明色。要使本示例正常运行，第一个形状必须是位图。*/
function test(){
　　　　let blueScreen = (0, 0, 255)
　　　　let myDocument = Application.ActivePresentation.Slides.Item(1)
　　　　let myPicture = myDocument.Shapes.Item(1).PictureFormat
 　　　　   myPicture.TransparentBackground = true
 　　　　   myPicture.TransparencyColor = blueScreen
　　　　myDocument.Shapes.Item(1).Fill.Visible = false
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/PictureFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# PlaceholderFormat 对象

## 简介

包含专门应用于占位符的属性，例如占位符类型

## 属性

| 序号 | 名称                                          | 说明                                                            |
|:----:|:----------------------------------------------|:----------------------------------------------------------------|
|  1   | [Application](#PlaceholderFormat.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。         |
|  2   | [Name](#PlaceholderFormat.Name)               | 返回或设置指定对象的名称。可读/写。                             |
|  3   | [Parent](#PlaceholderFormat.Parent)           | 返回指定对象的父对象。                                          |
|  4   | [Type](#PlaceholderFormat.Type)               | 返回一个 PpPlaceholderType 常量，该常量代表占位符的类型。只读。 |

## 成员属性

### PlaceholderFormat.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 PlaceholderFormat 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### PlaceholderFormat.Name

返回或设置指定对象的名称。可读/写。

**语法**

**express.Name**

\*express是一个代表 PlaceholderFormat 对象的变量。

### PlaceholderFormat.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 PlaceholderFormat 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
　　　　let myPicture = myShapes.AddShape(msoShapeOval,50,50,300,150).TextFrame
   　　　　 myPicture.TextRange.Text = "Test text"
   　　　　 myPicture.Parent.Rotation = 45
}
```

### PlaceholderFormat.Type

返回一个 **PpPlaceholderType** 常量，该常量代表占位符的类型。只读。

**语法**

**express.Type**

\*express是一个代表 PlaceholderFormat 对象的变量。

**说明**

| PpPlaceholderType 可以是下列 PpPlaceholderType 常量之一。 |
|-----------------------------------------------------------|
| ppPlaceholderBitmap                                       |
| ppPlaceholderBody                                         |
| ppPlaceholderCenterTitle                                  |
| ppPlaceholderChart                                        |
| ppPlaceholderDate                                         |
| ppPlaceholderFooter                                       |
| ppPlaceholderHeader                                       |
| ppPlaceholderMediaClip                                    |
| ppPlaceholderMixed                                        |
| ppPlaceholderObject                                       |
| ppPlaceholderOrgChart                                     |
| ppPlaceholderSlideNumber                                  |
| ppPlaceholderSubtitle                                     |
| ppPlaceholderTable                                        |
| ppPlaceholderTitle                                        |
| ppPlaceholderVerticalBody                                 |
| ppPlaceholderVerticalTitle                                |

> WPS 开发文档： [WPS 基础接口/演示 API 参考/PlaceholderFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ShapeNodes 对象

## 简介

指定的任意多边形中所有 ShapeNode 对象的集合。

## 说明

每个 ShapeNode 对象代表任意多边形中线段之间的一个节点或任意多边形曲线段的一个控制点。可以手动或使用 BuildFreeform 和 ConvertToShape 方法来创建任意多边形。

## 方法

| 序号 | 名称                                         | 说明                                                                                                                                                                            |
|:----:|:---------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Delete](#ShapeNodes.Delete)                 | 删除形状节点。                                                                                                                                                                  |
|  2   | [Insert](#ShapeNodes.Insert)                 | 将一个新线段插入到任意多边形的指定节点之后。                                                                                                                                    |
|  3   | [Item](#ShapeNodes.Item)                     | 从指定集合中返回单个对象。                                                                                                                                                      |
|  4   | [SetEditingType](#ShapeNodes.SetEditingType) | 设置由 Index 指定的节点的编辑类型。如果节点是曲线段的控制点，则该方法将设置与之相邻的连接两条线段的节点的编辑类型。请注意，由于编辑类型的不同，该方法可能会影响相邻节点的位置。 |
|  5   | [SetPosition](#ShapeNodes.SetPosition)       | 设置由 Index 指定的节点的位置。请注意，由于节点的编辑类型的不同，该方法可能会影响相邻节点的位置。                                                                               |
|  6   | [SetSegmentType](#ShapeNodes.SetSegmentType) | 设置由 Index 指定的节点后面的线段的段类型。如果节点是曲线段的控制点，该方法将设置该曲线的段类型。请注意，该方法可能会因插入或删除相邻的节点而影响节点总数。                     |

## 属性

| 序号 | 名称                                   | 说明                                                                                                                                                          |
|:----:|:---------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#ShapeNodes.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                       |
|  2   | [Count](#ShapeNodes.Count)             | 返回指定集合中的对象数目。                                                                                                                                    |
|  3   | [Creator](#ShapeNodes.Creator)         | 返回 Long 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。 |
|  4   | [Parent](#ShapeNodes.Parent)           | 返回指定对象的父对象。                                                                                                                                        |

## 成员方法

### ShapeNodes.Delete

删除形状节点。

**语法**

**express.Delete ( *Index* )**

\*express是一个代表 ShapeNodes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                                         |
|---------|-----------|----------|--------------------------------------------------------------------------------------------------------------|
| *Index* | 必选      | Long     | 指定要删除的节点。该节点后的段也将被删除。如果该节点是一条曲线的一个控制点，则该曲线及其所有节点都将被删除。 |

### ShapeNodes.Insert

将一个新线段插入到任意多边形的指定节点之后。

**语法**

**express.Insert ( *Index* , *SegmentType* , *EditingType* , *X1* , *Y1* , *X2* , *Y2* , *X3* , *Y3* )**

\*express是一个代表 ShapeNodes 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型       | 说明                                                                                                                                                                                                                         |
|---------------|-----------|----------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Index*       | 必选      | Long           | 要在其后插入新节点的节点。                                                                                                                                                                                                   |
| *SegmentType* | 必选      | MsoSegmentType | 要添加的线段的类型。                                                                                                                                                                                                         |
| *EditingType* | 必选      | MsoEditingType | 顶点的编辑属性。                                                                                                                                                                                                             |
| *X1*          | 必选      | Single         | 如果新线段的 EditingType 为 msoEditingAuto，则该参数指定从文档左上角到新线段终点的水平距离（以磅为单位）。如果新节点的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段第一个控点的水平距离（以磅为单位）。 |
| *Y1*          | 必选      | Single         | 如果新线段的 EditingType 为 msoEditingAuto，则该参数指定从文档左上角到新线段终点的垂直距离（以磅为单位）。如果新节点的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段第一个控点的垂直距离（以磅为单位）。 |
| *X2*          | 必选      | Single         | 如果新线段的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段的第二个控制点的水平距离（以磅为单位）。如果新线段的 EditingType 为 msoEditingAuto，请不要指定该参数的值。                                     |
| *Y2*          | 必选      | Single         | 如果新线段的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段的第二个控制点的垂直距离（以磅为单位）。如果新线段的 EditingType 为 msoEditingAuto，请不要指定该参数的值。                                     |
| *X3*          | 必选      | Single         | 如果新线段的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段终点的水平距离（以磅为单位）。如果新线段的 EditingType 为 msoEditingAuto，请不要指定该参数的值。                                               |
| *Y3*          | 必选      | Single         | 如果新线段的 EditingType 为 msoEditingCorner，则该参数指定从文档左上角到新线段终点的垂直距离（以磅为单位）。如果新线段的 EditingType 为 msoEditingAuto，请不要指定该参数的值。                                               |

**说明**

| MsoSegmentType 可以是下列 MsoSegmentType 常量之一。 |
|-----------------------------------------------------|
| msoSegmentCurve                                     |
| msoSegmentLine                                      |

| MsoEditingType 可以是下列 MsoEditingType 常量之一（不能是 msoEditingSmooth 或 msoEditingSymmetric）。 |
|-------------------------------------------------------------------------------------------------------|
| msoEditingAuto                                                                                        |
| msoEditingCorner                                                                                      |

**示例**

``` JavaScript
/*本示例在 myDocument 中第三个形状的第四个结点后添加一个带有一段曲线的平滑结点。第三个形状必须是至少有四个结点的任意多边形。*/
function test(){
　　　　let myDocument = ActivePresentation.Slides.Item(1)
　　　　let mynodes = myDocument.Shapes.Item(3).Nodes
　　　　mynodes.Insert (4, msoSegmentCurve, msoEditingSmooth, 210, 100)
}
```

### ShapeNodes.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 ShapeNodes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

**返回值**

ShapeNode

### ShapeNodes.SetEditingType

设置由 **Index** 指定的节点的编辑类型。如果节点是曲线段的控制点，则该方法将设置与之相邻的连接两条线段的节点的编辑类型。请注意，由于编辑类型的不同，该方法可能会影响相邻节点的位置。

**语法**

**express.SetEditingType ( *Index* , *EditingType* )**

\*express是一个代表 ShapeNodes 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型       | 说明                     |
|---------------|-----------|----------------|--------------------------|
| *Index*       | 必选      | Long           | 需要设置编辑类型的节点。 |
| *EditingType* | 必选      | MsoEditingType | 编辑类型。               |

**说明**

| MsoEditingType 可以是下列 MsoEditingType 常量之一。 |
|-----------------------------------------------------|
| msoEditingAuto                                      |
| msoEditingCorner                                    |
| msoEditingSmooth                                    |
| msoEditingSymmetric                                 |

**示例**

``` JavaScript
/*本示例将 myDocument 中第三个形状中的所有角结点更改为光滑结点。第三个形状必须是任意多边形。*/
function test(){
　　　　let myDocument = ActivePresentation.Slides.Item(1)
　　　　let mynodes = myDocument.Shapes.Item(3).Nodes
       for(let n = 1; n <= mynodes.Count; n++){
 　　　　   if(mynodes.Item(n).EditingType == msoEditingCorner){
     　　　　   mynodes.Item(n).SetEditingType (n, msoEditingSmooth)
  　　　　  }
　　　　}
}
```

### ShapeNodes.SetPosition

设置由 **Index** 指定的节点的位置。请注意，由于节点的编辑类型的不同，该方法可能会影响相邻节点的位置。

**语法**

**express.SetPosition ( *Index* , *X1, Y1* )**

\*express是一个代表 ShapeNodes 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                   |
|----------|-----------|----------|----------------------------------------|
| *Index*  | 必选      | Long     | 要设置位置的顶点。                     |
| *X1, Y1* | 必选      | Single   | 新顶点相对于文档左上角的位置（点数）。 |

**示例**

``` JavaScript
/*本示例将 myDocument 中第三个形状的第二个结点向右移 200 磅、向下移 300 磅。第三个形状必须是任意多边形。*/
function test(){
　　　　let myDocument = ActivePresentation.Slides.Item(1)
　　　　let mynodes = myDocument.Shapes.Item(3).Nodes
　　　　let pointsArray = mynodes.Item(2).Points
　　　　let currXvalue = pointsArray(1, 1)
　　　　let currYvalue = pointsArray(1, 2)
　　　　mynodes.SetPosition (2, currXvalue + 200, currYvalue + 300)
}
```

### ShapeNodes.SetSegmentType

设置由 **Index** 指定的节点后面的线段的段类型。如果节点是曲线段的控制点，该方法将设置该曲线的段类型。请注意，该方法可能会因插入或删除相邻的节点而影响节点总数。

**语法**

**express.SetSegmentType ( *Index* , *SegmentType* )**

\*express是一个代表 ShapeNodes 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型       | 说明                     |
|---------------|-----------|----------------|--------------------------|
| *Index*       | 必选      | Long           | 需要设置线段类型的顶点。 |
| *SegmentType* | 必选      | MsoSegmentType | 指明线段是直线还是曲线。 |

**说明**

| MsoSegmentType 可以是下列 MsoSegmentType 常量之一。 |
|-----------------------------------------------------|
| msoSegmentCurve                                     |
| msoSegmentLine                                      |

**示例**

``` JavaScript
/*本示例将 myDocument 中第三个形状的所有直线段更改为曲线段。第三个形状必须是任意多边形。*/
function test(){
　　　　let myDocument = ActivePresentation.Slides.Item(1)
　　　　let mynodes = myDocument.Shapes.Item(3).Nodes
　　　　let n = 1
　　　　while (n <= mynodes.Count){
   　　　　 if(mynodes.Item(n).SegmentType == msoSegmentLine){
 　　　　       mynodes.SetSegmentType (n, msoSegmentCurve)
　　　　    }
　　　　    n = n + 1
　　　　}
}
```

## 成员属性

### ShapeNodes.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 ShapeNodes 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres){
    pptPres.Slides.Add (1, 1)
    pptPres.SaveAs (pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test(){
　　　　let shpOle = ActivePresentation.Slides.Item(1).Shapes
       for(let i = 1; i <= shpOle.Count; i++){
   　　　　 if(shpOle.Item(i).Type == msoLinkedOLEObject){
     　　　　   MsgBox (shpOle.Item(i).OLEFormat.Application.Name)
　　　　    }
　　　　}
}
```

### ShapeNodes.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 ShapeNodes 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口。*/
function test(){
　　　　let wins = Application.Windows
    for(let i = 2; i <= wins.Count; i++){
    　　　　wins.Item(2).Close()
　　　　}
}
```

### ShapeNodes.Creator

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

\*express是一个代表 ShapeNodes 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

``` JavaScript
/*以下示例显示关于 myObject 创建者的消息。*/
function test(){
　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
　　　　if(myObject.Creator == "0x50575054"){
　　　　    MsgBox ("This is a WPP object")
　　　　}
　　　　else{
　　　　    MsgBox ("This is not a WPP object")
　　　　}
}
```

### ShapeNodes.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 ShapeNodes 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes
　　　　let myshape = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame
　　　　myshape.TextRange.Text = "Test text"
　　　　myshape.Parent.Rotation = 45}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/ShapeNodes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# ShapeRange 对象

## 简介

代表一个形状范围，即某个文档中的一组形状。形状范围可以只包含文档中的一个形状，也可以包含该文档中的所有形状。

## 说明

您可以使用任何所需的形状（从文档的所有形状中进行选择，或从选定内容的所有形状中进行选择）构造形状范围。例如，可以构造一个 ShapeRange 集合，其中包含文档中的前三个形状、文档中所有选定的形状，或文档中所有的任意多边形。

有关如何使用一个形状或同时使用多个形状的概述，请参阅 使用形状（图形对象）。

以下示例说明如何执行下列操作：

- 返回一组通过名称或索引序号指定的形状。

<!-- -->

- 返回文档中全部或部分选定的形状。

## 方法

| 序号 | 名称                                               | 说明                                                                                                                                                             |
|:----:|:---------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Align](#ShapeRange.Align)                         | 对齐指定形状区域中的形状。                                                                                                                                       |
|  2   | [Apply](#ShapeRange.Apply)                         | 适用于指定的形状格式，该格式已使用 PickUp 方法复制。                                                                                                             |
|  3   | [Copy](#ShapeRange.Copy)                           | 将指定对象复制到剪贴板。                                                                                                                                         |
|  4   | [Cut](#ShapeRange.Cut)                             | 删除指定对象并将其放到剪贴板中。                                                                                                                                 |
|  5   | [Delete](#ShapeRange.Delete)                       | 删除指定的对象。                                                                                                                                                 |
|  6   | [Duplicate](#ShapeRange.Duplicate)                 | 创建指定的 ShapeRange 对象的副本，将新的形状或形状范围添加到 Shapes 集合中，然后返回新的 ShapeRange 对象。副本对象位于 Shapes 集合末尾。                         |
|  7   | [Flip](#ShapeRange.Flip)                           | 绕水平或垂直轴翻转指定形状。                                                                                                                                     |
|  8   | [Group](#ShapeRange.Group)                         | 将指定区域中的形状形成一组。以单个 Shape 对象返回分组后的形状。                                                                                                  |
|  9   | [IncrementLeft](#ShapeRange.IncrementLeft)         | 以指定点数水平移动指定形状。                                                                                                                                     |
|  10  | [IncrementRotation](#ShapeRange.IncrementRotation) | 按指定度数改变指定形状绕 z 轴的旋转量。使用 Rotation 属性设置形状的绝对旋转量。                                                                                  |
|  11  | [IncrementTop](#ShapeRange.IncrementTop)           | 以指定点数垂直移动指定形状。                                                                                                                                     |
|  12  | [Item](#ShapeRange.Item)                           | 从指定集合中返回单个对象。                                                                                                                                       |
|  13  | [PickUp](#ShapeRange.PickUp)                       | 复制指定形状的格式。用 Apply 方法可将复制的格式应用于其他形状。                                                                                                  |
|  14  | [ScaleHeight](#ShapeRange.ScaleHeight)             | 以指定的比例缩放图形高度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于除图片和 OLE 对象以外的其他图形来说，缩放总是相对于当前高度而言。 |
|  15  | [ScaleWidth](#ShapeRange.ScaleWidth)               | 根据指定的系数缩放图形宽度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于图片和 OLE 对象以外的其他图形来说，总是相对于当前宽度进行缩放。 |
|  16  | [Select](#ShapeRange.Select)                       | 选择指定的对象。                                                                                                                                                 |
|  17  | [Ungroup](#ShapeRange.Ungroup)                     | 取消指定形状或者形状区域中任意组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。取消组合后的形状以单个 ShapeRange 对象的形式返回。                  |
|  18  | [ZOrder](#ShapeRange.ZOrder)                       | 将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-顺序中的位置）。                                                                                    |

## 属性

| 序号 | 名称                                               | 说明                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
|:----:|:---------------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ActionSettings](#ShapeRange.ActionSettings)       | 返回一个 ActionSettings 对象，该对象包含在幻灯片放映期间，当用户在指定形状或文本区域内单击或移动鼠标时所产生的动作的信息。只读。                                                                                                                                                                                                                                                                                                                      |
|  2   | [Application](#ShapeRange.Application)             | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                                                                                                                                                                                                                               |
|  3   | [Chart](#ShapeRange.Chart)                         | 返回当前 Shape 对象的 Chart 对象。只读。                                                                                                                                                                                                                                                                                                                                                                                                              |
|  4   | [Count](#ShapeRange.Count)                         | 返回指定集合中的对象数目。                                                                                                                                                                                                                                                                                                                                                                                                                            |
|  5   | [Fill](#ShapeRange.Fill)                           | 返回一个 FillFormat 对象，该对象包含指定形状的填充格式属性。只读。                                                                                                                                                                                                                                                                                                                                                                                    |
|  6   | [Glow](#ShapeRange.Glow)                           | 返回指定形状范围的发光格式。只读 msoGlowType 类型。                                                                                                                                                                                                                                                                                                                                                                                                   |
|  7   | [GroupItems](#ShapeRange.GroupItems)               | 返回一个 GroupShapes 对象，该对象代表指定形状组中的单个形状。使用 GroupShapes 对象的 Item 方法可返回形状组中的单个形状。适用于代表组合形状的 Shape 或 ShapeRange 对象。只读。                                                                                                                                                                                                                                                                         |
|  8   | [HasChart](#ShapeRange.HasChart)                   | 返回指定的对象代表的形状区域是否包含图表。只读。                                                                                                                                                                                                                                                                                                                                                                                                      |
|  9   | [HasTable](#ShapeRange.HasTable)                   | 返回指定的形状是否为表。只读。MsoTriState 类型。                                                                                                                                                                                                                                                                                                                                                                                                      |
|  10  | [HasTextFrame](#ShapeRange.HasTextFrame)           | 返回指定形状是否有文本框。只读。MsoTriState 类型。                                                                                                                                                                                                                                                                                                                                                                                                    |
|  11  | [Height](#ShapeRange.Height)                       | 以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。                                                                                                                                                                                                                                                                                                                                     |
|  12  | [Id](#ShapeRange.Id)                               | 返回一个 Long 类型值，该值标识形状或形状范围。只读。                                                                                                                                                                                                                                                                                                                                                                                                  |
|  13  | [Left](#ShapeRange.Left)                           | 返回或设置一个 Single 类型值，该值代表从形状范围中最左侧形状的左边缘到幻灯片左边缘的距离（以磅为单位）。可读/写。                                                                                                                                                                                                                                                                                                                                     |
|  14  | [Line](#ShapeRange.Line)                           | 返回一个 LineFormat 对象，该对象包含指定形状的线条格式属性。（对于线条来说，LineFormat 对象代表线条本身；而对于带有边框的形状来说，LineFormat 对象代表边框。）只读。                                                                                                                                                                                                                                                                                  |
|  15  | [MediaType](#ShapeRange.MediaType)                 | 返回 OLE 媒体类型。只读。PpMediaType 类型。                                                                                                                                                                                                                                                                                                                                                                                                           |
|  16  | [Name](#ShapeRange.Name)                           | 创建形状时，WPP 自动以 ShapeType Number 的形式为其分配一个名称，其中 ShapeType 指明形状或自选图形的类型，Number 是在幻灯片上形状的集合中具有唯一性的整数。例如，为幻灯片上的形状自动生成的名称可以为 Placeholder 1、Oval 2 和 Rectangle 3。为了避免与自动分配的名称发生冲突，对用户定义的名称请不要使用 ShapeType Number 形式，其中 ShapeType 为一个用于自动生成名称的值，Number 为任意的正整数。形状范围必须至少包含一个形状。String 类型，可读/写。 |
|  17  | [PictureFormat](#ShapeRange.PictureFormat)         | 返回一个 PictureFormat 对象，该对象包含指定形状的图片格式属性。适用于代表图片或 OLE 对象的 Shape 或 ShapeRange 对象。只读。                                                                                                                                                                                                                                                                                                                           |
|  18  | [PlaceholderFormat](#ShapeRange.PlaceholderFormat) | 返回一个 PlaceholderFormat 对象，该对象包含占位符特有的属性。只读。                                                                                                                                                                                                                                                                                                                                                                                   |
|  19  | [Shadow](#ShapeRange.Shadow)                       | 返回一个只读的 ShadowFormat 对象，该对象包含指定形状的阴影格式属性。                                                                                                                                                                                                                                                                                                                                                                                  |
|  20  | [ShapeStyle](#ShapeRange.ShapeStyle)               | 设置或返回指定对象的形状样式索引                                                                                                                                                                                                                                                                                                                                                                                                                      |
|  21  | [Table](#ShapeRange.Table)                         | 返回一个 Table 对象，该对象代表某个形状或形状区域内的一个表格。只读。                                                                                                                                                                                                                                                                                                                                                                                 |
|  22  | [TextEffect](#ShapeRange.TextEffect)               | 返回一个 TextEffectFormat 对象，该对象包含指定形状的文本效果格式属性。适用于代表艺术字的 Shape 或 ShapeRange 对象。                                                                                                                                                                                                                                                                                                                                   |
|  23  | [TextFrame](#ShapeRange.TextFrame)                 | 返回一个 TextFrame 对象，该对象包含指定形状或母版文本样式的对齐方式和定位属性。                                                                                                                                                                                                                                                                                                                                                                       |
|  24  | [TextFrame2](#ShapeRange.TextFrame2)               | 返回与指定的 ShapeRange 对象相关联的 TextFrame2 对象，前者包含指定的形状范围的对齐和定位属性。只读。                                                                                                                                                                                                                                                                                                                                                  |
|  25  | [ThreeD](#ShapeRange.ThreeD)                       | 返回一个 ThreeDFormat 对象，该对象包含指定形状的三维效果格式属性。只读。                                                                                                                                                                                                                                                                                                                                                                              |
|  26  | [Top](#ShapeRange.Top)                             | 返回或设置一个 Single 类型值，该值代表形状范围内最上方形状的上边缘到文档上边缘的距离。可读/写。                                                                                                                                                                                                                                                                                                                                                       |
|  27  | [Type](#ShapeRange.Type)                           | 返回一个 MsoShapeType 常量，该常量代表单个形状或形状范围内多个形状的类型。只读。                                                                                                                                                                                                                                                                                                                                                                      |
|  28  | [Visible](#ShapeRange.Visible)                     | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 MsoTriState 类型。                                                                                                                                                                                                                                                                                                                                                                        |
|  29  | [Width](#ShapeRange.Width)                         | 以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。                                                                                                                                                                                                                                                                                                                                     |
|  30  | [ZOrderPosition](#ShapeRange.ZOrderPosition)       | 返回指定的形状在 z-顺序中的位置。只读。Long 类型。                                                                                                                                                                                                                                                                                                                                                                                                    |

## 成员方法

### ShapeRange.Align

对齐指定形状区域中的形状。

**语法**

**express.Align ( *AlignCmd* , *RelativeTo* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型    | 说明                                   |
|--------------|-----------|-------------|----------------------------------------|
| *AlignCmd*   | 必选      | MsoAlignCmd | 指定某个指定形状范围中形状的对齐方式。 |
| *RelativeTo* | 必选      | MsoTriState | 决定形状是否与幻灯片边缘对齐。         |

**示例**

``` JavaScript
/*本示例将 myDocument 中指定范围的所有形状的左边与该范围内最左边的形状的左边对齐*/
function  test(){
    let myDocument = Application.ActivePresentation.Slides.Item(3)
    myDocument.Shapes.Range().Align(msoAlignLefts, msoFalse)
}
```

### ShapeRange.Apply

适用于指定的形状格式，该格式已使用 PickUp 方法复制。

**语法**

**express.Apply ()**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*本示例复制 myDocument 上第一个形状的格式，然后将该格式应用于第二个形状*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.Item(1).PickUp()
    myDocument.Shapes.Item(2).Apply()
}
```

### ShapeRange.Copy

将指定对象复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

``` JavaScript
/*本示例将当前演示文稿的第一张幻灯片复制到剪贴板。*/
Application.ActivePresentation.Slides.Item(1).Copy()

/*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy()
```

### ShapeRange.Cut

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut ()**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中*/
 Application.Windows.Item(1).Selection.Cut()

/*本示例从当前演示文稿的第一张幻灯片中删除第一个和第二个形状，并将它们的副本放到剪贴板，然后将该副本粘贴到第二张幻灯片*/
function test(){
  ActivePresentation.Slides.Item(1).Shapes.Range([1, 2]).Cut()
  ActivePresentation.Slides.Item(2).Shapes.Paste()
}
```

### ShapeRange.Delete

删除指定的对象。

**语法**

**express.Delete ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**示例**

``` JavaScript
/*本示例删除第一个窗口中选中的对象
Application.Windows.Item(1).Selection.Delete()
```

### ShapeRange.Duplicate

创建指定的 ShapeRange 对象的副本，将新的形状或形状范围添加到 Shapes 集合中，然后返回新的 ShapeRange 对象。副本对象位于 Shapes 集合末尾。

**语法**

**express.Duplicate ()**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Flip

绕水平或垂直轴翻转指定形状。

**语法**

**express.Flip ( *FlipCmd* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型   | 说明                           |
|-----------|-----------|------------|--------------------------------|
| *FlipCmd* | 必选      | MsoFlipCmd | 指定形状是水平翻转还是垂直翻转 |

**说明**

MsoFlipCmd 可以是下列 MsoFlipCmd 常量之一。 msoFlipHorizontal msoFlipVertical

### ShapeRange.Group

将指定区域中的形状形成一组。以单个 Shape 对象返回分组后的形状。

**语法**

**express.Group ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

由于一组形状作为单个形状处理，所以创建和分解形状组将更改 Shapes 集合中的项目数，并更改集合中受影响的项目之后的项目索引号。

**示例**

``` JavaScript
/*本示例将两个形状添加到第一张幻灯片，组合两个新形状，为该组设置填充，旋转该组并将其发送到绘图层后*/
function test(){
    let myShapes =  Application.ActivePresentation.Slides.Item(1).Shapes

    myShapes.AddShape(msoShapeCan, 50, 10, 100, 200).Name = "shpOne"
    myShapes.AddShape(msoShapeCube, 150, 250, 100, 200).Name = "shpTwo"

    let grouped = myShapes.Range(["shpOne", "shpTwo"]).Group()
    grouped.Fill.PresetTextured(msoTextureBlueTissuePaper)
    grouped.Rotation = 45
    grouped.ZOrder(msoSendToBack)
}
```

### ShapeRange.IncrementLeft

以指定点数水平移动指定形状。

**语法**

**express.IncrementLeft ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                               |
|-------------|-----------|----------|------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 以点为单位指定形状水平移动的距离。正值表示将形状向右移动，负值表示将形状向左移动。 |

### ShapeRange.IncrementRotation

按指定度数改变指定形状绕 z 轴的旋转量。使用 Rotation 属性设置形状的绝对旋转量。

**语法**

**express.IncrementRotation ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                             |
|-------------|-----------|----------|--------------------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 以度为单位指定形状的水平旋转量。正值表示使形状按顺时针方向旋转；负值表示使形状按逆时针方向旋转。 |

**说明**

要将三维形状绕 X 轴或 Y 轴旋转，请分别使用 IncrementRotationX 方法或 IncrementRotationY 方法。

**示例**

``` JavaScript
/*本示例复制myDocument上的第一个形状，并设置复制的填充，再向右移70磅、向上移50磅，然后顺时针旋转30度角*/
function test(){
    let myShape = ActivePresentation.Slides.Item(1).Shapes.Item(1).Duplicate()
    myShape.Fill.PresetTextured(msoTextureGranite)
    myShape.IncrementLeft(70)
    myShape.IncrementTop(-50)
    myShape.IncrementRotation(30)
}
```

### ShapeRange.IncrementTop

以指定点数垂直移动指定形状。

**语法**

**express.IncrementTop ( *Increment* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                                   |
|-------------|-----------|----------|----------------------------------------------------------------------------------------|
| *Increment* | 必选      | Single   | 以点为单位指定形状对象的垂直移动距离。正值表示将形状向下移动，负值表示将形状向上移动。 |

**说明**

要将三维形状绕 X 轴或 Y 轴旋转，请分别使用 IncrementRotationX 方法或 IncrementRotationY 方法。

**示例**

``` JavaScript
/*本示例复制 myDocument 上的第一个形状，并设置复制的填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角*/
function test(){
    let myShape = ActivePresentation.Slides.Item(1).Shapes.Item(1).Duplicate()
    myShape.Fill.PresetTextured(msoTextureGranite)
    myShape.IncrementLeft(70)
    myShape.IncrementTop(-50)
    myShape.IncrementRotation(30)
}
```

### ShapeRange.Item

从指定集合中返回单个对象。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                   |
|---------|-----------|----------|----------------------------------------|
| *Index* | 必选      | Variant  | 集合中要返回的单个对象的名称或索引号。 |

### ShapeRange.PickUp

复制指定形状的格式。用 **Apply** 方法可将复制的格式应用于其他形状。

**语法**

**express.PickUp ()**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*本示例复制 myDocument 上第一个形状的格式，然后将该格式应用于第二个形状*/
function test(){
    let myDocument =Application.ActivePresentation.Slides.Item(1)

    myDocument.Shapes.Item(1).PickUp()
    myDocument.Shapes.Item(2).Apply()
}
```

### ShapeRange.ScaleHeight

以指定的比例缩放图形高度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于除图片和 OLE 对象以外的其他图形来说，缩放总是相对于当前高度而言。

**语法**

**express.ScaleHeight ( *Factor* , *RelativeToOriginalSize* , *fScale* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型     | 说明                                                                                                   |
|--------------------------|-----------|--------------|--------------------------------------------------------------------------------------------------------|
| *Factor*                 | 必选      | Single       | 指定形状调整后的高度与当前或原始高度的比例。例如，若要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选      | MsoTriState  | 指定是否相对于形状的当前或原始尺寸来缩放形状。                                                         |
| *fScale*                 | 可选      | MsoScaleFrom | 在缩放形状时，形状中位置不变的部分。                                                                   |

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。 msoCTrue msoFalse：相对于形状的当前尺寸缩放。 msoTriStateMixed msoTriStateToggle msoTrue 相对于原始大小缩放该形状。仅当指定形状是图片或 OLE 对象时，才为此参数指定 msoTrue。 MsoScaleFrom 可以是下列 MsoScaleFrom 常量之一。 msoScaleFromBottomRight msoScaleFromMiddle msoScaleFromTopLeft：默认值。

**示例**

``` JavaScript
/*本示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)

    for(let i = 1; i <= myDocument.Shapes.Count; i++){
        switch(myDocument.Shapes.Item(i).Type){
            case msoEmbeddedOLEObject:
            case msoLinkedOLEObject:
            case msoOLEControlObject:
            case msoLinkedPicture:
            case msoPicture:
                myDocument.Shapes.Item(i).ScaleHeight(1.75, msoTrue)
                myDocument.Shapes.Item(i).ScaleWidth(1.75, msoTrue)
                break
            default:
                myDocument.Shapes.Item(i).ScaleHeight(1.75, msoFalse)
                myDocument.Shapes.Item(i).ScaleWidth(1.75, msoFalse)
                break
        }
    }
}
```

### ShapeRange.ScaleWidth

根据指定的系数缩放图形宽度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于图片和 OLE 对象以外的其他图形来说，总是相对于当前宽度进行缩放。

**语法**

**express.ScaleWidth ( *Factor* , *RelativeToOriginalSize* , *fScale* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称                     | 必选/可选 | 数据类型     | 说明                                                                                                   |
|--------------------------|-----------|--------------|--------------------------------------------------------------------------------------------------------|
| *Factor*                 | 必选      | Single       | 指定形状调整后的宽度与当前或原始宽度的比例。例如，若要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选      | MsoTriState  | 指定是否相对于形状的当前或原始尺寸来缩放形状。                                                         |
| *fScale*                 | 可选      | MsoScaleFrom | 在缩放形状时，形状中位置不变的部分。                                                                   |

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。 msoCTrue msoFalse：相对于形状的当前尺寸缩放。 msoTriStateMixed msoTriStateToggle msoTrue 相对于原始大小缩放该形状。仅当指定形状是图片或 OLE 对象时，才为此参数指定 msoTrue。 MsoScaleFrom 可以是下列 MsoScaleFrom 常量之一。 msoScaleFromBottomRight msoScaleFromMiddle msoScaleFromTopLeft：默认值。

**示例**

``` JavaScript
/*本示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)

    for(let i = 1; i <= myDocument.Shapes.Count; i++){
        switch(myDocument.Shapes.Item(i).Type){
            case msoEmbeddedOLEObject:
            case msoLinkedOLEObject:
            case msoOLEControlObject:
            case msoLinkedPicture:
            case msoPicture:
                myDocument.Shapes.Item(i).ScaleHeight(1.75, msoTrue)
                myDocument.Shapes.Item(i).ScaleWidth(1.75, msoTrue)
                break
            default:
                myDocument.Shapes.Item(i).ScaleHeight(1.75, msoFalse)
                myDocument.Shapes.Item(i).ScaleWidth(1.75, msoFalse)
                break
        }
    }
}
```

### ShapeRange.Select

选择指定的对象。

**语法**

**express.Select ( *Replace* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型    | 说明                                       |
|-----------|-----------|-------------|--------------------------------------------|
| *Replace* | 可选      | MsoTriState | 指定该选择是否替换之前所做的任何一项选择。 |

**说明**

如果做出不适用于当前视图的选择，此代码将失败。例如，可以在幻灯片浏览视图中而不能在幻灯片视图中选择幻灯片。 MsoTriState 可以是下列 MsoTriState 常量之一。 msoCTrue msoFalse 将该选择添加到前一项选择中。 msoTriStateMixed msoTriStateToggle msoTrue 默认值。该选择替换之前所做的任何一项选择。

**示例**

``` JavaScript
/*本示例选择活动演示文稿第一张幻灯片上的第一个和第三个形状*/
Application.ActivePresentation.Slides.Item(1).Shapes.Range([1, 3]).Select()
```

### ShapeRange.Ungroup

取消指定形状或者形状区域中任意组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。取消组合后的形状以单个 ShapeRange 对象的形式返回。

**语法**

**express.Ungroup ()**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

由于将一组形状作为单个形状来处理，对形状进行组合和取消组合将会改变 Shapes 集合中项目的数量，并且改变集合中受影响的项目之后的项目索引编号。

**示例**

``` JavaScript
/*本示例取消 myDocument 中的所有形状组合，并分解所有的图片或 OLE 对象*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)

    for(let i = 1; i <= myDocument.Shapes.Count; i++){
        myDocument.Shapes.Item(i).Ungroup()
    }
}

/*本示例取消 myDocument 中的所有形状组合，但不分解幻灯片中的图片或 OLE 对象*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)

    for(let i = 1; i <= myDocument.Shapes.Count; i++){
        if(myDocument.Shapes.Item(i).Type == msoGroup) {
           myDocument.Shapes.Item(i).Ungroup()
        }
    }
}
```

### ShapeRange.ZOrder

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-顺序中的位置）。

**语法**

**express.ZOrder ( *ZorderCmd* )**

\*express是一个代表 ShapeRange 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型     | 说明                                     |
|-------------|-----------|--------------|------------------------------------------|
| *ZorderCmd* | 必选      | MsoZOrderCmd | 指定针对其他形状将指定形状移动到的位置。 |

**说明**

MsoZOrderCmd 可以是下列 MsoZOrderCmd 常量之一。 msoBringForward msoBringInFrontOfText 仅用于 WPS。 msoBringToFront msoSendBackward msoSendBehindText 仅用于 WPS。 msoSendToBack 可用 ZOrderPosition 属性确定形状在 z-顺序中的当前位置。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加椭圆，如果幻灯片中有至少一个其他形状，则将该椭圆放置在 z-顺序中倒数第二的位置上*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myShape = myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300)

    while(myShape.ZOrderPosition > 2){
        myShape.ZOrder(msoSendBackward)
    }
}
```

## 成员属性

### ShapeRange.ActionSettings

返回一个 ActionSettings 对象，该对象包含在幻灯片放映期间，当用户在指定形状或文本区域内单击或移动鼠标时所产生的动作的信息。只读。

**语法**

**express.ActionSettings**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例设置在当前演示文稿中第二张幻灯片的第一个形状中单击或移动鼠标时产生的动作*/
function test(){
    let myShape = Application.ActivePresentation.Slides.Item(2).Shapes.Item(1)
    myShape.ActionSettings(ppMouseClick).Action = ppActionLastSlide
    myShape.ActionSettings(ppMouseOver).SoundEffect.Name = "applause.wav"
}
```

### ShapeRange.Application

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/
function test() {
    let pptPres = Application.ActivePresentation
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称*/
function test(){
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes

    for(let i = 1; i <= myShapes.Count; i++){
        if(myShapes.Item(i).Type == msoLinkedOLEObject){
            MsgBox(myShapes.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### ShapeRange.Chart

返回当前 Shape 对象的 Chart 对象。只读。

**语法**

**express.Chart**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除当前窗口外的所有窗口*/
function test(){
    for(let i = 2;i <= Application.Windows.Count;i++){
        Application.Windows.Item(2).Close()
    }
}
```

### ShapeRange.Fill

返回一个 FillFormat 对象，该对象包含指定形状的填充格式属性。只读。

**语法**

**express.Fill**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let filler = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill

    filler.ForeColor.RGB = (128, 0, 0)
    filler.BackColor.RGB = (170, 170, 170)
    filler.TwoColorGradient(msoGradientHorizontal, 1)
}
```

### ShapeRange.Glow

返回指定形状范围的发光格式。只读 msoGlowType 类型。

**语法**

**express.Glow**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.GroupItems

返回一个 GroupShapes 对象，该对象代表指定形状组中的单个形状。使用 GroupShapes 对象的 Item 方法可返回形状组中的单个形状。适用于代表组合形状的 Shape 或 ShapeRange 对象。只读。

**语法**

**express.GroupItems**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例向myDocument中添加三个三角形，将它们组合，再设置整个组的颜色，然后只更改第二个三角形的颜色*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)

    myDocument.Shapes.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"
    myDocument.Shapes.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"
    myDocument.Shapes.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"

    let grouped = myDocument.Shapes.Range(["shpOne", "shpTwo", "shpThree"]).Group()
    grouped.Fill.PresetTextured(msoTextureBlueTissuePaper)
    grouped.GroupItems.Item(2).Fill.PresetTextured(msoTextureGreenMarble)
}
```

### ShapeRange.HasChart

返回指定的对象代表的形状区域是否包含图表。只读。

**语法**

**express.HasChart**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.HasTable

返回指定的形状是否为表。只读。MsoTriState 类型。

**语法**

**express.HasTable**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。 msoCTrue msoFalse msoTriStateMixed msoTriStateToggle msoTrue：指定的形状是表。

**示例**

``` JavaScript
/*以下示例检查当前选定的形状是否为表格。如果是表格，此代码将第一列的宽度设置为 1 英寸*/
function test(){
    let myRange = ActiveWindow.Selection.ShapeRange
    if(myRange.HasTable == msoTrue){
        myRange.Table.Columns.Item(1).Width = 72
    }
}
```

### ShapeRange.HasTextFrame

返回指定形状是否有文本框。只读。MsoTriState 类型。

**语法**

**express.HasTextFrame**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。 msoCTrue msoFalse msoTriStateMixed msoTriStateToggle msoTrue：指定形状有文本框，因此可包含文本。

**示例**

``` JavaScript
/*下面的例子从第一张幻灯片上所有包含文本框的形状中提取文本，然后将这些形状的名称及其所包含的文本保存在一个数组中*/
function test(){
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes

    if(myShapes.Count > 1){
        let numTextShapes = 0
        let shpTextArray = []

        for(let i = 1; i <= myShapes.Count; i++){
            if(myShapes.Item(i).HasTextFrame){
                numTextShapes++
                shpTextArray[numTextShapes] = []
                shpTextArray[numTextShapes][0] = myShapes.Item(i).Name
                shpTextArray[numTextShapes][1] = myShapes.Item(i).TextFrame.TextRange.Text
            }
        }
    }
}
```

### ShapeRange.Height

以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Height**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

一个 HeightShape 对象的 ShapeHeight 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

**示例**

``` JavaScript
/*以下示例将第二个文档窗口的高度设置为应用程序窗口高度的一半*/
Application.Windows.Item(2).Height = Application.Height / 2

/*以下示例将指定表格中第二行的高度设为 100 磅*/
Application.ActivePresentation.Slides.Item(2).Shapes.Item(2).Table.Rows.Item(2).Height = 100
```

### ShapeRange.Id

返回一个 Long 类型值，该值标识形状或形状范围。只读。

**语法**

**express.Id**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例向当前演示文稿添加一个新形状，再根据 ID 属性的值填充该形状*/
function test(){
    let myShape = Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShape5pointStar, 100, 100, 100, 100)

    switch(true){
        case myShape.Id <= 500:
            myShape.Fill.ForeColor.RGB = (255, 0, 0)
            break
        case myShape.Id <= 1000:
            myShape.Fill.ForeColor.RGB = (255, 255, 0)
            break
        case myShape.Id <= 1500:
            myShape.Fill.ForeColor.RGB = (255, 0, 255)
            break
        case myShape.Id <= 2000:
            myShape.Fill.ForeColor.RGB = (0, 255, 0)
            break
        case myShape.Id <= 2500:
            myShape.Fill.ForeColor.RGB = (0, 255, 255)
            break
        default:
            myShape.Fill.ForeColor.RGB = (0, 0, 255)
    }
}
```

### ShapeRange.Left

返回或设置一个 Single 类型值，该值代表从形状范围中最左侧形状的左边缘到幻灯片左边缘的距离（以磅为单位）。可读/写。

**语法**

**express.Left**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Line

返回一个 LineFormat 对象，该对象包含指定形状的线条格式属性。（对于线条来说，LineFormat 对象代表线条本身；而对于带有边框的形状来说，LineFormat 对象代表边框。）只读。

**语法**

**express.Line**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加一条蓝色虚线*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myLine = myDocument.Shapes.AddLine(10, 10, 250, 250).Line

    myLine.DashStyle = msoLineDashDotDot
    myLine.ForeColor.RGB = (50, 0, 128)
}
```

### ShapeRange.MediaType

返回 OLE 媒体类型。只读。PpMediaType 类型。

**语法**

**express.MediaType**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

PpMediaType 可以是下列 PpMediaType 常量之一。 ppMediaTypeMixed ppMediaTypeMovie ppMediaTypeOther ppMediaTypeSound

**示例**

``` JavaScript
/*以下示例设置在放映幻灯片时，当前演示文稿内第一张幻灯片上的所有原始声音对象循环播放直到被手动停止*/
function test(){
    let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes

    for(let i = 1; i <= myShapes.Count; i++){
        if(myShapes.Item(i).Type == msoMedia){
            if(myShapes.Item(i).MediaType == ppMediaTypeSound){
                myShapes.Item(i).AnimationSettings.PlaySettings.LoopUntilStopped = true
            }
        }
    }
}
```

### ShapeRange.Name

创建形状时，WPP 自动以 ShapeType Number 的形式为其分配一个名称，其中 ShapeType 指明形状或自选图形的类型，Number 是在幻灯片上形状的集合中具有唯一性的整数。例如，为幻灯片上的形状自动生成的名称可以为 Placeholder 1、Oval 2 和 Rectangle 3。为了避免与自动分配的名称发生冲突，对用户定义的名称请不要使用 ShapeType Number 形式，其中 ShapeType 为一个用于自动生成名称的值，Number 为任意的正整数。形状范围必须至少包含一个形状。String 类型，可读/写。

**语法**

**express.Name**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

如果包含某对象的集合的 Item 方法采用 Variant 类型的参数，则可以将该对象的名称与 Item 方法联合使用来返回对该对象的引用。例如，如果某个形状的 Name 属性值为“Rectangle 2”，则 .Shapes.Item("Rectangle 2") 将返回对该形状的引用。

### ShapeRange.PictureFormat

返回一个 PictureFormat 对象，该对象包含指定形状的图片格式属性。适用于代表图片或 OLE 对象的 Shape 或 ShapeRange 对象。只读。

**语法**

**express.PictureFormat**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例设置了 myDocument 中第一个形状的亮度和对比度。第一个形状必须是图片或 OLE 对象*/
function test(){
    let myDocument = ActivePresentation.Slides.Item(1)
    myDocument.Shapes.Item(1).PictureFormat.Brightness = 0.3
    myDocument.Shapes.Item(1).PictureFormat.Contrast = .75
}
```

### ShapeRange.PlaceholderFormat

返回一个 PlaceholderFormat 对象，该对象包含占位符特有的属性。只读。

**语法**

**express.PlaceholderFormat**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例在当前演示文稿内第一张幻灯片的第一个占位符中添加文本（如果该占位符是水平标题占位符*/
function test(){
    let phs = Application.ActivePresentation.Slides.Item(1).Shapes.Placeholders

    if(phs.Count > 0){
        switch(phs.Item(1).PlaceholderFormat.Type){
            case ppPlaceholderTitle:
                phs.Item(1).TextFrame.TextRange.Text = "Title Text"
                break
            case ppPlaceholderCenterTitle:
                phs.Item(1).TextFrame.TextRange.Text = "Centered Title Text"
                break
            default:
                MsgBox("There's no horizontal title on this slide")
        }
    }
}
```

### ShapeRange.Shadow

返回一个只读的 ShadowFormat 对象，该对象包含指定形状的阴影格式属性。

**语法**

**express.Shadow**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*本示例在当前演示文稿第一张幻灯片中添加一个有阴影的矩形。蓝色浮凸效果的阴影距该矩形的右边 3 磅，下边 2 磅*/
function test(){
    let myShap = Application.ActivePresentation.Slides.Item(1).Shapes
    let myShad = myShap.AddShape(msoShapeRectangle, 10, 10, 150, 90).Shadow
        myShad.Type = msoShadow17
        myShad.ForeColor.RGB = (0, 0, 128)
        myShad.OffsetX = 3
        myShad.OffsetY = 2
}
```

### ShapeRange.ShapeStyle

设置或返回指定对象的形状样式索引

**语法**

**express.ShapeStyle**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Table

返回一个 Table 对象，该对象代表某个形状或形状区域内的一个表格。只读。

**语法**

**express.Table**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例将第三张幻灯片第二个形状中表格第二列的宽度设置为 60 磅*/
Application.ActivePresentation.Slides.Item(3).Shapes.Item(2).Table.Rows.Item(2).Height = 60
```

### ShapeRange.TextEffect

返回一个 TextEffectFormat 对象，该对象包含指定形状的文本效果格式属性。适用于代表艺术字的 Shape 或 ShapeRange 对象。

**语法**

**express.TextEffect**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例将 myDocument 上第三个形状的字体样式设置为加粗（如果该形状为艺术字）*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myShape = myDocument.Shapes.Item(3)
        if(myShape.Type == msoTextEffect) {
            myShape.TextEffect.FontBold = true
        }
}
```

### ShapeRange.TextFrame

返回一个 TextFrame 对象，该对象包含指定形状或母版文本样式的对齐方式和定位属性。

**语法**

**express.TextFrame**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

使用 TextFrame 对象的 TextRange 属性可返回文本框中的文本。 在应用 TextFrame 属性之前，可使用 HasTextFrame 属性来确定形状中是否包含文本框。

**示例**

``` JavaScript
/*以下示例在 myDocument 中添加一个矩形，向该矩形中添加文本，然后设置文本框的上边距*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myTextFrame = myDocument.Shapes 
            .AddShape(msoShapeRectangle, 180, 175, 350, 140).TextFrame
        myTextFrame.TextRange.Text = "Here is some test text"
        myTextFrame.MarginTop = 10
}
```

### ShapeRange.TextFrame2

返回与指定的 ShapeRange 对象相关联的 TextFrame2 对象，前者包含指定的形状范围的对齐和定位属性。只读。

**语法**

**express.TextFrame2**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.ThreeD

返回一个 ThreeDFormat 对象，该对象包含指定形状的三维效果格式属性。只读。

**语法**

**express.ThreeD**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
/*以下示例为应用于 myDocument 中第一个形状的三维效果设置深度、延伸色、延伸方向以及照明方向*/
function test(){
    let myDocument =Application.ActivePresentation.Slides.Item(1)
    let myThreeD = myDocument.Shapes.Item(1).ThreeD
        myThreeD.Visible = true
        myThreeD.Depth = 50
        //RGB value for purple
        myThreeD.ExtrusionColor.RGB = (255, 100, 255)
        myThreeD.SetExtrusionDirection(msoExtrusionTop)
        myThreeD.PresetLightingDirection = msoLightingLeft
}
```

### ShapeRange.Top

返回或设置一个 Single 类型值，该值代表形状范围内最上方形状的上边缘到文档上边缘的距离。可读/写。

**语法**

**express.Top**

\*express是一个代表 ShapeRange 对象的变量。

### ShapeRange.Type

返回一个 **MsoShapeType** 常量，该常量代表单个形状或形状范围内多个形状的类型。只读。

**语法**

**express.Type**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

MsoShapeType 可以是下列 MsoShapeType 常量之一。 msoAutoShape msoCallout msoCanvas msoChart msoComment msoDiagram msoEmbeddedOLEObject msoFormControl msoFreeform msoGroup msoLine msoLinkedOLEObject msoLinkedPicture msoMedia msoOLEControlObject msoPicture msoPlaceholder msoScriptAnchor msoShapeTypeMixed msoTable msoTextBox msoTextEffect

### ShapeRange.Visible

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。 **MsoTriState** 类型。

**语法**

**express.Visible**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** ：指定对象或对象格式是可见的。    |

**示例**

``` JavaScript
/*以下示例为当前演示文稿中第一张幻灯片的第三个形状设置阴影的水平偏移量和垂直偏移量，使其向右偏移 5 磅并向上偏移 3 磅。如果该形状原来没有阴影，本例将会为其添上阴影*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myShadow = myDocument.Shapes.Item(3).Shadow
        myShadow.Visible = msoTrue
        myShadow.OffsetX = 5
        myShadow.OffsetY = -3
}
```

### ShapeRange.Width

以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Width**

\*express是一个代表 ShapeRange 对象的变量。

**示例**

``` JavaScript
//以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所有水平可用空间。要使以下示例执行，只能有两个打开的演示窗口*/
function test(){
    //pplicatioAn.Windows.Arrange(ppArrangeTiled)
    let sngHeight = Application.ActivePresentation.Windows.Item(1).Height                     // available height
    let sngWidth = Application.ActivePresentation.Windows.Item(1).Width + Windows.Item(2).Width    // available width
    let win1 = Application.ActivePresentation.Windows.Item(1)
        win1.Width = sngWidth
        win1.Height = sngHeight / 2
        win1.Left = 0
    let win2 = Application.ActivePresentation.Windows.Item(2)
        win2.Width = sngWidth
        win2.Height = sngHeight / 2
        win2.Top = sngHeight / 2
        win2.Left = 0
}

/*以下示例将指定表中第一列的宽度设置为 80 磅*/
ActivePresentation.Slides.Item(2).Shapes.Item(2).Table.Columns.Item(1).Width = 80
```

### ShapeRange.ZOrderPosition

返回指定的形状在 z-顺序中的位置。只读。Long 类型。

**语法**

**express.ZOrderPosition**

\*express是一个代表 ShapeRange 对象的变量。

**说明**

Shapes(1) 返回 z-顺序中的最后一个形状，Shapes(Shapes.Count) 返回 z-顺序中的第一个形状。 若要设置形状在 z-顺序中的位置，请用 ZOrder 方法。 形状在 z-顺序中的位置与 Shapes 集合中形状的索引号相对应。例如，如果幻灯片上有四个形状，表达式 myDocument.Shapes(1) 在 z-顺序的后面返回该形状，表达式 myDocument.Shapes(4) 在 z-顺序的前面返回该形状。 如向一个集合添加新图形，其默认的添加位置是 z-顺序中的前面位置。

**示例**

``` JavaScript
/*以下示例向 myDocument 中添加椭圆，如果幻灯片中有至少一个其他形状，则将该椭圆放置在 z-顺序中倒数第二的位置上*/
function test(){
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    let myShape = 
     myDocument.Shapes.AddShape(msoShapeOval,  100, 100, 100, 300) 
     while(myShape.ZOrderPosition > 2) {
           myShape.ZOrder(msoSendBackward)
        }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/ShapeRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SlideShowSettings 对象

## 简介

代表演示文稿的幻灯片放映设置。

## 方法

| 序号 | 名称                          | 说明                                                          |
|:----:|:------------------------------|:--------------------------------------------------------------|
|  1   | [Run](#SlideShowSettings.Run) | 运行指定演示文稿的幻灯片放映。返回一个 SlideShowWindow 对象。 |

## 属性

| 序号 | 名称                                                      | 说明                                                                                                                                                                                                                                                                                     |
|:----:|:----------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AdvanceMode](#SlideShowSettings.AdvanceMode)             | 返回或设置一个值，该值指示幻灯片放映的切换方式。可读/写 PpSlideShowAdvanceMode 类型。                                                                                                                                                                                                    |
|  2   | [Application](#SlideShowSettings.Application)             | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                                                                  |
|  3   | [EndingSlide](#SlideShowSettings.EndingSlide)             | 返回或设置指定幻灯片放映中最后显示的幻灯片。可读/写 Long 类型。                                                                                                                                                                                                                          |
|  4   | [LoopUntilStopped](#SlideShowSettings.LoopUntilStopped)   | 决定指定幻灯片放映是否持续循环播放，直到用户按 Esc。 MsoTriState 类型，可读/写。                                                                                                                                                                                                         |
|  5   | [NamedSlideShows](#SlideShowSettings.NamedSlideShows)     | 返回一个 NamedSlideShows 集合，该集合代表指定演示文稿中所有命名幻灯片放映（自定义幻灯片放映）。每个命名幻灯片放映或自定义幻灯片放映都是指定演示文稿的用户选择的子集。只读。                                                                                                              |
|  6   | [Parent](#SlideShowSettings.Parent)                       | 返回指定对象的父对象。                                                                                                                                                                                                                                                                   |
|  7   | [PointerColor](#SlideShowSettings.PointerColor)           | 以 ColorFormat 对象的形式返回指定演示文稿的指针颜色。该颜色随演示文稿一起保存，并且是该演示文稿每次运行时的默认绘图笔颜色。只读。                                                                                                                                                        |
|  8   | [RangeType](#SlideShowSettings.RangeType)                 | 返回或设置运行幻灯片放映的类型。可读/写 PpSlideShowRangeType 类型。                                                                                                                                                                                                                      |
|  9   | [ShowScrollbar](#SlideShowSettings.ShowScrollbar)         | 属性值为 MsoTrue 时，在以浏览模式放映幻灯片的过程中显示滚动条。可读/写 MsoTriState 类型。                                                                                                                                                                                                |
|  10  | [ShowType](#SlideShowSettings.ShowType)                   | 返回或设置指定幻灯片放映的放映类型。可读/写 PpSlideShowType 类型。                                                                                                                                                                                                                       |
|  11  | [ShowWithAnimation](#SlideShowSettings.ShowWithAnimation) | 决定指定的幻灯片放映是否显示已分配动画设置的形状。可读/写 MsoTriState 类型。                                                                                                                                                                                                             |
|  12  | [ShowWithNarration](#SlideShowSettings.ShowWithNarration) | 决定显示指定幻灯片放映时是否有旁白。可读/写 MsoTriState 类型。                                                                                                                                                                                                                           |
|  13  | [SlideShowName](#SlideShowSettings.SlideShowName)         | 返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（ ActionSetting 对象）；返回或设置自定义幻灯片放映的名称以便打印（ PrintOptions 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（ PublishObject 对象）。 String 类型，可读/写。 |
|  14  | [StartingSlide](#SlideShowSettings.StartingSlide)         | 返回或设置指定幻灯片放映中要显示的第一张幻灯片。可读/写 Long 类型。                                                                                                                                                                                                                      |

## 成员方法

### SlideShowSettings.Run

运行指定演示文稿的幻灯片放映。返回一个 **SlideShowWindow** 对象。

**语法**

**express.Run ()**

\*express是一个代表 SlideShowSettings 对象的变量。

**返回值**

SlideShowWindow

**说明**

要运行自定义幻灯片放映，请将 **RangeType** 属性设置为 **ppShowNamedSlideShow** ，并将 **SlideShowName** 属性设置为要运行的自定义放映的名称。

**示例**

``` JavaScript
/*本示例为活动演示文稿播放全屏幻灯片放映，并禁用快捷键。*/
function test(){
　　　　ActivePresentation.SlideShowSettings.ShowType = ppShowSpeaker
　　　　ActivePresentation.SlideShowSettings.Run().View.AcceleratorsEnabled = false}
```

``` JavaScript
/*本示例运行名为“Quick Show”的幻灯片放映。*/
function test(){
　　　　ActivePresentation.SlideShowSettings.RangeType = ppShowNamedSlideShow
　　　　ActivePresentation.SlideShowSettings.SlideShowName = "Quick Show"
　　　　ActivePresentation.SlideShowSettings.Run()
}
```

## 成员属性

### SlideShowSettings.AdvanceMode

返回或设置一个值，该值指示幻灯片放映的切换方式。可读/写 **PpSlideShowAdvanceMode** 类型。

**语法**

**express.AdvanceMode**

\*express是一个代表 SlideShowSettings 对象的变量。

**说明**

| PpSlideShowAdvanceMode 可以是下列 PpSlideShowAdvanceMode 常量之一。 |
|---------------------------------------------------------------------|
| ppSlideShowManualAdvance                                            |
| ppSlideShowRehearseNewTimings                                       |
| ppSlideShowUseSlideTimings                                          |

### SlideShowSettings.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 SlideShowSettings 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test(){
　　　　let shpOle = ActivePresentation.Slides.Item(1).Shapes
       for(let i = 1; i <= shpOle.Count; i++) {
   　　　　 if(shpOle.Item(i).Type == msoLinkedOLEObject) {
    　　　　MsgBox(shpOle.Item(i).OLEFormat.Application.Name)
  　　　　  }
　　　　}
}
```

### SlideShowSettings.EndingSlide

返回或设置指定幻灯片放映中最后显示的幻灯片。可读/写 **Long** 类型。

**语法**

**express.EndingSlide**

\*express是一个代表 SlideShowSettings 对象的变量。

**示例**

``` JavaScript
/*本示例运行活动演示文稿的一个幻灯片放映，该放映从第二张幻灯片开始，到第四张幻灯片结束。*/
function test(){
　　　　let SliSet = ActivePresentation.SlideShowSettings
   　　　　 SliSet.RangeType = ppShowSlideRange
  　　　　  SliSet.StartingSlide = 2
  　　　　  SliSet.EndingSlide = 4
   　　　　 SliSet.Run()
}
```

### SlideShowSettings.LoopUntilStopped

决定指定幻灯片放映是否持续循环播放，直到用户按 Esc。 **MsoTriState** 类型，可读/写。

**语法**

**express.LoopUntilStopped**

\*express是一个代表 SlideShowSettings 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。    |
|--------------------------------------------------|
| msoCTrue                                         |
| msoFalse                                         |
| msoTriStateMixed                                 |
| msoTriStateToggle                                |
| msoTrue 指定幻灯片放映持续循环，直到用户按 Esc。 |

**示例**

``` JavaScript
/*本示例开始当前演示文稿的一个幻灯片放映，该演示文稿自动切换幻灯片（按存储的时间间隔），并循环放映直到用户按 Esc。*/
function test(){
　　　　let SliSet = ActivePresentation.SlideShowSettings
  　　　　  SliSet.AdvanceMode = ppSlideShowUseSlideTimings
  　　　　  SliSet.LoopUntilStopped = msoTrue
   　　　　 SliSet.Run()}
```

### SlideShowSettings.NamedSlideShows

返回一个 **NamedSlideShows** 集合，该集合代表指定演示文稿中所有命名幻灯片放映（自定义幻灯片放映）。每个命名幻灯片放映或自定义幻灯片放映都是指定演示文稿的用户选择的子集。只读。

**语法**

**express.NamedSlideShows**

\*express是一个代表 SlideShowSettings 对象的变量。

**说明**

使用 **NamedSlideShows** 对象的 **Add** 方法创建命名幻灯片放映。

**示例**

``` JavaScript
/*本示例为当前演示文稿添加一个命名幻灯片放映“Quick Show”（该放映包含第二、第七和第九张幻灯片），然后运行该放映。*/
function test(){
　　　　let qSlides = []
　　　　let SliID = ActivePresentation.Slides
   　　　　 qSlides[1] = SliID.Item(2).SlideID
  　　　　  qSlides[2] = SliID.Item(7).SlideID
    　　　　qSlides[3] = SliID.Item(9).SlideID

　　　　let SliSet = ActivePresentation.SlideShowSettings
   　　　　 SliSet.RangeType = ppShowNamedSlideShow
   　　　　 SliSet.NamedSlideShows.Add("Quick Show", qSlides)
  　　　　  SliSet.SlideShowName = "Quick Show"
   　　　　 SliSet.Run()
}
```

### SlideShowSettings.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 SlideShowSettings 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShapeOval,50,50,300,150).TextFrame
  　　　　  myShapes.TextRange.Text = "Test text"
  　　　　  myShapes.Parent.Rotation = 45
}
```

### SlideShowSettings.PointerColor

以 **ColorFormat** 对象的形式返回指定演示文稿的指针颜色。该颜色随演示文稿一起保存，并且是该演示文稿每次运行时的默认绘图笔颜色。只读。

**语法**

**express.PointerColor**

\*express是一个代表 SlideShowSettings 对象的变量。

**说明**

要将指针更改为绘图笔，可将 **PointerType** 属性设置为 **ppSlideShowPointerPen** 。

**示例**

``` JavaScript
/*本示例将当前演示文稿的默认绘图笔颜色设为蓝色，然后开始一个幻灯片放映，将指针改变为绘图笔，并设定绘图笔颜色仅在此幻灯片放映中为红色。*/
function test(){
　　　　let SliSet = ActivePresentation.SlideShowSettings
 　　　　   SliSet.PointerColor.RGB = (0, 0, 255)          //blue
  　　　　  SliSet.Run().View.PointerColor.RGB = (255, 0, 0)      //red
　　　　    SliSet.Run().View.PointerType = ppSlideShowPointerPen
}
```

### SlideShowSettings.RangeType

返回或设置运行幻灯片放映的类型。可读/写 **PpSlideShowRangeType** 类型。

**语法**

**express.RangeType**

\*express是一个代表 SlideShowSettings 对象的变量。

**说明**

| PpSlideShowRangeType 可以是下列 PpSlideShowRangeType 常量之一。 |
|-----------------------------------------------------------------|
| ppShowAll                                                       |
| ppShowNamedSlideShow                                            |
| ppShowSlideRange                                                |

若要打印在 **PrintRanges** 集合中定义的幻灯片范围，必须先将 **RangeType** 属性设置为 **ppPrintSlideRange** 。将 **RangeType** 设置为除 **ppPrintSlideRange** 之外的值会使在 **PrintRanges** 集合中定义的范围失效。但是这不会以任何方式影响 **PrintRanges** 集合的内容。也就是说，如果定义了一些打印范围，将 **RangeType** 属性设置为除 **ppPrintSlideRange** 之外的值，然后将 **RangeType** 重新设置为 **ppPrintSlideRange** ，则之前定义的打印范围将保持不变。

指定 **PrintOut** 方法的 *To* 和 *From* 参数将设置该属性的值。

**示例**

``` JavaScript
/*本示例运行名为“Quick Show”的幻灯片放映。*/
function test(){
　　　　ActivePresentation.SlideShowSettings.RangeType = ppShowNamedSlideShow
　　　　ActivePresentation.SlideShowSettings.SlideShowName = "Quick Show"
　　　　ActivePresentation.SlideShowSettings.Run()
}
```

### SlideShowSettings.ShowScrollbar

属性值为 **MsoTrue** 时，在以浏览模式放映幻灯片的过程中显示滚动条。可读/写 **MsoTriState** 类型。

**语法**

**express.ShowScrollbar**

\*express是一个代表 SlideShowSettings 对象的变量。

**说明**

在设置 **ShowScrollbar** 属性之前使用 **ShowType** 属性。

------------------------------------------------------------------------

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue 不用于此属性。                       |
| msoFalse                                      |
| msoTriStateMixed 不用于此属性。               |
| msoTriStateToggle 不用于此属性                |
| msoTrue                                       |

**示例**

``` JavaScript
/*本示例指定在某个窗口中显示当前演示文稿的幻灯片放映，并且在幻灯片放映时显示用于浏览幻灯片的滚动条。*/
function ShowSlideShowScrollBar() {
    ActivePresentation.SlideShowSettings.ShowType = ppShowTypeWindow
    ActivePresentation.SlideShowSettings.ShowScrollBar = msoTrue
}
```

### SlideShowSettings.ShowType

返回或设置指定幻灯片放映的放映类型。可读/写 **PpSlideShowType** 类型。

**语法**

**express.ShowType**

\*express是一个代表 SlideShowSettings 对象的变量。

**说明**

| PpSlideShowType 可以是下列 PpSlideShowType 常量之一。 |
|-------------------------------------------------------|
| ppShowTypeKiosk                                       |
| ppShowTypeSpeaker                                     |
| ppShowTypeWindow                                      |

**示例**

``` JavaScript
/*本示例在窗口中运行当前演示文稿的一个幻灯片放映，从第二张幻灯片开始，到第四张幻灯片结束。新的幻灯片放映窗口位于屏幕左上角，宽度和高度均为 300 磅。*/
function test(){
　　　　let SliSet = ActivePresentation.SlideShowSettings
   　　　　 SliSet.RangeType = ppShowSlideRange
   　　　　 SliSet.StartingSlide = 2
  　　　　  SliSet.EndingSlide = 4
 　　　　   SliSet.ShowType = ppShowTypeWindow
　　　　let SliSetR = SliSet.Run()
  　　　　  SliSetR.Left = 0
　　　　    SliSetR.Top = 0
  　　　　  SliSetR.Width = 300
   　　　　 SliSetR.Height = 300
}
```

### SlideShowSettings.ShowWithAnimation

决定指定的幻灯片放映是否显示已分配动画设置的形状。可读/写 **MsoTriState** 类型。

**语法**

**express.ShowWithAnimation**

\*express是一个代表 SlideShowSettings 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。      |
|----------------------------------------------------|
| msoCTrue                                           |
| msoFalse                                           |
| msoTriStateMixed                                   |
| msoTriStateToggle                                  |
| msoTrue 指定的幻灯片放映显示已分配动画设置的形状。 |

**示例**

``` JavaScript
/*本示例运行活动演示文稿的幻灯片放映，关闭动画和旁白。*/
function test(){
　　　　ActivePresentation.SlideShowSettings.ShowWithAnimation = msoFalse
　　　　ActivePresentation.SlideShowSettings.ShowWithNarration = msoFalse
　　　　ActivePresentation.SlideShowSettings.Run()
}
```

### SlideShowSettings.ShowWithNarration

决定显示指定幻灯片放映时是否有旁白。可读/写 **MsoTriState** 类型。

**语法**

**express.ShowWithNarration**

\*express是一个代表 SlideShowSettings 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue 显示指定幻灯片放映时有旁白。          |

**示例**

``` JavaScript
/*本示例运行活动演示文稿的幻灯片放映，关闭动画和旁白。*/
function test(){
　　　　ActivePresentation.SlideShowSettings.ShowWithAnimation = msoFalse
　　　　ActivePresentation.SlideShowSettings.ShowWithNarration = msoFalse
　　　　ActivePresentation.SlideShowSettings.Run()
}
```

### SlideShowSettings.SlideShowName

返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（ **ActionSetting** 对象）；返回或设置自定义幻灯片放映的名称以便打印（ **PrintOptions** 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（ **PublishObject** 对象）。 **String** 类型，可读/写。

**语法**

**express.SlideShowName**

\*express是一个代表 SlideShowSettings 对象的变量。

**说明**

必须将 **RangeType** 属性设置为 **ppPrintNamedSlideShow** 才能打印自定义幻灯片放映。

**示例**

``` JavaScript
/*本示例打印名为“tech talk”的现有自定义幻灯片放映。*/
function test(){
　　　　ActivePresentation.PrintOptions.RangeType = ppPrintNamedSlideShow
　　　　ActivePresentation.PrintOptions.SlideShowName = "tech talk"
　　　　ActivePresentation.PrintOut()
}
```

``` JavaScript
/*下例将当前演示文稿保存为 HTML 4.0 文件格式，文件名为“mallard.htm”。然后显示一条消息，说明正将当前名称的演示文稿保存为WPP 和 HTML 两种格式。*/
function test(){
　　　　let ActivePub = ActivePresentation.PublishObjects.Item(1)
　　　　let PresName = ActivePub.SlideShowName
　　　　    ActivePub.SourceType = ppPublishAll
　　　　    ActivePub.FileName = "C:\\HTMLPres\\mallard.htm"
　　　　    ActivePub.HTMLVersion = ppHTMLv4
　　　　    MsgBox("Saving presentation " + "'" + PresName + "'" + " inWPP" + "\n" + "\r" + " format and HTML version 4.0 format")
　　　　    ActivePub.Publish
}
```

### SlideShowSettings.StartingSlide

返回或设置指定幻灯片放映中要显示的第一张幻灯片。可读/写 **Long** 类型。

**语法**

**express.StartingSlide**

\*express是一个代表 SlideShowSettings 对象的变量。

**示例**

``` JavaScript
/*本示例运行活动演示文稿的一个幻灯片放映，该放映从第二张幻灯片开始，到第四张幻灯片结束。*/
function test(){
　　　　ActivePresentation.SlideShowSettings.RangeType = ppShowSlideRange
　　　　ActivePresentation.SlideShowSettings.StartingSlide = 2
　　　　ActivePresentation.SlideShowSettings.EndingSlide = 4
　　　　ActivePresentation.SlideShowSettings.Run()
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/SlideShowSettings](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SlideShowWindows 对象

## 简介

代表WPP 中打开的幻灯片放映的所有 SlideShowWindow 对象的集合。

## 方法

| 序号 | 名称                           | 说明                             |
|:----:|:-------------------------------|:---------------------------------|
|  1   | [Item](#SlideShowWindows.Item) | 集合中要返回的单个对象的索引号。 |

## 属性

| 序号 | 名称                                         | 说明                                                    |
|:----:|:---------------------------------------------|:--------------------------------------------------------|
|  1   | [Application](#SlideShowWindows.Application) | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
|  2   | [Count](#SlideShowWindows.Count)             | 返回指定集合中的对象数目。                              |

## 成员方法

### SlideShowWindows.Item

集合中要返回的单个对象的索引号。

**语法**

**express.Item ( *Index* )**

\*express是一个代表 SlideShowWindows 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                             |
|---------|-----------|----------|----------------------------------|
| *Index* | 必选      | Long     | 集合中要返回的单个对象的索引号。 |

## 成员属性

### SlideShowWindows.Application

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 SlideShowWindows 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/
function test(){
    let pptPres = Application.ActivePresentation
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") 
}


/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称*/
function test(){
    let shpOle = ActivePresentation.Slides.Item(1).Shapes
    for(let i = 1; i <= shpOle.Count; i++) {
        if(shpOle.Item(i).Type == msoLinkedOLEObject) {
            MsgBox(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### SlideShowWindows.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 SlideShowWindows 对象的变量。

**示例**

``` JavaScript
/*以下示例关闭除窗口1以外的所有窗口*/
function test(){
    for(let i = 2; i <= Application.Windows.Count; i++) {
        Application.Windows.Item(2).Close()
    }
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/SlideShowWindows](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# TextRange 对象

## 简介

包含附加到形状的文本，以及用于操作文本的属性和方法。

## 说明

以下示例说明如何执行下列操作：

- 返回任意指定形状中的文本范围。
- 返回选定范围中的文本范围。
- 返回文本范围中的特定字符、单词、行、句子或段落。
- 查找和替换文本范围内的文本。
- 向文本范围中插入文本、日期和时间或幻灯片编号。
- 将光标定位到文本范围内所需的任意位置。

``` JavaScript
/*使用 TextFrame 对象的 TextRange 属性返回任意指定形状的 TextRange 对象。使用 Text 属性返回 TextRange 对象中的文本字符串。以下示例向 myDocument 中添加一个矩形并设置其包含的文本。*/
function test() {
    let myDocument = Application.ActivePresentation.Slides.Item(1)
    myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame.TextRange.Text = "Here is some test text"
}
```

``` JavaScript
/*因为 Text 属性是 TextRange 对象的默认属性，所以以下两个语句是等效的。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Text = "Here is some test text"
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange = "Here is some test text"
```

使用 HasTextFrame 属性判断形状是否含有文本框，然后使用 HasText 属性判断该文本框是否包含文本。

``` JavaScript
/*使用 Selection 对象的 TextRange 属性返回当前选定的文字。以下示例将选定内容复制到剪贴板。*/
Application.ActiveWindow.Selection.TextRange.Copy()
```

使用下列方法之一可返回 TextRange 对象中的部分文本： Characters 、 Lines 、 Paragraphs 、 Runs 、 Sentences 或 Words 。

使用 Find 和 Replace 方法可查找和替换文本范围内的文本。

使用下列方法之一可向 TextRange 对象中插入字符： InsertAfter 、 InsertBefore 、 InsertDateTime 、 InsertSlideNumber 或 InsertSymbol 。

## 方法

| 序号 | 名称                                              | 说明                                                                                                                                                 |
|:----:|:--------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ChangeCase](#TextRange.ChangeCase)               | 更改指定文本的大小写。                                                                                                                               |
|  2   | [Characters](#TextRange.Characters)               | 返回一个 TextRange 对象，该对象代表指定的文本字符子集。有关计算或浏览一段文本中的字符的详细信息，请参阅 TextRange 对象。                             |
|  3   | [Copy](#TextRange.Copy)                           | 将指定对象复制到剪贴板。                                                                                                                             |
|  4   | [Cut](#TextRange.Cut)                             | 删除指定对象并将其放到剪贴板中。                                                                                                                     |
|  5   | [Delete](#TextRange.Delete)                       | 删除指定的对象。                                                                                                                                     |
|  6   | [Find](#TextRange.Find)                           | 在一个文本范围内查找指定的文本，并返回 TextRange 对象，该对象代表在其中找到指定的文本的第一个文本范围。如果找不到指定的文本，则返回 Nothing 。       |
|  7   | [InsertAfter](#TextRange.InsertAfter)             | 在指定文本范围的末尾添加一个字符串。返回一个代表添加文本的 TextRange 对象。省略参数时，该方法返回一个在指定范围末尾的长度为零的字符串。              |
|  8   | [InsertBefore](#TextRange.InsertBefore)           | 在指定文本范围的开头添加一个字符串。返回一个代表添加文本的 TextRange 对象。省略参数时，该方法返回一个在指定范围末尾的长度为零的字符串。              |
|  9   | [InsertDateTime](#TextRange.InsertDateTime)       | 在指定的文本范围中插入日期和时间。返回一个代表所插入文本的 TextRange 对象。                                                                          |
|  10  | [InsertSlideNumber](#TextRange.InsertSlideNumber) | 将当前幻灯片的幻灯片编号插入指定的文本范围。返回一个代表幻灯片编号的 TextRange 对象。                                                                |
|  11  | [InsertSymbol](#TextRange.InsertSymbol)           | 返回一个 TextRange 对象，该对象代表一个插入到指定文本范围的符号。                                                                                    |
|  12  | [Lines](#TextRange.Lines)                         | 返回一个 TextRange 对象，该对象代表指定的文本行子集。有关计算或浏览文本范围中文本行的信息，请参阅 TextRange 对象。                                   |
|  13  | [Paragraphs](#TextRange.Paragraphs)               | 返回一个 TextRange 对象，该对象代表文本段落的指定子集。有关统计或遍历文本范围内的段落的信息，请参阅 TextRange 对象。                                 |
|  14  | [Paste](#TextRange.Paste)                         | 将剪贴板上的文本粘贴到指定的文本范围中，并返回一个代表粘贴文本的 TextRange 对象。                                                                    |
|  15  | [Replace](#TextRange.Replace)                     | 在文本范围内查找特定文本，用指定的字符串替换查找到的文本，返回代表查找文本的第一处匹配内容的 TextRange 对象。如果找不到匹配的内容，则返回 Nothing 。 |
|  16  | [Select](#TextRange.Select)                       | 选择指定的对象。                                                                                                                                     |

## 属性

| 序号 | 名称                                          | 说明                                                                                                                     |
|:----:|:----------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#TextRange.Application)         | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                  |
|  2   | [BoundHeight](#TextRange.BoundHeight)         | 以磅为单位返回指定文本框的文本 <span class="glossary">边界框</span> 的高度。只读 Single 类型。                           |
|  3   | [BoundLeft](#TextRange.BoundLeft)             | 以磅为单位返回指定文本框的文本 <span class="glossary">边界框</span> 的左边界到幻灯片的左边界的距离。只读。 Single 类型。 |
|  4   | [BoundTop](#TextRange.BoundTop)               | 以磅为单位返回指定文本框的文本 <span class="glossary">边界框</span> 的上边界到幻灯片的上边界的距离。只读。 Single 类型。 |
|  5   | [BoundWidth](#TextRange.BoundWidth)           | 以磅为单位返回指定文本框的文本 <span class="glossary">边界框</span> 的宽度。只读 Single 类型。                           |
|  6   | [Count](#TextRange.Count)                     | 返回指定集合中的对象数目。                                                                                               |
|  7   | [Font](#TextRange.Font)                       | 返回一个代表字符格式的 Font 对象。只读。                                                                                 |
|  8   | [ParagraphFormat](#TextRange.ParagraphFormat) | 返回一个 ParagraphFormat 对象，该对象代表指定文本的段落格式。只读。                                                      |
|  9   | [Text](#TextRange.Text)                       | 返回或设置一个 String 类型值，该值代表指定对象中包含的文本。可读/写。                                                    |

## 成员方法

### TextRange.ChangeCase

更改指定文本的大小写。

**语法**

**express.ChangeCase ( *Type* )**

\*express是一个代表 TextRange 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型     | 说明                   |
|--------|-----------|--------------|------------------------|
| *Type* | 必选      | PpChangeCase | 指定更改大小写的方法。 |

**示例**

``` JavaScript
本示例将当前演示文稿的第一张幻灯片的标题设为大写。
Application.ActivePresentation.Slides.Item(1).Shapes.Title.TextFrame.TextRange.ChangeCase(ppCaseTitle)
```

### TextRange.Characters

返回一个 **TextRange** 对象，该对象代表指定的文本字符子集。有关计算或浏览一段文本中的字符的详细信息，请参阅 **TextRange** 对象。

**语法**

**express.Characters ( *Start* , *Length* )**

\*express是一个代表 TextRange 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                       |
|----------|-----------|----------|----------------------------|
| *Start*  | 可选      | Long     | 返回的范围中的第一个字符。 |
| *Length* | 可选      | Long     | 要返回的字符数。           |

**说明**

如果 **Start** 和 **Length** 都被忽略，则返回的范围从指定范围中的第一个字符开始，到最后一段结束。

如果指定 **Start** 但忽略 **Length** ，则返回的范围包含一个字符。

如果指定 **Length** 但忽略 **Start** ，则返回的范围从指定范围中的第一个字符开始。

如果 **Start** 大于指定文本中的字符数，则返回的范围从指定范围中的最后一个字符开始。

如果 **Length** 大于从文本的指定开始字符到末尾的字符数，则返回的范围包含所有这些字符。

**示例**

``` JavaScript
/*本示例设置当前演示文稿中第一张幻灯片的第二个形状的文本，并将第二个字符设为偏移 20% 的下标。*/
function test() {
  let myShape = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2)
  let charRange = myShape.TextFrame.TextRange.InsertBefore("H2O")
  charRange.Characters(2).Font.BaselineOffset = -0.2
}

/*本示例会将第一张幻灯片上第二个形状的所有下标字符设为加粗。*/
function test() {
  let Trng = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).TextFrame2.TextRange
  for(let i=1;i <= Trng.Characters().Count;i++) {
      let ft = Trng.Characters(i).Font
      if(ft.Subscript) {
          ft.Bold = true
      }
  }
}
```

### TextRange.Copy

将指定对象复制到剪贴板。

**语法**

**express.Copy ()**

\*express是一个代表 TextRange 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

``` JavaScript
/*本示例在第一个窗口中拷贝选定幻灯片到剪贴板然后将其粘贴到第二窗口的视图中。如果剪贴板的内容不能粘贴到第二个窗口的视图中（例如，如果试图将一个形状粘贴到幻灯片浏览视图中），则本示例失败。*/
function test() {
  Windows.Item(1).Selection.Copy()
  Windows.Item(2).View.Paste()
}

/*本示例将当前演示文稿中第一张幻灯片的第一个和第二个形状复制到剪贴板，然后将该副本粘贴到第二张灯片。*/
function test() {
  ActivePresentation.Slides.Item(1).Shapes.Range([1, 2]).Copy()
  ActivePresentation.Slides.Item(2).Shapes.Paste()
}

/*本示例将当前演示文稿的第一张幻灯片复制到剪贴板。*/
ActivePresentation.Slides.Item(1).Copy()

/*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板。*/
ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy()
```

### TextRange.Cut

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut ()**

\*express是一个代表 TextRange 对象的变量。

**示例**

``` JavaScript
/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中。*/
Windows.Item(1).Selection.Cut()

/*本示例从当前演示文稿的第一张幻灯片中删除第一个和第二个形状，并将它们的副本放到剪贴板，然后将该副本粘贴到第二张幻灯片。*/
function test() {
  ActivePresentation.Slides.Item(1).Shapes.Range([1, 2]).Cut()
  ActivePresentation.Slides.Item(2).Shapes.Paste()
}

/*本示例从当前演示文稿中删除第一张幻灯片并将它的一个副本放到剪贴板中。*/
ActivePresentation.Slides.Item(1).Cut()

/*本示例删除当前演示文稿中第一张幻灯片的第一个形状中的文本，并将它的一个副本放到剪贴板*/
ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Cut()
```

### TextRange.Delete

删除指定的对象。

**语法**

**express.Delete ()**

\*express是一个代表 TextRange 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

### TextRange.Find

在一个文本范围内查找指定的文本，并返回 TextRange 对象，该对象代表在其中找到指定的文本的第一个文本范围。如果找不到指定的文本，则返回 **Nothing** 。

**语法**

**express.Find ( *FindWhat* , *After* , *MatchCase* , *WholeWords* )**

\*express是一个代表 TextRange 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型    | 说明                                                                                                                                                                                            |
|--------------|-----------|-------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *FindWhat*   | 必选      | String      | 要搜索的文本。                                                                                                                                                                                  |
| *After*      | 可选      | Long        | 指定文本范围内的字符位置，将从该字符开始搜索下一处 FindWhat 匹配内容。例如，如果要从文本范围的第五个字符开始搜索，可指定 After 为 4。如果忽略该参数，则将文本范围的第一个字符作为搜索的起始点。 |
| *MatchCase*  | 可选      | MsoTriState | 如果此属性为 MsoTrue，则搜索时区分字符的大小写。                                                                                                                                                |
| *WholeWords* | 可选      | MsoTriState | 如果此属性为 MsoTrue，则搜索时仅查找整个词，而不搜索较长单词的部分字符。                                                                                                                        |

**示例**

``` JavaScript
本示例在当前演示文稿中查找所有“CompanyX”字符串，并将其格式设为加粗。
function test() {
  for(let i=1;i <= Application.ActivePresentation.Slides.Count;i++) {
      let sli = Application.ActivePresentation.Slides
      for(let j=1;j <= sli.Item(i).Shapes.Count;j++) {
          if(sli.Item(i).Shapes.Item(j).HasTextFrame) {
              let txtRng = sli.Item(i).Shapes.Item(j).TextFrame.TextRange
              let foundText = txtRng.Find("CompanyX")
              while(foundText != null) {
                  foundText.Font.Bold = true
                  let foundText = txtRng.Find("CompanyX", foundText.Start + foundText.Length-1)
              }
          }
      }
  }
}
```

### TextRange.InsertAfter

在指定文本范围的末尾添加一个字符串。返回一个代表添加文本的 **TextRange** 对象。省略参数时，该方法返回一个在指定范围末尾的长度为零的字符串。

**语法**

**express.InsertAfter ( *NewText* )**

\*express是一个代表 TextRange 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                 |
|-----------|-----------|----------|--------------------------------------|
| *NewText* | 可选      | String   | 要插入的文字。默认值为一个空字符串。 |

**示例**

``` JavaScript
/*本示例将字符串“: Test version”追加到活动演示文稿第一张幻灯片标题的末尾。*/
function test() {
  let myShape = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
  myShape.TextFrame.TextRange.InsertAfter(": Test version")
}

/*本示例将剪贴板的内容追加到第一张幻灯片标题的末尾。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.InsertAfter().Paste()
```

### TextRange.InsertBefore

在指定文本范围的开头添加一个字符串。返回一个代表添加文本的 **TextRange** 对象。省略参数时，该方法返回一个在指定范围末尾的长度为零的字符串。

**语法**

**express.InsertBefore ( *NewText* )**

\*express是一个代表 TextRange 对象的变量。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明                                 |
|-----------|-----------|----------|--------------------------------------|
| *NewText* | 可选      | String   | 要添加的文本。默认值为一个空字符串。 |

**示例**

``` JavaScript
/*本示例将字符串“Test version:”添加到活动演示文稿第一张幻灯片标题的开头。*/
function test() {
  let myShape = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)
  myShape.TextFrame.TextRange.InsertBefore("Test version: ")
}

/*本示例将剪贴板中的内容添加到活动演示文稿第一张幻灯片标题的开头。*/
Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.InsertBefore().Paste()
```

### TextRange.InsertDateTime

在指定的文本范围中插入日期和时间。返回一个代表所插入文本的 **TextRange** 对象。

**语法**

**express.InsertDateTime ( *DateTimeFormat* , *InsertAsField* )**

\*express是一个代表 TextRange 对象的变量。

**参数**

| 名称             | 必选/可选 | 数据类型         | 说明                                               |
|------------------|-----------|------------------|----------------------------------------------------|
| *DateTimeFormat* | 必选      | PpDateTimeFormat | 日期和时间的格式。                                 |
| *InsertAsField*  | 可选      | MsoTriState      | 决定是否每次打开演示文稿时都更新插入的日期和时间。 |

**说明**

|                                                         |
|---------------------------------------------------------|
| PpDateTimeFormat 可以是下列 PpDateTimeFormat 常量之一。 |
| **ppDateTimeddddMMMMddyyyy**                            |
| **ppDateTimedMMMMyyyy**                                 |
| **ppDateTimedMMMyy**                                    |
| **ppDateTimeFormatMixed**                               |
| **ppDateTimeHmm**                                       |
| **ppDateTimehmmAMPM**                                   |
| **ppDateTimeHmmss**                                     |
| **ppDateTimehmmssAMPM**                                 |
| **ppDateTimeMdyy**                                      |
| **ppDateTimeMMddyyHmm**                                 |
| **ppDateTimeMMddyyhmmAMPM**                             |
| **ppDateTimeMMMMdyyyy**                                 |
| **ppDateTimeMMMMyy**                                    |
| **ppDateTimeMMyy**                                      |

|                                                        |
|--------------------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。          |
| **msoCTrue**                                           |
| **msoFalse** 默认值。                                  |
| **msoTriStateMixed**                                   |
| **msoTriStateToggle**                                  |
| **msoTrue** 每次打开演示文稿时，更新插入的日期和时间。 |

**示例**

``` JavaScript
本示例将日期和时间插入在活动演示文稿第一张幻灯片第二个形状的第一段第一句之后。
function test() {
  let sh = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2)
  let sentOne = sh.TextFrame.TextRange.Paragraphs(1).Sentences(1)
  sentOne.InsertAfter().InsertDateTime(ppDateTimeMdyy)
}
```

### TextRange.InsertSlideNumber

将当前幻灯片的幻灯片编号插入指定的文本范围。返回一个代表幻灯片编号的 **TextRange** 对象。

**语法**

**express.InsertSlideNumber ()**

\*express是一个代表 TextRange 对象的变量。

**说明**

当前幻灯片的幻灯片编号改变时，已插入的幻灯片编号将自动更新。

**示例**

``` JavaScript
本示例将当前幻灯片的编号插入到活动演示文稿第一张幻灯片第二个形状的第一段第一句之后。
function test() {
  let sh = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2)
  let sentOne = sh.TextFrame.TextRange.Paragraphs(1).Sentences(1)
  sentOne.InsertAfter().InsertSlideNumber()
}
```

### TextRange.InsertSymbol

返回一个 TextRange 对象，该对象代表一个插入到指定文本范围的符号。

**语法**

**express.InsertSymbol ( *FontName* , *CharNumber* , *Unicode* )**

\*express是一个代表 TextRange 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型    | 说明                                                   |
|--------------|-----------|-------------|--------------------------------------------------------|
| *FontName*   | 必选      | String      | 字体名称。                                             |
| *CharNumber* | 必选      | Long        | Unicode 或 ASCII 字符编号。                            |
| *Unicode*    | 必选      | MsoTriState | 指定 CharNumber 参数代表 ASCII 字符还是 Unicode 字符。 |

**说明**

|                                                                    |
|--------------------------------------------------------------------|
| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。              |
| **msoCTrue** 不适用于此方法。                                      |
| **msoFalse** 默认值。 **CharNumber** 参数代表一个 ASCII 字符编号。 |
| **msoTriStateMixed** 不适用于此方法。                              |
| **msoTriStateToggle** 不适用于此方法。                             |
| **msoTrue** **CharNumber** 参数代表一个 Unicode 字符。             |

**示例**

``` JavaScript
本示例在活动演示文稿第一张幻灯片上的新文本框的第一段第一句后插入注册商标符号。
function test() {
    //Add text box
    let txtBox = Application.ActivePresentation.Slides.Item(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 100, 100)
    //Add symbol to text box
    txtBox.TextFrame.TextRange.InsertSymbol("Symbol", 226)
}
```

### TextRange.Lines

返回一个 TextRange 对象，该对象代表指定的文本行子集。有关计算或浏览文本范围中文本行的信息，请参阅 **TextRange** 对象。

**语法**

**express.Lines ( *Start* , *Length* )**

\*express是一个代表 TextRange 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                   |
|----------|-----------|----------|------------------------|
| *Start*  | 必选      | Long     | 返回的范围中的第一行。 |
| *Length* | 必选      | Long     | 要返回的行数。         |

**说明**

如果 **Start** 和 **Length** 都被忽略，则返回的范围从指定范围中的第一行开始，到最后一段结束。

如果指定 **Start** 但忽略 **Length** ，则返回的范围包含一行文本。

如果指定 **Length** 但忽略 **Start** ，则返回的范围从指定范围中的第一行开始。

如果 **Start** 大于指定文本中的行数，则返回的范围从指定范围中的最后一行开始。

如果 **Length** 大于从文本的指定起始行到末尾的行数，则返回的范围包含所有这些行。

**示例**

``` JavaScript
本示例将当前演文示稿第一张幻灯片第二个形状中第二段的前两行设为斜体。
Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).TextFrame.TextRange.Paragraphs(2).Lines(1, 2).Font.Italic = true
```

### TextRange.Paragraphs

返回一个 **TextRange** 对象，该对象代表文本段落的指定子集。有关统计或遍历文本范围内的段落的信息，请参阅 TextRange 对象。

**语法**

**express.Paragraphs ( *Start* , *Length* )**

\*express是一个代表 TextRange 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                   |
|----------|-----------|----------|------------------------|
| *Start*  | 可选      | Long     | 返回的范围中的第一段。 |
| *Length* | 可选      | Long     | 要返回的段落数。       |

**说明**

有关统计或遍历文本范围内的段落的信息，请参阅 TextRange 对象。

如果 *Start* 和 *Length* 都省略，则返回的范围从第一段开始，到指定范围的最后一段结束。

如果指定 *Start* 但忽略 *Length* ，则返回的范围包含一个段落。

如果指定 *Length* 但忽略 *Start* ，则返回的范围从指定范围中的第一段开始。

如果 *Start* 参数大于指定文本的段落数，则返回的范围从指定范围中的最后一段开始。

如果 *Length* 大于从指定起始段落到文字结尾的段落数，则返回的范围包含所有这些段落。

**示例**

``` JavaScript
以下示例将当前演文示稿第一张幻灯片第二个形状的第二段前两行设为倾斜。
Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).TextFrame.TextRange.Paragraphs(2).Lines(1, 2).Font.Italic = true
```

### TextRange.Paste

将剪贴板上的文本粘贴到指定的文本范围中，并返回一个代表粘贴文本的 **TextRange** 对象。

**语法**

**express.Paste ()**

\*express是一个代表 TextRange 对象的变量。

**说明**

将剪贴板内容粘贴到窗口中之前，请使用 **ViewType** 属性设置窗口视图。下表显示了可以粘贴到每个视图中的内容。

| 视图                   | 可以从剪贴板中粘贴以下内容                                                                                                                                                                                                                                                                                                |
|------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 幻灯片视图或备注页视图 | 形状、文本或整张幻灯片。如果从剪贴板粘贴一个幻灯片，该幻灯片的图像将作为嵌入对象被插入到幻灯片、母版或备注页中；如果选中形状，粘贴的文本将附加到形状文本之后；如果选中文本，粘贴的文本将替换选中的文本；如果选中任何其他对象，粘贴的文本将被放到它自己的文本框中。粘贴的形状将被放到 z 顺序的最上端且不会替换选中的形状。 |
| 大纲视图               | 文本或整张幻灯片。不能向大纲视图粘贴形状。粘贴的幻灯片将被插到光标所在的幻灯片之前。                                                                                                                                                                                                                                      |
| 幻灯片浏览视图         | 整张幻灯片。不能向幻灯片浏览视图粘贴形状或文本。粘贴的幻灯片将被插到光标处或演示文稿中最后选中的一张幻灯片之后。                                                                                                                                                                                                          |

### TextRange.Replace

在文本范围内查找特定文本，用指定的字符串替换查找到的文本，返回代表查找文本的第一处匹配内容的 **TextRange** 对象。如果找不到匹配的内容，则返回 **Nothing** 。

**语法**

**express.Replace ( *FindWhat* , *ReplaceWhat* , *After* , *MatchCase* , *WholeWords* )**

\*express是一个代表 TextRange 对象的变量。

**参数**

| 名称          | 必选/可选 | 数据类型    | 说明                                                                                                                                                                                |
|---------------|-----------|-------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *FindWhat*    | 必选      | String      | 要搜索的文本。                                                                                                                                                                      |
| *ReplaceWhat* | 必选      | String      | 用来替换查找到的文本的文本。                                                                                                                                                        |
| *After*       | 可选      | Integer     | 指定文本范围内开始搜索下一处 FindWhat 匹配内容的字符位置。例如，如果要从文本范围的第五个字符开始搜索，则可指定 After 为 4。如果省略此参数，则将文本范围的第一个字符作为搜索的起点。 |
| *MatchCase*   | 可选      | MsoTriState | 确定是否区分大小写。                                                                                                                                                                |
| *WholeWords*  | 可选      | MsoTriState | 确定是否只查找全字匹配的内容。                                                                                                                                                      |

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse** 默认值。                         |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 区分字符的大小写。                |

**示例**

``` JavaScript
本示例在活动文档内的所有形状中，查找每一处全字匹配“like”的内容，并将其替换为“NOT LIKE”。
function test() {
    let oSld = Application.ActivePresentation.Slides.Item(1)   
    for(let i=1;i <= oSld.Shapes.Count;i++) {
        let oTxtRng = oSld.Shapes.Item(i).TextFrame.TextRange
        let oTmpRng = oTxtRng.Replace("like","NOT LIKE",undefined,undefined,true)
        while(oTmpRng != null) {
            let oTxtRng = oTxtRng.Characters(oTmpRng.Start + oTmpRng.Length, oTxtRng.Length)
            let oTmpRng = oTxtRng.Replace("like","NOT LIKE",undefined,undefined,true)
        }
    }
}
```

### TextRange.Select

选择指定的对象。

**语法**

**express.Select ()**

\*express是一个代表 TextRange 对象的变量。

**说明**

如果做出不适用于当前视图的选择，此代码将失败。例如，可以在幻灯片浏览视图中而不能在幻灯片视图中选择幻灯片。

**示例**

``` JavaScript
/*本示例选择活动演示文稿中第一张幻灯片标题的前五个字符。*/
ActivePresentation.Slides.Item(1).Shapes.Title.TextFrame.TextRange.Characters(1, 5).Select()
/*本示例选择活动演示文稿中的第一张幻灯片。*/
ActivePresentation.Slides.Item(1).Select()
/*本示例选取一个已添加到演示文稿的新幻灯片中的表格。该表格具有三行和三列。*/
function test() {
  let sli = Presentations.Add().Slides
  sli.Add(1, ppLayoutBlank).Shapes.AddTable(3, 3).Select()
}
```

## 成员属性

### TextRange.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 TextRange 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test() {
  let shpOle = ActivePresentation.Slides.Item(1).Shapes
  for(let i = 1; i <= shpOle.Count; i++) {
      if(shpOle.Item(i).Type == msoLinkedOLEObject) {
          MsgBox(shpOle.Item(i).OLEFormat.Application.Name)
      }
  }
}
```

### TextRange.BoundHeight

以磅为单位返回指定文本框的文本 <span class="glossary">边界框</span> 的高度。只读 **Single** 类型。

**语法**

**express.BoundHeight**

\*express是一个代表 TextRange 对象的变量。

**示例**

``` JavaScript
本示例在活动演示文稿的第一张幻灯片中添加一个圆角矩形，该矩形与第一个形状的文本边界框的大小相同。
function test() {
  let shape = Application.ActivePresentation.Slides.Item(1).Shapes
  let tr = shape.Item(1).TextFrame.TextRange
  let roundRect = shape.AddShape(msoShapeRoundedRectangle, tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight)
  let fill = roundRect.Fill
  fill.ForeColor.RGB = 255, 0, 128
  fill.Transparency = 0.75
}
```

### TextRange.BoundLeft

以磅为单位返回指定文本框的文本 <span class="glossary">边界框</span> 的左边界到幻灯片的左边界的距离。只读。 **Single** 类型。

**语法**

**express.BoundLeft**

\*express是一个代表 TextRange 对象的变量。

**示例**

``` JavaScript
本示例在活动演示文稿的第一张幻灯片中添加一个圆角矩形，该矩形与第一个形状的文本边界框的大小相同。
function test() {
  let shape = Application.ActivePresentation.Slides.Item(1).Shapes
  let tr = shape.Item(1).TextFrame.TextRange
  let roundRect = shape.AddShape(msoShapeRoundedRectangle, tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight)
  let fill = roundRect.Fill
  fill.ForeColor.RGB = 255, 0, 128
  fill.Transparency = 0.75
}
```

### TextRange.BoundTop

以磅为单位返回指定文本框的文本 <span class="glossary">边界框</span> 的上边界到幻灯片的上边界的距离。只读。 **Single** 类型。

**语法**

**express.BoundTop**

\*express是一个代表 TextRange 对象的变量。

**示例**

``` JavaScript
本示例在活动演示文稿的第一张幻灯片中添加一个圆角矩形，该矩形与第一个形状的文本边界框的大小相同。
function test() {
  let shape = Application.ActivePresentation.Slides.Item(1).Shapes
  let tr = shape.Item(1).TextFrame.TextRange
  let roundRect = shape.AddShape(msoShapeRoundedRectangle, tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight)
  let fill = roundRect.Fill
  fill.ForeColor.RGB = 255, 0, 128
  fill.Transparency = 0.75
}
```

### TextRange.BoundWidth

以磅为单位返回指定文本框的文本 <span class="glossary">边界框</span> 的宽度。只读 **Single** 类型。

**语法**

**express.BoundWidth**

\*express是一个代表 TextRange 对象的变量。

**示例**

``` JavaScript
本示例在活动演示文稿的第一张幻灯片中添加一个圆角矩形，该矩形与第一个形状的文本边界框的大小相同。
function test() {
  let shape = Application.ActivePresentation.Slides.Item(1).Shapes
  let tr = shape.Item(1).TextFrame.TextRange
  let roundRect = shape.AddShape(msoShapeRoundedRectangle, tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight)
  let fill = roundRect.Fill
  fill.ForeColor.RGB = 255, 0, 128
  fill.Transparency = 0.75
}
```

### TextRange.Count

返回指定集合中的对象数目。

**语法**

**express.Count**

\*express是一个代表 TextRange 对象的变量。

**示例**

``` JavaScript
以下示例关闭除当前窗口外的所有窗口。
function test() {
  let win = Application.Windows
  for (let i = 2; i <= win.Count; i++) {
      win.Item(2).Close()
  }
}
```

### TextRange.Font

返回一个代表字符格式的 Font 对象。只读。

**语法**

**express.Font**

\*express是一个代表 TextRange 对象的变量。

**示例**

``` JavaScript
/*以下示例为当前演示文稿第一张幻灯片的第一个形状中的文本设置格式。*/
function test() {
  let shape = ActivePresentation.Slides.Item(1).Shapes.Item(1)
  let font = shape.TextFrame.TextRange.Font
  font.Size = 48
  font.Name = "Palatino"
  font.Bold = true
  font.Color.RGB = 55, 127, 255
}
/*以下示例设置第一张幻灯片第二个形状中项目符号的颜色和字体名称。*/
function test() {
  let shape = ActivePresentation.Slides.Item(1).Shapes.Item(2)
  let bullet = shape.TextFrame.TextRange.ParagraphFormat.Bullet
  bullet.Visible = true
  let font = bullet.Font
  font.Name = "Palatino"
  font.Color.RGB = 0, 0, 255
}
```

### TextRange.ParagraphFormat

返回一个 **ParagraphFormat** 对象，该对象代表指定文本的段落格式。只读。

**语法**

**express.ParagraphFormat**

\*express是一个代表 TextRange 对象的变量。

**示例**

``` JavaScript
以下示例设置当前演示文稿第一张幻灯片第二个形状中每段的段前、段内和段后行间距。
function test() {
  let shape = Application.ActivePresentation.Slides.Item(2).Shapes.Item(2)
  let format = shape.TextFrame.TextRange.ParagraphFormat
  format.LineRuleWithin = msoTrue
  format.SpaceWithin = 1.4
  format.LineRuleBefore = msoTrue
  format.SpaceBefore = 0.25
  format.LineRuleAfter = msoTrue
  format.SpaceAfter = 0.75
}
```

### TextRange.Text

返回或设置一个 **String** 类型值，该值代表指定对象中包含的文本。可读/写。

**语法**

**express.Text**

\*express是一个代表 TextRange 对象的变量。

**示例**

``` JavaScript
本示例设置当前演示文稿第一张幻灯片中标题的文本和字形。
function test() {
  let myPres = Application.ActivePresentation
  let range = myPres.Slides.Item(1).Shapes.Title.TextFrame.TextRange
  range.Text = "Welcome!"
  range.Font.Italic = true
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/TextRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# CommandBarControl 对象

## 简介

代表一个命令栏控件。 CommandBarControl 对象是 CommandBarControls 集合的成员。 CommandBarControl 对象与 CommandBarButton 、 CommandBarComboBox 和 CommandBarPopup 对象具有同样的属性和方法。

## 说明

在编写处理自定义命令栏控件的 Visual Basic 代码时，请使用 CommandBarButton 、 CommandBarComboBox 和 CommandBarPopup 对象。在编写用于处理容器应用程序中的内置控件的代码时，如果该控件不能用上述三个对象中的任意一个来代表，则可以使用 CommandBarControl 对象。使用 Controls ( *index* ) 可返回一个 CommandBarControl 对象，其中 *index* 是控件的索引号。（该控件的 Type 属性必须是 msoControlLabel 、 msoControlExpandingGrid 、 msoControlSplitExpandingGrid 、 msoControlGrid 或 msoControlGauge ）。声明为 CommandBarControl 的变量可赋予 CommandBarButton 、 CommandBarComboBox 和 CommandBarPopup 值

## 方法

| 序号 | 名称                                  | 说明                                                    |
|:----:|:--------------------------------------|:--------------------------------------------------------|
|  1   | [Copy](#CommandBarControl.Copy)       | 将一个命令栏控件复制到已有的命令栏中。                  |
|  2   | [Delete](#CommandBarControl.Delete)   | 将 CommandBarControl 对象从其集合中删除。               |
|  3   | [Execute](#CommandBarControl.Execute) | 运行分配给指定 CommandBarControl 控件的过程或内置命令。 |
|  4   | [Move](#CommandBarControl.Move)       | 将指定的 CommandBarControl 移动到已有的命令栏。         |

## 属性

| 序号 | 名称                                                  | 说明                                                                                                  |
|:----:|:------------------------------------------------------|:------------------------------------------------------------------------------------------------------|
|  1   | [Caption](#CommandBarControl.Caption)                 | 获取或设置命令栏控件的标题文字。可读写。                                                              |
|  2   | [DescriptionText](#CommandBarControl.DescriptionText) | 获取或设置命令栏控件的说明。可读写。                                                                  |
|  3   | [Enabled](#CommandBarControl.Enabled)                 | 获取或设置指定是否启用 CommandBarControl 的 Boolean 值。可读写。                                      |
|  4   | [Id](#CommandBarControl.Id)                           | 获取内置的 CommandBarControl 的 ID。只读。                                                            |
|  5   | [Tag](#CommandBarControl.Tag)                         | 获取或设置有关 CommandBarControl 的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。   |
|  6   | [Visibl](#CommandBarControl.Visibl)                   | 获取或设置 CommandBarControl 的 Visible 属性。如果 CommandBarControl 可见，则此属性为 true 。可读写。 |

## 成员方法

### CommandBarControl.Copy

将一个命令栏控件复制到已有的命令栏中。

**语法**

**express.Copy ( *Bar* , *Before* )**

\*express是一个代表 CommandBarControl 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                                                   |
|----------|-----------|----------|------------------------------------------------------------------------------------------------------------------------|
| *Bar*    | 可选      | Variant  | 一个代表目标命令栏的 CommandBar 对象。如果省略此参数，控件将被复制到它原来所在的命令栏中。                             |
| *Before* | 可选      | Variant  | 一个指示新控件在命令栏上的位置的数值。新控件将插入到位于此位置的控件之前。如果省略此参数，控件将被复制到命令栏的末尾。 |

### CommandBarControl.Delete

将 **CommandBarControl** 对象从其集合中删除。

**语法**

**express.Delete ( *Temporary* )**

\*express是一个代表 CommandBarControl 对象的变量。

**参数**

| 名称        | 必选/可选 | 数据类型 | 说明                                                                        |
|-------------|-----------|----------|-----------------------------------------------------------------------------|
| *Temporary* | 可选      | Variant  | 如果为 true，则从当前会话中删除此控件。应用程序在下次会话中将再次显示该控件 |

### CommandBarControl.Execute

运行分配给指定 **CommandBarControl** 控件的过程或内置命令。

**语法**

**express.Execute ()**

\*express是一个代表 CommandBarControl 对象的变量。

**示例**

``` JavaScript
/*
本 ET 示例创建一个命令栏，然后向其中添加内置命令栏按钮控件。该按钮将执行 ET AutoSum 函数。本示例使用 Execute 方法在显示该命令栏时计算选定单元格区域的总计。
*/

function test()
{
    let cmdBars = Application.ActiveDocument.CommandBars
    let cbrCustBar = cmdBars.Add("Custom")
    let ctlAutoSum = cbrCustBar.Controls.Add(Application.Enum.msoControlButton, cmdBars.Item("Standard").Controls.Item("AutoSum").Id)
    cbrCustBar.Visible = true 
    ctlAutoSum.Execute()
}
```

### CommandBarControl.Move

将指定的 **CommandBarControl** 移动到已有的命令栏。

**语法**

**express.Move ( *Bar* , *Before* )**

\*express是一个代表 CommandBarControl 对象的变量。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                                                                          |
|----------|-----------|----------|-----------------------------------------------------------------------------------------------|
| *Bar*    | 可选      | Variant  | 一个代表控件的目标命令栏的 Command 对象。如果忽略该参数，则控件将移动到当前所在命令栏的末端。 |
| *Before* | 可选      | Variant  | 表示控件位置的数字。控件将插到该位置的控件之前。如果忽略该参数，控件插入到同一命令栏。        |

**示例**

``` JavaScript
/*
本示例将名为 Custom 的命令栏上的第一个组合框控件移动到该命令栏上的第七个控件之前。本示例将 Tag 设置为“Selection box”并赋予控件较低的优先级，以便在一行容纳不下所有控件时将其隐藏。
*/
function test
{
    let allcontrols = Application.ActiveDocument.CommandBars.Item("Custom").Controls
    for(let i = 1; i <= allControls.Count; i++){
    if(allControls.Item(i).Type == Application.Enum.msoControlComboBox){
        let newallcontrols = allControls.Item(i)
            newallcontrols.Move(null, 7)
            newallcontrols.Tag = "Selection box"
            newallcontrols.Priority = 5
        break;
    }
}   
}
```

## 成员属性

### CommandBarControl.Caption

获取或设置命令栏控件的标题文字。可读写。

**语法**

**express.Caption**

\*express是一个代表 CommandBarControl 对象的变量。

**示例**

``` JavaScript
/*
本示例在自定义命令栏中添加一个带拼写检查按钮图标的命令栏控件，然后将其标题设置为“Spelling checker”。
*/

function test()
{
    let myBar = Application.ActiveDocument.CommandBars.Add("Custom", Application.Enum.msoBarTop, true)
    myBar.Visible = true 
    let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, 2)
    myControl.DescriptionText = "Starts the spelling checker"
    myControl.Caption = "Spelling checker"
}
```

### CommandBarControl.DescriptionText

获取或设置命令栏控件的说明。可读写。

**语法**

**express.DescriptionText**

\*express是一个代表 CommandBarControl 对象的变量。

**说明**

该说明不显示给用户，但有助于其他开发者将控件行为归档。

**示例**

``` JavaScript
/*
本示例在自定义命令栏中添加一个带控件行为说明的控件。
*/
function test()
{
    let commandbars = Application.ActiveDocument.CommandBars
    let myBar = commandbars.Add("Custom", Application.Enum.msoBarTop, true)
    myBar.Visible = true 
    let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, commandbars.Item("Standard").Controls.Item("Paste").Id)
    myControl.DescriptionText = "Pastes the contents of the Clipboard"
        myControl.Caption = "Paste"
}
```

### CommandBarControl.Enabled

获取或设置指定是否启用 **CommandBarControl** 的 **Boolean** 值。可读写。

**语法**

**express.Enabled**

\*express是一个代表 CommandBarControl 对象的变量。

**说明**

对于命令栏，如果将该属性设置为 **True** ，则该命令栏的名称将出现在可用命令栏列表中。

对于内置控件，如果将 **Enabled** 属性设置为 **True** ，则将由应用程序确定其状态。但如果将该属性设置为 **False** ，则强行禁用该控件。

### CommandBarControl.Id

获取内置的 **CommandBarControl** 的 ID。只读。

**语法**

**express.Id**

\*express是一个代表 CommandBarControl 对象的变量。

### CommandBarControl.Tag

获取或设置有关 **CommandBarControl** 的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。

**语法**

**express.Tag**

\*express是一个代表 CommandBarControl 对象的变量。

### CommandBarControl.Visibl

获取或设置 **CommandBarControl** 的 **Visible** 属性。如果 **CommandBarControl** 可见，则此属性为 **true** 。可读写。

**语法**

**express.Visibl**

\*express是一个代表 CommandBarControl 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/CommandBarControl](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# FileDialog 对象

## 简介

提供文件对话框，其功能与 WPS Office应用程序中标准的 “打开” 和 “保存” 对话框类似。

## 说明

使用 FileDialog 属性返回一个 FileDialog 对象。 FileDialog 属性位于每个单独 Office 应用程序的 Application 对象中。该属性使用一个参数 *DialogType* 确定该属性返回的 FileDialog 对象类型。 FileDialog 对象有四种类型：

- “打开” 对话框 - 允许用户选择一个或多个可以在宿主应用程序中使用 Execute 方法打开的文件。
- “另存为” 对话框 - 允许用户选择一个文件，然后可以使用 Execute 方法将当前文件另存为该文件。
- “文件选取器” 对话框 - 允许用户选择一个或多个文件。用户选择的文件路径将捕获到 FileDialogSelectedItems 集合中。
- “文件夹选取器” 对话框 - 允许用户选择一个路径。用户选择的路径将捕获到 FileDialogSelectedItems 集合中。

每个宿主应用程序只能创建一个 FileDialog 对象实例。因此，即使创建多个 FileDialog 对象， FileDialog 对象的很多属性也会保持不变。所以，在显示对话框之前请确保已经针对用途适当地设置了所有属性。

## 方法

| 序号 | 名称                           | 说明                                                                                                                                                                                                                                                    |
|:----:|:-------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Execute](#FileDialog.Execute) | 在调用 Show 方法后立即执行用户的操作。                                                                                                                                                                                                                  |
|  2   | [Show](#FileDialog.Show)       | 显示文件对话框并返回一个 Long 类型的值，指示用户按下的是 “操作” 按钮 (-1) 还是 “取消” 按钮 (0)。在调用 Show 方法时，在用户关闭文件对话框之前不会执行其他代码。在 “打开” 和 “另存为” 对话框中，在使用了 Show 方法后会立即使用 Execute 方法执行用户操作。 |

## 属性

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.AllowMultiSelect">AllowMultiSelect</a></td>
<td style="text-align: left;" data-valian="middle"><p>如果允许用户从文件对话框中选择多个文件，则为 True 。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.Application">Application</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取一个 Application 对象，代表 FileDialog 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.ButtonName">ButtonName</a></td>
<td style="text-align: left;" data-valian="middle"><p>设置或获取代表文件对话框中动作按钮上所显示文本的 String 类型的值。可读/写。</p>
<table>
<thead>
<tr class="header">
<th>注释</th>
</tr>
</thead>
<tbody>
<tr class="odd">
<td>WPS Office Assistant 在 WPS Office 2015 版中已被淘汰。</td>
</tr>
</tbody>
</table></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.Creator">Creator</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取一个 32 位整数，指示创建 FileDialog 对象时所使用的应用程序。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.DialogType">DialogType</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回一个 MsoFileDialogType 常量，代表 FileDialog 对象被设置为要显示的文件对话框的类型。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.FilterIndex">FilterIndex</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取或设置一个 Long 类型的值，指示文件对话框的默认文件筛选器。默认筛选器决定首次打开文件对话框时显示的文件类型。可读/写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.Filters">Filters</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取一个 FileDialogFilters 集合。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">8</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.InitialFileName">InitialFileName</a></td>
<td style="text-align: left;" data-valian="middle"><p>设置或返回一个 String 类型的值，代表文件对话框中初始显示的路径或文件名。可读/写。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">9</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.InitialView">InitialView</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取或设置一个 MsoFileDialogView 常量，代表文件对话框中文件和文件夹的初始表示形式。可读/写。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">10</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.Item">Item</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取与对象关联的文本。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">11</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.Parent">Parent</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取 FileDialog 对象的 Parent 对象。只读。</p></td>
</tr>
<tr class="even" data-editname="properties">
<td style="text-align: center;" data-valian="middle">12</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.SelectedItems">SelectedItems</a></td>
<td style="text-align: left;" data-valian="middle"><p>获取一个 FileDialogSelectedItems 集合。此集合包含用户在使用 FileDialog 对象的 Show 方法显示的文件对话框中所选的文件的路径列表。只读。</p></td>
</tr>
<tr class="odd" data-editname="properties">
<td style="text-align: center;" data-valian="middle">13</td>
<td style="text-align: left;" data-valian="middle"><a href="#FileDialog.Title">Title</a></td>
<td style="text-align: left;" data-valian="middle"><p>设置或获取使用 FileDialog 对象显示的文件对话框的标题。可读/写。</p></td>
</tr>
</tbody>
</table>

## 成员方法

### FileDialog.Execute

在调用 **Show** 方法后立即执行用户的操作。

**语法**

**express.Execute ()**

\*express是一个代表 FileDialog 对象的变量。

### FileDialog.Show

显示文件对话框并返回一个 **Long** 类型的值，指示用户按下的是 **“操作”** 按钮 (-1) 还是 **“取消”** 按钮 (0)。在调用 **Show** 方法时，在用户关闭文件对话框之前不会执行其他代码。在 **“打开”** 和 **“另存为”** 对话框中，在使用了 **Show** 方法后会立即使用 **Execute** 方法执行用户操作。

**语法**

**express.Show ()**

\*express是一个代表 FileDialog 对象的变量。

**示例**

## 成员属性

### FileDialog.AllowMultiSelect

如果允许用户从文件对话框中选择多个文件，则为 **True** 。可读/写。

**语法**

**express.AllowMultiSelect**

\*express是一个代表 FileDialog 对象的变量。

**说明**

此属性对 **“文件夹选取器”** 对话框或 **“另存为”** 对话框无效，因为用户永远无法从这些类型的文件对话框中选择多个文件。

### FileDialog.Application

获取一个 **Application** 对象，代表 **FileDialog** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 FileDialog 对象的变量。

### FileDialog.ButtonName

设置或获取代表文件对话框中动作按钮上所显示文本的 **String** 类型的值。可读/写。

| 注释                                                   |
|--------------------------------------------------------|
| WPS Office Assistant 在 WPS Office 2015 版中已被淘汰。 |

**语法**

**express.ButtonName**

\*express是一个代表 FileDialog 对象的变量。

**说明**

默认情况下，此属性设置为该文件对话框类型的标准文本。例如，在 **“打开”** 对话框中，该属性默认设置为“打开”。此字符串所允许的最大长度为 51 个字符。

### FileDialog.Creator

获取一个 32 位整数，指示创建 **FileDialog** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 FileDialog 对象的变量。

### FileDialog.DialogType

返回一个 **MsoFileDialogType** 常量，代表 **FileDialog** 对象被设置为要显示的文件对话框的类型。只读。

**语法**

**express.DialogType**

\*express是一个代表 FileDialog 对象的变量。

**示例**

下面的示例将 **FileDialog** 对象视为未知类型，并且如果是 **“另存为”** 对话框或 **“打开”** 对话框，则运行 **Execute** 方法。

``` JavaScript
function test(fd){
    // If the user presses the action button...
    if(fd.Show == -1){
        // Use the DialogType property to determine whether to
        // use the Execute method.
        switch(fd.DialogType){
            case Application.Enum.msoFileDialogOpen:
            case Application.Enum.msoFileDialogSaveAs:
                fd.Execute()
                break
        }

        // If the user presses Cancel...
    }
}
```

### FileDialog.FilterIndex

获取或设置一个 **Long** 类型的值，指示文件对话框的默认文件筛选器。默认筛选器决定首次打开文件对话框时显示的文件类型。可读/写。

**语法**

**express.FilterIndex**

\*express是一个代表 FileDialog 对象的变量。

**说明**

如果试图将该属性设置为大于筛选数量的数字，将选定最后一个可用筛选。

**示例**

下面的示例使用 **FileDialog** 对象显示一个 **“File Picker”** 对话框，并在消息框中显示每个选定的文件。此示例还演示如何添加新文件筛选器并使其成为默认筛选器。

``` JavaScript
function test() {

    //Declare a variable as a FileDialogFilters collection.

    //Declare a variable as a FileDialogFilter object.

    //Set the FileDialogFilters collection variable to
    //the FileDialogFilters collection of the SaveAs dialog box.
    let fdfs = Application.FileDialog(Application.Enum.msoFileDialogFilePicker).Filters

    //Iterate through the description and extensions of each
    //default filter in the SaveAs dialog box.
    for(let i = 1;i <= fdfs.Count;i++) {
        let fdf = fdfs.Item(i)
        //Display the description of filters that include
        //ET files
        if(fdf.Extensions.indexOf("xls") > -1) {
            alert("Description of filter: " + fdf.Description)
        }
    }

}
```

### FileDialog.Filters

获取一个 **FileDialogFilters** 集合。只读。

**语法**

**express.Filters**

\*express是一个代表 FileDialog 对象的变量。

### FileDialog.InitialFileName

设置或返回一个 **String** 类型的值，代表文件对话框中初始显示的路径或文件名。可读/写。

**语法**

**express.InitialFileName**

\*express是一个代表 FileDialog 对象的变量。

**说明**

在指定文件名时可以使用 **'\*'** 和 **'?'** 通配符，但是指定路径时不能使用这些通配符。 **'\*'** 符号代表任意数量的连续字符，而 **'?'** 代表单个字符。例如， **.InitialFileName = "c:\c\*s.txt"** 将返回“charts.txt”和“checkregister.txt”。

如果指定了路径而没有指定文件名，则对话框中将显示文件筛选器所允许的所有文件。

如果指定了位于初始文件夹中的某个文件，则对话框中只显示该文件。

如果指定了初始文件夹中不存在的某个文件名，则对话框中将不包含文件。在 **InitialFileName** 属性中指定的文件类型将覆盖文件筛选器的设置。

如果指定了无效路径，则使用上次使用的路径。如果使用无效路径，则会向用户显示警告消息。

将此属性设置为长度大于 256 个字符的字符串将导致运行时错误。

### FileDialog.InitialView

获取或设置一个 **MsoFileDialogView** 常量，代表文件对话框中文件和文件夹的初始表示形式。可读/写。

**语法**

**express.InitialView**

\*express是一个代表 FileDialog 对象的变量。

### FileDialog.Item

获取与对象关联的文本。只读。

**语法**

**express.Item**

\*express是一个代表 FileDialog 对象的变量。

### FileDialog.Parent

获取 **FileDialog** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

\*express是一个代表 FileDialog 对象的变量。

**说明**

### FileDialog.SelectedItems

获取一个 **FileDialogSelectedItems** 集合。此集合包含用户在使用 **FileDialog** 对象的 **Show** 方法显示的文件对话框中所选的文件的路径列表。只读。

**语法**

**express.SelectedItems**

\*express是一个代表 FileDialog 对象的变量。

### FileDialog.Title

设置或获取使用 **FileDialog** 对象显示的文件对话框的标题。可读/写。

**语法**

**express.Title**

\*express是一个代表 FileDialog 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/FileDialog](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
# SmartDocument 对象

## 简介

WPS 中的 Document 对象和 ET 中的 Workbook 对象的 SmartDocument 属性返回一个 SmartDocument 对象。

## 说明

使用 SmartDocument 对象可管理附加到活动文档上的 XML 扩展包。

使用 SmartDocument 对象的 SolutionID 和 SolutionURI 属性可检索有关附加到活动文档或工作簿上的 XML 扩展包的信息。使用 PickSolution 方法可允许用户从列表中选择可用的 XML 扩展包以便附加到活动文档或工作簿中。使用 RefreshPane 方法可刷新智能文档的 “文档操作” 任务窗格。

无论文档是否附加了 XML 扩展包， SmartDocument 对象模型均可用。当活动文档没有附加 XML 扩展包时， Document 或 Workbook 对象的 SmartDocument 属性不返回 Nothing 。检查 SolutionID 属性可确定活动文档是否附加了 XML 扩展包。

## 方法

| 序号 | 名称                                        | 说明                                                                                       |
|:----:|:--------------------------------------------|:-------------------------------------------------------------------------------------------|
|  1   | [PickSolution](#SmartDocument.PickSolution) | 显示一个对话框，允许用户选择可用的 XML 扩展包，以便将其附加到 WPS 活动文档或 ET 工作簿中。 |
|  2   | [RefreshPane](#SmartDocument.RefreshPane)   | 为 WPS 中的活动文档或 ET 中的工作簿刷新 “文档操作” 任务窗格。                              |

## 属性

| 序号 | 名称                                      | 说明                                                                                                                                 |
|:----:|:------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#SmartDocument.Application) | 获取一个 Application 对象，代表 SmartDocument 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Creator](#SmartDocument.Creator)         | 获取一个 32 位整数，指示创建 SmartDocument 对象时所使用的应用程序。只读。                                                            |
|  3   | [SolutionID](#SmartDocument.SolutionID)   | 获取或设置 ID，通常为全局唯一标识符 (GUID)，它标识附加在活动的 WPS 文档或者 ET 工作簿上的 XML 扩展包。可读写。                       |
|  4   | [SolutionURL](#SmartDocument.SolutionURL) | 获取或设置一个绝对 URL，该 URL 提供指向附加到活动的 WPS 文档或者 ET 工作簿上的 XML 扩展包文件的完整路径。可读写。                    |

## 成员方法

### SmartDocument.PickSolution

显示一个对话框，允许用户选择可用的 XML 扩展包，以便将其附加到 WPS 活动文档或 ET 工作簿中。

**语法**

**express.PickSolution ( *ConsiderAllSchemas* )**

\*express是一个代表 SmartDocument 对象的变量。

**参数**

| 名称                 | 必选/可选 | 数据类型 | 说明                                                                                                                              |
|----------------------|-----------|----------|-----------------------------------------------------------------------------------------------------------------------------------|
| *ConsiderAllSchemas* | 必选      | Boolean  | 如果为 True，则会显示用户计算机上安装的所有可用 XML 扩展包。如果为 False，则会仅显示可用于活动文档的 XML 扩展包。默认值为 False。 |

**说明**

使用 **PickSolution** 方法可允许用户从列表中选择 XML 扩展包。附加在活动文档或工作簿上的架构将确定可用的 XML 扩展包。

**PickSolution** 方法不会返回一个值以表明用户是选择了 XML 扩展包还是单击了对话框中的 **Cancel** 。在调用 **SolutionID** 后，检查 **PickSolution** 属性可确定是否已附加了 XML 扩展包。

如果智能文档的开发人员未能在 XML 扩展包清单文件中指定“targetApplication”，则 **PickSolution** 所显示的列表中可能会包含未面向活动应用程序的 XML 扩展包。例如，ET 用户可能会看到专门面向 WPS 的 XML 扩展包。在这些情况下，用户可能会选择不适合活动应用程序的 XML 扩展包。

关于智能文档或智能文档 XML 扩展包的详细信息，请参阅WPS开发网站上的 Smart Document Software Development Kit (SDK)。

**示例**

``` JavaScript
/*下面的示例将检查 SolutionID 属性以确定活动 WPS 文档是否已有附加的 XML 扩展包。如果没有，它会显示一个对话框，允许用户选择可用的 XML 扩展包，然后显示智能文档的属性。*/
function test(){
    var objSmartDoc = Application.ActiveWorkbook.SmartDocument;
    if(objSmartDoc.SolutionID == "None" || objSmartDoc.SolutionID == "")
        objSmartDoc.PickSolution(true);
    if(objSmartDoc.SolutionID > "None" && objSmartDoc.SolutionID > "") 
    {
        　　　　let strSmartDocInfo = "SolutionID: " + objSmartDoc.SolutionID + "\r\n" + "SolutionURL: " + objSmartDoc.SolutionURL;
        　　　　alert(strSmartDocInfo + "Smart Doc Properties");
    }
    else 
        　　　　alert("The user clicked Cancel.");
}
```

### SmartDocument.RefreshPane

为 WPS 中的活动文档或 ET 中的工作簿刷新 **“文档操作”** 任务窗格。

**语法**

**express.RefreshPane ()**

\*express是一个代表 SmartDocument 对象的变量。

**说明**

如果活动文档没有附加 XML 扩展包， **RefreshPane** 方法就会产生错误。

**示例**

``` JavaScript
/*下面的示例确定活动的 ET 工作簿是否附加了 XML 扩展包。如果是，它将刷新智能文档的“文档操作”任务窗格。*/
function test(){
    var objSmartDoc = Application.ActiveWorkbook.SmartDocument;
    if(objSmartDoc.SolutionID > "None") 
        objSmartDoc.RefreshPane;
    else
        alert("No XML expansion pack attached.");
}
```

## 成员属性

### SmartDocument.Application

获取一个 **Application** 对象，代表 **SmartDocument** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 SmartDocument 对象的变量。

### SmartDocument.Creator

获取一个 32 位整数，指示创建 **SmartDocument** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 SmartDocument 对象的变量。

### SmartDocument.SolutionID

获取或设置 ID，通常为全局唯一标识符 (GUID)，它标识附加在活动的 WPS 文档或者 ET 工作簿上的 XML 扩展包。可读写。

**语法**

**express.SolutionID**

\*express是一个代表 SmartDocument 对象的变量。

**说明**

当活动文档没有附加 XML 扩展包时， **SolutionID** 属性返回空字符串或“None”。

通过为 **SolutionID** 和 **SolutionURL** 属性提供适当的值以将可用的 XML 扩展包附加到活动文档上，可以在不使用 **PickSolution** 方法的情况下将该文档转换为智能文档。通过将 **SolutionID** 和 **SolutionUrl** 属性设置为空字符串，可以删除附加的 XML 扩展包。

**示例**

``` JavaScript
/*下面的示例将通过检查 SolutionID 属性来确定是否在活动 ET 工作簿上附加了 XML 扩展包。*/
function test(){
    var objSmartDoc = Application.ActiveWorkbook.SmartDocument;
    if(objSmartDoc.SolutionID == "None" || objSmartDoc.SolutionID == "")
        alert("No XML expansion pack attached.");
    else 
        alert("Smart document Solution ID: " + objSmartDoc.SolutionID);
}
```

### SmartDocument.SolutionURL

获取或设置一个绝对 URL，该 URL 提供指向附加到活动的 WPS 文档或者 ET 工作簿上的 XML 扩展包文件的完整路径。可读写。

**语法**

**express.SolutionURL**

\*express是一个代表 SmartDocument 对象的变量。

**说明**

当没有在活动文档上附加 XML 扩展包时， **SolutionUrl** 属性返回一个空字符串。

通过为 **SolutionID** 和 **SolutionUrl** 属性提供适当的值以将可用的 XML 扩展包附加到活动文档上，可以在不使用 **PickSolution** 方法的情况下将该文档转换为智能文档。通过将 **SolutionID** 和 **SolutionUrl** 属性设置为空字符串，可以删除附加的 XML 扩展包。

**示例**

``` JavaScript
/*下面的示例先确定是否在活动 WPS 文档上附加了 XML 扩展包，然后显示智能文档的解析 URL。*/
function test(){
    var objSmartDoc = Application.ActiveWorkbook.SmartDocument;
    if(objSmartDoc.SolutionID == "None" || objSmartDoc.SolutionID == "")
        alert("No XML expansion pack attached.");
    else 
        alert("Smart document Solution URL: " + objSmartDoc.SolutionURL);
}
```

> WPS 开发文档： [WPS 基础接口/通用 API 参考/SmartDocument](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
