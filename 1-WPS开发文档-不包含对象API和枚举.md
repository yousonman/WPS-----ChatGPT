# 加载项概述

WPS 加载项是一套基于 Web 技术用来扩展 WPS 应用程序的解决方案。每个 WPS 加载项都对应打开了一个网页，并通过调用网页中 JavaScript 方法来完成其功能逻辑。 WPS 加载项打开的网页可以直接与 WPS 应用程序进行交互，同时一个 WPS 加载项中的多个网页形成了一个整体， 相互之间可以进行数据共享。 开发者不必关注浏览器兼容的问题，因为 WPS 加载项的底层是以 Chromium 开源浏览器项目为基础进行的优化扩展。 WPS 加载项具备快速开发、轻量化、跨平台的特性，目前已针对Windows/Linux操作系统进行适配。 WPS 加载项功能特点如下:

- 完整的功能。可通过多种不同的方法对文档、电子表格和演示文稿进行创作、格式设置和操控；通过鼠标、键盘执行的操作几乎都能通过WPS 加载项 完成；可以轻松地执行重复任务，实现自动化。
- 三种交互方式。 自定义功能区 ，采用公开的CustomUI标准，快速组织所有功能； 任务窗格 ，展示网页，内容更丰富； Web 对话框 ，结合事件监听，实现自由交互。
- 标准化集成。不影响 JavaScript 语言特性，网页运行效果和在浏览器中完全一致；WPS 加载项开发文档完整，接口设计符合 JavaScript 语法规范，避免不必要的学习成本，缩短开发周期。

**WPS 加载项介绍** 了解有关WPS 加载项的基本概念

- 自定义功能区
- 任务窗格
- 使用 Web 对话框
- WPS 加载项可用性
- 使用范例：OA 助手

**WPS 加载项开发** 快速学习开发技能

- 生成首个 WPS 加载项
- WPS 加载项开发说明
- WPS 加载项发布说明
- 事件
- 加载项数据
- WPS 加载项性能

更多核心能力，敬请期待！

有任何疑问，请加QQ：3253920855 获取帮助。

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/加载项概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/%E5%8A%A0%E8%BD%BD%E9%A1%B9%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# 生成首个 WPS 加载项

在本教程中，将创建一个 WPS 加载项，该加载项将：

- 设计自定义功能区
- 打开对话框
- 创建自定义任务窗格

## 准备开发环境

- 安装wps
- 安装Node.js
- 安装代码编辑器 Visual Studio Code

## 新建 WPS 加载项

1、 **管理员权限**(如果安装的是wps个人版，不需要管理员权限) 启动命令行，通过npm全局安装 **wpsjs** 开发工具包: 安装命令： **npm install -g wpsjs** , 如果之前已经安装了，可以检查下wpsjs版本，更新wpsjs的命令为: **npm update -g wpsjs**

![](Base64图像/Base64图像1来自_WPS%20集成模式_WPS%20加载项开发_生成首个%20WPS%20加载项.png)

2、新建一个wps加载项，假设这个wps加载项取名为"HelloWps"。 输入命令： **wpsjs create HelloWps** , 会出现如下图的几个选项:

![](Base64图像/Base64图像2来自_WPS%20集成模式_WPS%20加载项开发_生成首个%20WPS%20加载项.png)

通过上下方向键可以选择要创建的wps加载项的类型，如果选择“文字”，则创建的加载项会在wps文字程序中加载并运行， 同理选择“电子表格”，则会在wps表格中运行，这里假设我们选择的是“文字”，按Enter健确定。

![](Base64图像/Base64图像3来自_WPS%20集成模式_WPS%20加载项开发_生成首个%20WPS%20加载项.png)

3、选择示例代码的代码风格类型 wpsjs工具包提供了两种不同代码风格的示例，“无”代表示例代码中都是原生的js及html代码，没有集成vue\react等流行的前端框架。 "Vue"代表生成的示例代码集成了Vue相关的脚手架，在实际的项目中选用Vue基于示例代码可能更适合做工程化的开发，感兴趣的同学可以两种都尝试一下。 这里我们选择“无”，按Enter健确认。 确认后wpsjs工具包会在当前目录下生成一个HelloWps的文件夹，我们进入到此文件夹，可以看到HelloWps的相关代码已经生成：

![](Base64图像/Base64图像4来自_WPS%20集成模式_WPS%20加载项开发_生成首个%20WPS%20加载项.png)

4、开始调试并愉快的写代码 执行命令： **wpsjs debug** 执行此命令后即可开始调试，wpsjs工具包会自动启动wps并加载HelloWps这个加载项，同时wpsjs工具包启了一个http服务,此服务主要提供两方面的能力: a.提供前端页的的热更新服务，wpsjs工具包检测到网页数据变化时，自动刷新页面。 b.提供wps加载项的在线服务，wpsjs生成的代码示例是一个在线模式，wps客户端程序实际上是通过http服务来请求在线的wps加载项相关代码和资源的。 最后，可以用visual studio code打开示例代码，开始愉快的写代码了。

![](Base64图像/Base64图像5来自_WPS%20集成模式_WPS%20加载项开发_生成首个%20WPS%20加载项.png)

![](Base64图像/Base64图像6来自_WPS%20集成模式_WPS%20加载项开发_生成首个%20WPS%20加载项.png)

备注：wpsjs工具包为示例代码中有一个package.json文件，这是node工具标准的配置文件，其中有一个依赖包为wps-jsapi, 这个依赖包是wps支持的全部接口的TypeScript描述，方便在vscode中敲代码时，提供代码联想功能，由于wps接口会跟随业务 需求不断更新，因此当发现代码联想对于有些接口不支持时，通过 **npm update --save-dev wps-jsapi** 命令定期更新这个包。

以上展示了如何开始第一个wps加载项的制作，在实际用况中，如果需要让企业的业务系统与wps加载项进行集成，可以参考以下示例: [OA助手示例](https://code.aliyun.com/zouyingfeng/wps/tree/master/oaassist)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/生成首个 WPS 加载项](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/%E7%94%9F%E6%88%90%E9%A6%96%E4%B8%AA%20WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9.html)

------------------------------------------------------------------------
# WPS 加载项开发说明

## WPS 加载项结构

WPS 加载项由 自定义功能区 和网页两部分组成。自定义功能区只需要一个配置文件，对应 WPS 加载项目录中的ribbon.xml文件； 网页部分负责执行自定义功能区对应的逻辑功能。因为不需要显示网页，所以省略了 HTML 文件，并用main.js来引入所有的外部 JavaScript 文件； 在这些 JavaScript 文件中通常包含了一系列用 JavaScript 实现的函数，这些函数与自定义功能区的功能一一对应，我们称之为接口函数。

## 启动流程

WPS 加载项启动时，首先在 WPS 加载项对应文件夹中自动创建index.html网页并打开，index.html从当前路径引入main.js，从而能够在接下来的过程中执行接口函数。 当网页打开成功之后，开始解析ribbon.xml生成自定义功能区，解析过程中会调用若干次接口函数，最终完成加载。 **注意** ，开发者应当避免在该目录下创建index.html。

## WPS 加载项 API 使用

WPS 加载项 API 通过对 JavaScript 功能进行的扩展，实现了网页与 WPS 应用程序交互的能力。这些 API 被集中在window.wps对象下，而我们在开发中通常会省略掉window，直接以wps开始。

## 调试

WPS 加载项调试是对其中的一个网页单独进行的调试。调试时会弹出一个独立调试器对话框，除此之外和网页调试基本一致。 可以在调试器的 Console 中直接查看任意的 API 属性和调用 API 方法。调试自动生成的index.html网页，使用快捷键 ALT + F12。 注意调试过程中需要先关闭alert或其它同步弹框，才能继续向下调试。

![](服务器端图像/debuger.png)

## 系统集成

用户可以在自己的浏览器中调用 WPS 加载项的 JavaScript 方法。 wps_sdk.js 对调用进行了封装，让开发者可以快速调用， wps_sdk.js对Chrome、Edge、IE8及IE8以上浏览器进行了支持。其它方式的集成请参考wps_sdk.js实现。 点击 <span class="showLink">这里</span> 试一下吧。

``` JavaScript
WpsInvoke.InvokeAsHttp(WpsInvoke.ClientType.wps, 'JsDemo', 'OnbtnShowDialogClick', {},  
   function (res) {
      if (res.status == 0)
    alert('finish')
      else
    alert(res.message)
})
```

**接口定义**

**WpsInvoke.InvokeAsHttp** ( *type*, *name*, *func*, *params*, *callBack* )

**参数**

其中 WpsInvoke 是wps_sdk.js封装的对象，InvokeAsHttp是启动 WPS 应用程序的接口。

- *type* ： WPS 应用程序的类型，类型的定义在ClientType中，ClientType.wps代表文字、ClientType.et代表表格、ClientType.wpp代表演示。
- *name* ： WPS 加载项名称。
- *func* ： 执行的 JavaScript 方法。
- *params* ： 传递给方法func的参数。
- *callBack* ： WpsInvoke.InvokeAsHttp执行的回调函数。

**说明**

启动 WPS 应用程序需要用户在浏览器点击允许启动 WPS Office。WpsInvoke.InvokeAsHttp执行是异步的，调用后立刻返回。 等到WpsInvoke.InvokeAsHttp执行完成后，执行callBack回调函数，并给回调函数传参

``` JavaScript
{
    status: 0,  //返回状态。0 代表成功；1 代表上次请求没有完成；2 代表没有允许执行。
    message: "" //返回状态描述信息
}
```

## 发布部署

加载项开发完成后，很多开发者可能会有这样的一些问题：加载项如何部署呢、用户如何访问部署后的加载项呢、用户需要安装什么、用户是否需要去手动修改配置呢

目前我们提供两种部署方式，jsplugins.xml模式和publish.xml模式，两种都不需要用户去手动配置什么。只需要本地安装好相应版本的WPS就行。

WPS可以同时支持这两种模式：

### publish模式

#### 模式介绍

publish模式是通过wpsjs工具包的wpsjs publish命令打包，将生成的文件夹下的所有文件部署到打包时填写服务器地址去。告知用户publish.html地址，业务系统开发商可将publish.html的功能按需整合到自己的页面中，便于做基础环境监测。也可以复用此页面给到用户，用户可自己控制启用和禁用哪些加载项。

#### 部署

- 使用wpsjs包的wpsjs publish命令进行打包
- 将目录wps-addon-build下的文件署到服务器
- 将wps-addon-publish下的publish.html文件部署到服务器上，一般与加载项分开部署。
- 告知用户publish.html文件地址。 [wpsjs工具包使用](https://kdocs.cn/l/cASCu9B0G)

#### 加载项加载流程

- 用户在浏览器中打开publish.html文件
- 校验加载项是否正常
  - 在线模式：去请求地址+/ribbon.xml和 地址+/index.html
  - 离线模式：校验改地址是否能够访问到压缩包
- 点击安装或卸载的加载项
- 本地自动生成publish.xml文件
  - window：%appdata%/kingsoft/wps/jsaddons
  - linux: ~/.local/share/Kingsoft/wps/jsaddons
- 启动WPS
- 读取本地publish.xml文件
- 加载对应组件的所有加载项
- 根据业务系统指定的加载项名称，使用该加载项来接收参数
- 业务开发方可将此页面的方法按需整合到自己的需要调起WPS的业务场景中，从而达到自动化的环境配置。

#### 适用场景

- 只交付WPS基础包，无需二次打包
- 实现集成松耦合
- 便于业务开发方按需定义集成场景

#### 版本支持情况

- Windows：企业版20200425分支之后版本
- Linux：企业版20200530分支之后版本

### jsplugins.xml模式

#### 模式介绍

jsplugins.xml模式是通过设置oem.ini配置文件的JSPluginsServer的值为加载项管理文件jsplugins.xml来控制加载项的加载（相当于WPS加载项列表文件），二次打包时，业务开发商需要告知我们JSPluginsServer的配置地址，，将起配置到oem.ini文件中，业务开发商再做安装包分发。

后续的加载项的控制用，业务开发商可以自由的更改jsplugins.xml文件，实现加载项的新增，修改。

[jsplugins.xml文件的配置](https://kdocs.cn/l/cBk8tsBIf)

#### 部署

- 使用wpsjs包的wpsjs build进行加载项打包
- 将目录wps-addon-build下的文件署到服务器
- 配置jsplugins.xml文件
- 部署jsplugins.xml文件，一般和加载项项目分开部署
- 告知WPS获取jsplugins.xml的网络地址，即JSPluginsServer。（只需要打一次包，后续更改jsplugins.xml的内容不需要再次打包）[wpsjs工具包使用](https://kdocs.cn/l/cASCu9B0G)

#### 加载项加载流程

- 启动WPS
- 读取oem.ini，去下载JSPluginsServer对应地址的jsplugins.xml文件
  - window：%appdata%/kingsoft/wps/jsaddons
  - linux: ~/.local/share/Kingsoft/wps/jsaddons
- 读取jsplugins.xml文件
- 加载对应组件的所有加载项
- 根据业务系统指定的加载项名称，使用该加载项来接收参数

#### 适用场景

- 本来就需要做二次打包的项目

#### 版本支持情况

- Windows：2019版本，2020发的版本
- Linux：2019版本，2020年3月后发的版本

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/WPS 加载项开发说明](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91%E8%AF%B4%E6%98%8E.html)

------------------------------------------------------------------------
# WPS加载项 集成 业务系统

通过使用wpsjsrpcsdk.js提供的接口，开发者可以很容易的构筑一个web页面；用户通过点击web前端页面上的按钮，让WPS加载项执行一系列操作

## wpsjsrpcsdk.js提供的能力

wpsjsrpcsdk.js用于wpsjs加载项，web端前端页面的编写，目前提供三种模式的支持：

- 单进程模式 ：用户窗口下仅启动一个wps客户端，用户在web端进行的操作都会反映到该wps客户端下
- 多进程模式 ：用户窗口下可以启动多个wps客户端，通过使用创建的WpsClient对象调用wpsjsrpcsdk.js提供的接口，可以将操作发送到该对象对应的wps客户端上，不同的客户端可以执行不同的操作，互不影响
- 多用户模式（目前仅windows）：多个用户同时登陆一个服务器时，每个用户可以通过web端在自己的窗口下独立地操作wps客户端，不同用户的操作互不影响，支持单进程模式和多进程模式

## wpsjsrpcsdk.js开放的接口

### 管理加载项

web开发人员可以通过下面的接口来管理加载项

| 接口                     | 说明                           |
|--------------------------|--------------------------------|
| WpsAddonMgr.getAllConfig | 获取publish.xml的内容          |
| WpsAddonMgr.verifyStatus | 检查ribbon.xml文件内容是否有效 |
| WpsAddonMgr.enable       | 启用加载项                     |
| WpsAddonMgr.disable      | 禁用加载项                     |

### 单进程模式

web开发人员可以通过下面的接口，以单进程模式来唤起wps客户端并跟加载项进行交互，

| 接口                      | 说明                                  |
|---------------------------|---------------------------------------|
| WpsInvoke.InvokeAsHttp    | 以http协议启动wps客户端               |
| WpsInvoke.InvokeAsHttps   | 以https协议启动wps客户端              |
| WpsInvoke.RegWebNotify    | 注册一个前端页面接收WPS传来消息的方法 |
| WpsInvoke.CreateXHR       | 创建一个xhr对象                       |
| WpsInvoke.IsClientRunning | 判断当前客户端是否在运行              |
| WpsInvoke.ClientType      | 指定加载项类型， wps / et / wpp       |

### 多进程模式

web开发人员可以通过下面的接口，以多进程模式来唤起wps客户端并跟加载项进行交互

| 接口                           | 说明                                |
|--------------------------------|-------------------------------------|
| WpsClient                      | 构造函数，用来创建一个WpsClient对象 |
| wpsClient.InvokeAsHttp         | 以http启动wps客户端                 |
| wpsClient.InvokeAsHttps        | 以https启动wps客户端                |
| wpsClient.onMessage            | 消息注册函数的回调函数              |
| wpsClient.StartWpsInSilentMode | 以静默模式启动wps客户端             |
| wpsClient.ShowToFront          | 显示wps客户端到最前面               |
| wpsClient.CloseSilentClient    | 关闭未显示出来的静默启动的wps客户端 |
| wpsClient.IsClientRunning      | 判断当前客户端是否在运行            |
| wpsClient.jsPluginsXml         | 指定加载项路径                      |

### 多用户模式

web开发人员可以通过下面的接口，启用多用户模式

| 接口              | 说明           |
|-------------------|----------------|
| EnableMultiUser() | 启用多用户模式 |

## wpsjsrpcsdk.js的使用

### 管理加载项

开发者可以通过我们注册到window上的一个对象 WpsAddonMgr，来管理WPS加载项，包括安装、卸载、检查加载项可用性以及获取加载项配置的功能

注册到 WpsAddonMgr对象上的接口有以下四个，开发者可以参考下面的示例代码 复制 和参数说明

通过 WpsAddonMgr，可以调用的接口包括：

#### WpsAddonMgr.getAllConfig

通过该接口可以获得publish.xml的内容，查看加载项配置，包含一个参数

**代码示例**

``` JavaScript
function installWpsAddin(callBack){
    WpsAddonMgr.getAllConfig(function(e){
        if(!e.response||e.response.indexOf("null")>=0){//本地没有加载项，直接安装
            if(curList.length>0){
                installWpsAddinOne(callBack);
            } 
        }else{//本地有加载项，先卸载原有加载项，然后再安装
            localList=JSON.parse(e.response)
            unInstallWpsAddinOne(callBack)
        }
    })
}
```

**参数说明**

|  参数名  | 数据类型 |         描述         | 必选（默认值） |
|:--------:|:--------:|:--------------------:|:--------------:|
| *callback* | function | 执行成功后的回调函数 |   否（null）   |

#### WpsAddonMgr.verifyStatus

通过该接口可以获得加载项ribbon.xml的内容，通过格式检查可以判断该加载项是否有效，包含两个参数

**代码示例**

``` JavaScript
var element = {
    "name": "WpsOAAssist",
    "addonType": "wps",
    "online": "true",
    "url": "http://127.0.0.1:3888/plugin/wps/"
}
WpsAddonMgr.verifyStatus(element, function(result) {
    console.log(result.msg)
})
```

**参数说明**

|  参数名  | 数据类型 |           描述           | 必选（默认值） |
|:--------:|:--------:|:------------------------:|:--------------:|
| *element*  |  object  | JSON对象，包含加载项信息 |       是       |
| *callback* | function |   执行成功后的回调函数   |   否（null）   |

#### WpsAddonMgr.enable

通过该接口可以安装指定的加载项

**代码示例**

``` JavaScript
var element = {
    "name": "WpsOAAssist",
    "addonType": "wps",
    "online": "true",
    "url": "http://127.0.0.1:3888/plugin/wps/"
}
WpsAddonMgr.enable(element, function (result) {
    if (result.status) {
        console.log(result.msg)
    } else {
        console.log("安装成功")
    }
})
```

**参数说明**

|  参数名  | 数据类型 |           描述           | 必选（默认值） |
|:--------:|:--------:|:------------------------:|:--------------:|
| *element*  |  object  | JSON对象，包含加载项信息 |       是       |
| *callback* | function |   执行成功后的回调函数   |   否（null）   |

#### WpsAddonMgr.disable

通过该接口可以卸载指定的加载项

**代码示例**

``` JavaScript
var element = {
    "name": "WpsOAAssist",
    "addonType": "wps",
    "online": "true",
    "url": "http://127.0.0.1:3888/plugin/wps/"
}
WpsAddonMgr.disable(element, function (result) {
    if (result.status) {
        console.log(result.msg)
    } else {
        console.log("卸载成功")
    }
})
```

**参数说明**

|  参数名  | 数据类型 |           描述           | 必选（默认值） |
|:--------:|:--------:|:------------------------:|:--------------:|
| *element*  |  object  | JSON对象，包含加载项信息 |       是       |
| *callback* | function |   执行成功后的回调函数   |   否（null）   |

### 单进程模式

开发者可以通过我们注册到window上的一个对象WpsInvoke，来使用WPS加载项的单进程模式，该对象包含了我们注册进去的，web页面可以调用的接口

注册到WpsInvoke对象上的接口有以下六个，开发者可以参考下面的示例代码 复制 和参数说明

通过WpsInvoke，可以调用的接口包括：

#### WpsInvoke.InvokeAsHttp

该接口通过http协议发送请求，包含八个参数

**代码示例**

``` JavaScript
function openOfficeFile(param) {
    var invokeParam = {
        flag: "openOfficeFile",
        filepath: param
    }
    WpsInvoke.InvokeAsHttp(
        projInfo.type,
        projInfo.name,
        "openOfficeFileFromSystemDemo",
        JSON.stringify(invokeParam),
        callbackFunc)
}
```

**参数说明**

|    参数名     | 数据类型  |             描述             | 必选（默认值） |
|:-------------:|:---------:|:----------------------------:|:--------------:|
|  *clientType*   |  string  |         加 载项类型          |       是       |
|     *name*     |  string  |          加载项名称          |       是       |
|     *func*     |  string  | 要调用的wps加载项中的函数名  |       是       |
|    *param*     |   json   | func 调用的函数所需要的参数  |       是       |
|   *callback*   | function |   接口执行成功后的回调函数   |   否（null）   |
| *showToFront*  |   bool   | 是否将wps客户端显示到前面来  |   否（true）   |
| *jsPluginsXml* |  string  | 指定jsplugins.xml文件的地 址 |   否（null）   |
|  *silentMode*  |   bool   |         是否静默启动         |  否（false）   |

#### WpsInvoke.InvokeAsHttps

该接口通过http协议发送请求，包含八个参数

**代码示例**

``` JavaScript
function openOfficeFile(param) {
    var invokeParam = {
        flag: "openOfficeFile",
        filepath: param
    }
    WpsInvoke.InvokeAsHttps(
        projInfo.type,
        projInfo.name,
        "openOfficeFileFromSystemDemo",
        JSON.stringify(invokeParam),
        callbackFunc)
}
```

**参数说明**

|    参数 名    | 数据类型  |            描述             | 必选（默认值） |
|:-------------:|:---------:|:---------------------------:|:--------------:|
|  *clientType*  |  string   |         加载项类型          |       是       |
|     *name*     |  string   |         加载项名称          |       是       |
|     *func*     |  string  | 要调用的wps加载项中的函数名 |       是       |
|    *param*     |   json   | func调用的函数所需要的参数  |       是       |
|   *callback*   | function |  接口执行成功后的回调函数   |   否（null）   |
| *showToFront*  |   bool   | 是否将wps客户端显示到前面来 |   否（true）   |
| *jsPluginsXml* |  string  | 指定jsplugins.xml文件的地址 |   否（null）   |
|  *silentMode*  |   bool   |        是否静默启动         |  否（false）   |

#### WpsInvoke.RegWebNotify

该接口发送注册请求，用于接收wps客户端发回来的消息，包含三个参数

**代码示例**

``` JavaScript
WpsInvoke.RegWebNotify(WpsInvoke.ClientType.wps, pluginName, function (messageText) {
    var span = document.getElementById("webnotifyspan")
    span.innerHTML = "(次数：" + ++WebNotifycount + ")：" + messageText;
})
```

**参数说明**

|   参数名    | 数据类型 |         描述         | 必选 （默认值） |
|:-----------:|:--------:|:--------------------:|:---------------:|
| *clientType* |  string  |      加载项名称      |       是        |
|    *name*    |  string  |      加载项类型      |       是        |
|  *callback*  | function | 收到消息后的回调函数 |   否（null）    |

#### WpsInvoke.CreateXHR

该接口返回一个可用于进行数据传输的http对象，无参数

**代码示例**

``` JavaScript
/*创建一个http请求对象*/
var httpObj = WpsInvoke.CreateXHR()
```

#### WpsInvoke.IsClientRunning

该接口用于查看wps客户端是否在运行 ，包含两个参数

**代码示例**

``` JavaScript
WpsInvoke.IsClientRunning(WpsInvoke.ClientType.wps, function (status) {
    if (status.response == "Client is running.")
        alert("任务发送失败，WPS 正在执行其他任务，请前往WPS完成当前任务")
    else
        alert("WPS 客户端被关闭了!")
})
```

**参数说明**

|   参数名   | 数据类型 |         描述         | 必选 （默认值） |
|:----------:|:--------:|:--------------------:|:---------------:|
| *clientType* |  string  |      加载项类型      |       是        |
| *callback*  | function | 执行成功后的回调函数 |   否（null）    |

#### WpsInvoke.ClientType

该接口用于获取加载项类型，返回值为对应的加载项类型，无参数

**代码示例**

``` JavaScript
WpsInvoke.ClientType.wps
WpsInvoke.ClientType.wpp
WpsInvoke.ClientType.et
```

### 多进程模式

要使用多进程模式，先要通过我们注册到window上的WpsClient方法，构造一个wpsclient对象，然后通过该对象调用WpsClient里面的接口

首先，通过WpsClient创建一个对象

**代码示例**

``` JavaScript
var wpsClient = new WpsClient(WpsInvoke.ClientType.wps);
```

通过上面的代码，可以创建一个wpsClient对象，该对象可以调用WpsClient里面的接口，需要传递加载项类型作为必填参数

使用创建的wpsClient对象，我们可以调用的接口包括：

#### wpsClient.InvokeAsHttp

该接口通过http协议发送请求 ，包含五个参数

**代码示例**

``` JavaScript
var invokeParam = {
    Index: 'OpenFile',
    AppType: projInfo.type,
    filepath: url,
    params: {
        date: "2020/02/20"
    }
}
wpsClient.InvokeAsHttp(
    projInfo.name,
    "InvokeFromSystemDemo",
    JSON.stringify(invokeParam),
    openFileCallbackFunc,
    false)
```

**参数说明**

|    参数名    | 数据类型 |             描述              | 必选 （默认值） |
|:------------:|:--------:|:-----------------------------:|:---------------:|
|    *name*     |  string  |         加 载项名 称          |       是        |
|    *func*     |  string  | 要 调用的wps加载项 中的函数名 |       是        |
|    *param*    |  string  |  func 调用的函数所需要的参数  |       是        |
|  *callback*   | function |   接口执行成功后的回调函数    |   否（null）    |
| *showToFront* |   bool   |  是否将wps客户端显示到前面来  |   否（true）    |

#### wpsClient.InvokeAsHttps

该接口通过https协议发送请求，包含五个参数

**代码示例**

``` JavaScript
var invokeParam = {
    Index: 'OpenFile',
    AppType: projInfo.type,
    filepath: url,
    params: {
        date: "2020/02/20"
    }
}
wpsClient.InvokeAsHttps(
    projInfo.name,
    "InvokeFromSystemDemo",
    JSON.stringify(invokeParam),
    openFileCallbackFunc,
    false)
```

**参数说明**

|    参数名    | 数据类型 |                    描述                     | 必选（默认值） |
|:------------:|:--------:|:-------------------------------------------:|:--------------:|
|    *name*     |  string  |                 加载项名称                  |       是       |
|    *func*     |  string  |        要 调用的wps加载项中的函数名         |       是       |
|    *param*    |  string  | func 调用的函数所需要的参数，是一个json对象 |       是       |
|  *callback*   | function |          接口执行成功后的回调函数           |   否（null）   |
| *showToFront* |   bool   |         是否将wps客户端显示到前面来         |   否（true）   |

#### wpsClient.onMessage

通过该接口指定收到客户端消息之后的回调函数，无参数

**代码示例**

``` JavaScript
wpsClient.onMessage = function (messageText) {
    ++WebNotifycount; var spanDiv = document.getElementById("webnotifyCount" + eleId);
    spanDiv.innerText = "(次数：" + WebNotifycount + ")："; 
    var span = document.getElementById("webnotifyspan" + eleId);
    span.innerHTML = messageText;
} 
```

#### wpsClient.StartWpsInSilentMode

通过该接口静默启动多进程客户端，包含两个参数

**代码示例**

``` JavaScript
wpsClient.StartWpsInSilentMode(projInfo.name, function () {
    var invokeParam = {
        Index: 'OpenFile',
        AppType: projInfo.type,
        filepath: url,
        data: {
            date: "2020/02/20"
        }
    }
    wpsClient.InvokeAsHttp(
        projInfo.name,
        "InvokeFromSystemDemo",
        JSON.stringify(invokeParam),
        openFileCallbackFunc,
        false)
})
```

**参数说明**

|  参数名   | 数据类型 |           描述           | 必选（默认值） |
|:---------:|:--------:|:------------------------:|:--------------:|
|   *name*   |  string  |        加载项名称        |       是       |
| *callback* | function | 接口执行成功后的回调函数 |   否（null）   |

#### wpsClient.ShowToFront

通过该接口将客户端显示到最前面，包含两个参数

**代码示例**

``` JavaScript
wpsClient.ShowToFront(projInfo.name, function () {
    setProgress(0) 
    var dlg = document.getElementById("message");
    dlg.style.display = 'none'; 
    document.body.style.pointerEvents = "auto"
});
```

**参数说明**

|  参数名   | 数据类型  |           描述           | 必选（默认值） |
|:---------:|:---------:|:------------------------:|:--------------:|
|   *name*   |  string  |       加载 项 名称       |       是       |
| *callback* | function | 接口执行成功后的回调函数 |   否（null）   |

#### wpsClient.CloseSilentClient

通过该接口可以关闭未显示出来的启动的静默启动客户端，包含两个参数

**代码示例**

``` JavaScript
wpsClient.CloseSilentClient(projInfo.name, function () { 
    setProgress(0) 
    var dlg = document.getElementById("message");
    dlg.style.display = 'none'; 
    document.body.style.pointerEvents = "auto" 
});
```

**参数说明**

|  参数名   | 数据类型 |           描述           | 必选（默认值） |
|:---------:|:--------:|:------------------------:|:--------------:|
|   *name*   |  string  |        加载项名称        |       是       |
| *callback* | function | 接口执行成功后的回调函数 |   否（null）   |

#### wpsClient.IsClientRunning

通过该接口可以判断启动的客户端是否已关闭，包含一个参数

**代码示例**

``` JavaScript
wpsClient.IsClientRunning(function (status) {
    if (status.response == "Client is running.")
        alert("任务发送失败，WPS 正在执行其他任务，请前往WPS完成当前任务")
    else
        alert("WPS 客户端被关闭了!")
})
```

**参数说明**

|  参数名  | 数据类型 |           描述           | 必选（默认值） |
|:--------:|:--------:|:------------------------:|:--------------:|
| *callback* | function | 接口执行成功后的回调函数 |   否（null）   |

#### wpsClient.jsPluginsXml

通过该接口可以指定jsplugins.xml文件的地址，让客户端加载 jsplugins.xml里的加载项，无参数

**代码示例：**

``` JavaScript
wpsClient.jsPluginsXml = "http://127.0.0.1:8080/jsplugins.xml"
```

### 多用户模式

目前多用户模式仅支持windows，且至少需要客户端达到企业版20210525版本

多用户模式默认处于关闭状态，如果有场景需要用的多用户，开发者可以在编写前端代码时调用接口一次，即可开启多用户模式

``` JavaScript
EnableMultiUser();
```

开启多用户模式后，参考上面描述的方式，书写单进程或多进程代码即可，无需进行其他操作，多用户场景下即可自动支持

### 返回值处理

wpsjsrpcsdk.js返回给web前端的数据格式为JSON

分为两种格式：

当wps加载项成功执行了业务系统发送的操作时

``` JavaScript
{
    status: 0, 
    response: responseStr
}
```

返回值的status == 0，加载项的操作的返回值为responseStr

当wps加载项没有成功执行业务系统发送的操作时

``` JavaScript
{
    status: 1 / 2 / 3 / 4, 
    message: errorMessage
}
```

返回值的status != 0，返回的错误信息为errorMessage

因此，可以根据返回值的status，判断操作是否执行成功，然后进行处理

**代码示例**

``` JavaScript
function clientCallbackFunc(result) {
    if (result.status !== 0) {
        if (result.message == "Failed to send message to WPS.") {
            wpsClient.IsClientRunning(function (status) {
                if (status.response == "Client is running.")
                    alert("任务发送失败，WPS 正在执行其他任务，请前往WPS完成当前任务")
                else
                    alert("WPS 客户端被关闭了!")
            })
        } else {
            alert(result.message)
        }
    } else {
        alert(result.response)
    }
}
```

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/WPS 加载项集成业务系统开发](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E9%9B%86%E6%88%90%E4%B8%9A%E5%8A%A1%E7%B3%BB%E7%BB%9F%E5%BC%80%E5%8F%91.html)

------------------------------------------------------------------------
# OA 助手概述

## 基本介绍

OA助手是OA系统与 WPS 集成的解决方案，采用了 WPS 加载项技术，具有跨浏览器、跨平台的特性。

## 使用场景

OA助手集成了常用的 OA 系统功能，形成了一套标准。通过参数配置、功能组合能够快速完成对 OA 厂商的集成。 可以通过 [OAAssist](https://code.aliyun.com/zouyingfeng/wps/tree/master/oaassist) 示例对 OA 助手有一个直观了解 。

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA 助手概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%20%E5%8A%A9%E6%89%8B%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# 新建文档

创建空白文档，可以通过参数指定名称，格式转换保存，设置用户名等

``` JavaScript
_WpsStartUp([{
    "NewDoc": {}
}])
```

| 参数名 | 数据类型 | 描述                                     | 必选 |
|--------|----------|------------------------------------------|------|
| *NewDoc* | String   | 方法名，对应于OA助手dispatcher支持的方法 | 是   |

## 创建空白文档场景展现

![](服务器端图像/新建文件.gif)

## 业务流程（执行逻辑）

![](服务器端图像/文档流程图.png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/新建文档](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E6%96%B0%E5%BB%BA%E6%96%87%E6%A1%A3.html)

------------------------------------------------------------------------
# 打开文档

打开指定文档，可以通过参数指定保护模式，格式转换保存，设置用户名等

``` JavaScript
_WpsStartUp([{
    "OpenDoc": {
        "docId": "123",
        "uploadPath": uploadPath,
        "fileName": filePath,
        "docPassword": {
            "docPassword": docPassword
        }
    }
}])
```

| 参数名      | 数据类型 | 描述                                     | 必选 |
|-------------|----------|------------------------------------------|------|
| *OpenDoc*     | String   | 方法名，对应于OA助手dispatcher支持的方法 | 是   |
| *docId*       | String   | 文档ID                                   | 是   |
| *uploadPath*  | String   | 保存文档上传路径                         | 是   |
| *fileName*    | String   | 打开的文档路径                           | 是   |
| *docPassword* | String   | 文档密码                                 | 是   |

## 打开指定文档场景展现

![](服务器端图像/打开指定文件.gif)

## 业务流程（执行逻辑）

![](服务器端图像/文档流程图.png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/打开文档](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E6%89%93%E5%BC%80%E6%96%87%E6%A1%A3.html)

------------------------------------------------------------------------
# 文档上传

打开指定文档，需传入uploadPath等必须的参数,根据一系列的判断后,将文档保存到OA服务器上

``` JavaScript
_wps.openDoc({
    "docId" :"c2de1fcd1d3e4ac0b3cda1392c36c9",
    "uploadPath" : "***", //文档下载地址
    "fileName" : "***" //文档上传地址
});
```

## 输入参数

| 参数名     | 数据类型 | 描述                                          | 必选 |
|------------|----------|-----------------------------------------------|------|
| *docId*      | String   | 文档ID，文档在数据库中的ID                    | 是   |
| *fileName*   | String   | 获取服务器文档接口（不传即为新建空文档）      | 是   |
| *uploadPath* | String   | 保存文档接口,将文档保存到服务器上的服务器地址 | 是   |

## 保存功能场景展现

![](服务器端图像/保存文档到OA.gif)

## 业务流程（执行逻辑）

![](服务器端图像/保存文档到服务器.png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/文档上传](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E6%96%87%E6%A1%A3%E4%B8%8A%E4%BC%A0.html)

------------------------------------------------------------------------
# 显示修订

打开文档，文档打开修订模式，通过参数指定是否显示痕迹

**\_wps.openDoc** ( *params* )

``` JavaScript
_WpsStartUp([{
    "OpenDoc": { // OpenDoc 对应于OA助手dispatcher中的方法名
        "uploadPath": uploadPath, // 保存文档上传路径
        "fileName": filePath,
        "userName": '李雷', //用户名
        "revisionCtrl": {
            "bOpenRevision": true, //true 打开修订，false关闭修订
            "bShowRevision": true //true 显示修订，false隐藏修订
        }
    }
}])
```

## 输入参数

| 参数名        | 数据类型 | 描述                                              | 必选 |
|---------------|----------|---------------------------------------------------|------|
| *docId*         | String   | 文档ID，文档在数据库中的ID                        | 是   |
| *fileName*      | String   | 获取服务器文档接口（不传即为新建空文档）          | 是   |
| *uploadPath*    | String   | 新建空文档文件名,不传默认新建文档以系统时间为名称 | 是   |
| *revisionCtrl*  | JSON     | 痕迹控制 ，不传正常打开                           | 是   |
| *bOpenRevision* | Boolean  | true(打开)false(关闭)修订                         | 是   |
| *bShowRevision* | Boolean  | true(显示)/false(关闭)痕迹                        | 是   |

## 文档修订模式场景展现

![](服务器端图像/打开修订模式.gif)

## 业务流程（执行逻辑）

![](服务器端图像/控制公文的修订显示.png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/显示修订](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E6%98%BE%E7%A4%BA%E4%BF%AE%E8%AE%A2.html)

------------------------------------------------------------------------
# 文档自动填充

打开一个包含书签的模板文件后，OA助手在后台调用templateDataUrl获取模板数据的json字符串（书签定义及值）填充到对应的模板位置中。支持文字、链接、图片三种填充类型数值。

**\_wps.fillTemplate** ( *params* )

``` JavaScript
_WpsStartUp([{
    "OpenDoc": {
        "docId": "123",
        "uploadPath": uploadPath,
        "fileName": filePath,
        "picPath" : GetDemoPngPath(),
        "copyUrl" : backupPath,
    }
}])
```

## 输入参数

| 参数名     | 数据类型 | 描述                                             | 必选 |
|------------|----------|--------------------------------------------------|------|
| *OpenDoc*    | String   | 方法名，对应于OA助手dispatcher支持的方法         | 是   |
| *docId*      | String   | 文档ID，OA助手用以标记文档的信息，以区分其他文档 | 是   |
| *uploadPath* | String   | 保存文档上传路径                                 | 是   |
| *fileName*   | String   | 打开的文档路径                                   | 是   |
| *picPath*    | String   | 插入图片的路径                                   | 否   |
| *copyUrl*    | String   | 备份的服务器路径                                 | 否   |

## 文档填充公文场景展现

![](服务器端图像/填充模板.gif)

## 业务流程（执行逻辑）

![](服务器端图像/填充模板.png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/文档自动填充](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E6%96%87%E6%A1%A3%E8%87%AA%E5%8A%A8%E5%A1%AB%E5%85%85.html)

------------------------------------------------------------------------
# 文档格式转换

打开文档，将文档转换成PDF格式或OFD格式，再进行上传到服务器

``` JavaScript
_wps.openDoc({
    "docId" :"c2de1fcd1d3e4ac0b3cda1392c36c9",
    "uploadPath" :  "https://***",
    "fileName" :  "https://***",
    "suffix" : ".pdf ",
    "uploadAppendPath" :  "https://***",
    "uploadFieldName" : "file"
});
```

## 输入参数

| 参数名           | 数据类型 | 描述                                                       | 必选 |
|------------------|----------|---------------------------------------------------------------------------|------|
| *docId*            | String   | 文档ID，文档在数据库中的ID                                            | 是   |
| *fileName*         | String   | 获取服务器文档接口（不传即为新建空文档）                                  | 是   |
| *uploadPath*       | String   | 新建空文档文件名,不传默认新建文档以系统时间为名称                            | 是   |
| *suffix*           | String   | 要由Doc转到其他格式的后缀。包括：pdf、out、ofd三种。                              | 是   |
| *uploadAppendPath* | String   | 用户转其他文件格式时（如转PDF格式保存），OA助手会把转换格式后的文件，上载到这个路径上。不传默认为uploadPath的值 | 否   |
| *uploadFieldName*  | String   | 配合上传路径，包装上传form表单时，对应上传的字段域名称。不传默认为“file”   | 否   |

## 文档转换格式上传场景展现

![](服务器端图像/转换格式上传.gif)

## 业务流程（执行逻辑）

![](服务器端图像/文档转换上传.png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/文档格式转换](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E6%96%87%E6%A1%A3%E6%A0%BC%E5%BC%8F%E8%BD%AC%E6%8D%A2.html)

------------------------------------------------------------------------
# 公文套红头

打开指定文档，将红头文件插入到当前文档的顶部，若模板中存在对应书签， 则会将当前文档的正文插入到书签所在位置

``` JavaScript
_WpsStartUp([{
    "OpenDoc": {
        "docId": "123",
        "fileName": filePath,
        "insertFileUrl": templateURL,
        "bkInsertFile": 'Content', //红头模板中填充正文的位置书签名
    }
}])
```

## 输入参数

| 参数名        | 数据类型 | 描述                                   | 必选 |
|---------------|----------|----------------------------------------|------|
| *OpenDoc*       | String   | 方法名，对应于OA助手dispatcher中的方法 | 是   |
| *docId*         | String   | 文档ID                                 | 是   |
| *fileName*      | String   | 打开的文档路径                         | 否   |
| *insertFileUrl* | String   | 指定的红头模板                         | 否   |
| *bkInsertFile*  | String   | 书签名，红头模板插入正文的位置         | 否   |

## 公文套红头场景展现

![](服务器端图像/套红头(OA-Web端).gif)

## 业务流程（执行逻辑）

![](服务器端图像/套红头(OA-Web端).png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/公文套红头](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E5%85%AC%E6%96%87%E5%A5%97%E7%BA%A2%E5%A4%B4.html)

------------------------------------------------------------------------
# 在线编辑

点击按钮，输入要打开的文档路径，输入文档上传路径，如果传的不是有效的服务端地址，将无法使用保存上传功能。打开WPS文字后,将根据文档路径在线打开对应的文档，保存将自动上传指定服务器地址

| 参数名        | 数据类型 | 描述                                     | 必选 |
|---------------|----------|------------------------------------------|------|
| *onlineEditDoc* | String   | 方法名，对应于OA助手dispatcher支持的方法 | 是   |
| *docId*         | String   | 文档ID                                   | 是   |
| *uploadPath*    | String   | 保存文档上传路径                         | 是   |
| *fileName*      | String   | 打开的文档路径                           | 是   |
| *buttonGroups*  | String   | 屏蔽OA助手功能按钮                       | 否   |

## 创建空白文档场景展现

![](服务器端图像/新建文件.gif)

## 业务流程（执行逻辑）

![](服务器端图像/文档流程图.png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/在线编辑](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E5%9C%A8%E7%BA%BF%E7%BC%96%E8%BE%91.html)

------------------------------------------------------------------------
# WPS 宏编辑器概述

WPS Office 宏编辑器，是一个为二次开发者设计的编写和编辑JavaScript代码调用或是扩展WPS功能的程序。除了可通过编写 JS脚本来加速执行日常任务外，还可以使用 JS为 WPS Office添加新功能，或以特定于业务需要的方式来提示文档用户并与之交互。

宏编辑器为用户提供了常用代码编辑器，录制宏，运行宏，自定义公式，设计UI控件执行宏以及调试器等功能。开发者可以通过代码编辑器用JS的语法，把想要执行的动作编写成宏，也可以通过录制宏功能，把常见的动作自动转换成宏代码，制作出包含宏的文档。文档使用者在有工作需要时，可以通过打开宏对话框去运行宏，也可以通过结合可视化的控件，在合适的控件响应时执行宏。在表格里面，还可以通过编写设计宏来扩展公式。

宏编辑器界面包含由以下部分组成：

- 菜单栏
- 工具栏
- 状态栏
- 主体功能窗口区

宏编辑器的开发：

- 切换宏编辑器开发环境
- 录制宏
- 运行宏
- 在表格中创建和使用自定义函数

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/宏编辑器概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# 切换宏编辑器开发环境

对于部分VBA和WPS宏编辑器共存的WPS版本，默认的开发环境是VBA，您可以通过点击“开发工具”选项卡的“切换到JS环境”来切换当前的开发环境，如下图：

![](Base64图像/Base64图像7来自_WPS%20集成模式_WPS宏编辑器开发_切换宏编辑器开发环境.png)

![](Base64图像/Base64图像8来自_WPS%20集成模式_WPS宏编辑器开发_切换宏编辑器开发环境.png)

也可以通过“选项”——“视图”——勾选的“功能区选项”的“默认JS开发环境”，来设置默认的开发环境：

![](Base64图像/Base64图像9来自_WPS%20集成模式_WPS宏编辑器开发_切换宏编辑器开发环境.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/切换宏编辑器开发环境](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E5%88%87%E6%8D%A2%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91%E7%8E%AF%E5%A2%83.html)

------------------------------------------------------------------------
# 自定义函数

尽管 WPS 表格包含大量内置工作表函数，但它可能没有适用于您执行的每种类型的计算的函数，因此 WPS 宏编辑器提供了创建自定义函数的功能。

宏编辑器里的自定义函数和宏类似，不同的是

1.它们执行计算而不是操作，所以某些类型的语句，例如设置单元格的值、修改选区等不建议写在自定义函数中。

2.宏通常不会有参数和返回值，而函数在参数上没有限制，往往是有返回值的。

下面以求矩形面积的函数为例。

## 插入代码

在代码模块中插入代码：

``` JavaScript
function 计算面积(长, 宽)
{
       if (宽 == undefined)
              return 长*长
       else
              return 长*宽
}
```

![](Base64图像/Base64图像10来自_WPS%20集成模式_WPS宏编辑器开发_在表格中创建和使用自定义函数.png)

## 使用函数

点击编辑栏的插入函数按钮，在弹出的插入函数对话框中，选择类别为“用户定义”，选中刚刚新增的“计算面积”函数，点击“确定”，在函数参数中输入长、宽的值，或是引用对应的单元格，点击"确定”即可完成函数插入操作，并自动计算

![](Base64图像/Base64图像11来自_WPS%20集成模式_WPS宏编辑器开发_在表格中创建和使用自定义函数.png)

![](Base64图像/Base64图像12来自_WPS%20集成模式_WPS宏编辑器开发_在表格中创建和使用自定义函数.png)

![](Base64图像/Base64图像13来自_WPS%20集成模式_WPS宏编辑器开发_在表格中创建和使用自定义函数.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/在表格中创建和使用自定义函数](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E5%9C%A8%E8%A1%A8%E6%A0%BC%E4%B8%AD%E5%88%9B%E5%BB%BA%E5%92%8C%E4%BD%BF%E7%94%A8%E8%87%AA%E5%AE%9A%E4%B9%89%E5%87%BD%E6%95%B0.html)

------------------------------------------------------------------------
# 录制宏

WPS 宏编辑器提供了一个自动生成 JS 代码的功能，即录制宏。录制宏功能通过宏录制器捕捉用户与 WPS 交互的操作，并以 JS 代码的形式记录下来，整个过程是自动的，不需要用户写代码。在实际的 WPS 二次开发过程中，此方法使用频率很高。当你不知道如何写 JS 代码时，只需录制下来，打开编辑器，查看代码即可。

因为宏录制器是通过捕获你大部分操作来生成代码，因此如果操作顺序中出现错误，例如点击了一个本不打算点击的按钮，宏录制器也会录制该错误操作。解决方法是重新录制整个操作，或直接修改 JS 代码。因此，录制内容前，最好先确保你的操作熟练，这样录制时操作越流畅，生成的宏越精简，运行宏时就越高效。

操作步骤：

1.切换到“开发工具”选项卡——点击“切换到 JS 环境”（如果有需要的话）——点击“录制新宏”按钮。——在弹出的“录制 JS 宏”对话框中，填写宏名，说明等宏信息，或者使用默认的信息。然后点击确定，进入宏录制状态。

![](Base64图像/Base64图像14来自_WPS%20集成模式_WPS宏编辑器开发_录制宏.png)

2.进入宏录制状态后，原有的按钮变成 “停止录制”，这是我们就可以进行操作了，这里我们以设置单元格边框为例。操作完成后，不要进行其他操作，切换到“开发工具”选项卡，点击“停止录制”按钮，录制宏就完成了。

![](Base64图像/Base64图像15来自_WPS%20集成模式_WPS宏编辑器开发_录制宏.png)

3.如需修改或是查看，可以切换到“开发工具”选项卡，点击“ WPS 宏编辑器”按钮，打开宏编辑器的主界面，查看代码。

![](Base64图像/Base64图像16来自_WPS%20集成模式_WPS宏编辑器开发_录制宏.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/录制宏](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E5%BD%95%E5%88%B6%E5%AE%8F.html)

------------------------------------------------------------------------
# 用户窗体设计

在工程中创建用户窗体后，可以在该窗体上添加工具箱内任意控件。该模块结合属性编辑器可以轻松设置控件属性；该模块结合代码编辑器可调用控件支持的事件，并能够通过代码获取和设置控件属性。

## 设计用户窗体

- 每个窗体窗口有最大化、最小化及关闭按钮。
- 可以拖动“工具箱”中的控件到用户窗体上进行布局。 控件的类别 详细见 工具箱 。
- 双击控件，将会生成响应该控件默认事件的函数，并进入代码编辑，也可以在已有或是新建的代码模块里面编写控件响应事件的函数。代码编辑器使用详细见 代码编辑窗口 ，控件的使用详细见控件的api。

## 运行用户窗体

选中用户窗体，通过点击快捷键 F5 或快捷菜单运行按钮可以运行用户窗体。也可以通过接口调用UserForm的Show方法运行窗体。

![](Base64图像/Base64图像17来自_WPS%20集成模式_WPS宏编辑器开发_用户窗体设计.png)

![](Base64图像/Base64图像18来自_WPS%20集成模式_WPS宏编辑器开发_用户窗体设计.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/用户窗体设计](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E7%94%A8%E6%88%B7%E7%AA%97%E4%BD%93%E8%AE%BE%E8%AE%A1.html)

------------------------------------------------------------------------
# 宏安全性

出于安全考虑， WPS 提供宏安全性来控制打开文档宏的运行。您可以通过“开发工具”选项卡上的“宏安全性”入口，来对宏的执行行为进行控制。如果您安装的 WPS 支持 VBA, 此处的设置跟 VBA 宏的安全性是共享的 。

![](Base64图像/Base64图像19来自_WPS%20集成模式_WPS宏编辑器开发_运行宏_宏安全性.png)

| 级别   | 效果                                                                     |
|--------|---------------------------------------------------------|
| 低     | 您将不受保护，而某些宏具有潜在的不安全因素。只有在安装了防毒软件或检查了所有要打开的文档的安全性时，才能使用这项设置。 |
| 中     | 您可以选择是否运行可能不安全的宏。                                        |
| 高     | 只允许运行可靠来源签署的宏，未经签署的宏会自动取消。                                                                   |
| 非常高 | 只允许运行安装在受信任位置的宏。所有其他签署的和未经签署的宏都将被禁用。                                               |

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/运行宏/宏安全性](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E8%BF%90%E8%A1%8C%E5%AE%8F/%E5%AE%8F%E5%AE%89%E5%85%A8%E6%80%A7.html)

------------------------------------------------------------------------
# 运行宏对话框

切换到“开发工具”选项卡——点击“切换到 JS 环境”（如果有需要的话）——点击“ JS 宏”按钮。——在弹出的“ JS 宏”对话框中，可以对当前环境下的 JS 宏进行编辑、删除、运行、修改信息。点击运行按钮就可以执行宏了。

![](Base64图像/Base64图像20来自_WPS%20集成模式_WPS宏编辑器开发_运行宏_运行宏对话框.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/运行宏/运行宏对话框](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E8%BF%90%E8%A1%8C%E5%AE%8F/%E8%BF%90%E8%A1%8C%E5%AE%8F%E5%AF%B9%E8%AF%9D%E6%A1%86.html)

------------------------------------------------------------------------
# 宏控件响应宏

开发者可以切换到“开发工具”选项卡下，点击合适的 WPS 宏控件，往 WPS 文档插入控件。与宏编辑器内设计 用户窗体 类似，双击控件可以生成响应该控件默认事件的函数，并进入代码编辑，也可以在已有或是新建的代码模块里面编写控件响应事件的函数。

![](Base64图像/Base64图像21来自_WPS%20集成模式_WPS宏编辑器开发_运行宏_宏控件响应宏.png)

![](Base64图像/Base64图像22来自_WPS%20集成模式_WPS宏编辑器开发_运行宏_宏控件响应宏.png)

![](Base64图像/Base64图像23来自_WPS%20集成模式_WPS宏编辑器开发_运行宏_宏控件响应宏.png)

文档中支持的插入的宏控件和宏控件具体支持的事件和说明参照 空间接口 。

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/运行宏/宏控件响应宏](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E8%BF%90%E8%A1%8C%E5%AE%8F/%E5%AE%8F%E6%8E%A7%E4%BB%B6%E5%93%8D%E5%BA%94%E5%AE%8F.html)

------------------------------------------------------------------------
# 旧窗体控件运行宏

WPS 文字和表格组件支持插入旧式窗体，在窗体属性对话框（文字）或是指定宏对话框（表格）中可以选择窗体在不同时机的运行宏。

![](Base64图像/Base64图像24来自_WPS%20集成模式_WPS宏编辑器开发_运行宏_旧窗体控件运行宏.png)

![](Base64图像/Base64图像25来自_WPS%20集成模式_WPS宏编辑器开发_运行宏_旧窗体控件运行宏.png)

![](Base64图像/Base64图像26来自_WPS%20集成模式_WPS宏编辑器开发_运行宏_旧窗体控件运行宏.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/运行宏/旧窗体控件运行宏](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E8%BF%90%E8%A1%8C%E5%AE%8F/%E6%97%A7%E7%AA%97%E4%BD%93%E6%8E%A7%E4%BB%B6%E8%BF%90%E8%A1%8C%E5%AE%8F.html)

------------------------------------------------------------------------
# 事件响应宏

用户在使用WPS进行某些操作时，会触发所谓的事件。例如用户新建一个工作表，就有触发工作表的新建事件。WPS宏编辑器支持开发者编写事件的响应宏，来对某些操作自动执行宏。下例的代码是在当前应用程序切换工作表时调整该工作表 A 列到 F 列的大小的示例。

``` JavaScript
function Application_SheetActivate(Sh)
{
    Sh.Columns("A:F").AutoFit()
}
```

WPS 各个组件支持的事件可以通过 代码编辑窗口的导航栏 查看，具体支持的事件和说明参照WPS基础接口。

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/运行宏/事件响应宏](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E8%BF%90%E8%A1%8C%E5%AE%8F/%E4%BA%8B%E4%BB%B6%E5%93%8D%E5%BA%94%E5%AE%8F.html)

------------------------------------------------------------------------
# 菜单栏

- 文件菜单：包含保存文件、导入、导出项目文件、删除模块等功能。
- 编辑菜单：提供撤消、恢复、剪切、复制、粘贴、全选、查找替换等常见文本编辑操作，以及切换注释、缩进凸出、代码提示，书签等常见代码工具。
- 视图菜单：提供立即窗口、本地窗口、监视窗口、调用堆栈、查找结果工程资源管理器、属性窗口、工具箱等窗口的显示功能。
- 插入菜单：可以插入宏、窗体、模块和文件。
- 格式菜单：提供窗体控件布局、顺序功能。
- 调试菜单：提供编译、步进步出、切换断点、添加监视的能力。
- 运行菜单：运行宏、中断、重新设置、设计模式。
- 工具菜单：提供运行宏对话框、设置选项和工程属性对话框功能。
- 窗口菜单：提供新建窗口、拆分窗口和窗口切换功能。
- 帮助菜单：提供帮助文档、宏编辑器版本和版权信息功能。

![](Base64图像/Base64图像27来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_菜单栏.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/界面组成/菜单栏](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E7%95%8C%E9%9D%A2%E7%BB%84%E6%88%90/%E8%8F%9C%E5%8D%95%E6%A0%8F.html)

------------------------------------------------------------------------
# 工具栏

工具栏由 文件工具栏、 编辑工具栏、 调试工具栏、 运行工具栏、 布局工具栏组成，包含以下功能。

- 文件工具栏：切回主程序视图、导入、保存。
- 编辑工具栏：撤消、恢复、复制、剪切、粘贴、切换注释。
- 调试工具栏：编译。
- 运行工具栏：运行宏、中断、重新设置、设计模式。
- 布局工具栏：控件布局。

![](Base64图像/Base64图像28来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_工具栏.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/界面组成/工具栏](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E7%95%8C%E9%9D%A2%E7%BB%84%E6%88%90/%E5%B7%A5%E5%85%B7%E6%A0%8F.html)

------------------------------------------------------------------------
# 工程资源管理器

## 功能

工程资源管理器能够展示工程及工程包含的所有代码模块和用户窗体，并对代码模块和窗体进行管理。

![](Base64图像/Base64图像29来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_工程资源管理器.png)

## 窗口元素

- 工程名称：默认以当前文档名称作为工程名，且不可修改。
- 三角箭头：点击可展开或折叠列表。
- 代码：工程内包含的代码模块。
- 窗体：工程内包含的用户窗体。
- 用户窗体 (UserForm): 用于存放窗体控件的容器。

## 操作

1、右键菜单：选中工程名称 / 模块名 / 窗体名，点击鼠标右键，可进行插入、删除、重命名和切除文件夹操作。

2、键盘快捷方式：

- 方向向左键：选中工程工程，按此键可以收起列表。

- 方向向右键：选中工程工程，按此键可以收起列表。

- 方向向上键：每次上移一项。

- 方向向下键：每次下移一项。

3、热键：使用 ALT 键 + 辅助键可快速使用某项功能，辅助键在该项功能菜单后均有显示。

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/界面组成/主体功能窗口区/工程资源管理器](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E7%95%8C%E9%9D%A2%E7%BB%84%E6%88%90/%E4%B8%BB%E4%BD%93%E5%8A%9F%E8%83%BD%E7%AA%97%E5%8F%A3%E5%8C%BA/%E5%B7%A5%E7%A8%8B%E8%B5%84%E6%BA%90%E7%AE%A1%E7%90%86%E5%99%A8.html)

------------------------------------------------------------------------
# 代码编辑窗口

可以使用代码编辑窗口来查看、编辑 JS 代码。打开各模块的代码窗口后，支持在不同代码窗口中进行复制、粘贴等操作。

![](Base64图像/Base64图像30来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_代码编辑窗口.png)

## 对象框

导航栏上的对象框中包含此工程中包含事件的 JSAPI 对象及所有窗体对象。

![](Base64图像/Base64图像31来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_代码编辑窗口.png)

## 事件框

当从“对象框”中选择了一个对象后，所有与该对象相关的事件则会显示在“事件框”中。

![](Base64图像/Base64图像32来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_代码编辑窗口.png)

从事件框中点击选择一个事件后，该事件对应的方法会自动填充到代码编辑区域中。

![](Base64图像/Base64图像33来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_代码编辑窗口.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/界面组成/主体功能窗口区/代码编辑窗口](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E7%95%8C%E9%9D%A2%E7%BB%84%E6%88%90/%E4%B8%BB%E4%BD%93%E5%8A%9F%E8%83%BD%E7%AA%97%E5%8F%A3%E5%8C%BA/%E4%BB%A3%E7%A0%81%E7%BC%96%E8%BE%91%E7%AA%97%E5%8F%A3.html)

------------------------------------------------------------------------
# 用户窗体窗口

用户窗体窗口是用户窗体的容器，开发者可以在此窗口对用户窗体及其他控件双击、拖拽、右键进行设计修改。

![](Base64图像/Base64图像34来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_用户窗体窗口.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/界面组成/主体功能窗口区/用户窗体窗口](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E7%95%8C%E9%9D%A2%E7%BB%84%E6%88%90/%E4%B8%BB%E4%BD%93%E5%8A%9F%E8%83%BD%E7%AA%97%E5%8F%A3%E5%8C%BA/%E7%94%A8%E6%88%B7%E7%AA%97%E4%BD%93%E7%AA%97%E5%8F%A3.html)

------------------------------------------------------------------------
# 工具箱

## 功能描述

![](Base64图像/Base64图像35来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_工具箱.png)

## 工具箱控件

| 控件类型       | 功能                 |
|----------------|-----------------------------------------------|
| 选定对象   | 用于选中或是移动控件对象。                                            |
| 命令按钮   | 创建一个可以让用户选择，以完成一个命令的按钮                          |
| 标签       | 用于显示用户不能编辑的文本或图像，例如图形下的标题文本。              |
| 文本框     | 用于输入或改变文本内容。                                              |
| 复合框     | 复合框由一个文本输入控件和一个下拉菜单组成，用户可以从预先定义的列表中的选择一个选项，同时也可以直接通过文本框输入     |
| 列表框     | 用来显示用户可以选择的 项目 列表。如果不能一次显示全部 项目 的话，列表可以滚动。                                       |
| 复选框     | 可以显示出多重选择，用户可以选择一个或者多个选项，如果选中多个选项则显示出多重选择。                                   |
| 选项按钮   | 可以显示出多重选择，用户只能选择一个。                                |
| 滚动条     | 提供在长列表工程或大量信息中快速浏览的图形工具，以比例方式指示出当前位置，或是做为一个输入设备，或速度及数量的指示器。 |
| 水平布局器 | 将内部的控件按照水平方向排布，一列一个。                              |
| 垂直布局器 | 将内部的控件按照垂直方向排布，一行一个。                              |
| 水平间隔   | 填充水平方向上无用的空隙。                                            |
| 垂直间隔   | 填充垂直方向上无用的空隙。                                            |
| 旋转按钮   | 一种与其它控件并用微调控制的控件，也可以用来对一个范围的值或工程列表做向前或向后的遍历。                               |
| 切换按钮   | 创建一个切换开关的按钮。                                              |
| 框架       | 可以创建一个图形或控件的功能组。要为多个控件创建组，可先画框架，接着在框架中添加控件。                                 |
| 多页       | 包含一个选项卡栏和一个页面区的容器类组件，通过切换选项卡来切换不同页面，从而达到在同一个窗口中自由切换不同的内容。     |

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/界面组成/主体功能窗口区/工具箱](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E7%95%8C%E9%9D%A2%E7%BB%84%E6%88%90/%E4%B8%BB%E4%BD%93%E5%8A%9F%E8%83%BD%E7%AA%97%E5%8F%A3%E5%8C%BA/%E5%B7%A5%E5%85%B7%E7%AE%B1.html)

------------------------------------------------------------------------
# 属性编辑器

属性编辑器能够展示及设置对象的属性。

![](Base64图像/Base64图像36来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_属性编辑器.png)

## 对象框

![](Base64图像/Base64图像37来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_属性编辑器.png)

选中用户窗体后，属性编辑器的对象框内会展示选中窗体中的所有对象。选择窗体对象后，属性窗口展示该控件的属性。如果没有选中用户窗体，属性编辑器处于不可以选择状态。

## 属性列表选项框

**“按字母序”选项卡** — 按字母顺序列出所选的对象的所有属性。

**“按分类序”选项卡** — 根据性质分组，列出所选对象的所有属性。

点击属性的”值“列的值可以对属性进行编辑修改。

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/界面组成/主体功能窗口区/属性编辑器](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E7%95%8C%E9%9D%A2%E7%BB%84%E6%88%90/%E4%B8%BB%E4%BD%93%E5%8A%9F%E8%83%BD%E7%AA%97%E5%8F%A3%E5%8C%BA/%E5%B1%9E%E6%80%A7%E7%BC%96%E8%BE%91%E5%99%A8.html)

------------------------------------------------------------------------
# 代码调试

宏编辑器提供设置断点、查看堆栈、查看变量、单步、步进、步出等功能来辅助开发者查看代码执行情况，从而便于他们快速定位代码问题。

## 设置断点

在代码行号左侧空白处状态区，点击即可切换行断点状态，右键单击不仅可以设置行断点，还可以设置条件断点。

![](Base64图像/Base64图像38来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_代码调试.png)

![](Base64图像/Base64图像39来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_代码调试.png)

## 查看调用堆栈

当代码执行中断时，例如在断点中断时，即可使用堆栈功能。

在堆栈窗口中能够展示堆栈上每一帧所在代码的文档名、代码模块名称、具体函数名称以及行号信息。双击任一帧将会激活帧所在代码的窗口，并定位到对应的行。

![](Base64图像/Base64图像40来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_代码调试.png)

## 查看变量

中断时，在监视窗口或是立即窗口输入变量名后回车，可以查看变量的值：

![](Base64图像/Base64图像41来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_代码调试.png)

打开局部变量窗口，查看当前帧代码所在作用域的局部变量的值信息：

![](Base64图像/Base64图像42来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_代码调试.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/界面组成/主体功能窗口区/代码调试](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E7%95%8C%E9%9D%A2%E7%BB%84%E6%88%90/%E4%B8%BB%E4%BD%93%E5%8A%9F%E8%83%BD%E7%AA%97%E5%8F%A3%E5%8C%BA/%E4%BB%A3%E7%A0%81%E8%B0%83%E8%AF%95.html)

------------------------------------------------------------------------
# 立即窗口

立即窗口能够执行或调用 JSAPI 代码、自定义函数、执行窗体相关方法和属性等，但立即窗口中的代码不能够被存储。

![](Base64图像/Base64图像43来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_主体功能窗口区_立即窗口.png)

## 运行代码

在立即窗口中编辑或粘贴一行代码，然后按下回车 (Enter) 键来执行该段代码。

## 窗口操作

- 立即窗口可以拖放到屏幕中的任何地方，或是关闭。
- 拖放窗体位置后，通过点击工具栏“视图” - “立即窗口”，可将立即窗口恢复到默认位置。

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/界面组成/主体功能窗口区/立即窗口](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E7%95%8C%E9%9D%A2%E7%BB%84%E6%88%90/%E4%B8%BB%E4%BD%93%E5%8A%9F%E8%83%BD%E7%AA%97%E5%8F%A3%E5%8C%BA/%E7%AB%8B%E5%8D%B3%E7%AA%97%E5%8F%A3.html)

------------------------------------------------------------------------
# 状态栏

状态栏提供了代码内容的字符数、行数，当前光标定位的行、列、选中内容字符数、和行数、输入模式。

![](Base64图像/Base64图像44来自_WPS%20集成模式_WPS宏编辑器开发_界面组成_状态栏.png)

> WPS 开发文档： [WPS 集成模式/WPS宏编辑器开发/界面组成/状态栏](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8%E5%BC%80%E5%8F%91/%E7%95%8C%E9%9D%A2%E7%BB%84%E6%88%90/%E7%8A%B6%E6%80%81%E6%A0%8F.html)

------------------------------------------------------------------------
# Java 应用集成概述

## 基本介绍

WPS JavaAPI 属于WPS的跨平台二次开发解决方案的一部分，使用 Java 作为开发语言，具备快速开发、跨平台等优良特性。使用Java原生swing和awt作为窗体容器，大大提升了对平台和操作系统的兼容性。在窗体中通过Java API直接与WPS进行交互，特性如下：

- 具有完整的功能。可通过多种不同的方法对文档、表单和演示文稿进行创作、格式设置和操控；通过鼠标、键盘执行的操作几乎都能通过 Java API 完成；结合Java程序语言的编程逻辑，可以轻松实现自动化文档编辑
- 交互方式灵活。支持通过的接口调用的方式实现功能，还支持事件注册回调的方式实现自由交互
- 标准化集成。API 文档完整，接口设计符合 Java 语法规范，上手难度低，缩短开发周期

## 使用场景

在第三方开发的系统中想要实现一些文档编辑功能，然而自行开发一套文档编辑系统成本非常高也不现实，所以可以直接将wps集成到第三方的系统中，通过调用WPS二次开发接口实现想要的文档编辑功能。

## 环境搭建

- 系统要求：linux
- 开发工具：eclipse、idea
- JDK:1.7版本及以上
- WPS版本：2019

## 范例使用说明

![](服务器端图像/javademo.png)

更多核心能力，敬请期待！

有任何疑问，请加QQ：3253920855 获取帮助。

> WPS 开发文档： [WPS 集成模式/Java 应用集成 WPS 指南/Java 应用集成概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/Java%20%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%20WPS%20%E6%8C%87%E5%8D%97/Java%20%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# 创建首个 Java 应用示例

## 环境搭建

1）新建空的Java项目。 2）导入依赖jar包，以idea为例： File\>\>Project Structure\>\>Libraries\>\>Add(+)\>\>Java\>\>选择所有jar包\>\>apply\>\>ok 。

## 编写代码测试主类创建空白swing窗体

``` Java
import javax.swing.*;
public class TestMain {
     public static void main(String[] args) {
    //Linux下必须加这一句
    java.lang.System.setProperty("sun.awt.xembedserver", "true");

    SwingUtilities.invokeLater(new Runnable() {
        @Override
        public void run() {
            JFrame mainFrame = new JFrame();
                //设置显示窗口标题
                mainFrame.setTitle("WPS JAVA接口调用演示");
                //设置窗口显示尺寸
                mainFrame.setSize(1524, 768);
                    //置窗口是否可以关闭
                    mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
                //禁止缩放
                mainFrame.setResizable(false);
                //设置窗口是否可见
                mainFrame.setVisible(true);
            }
        });
    }
}
```

## 添加按钮事件

``` Java
Panel wpsPanel = new JPanel();
wpsPanel.setLayout(new BorderLayout());

//初始化一个空白的画布用于装载WPS
Canvas client = new Canvas() {
        public void paint(Graphics g) {
        super.paint(g);
    }
};
//添加画布
wpsPanel.add(client, BorderLayout.CENTER);
//添加JPane
mainFrame.add(wpsPanel);
```

## 将WPS嵌入到窗体内

``` Java
//获取用于装载WPS的Canves的句柄
WindowIDProvider pid = (WindowIDProvider)client.getPeer();
XEmbedCanvasPeer peer = (XEmbedCanvasPeer)pid;
peer.removeXEmbedDropTarget();
//将画布句柄和长宽等信息作为参数初始化WPS窗体
WpsArgs args = WpsArgs.ARGS_MAP.get(WpsArgs.App.WPS);
args.setWinid(pid.getWindow());
args.setHeight(client.getHeight());
args.setWidth(client.getWidth());
Application app = ClassFactory.createApplication();// 创建WPS Application对象
```

## 调用api接口

``` Java
//新建文档
app.get_Documents().Add(Variant.getMissing(),
    Variant.getMissing(), Variant.getMissing(), Variant.getMissing());
//获取版本号
String version = app.get_Build();
//光标位置插入文本，此处写入版本号
app.get_Selection().TypeText(version);
//关闭文档
app.get_ActiveDocument().Close(false, 
Variant.getMissing(), Variant.getMissing());
```

## 完整范例

1.[一个简单范例](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20集成模式/Java%20应用集成%20WPS%20指南/smalldemo.htm)

2.[获取更多范例](https://code.aliyun.com/zouyingfeng/wps/tree/master/java)

> WPS 开发文档： [WPS 集成模式/Java 应用集成 WPS 指南/创建首个 Java 应用示例](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/Java%20%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%20WPS%20%E6%8C%87%E5%8D%97/%E5%88%9B%E5%BB%BA%E9%A6%96%E4%B8%AA%20Java%20%E5%BA%94%E7%94%A8%E7%A4%BA%E4%BE%8B.html)

------------------------------------------------------------------------
# 最佳做法

## 默认参数

Java中不提供函数默认参数，因而调用接口时必须给定义的所有参数都传参，如果某个对应位置的参数需要使用默认值，请传递 **Variant.getMissing()** ，例如：

``` Java
app.get_ActiveDocument().Protect(WdProtectionType.wdAllowOnlyReading,
　　Variant.getMissing(),
　　Variant.getMissing(),
　　Variant.getMissing(),
　　Variant.getMissing()
);
```

## 函数返回值

返回值是api对象类型的接口可以直接判断返回值是否为null进行容错校验，如 **Application.get_ActiveDocument()** 无返回值或返回值是数字，字符串等类型的接口，需捕获Exception异常，以此判断操作是否执行成功。

## 属性获取和设置

所有对象的属性在设置和获取时，在java中均需使用加上get_或put_前缀的对应方法，无法通过“对象名.属性”直接调用，例如：

``` Java
CommandBarControl controlOfd = commandBar.get_Controls().get_Item(26);
controlOfd.put_Visible(false);
// 不能使用 controlOfd.visible = false;
```

## 关于lcid参数

遇到接口参数名为 lcid 的参数，默认传递2052即可。

## Canves句柄说明

1.获取用于装载WPS的Canves的句柄(pid和peer)的语句，必须在 mainFrame.setVisible(true) 之后调用，否则拿到空指针。

2.获取用于装载WPS的Canves的句柄(pid和peer)的语句，不能写在JPanel的构造函数里，必须在JPanel对象创建完成后再获取，否则拿到空指针。

## 注册事件回调流程

1.注册关闭事件。

``` Java
EventCookie cookie = 
    app.advise(ApplicationEvents4.class, new ApplicationEvents4() {
        @Override
        public void DocumentBeforeClose(Document doc, Holder<Boolean> cancel) {
            JOptionPane.showMessageDialog(null,"关闭事件被拦截");
            cancel.value = true; //cancel为true，代表取消关闭
        }
    });
```

2.取消注册关闭事件。

3.添加一个按钮，给按钮注册事件

``` Java
cookie.close();
cookie = null; 
```

``` Java
//添加一个按钮
CommandBar bar = app.get_CommandBars().Add(
    "test", 1, Variant.getMissing(), Variant.getMissing());
CommandBarControl control = 
    bar.get_Controls().Add(1, 1, "test", 1, "test");
bar.put_Visible(true);
control.put_Visible(true);
control.put_Caption("test");

//给按钮注册事件
EventCookie cookie = control.advise(_CommandBarButtonEvents.class,
    new _CommandBarButtonEvents() {
        @Override
        public void Click(CommandBarButton ctrl, Holder<Boolean> cancelDefault) {
            JOptionPane.showMessageDialog(null,"你点击了按钮 test");
            cancelDefault.value = true;
        }
    });
```

4.取消注册按钮时间

``` Java
cookie.close(); cookie = null;
```

## 调用WPS 2016程序的说明

JavaAPI Demo中默认开启了加密，在WPS2016下会导致新建文档阻塞。使用WPS2016的用户需要将代码中这一行前面的注释去掉：

``` Java
args.setCrypted(false); //wps2016需要关闭加密
```

> WPS 开发文档： [WPS 集成模式/Java 应用集成 WPS 指南/最佳做法](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/Java%20%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%20WPS%20%E6%8C%87%E5%8D%97/%E6%9C%80%E4%BD%B3%E5%81%9A%E6%B3%95.html)

------------------------------------------------------------------------
# C++ 应用集成概述

## 基本介绍

欢迎使用WPS二次开发C++接口帮助文档。本文档包含C++接口和API事件的详细介绍、示例和参考资料，用于指导您开发基于 WPS 的解决方案。

本文档包含下列部分： 1) 提供 按对象 分类并按字母顺序列出的新成员列表。 2) 提供接口对象和枚举的列表。 3) 提供通过文件索引功能进行成员对象和枚举的查找操作。

## 使用说明

- WPS对象的属性总共分为两种，分别是只读属性和可读写属性。只读属性只提供了读取接口，接口名为 **"get\_"+属性名** ，例如：Application对象的Version属性是只读属性，若要获取此属性值，调用接口"get_Version"即可。可读写属性提供了读取和设置接口，读取属性的接口名与上述只读属性的一样，设置属性的接口名为 **"put\_"+属性名** ，例如：Application对象的Caption 属性是可读写属性，此属性的读取接口为"get_Caption"，设置属性的接口为"put_Caption"。详细用法请参考 [接口使用示例](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20集成模式/C++%20应用集成%20WPS%20指南/Demo.html) 。

- 调用接口获取属性值（设置属性）时，获取的值以参数形式传出（传入），接口调用的返回值作为判断调用成功与否的依据，只有当调用成功的时候，获取的值才是有效值（设置的值才生效）。 **若返回值为0，表示调用成功，非0为调用失败**， 具体的返回值表示的信息请参考 [接口错误码](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20集成模式/C++%20应用集成%20WPS%20指南/ErrorCode.html) 。

- 在WPS对象模型中，Application对象代表WPS 应用程序，它是所有其他WPS对象的最顶层对象。Application对象包含可返回最高级对象的属性和方法，通过使用这些属性和方法可以控制整个WPS环境。在程序中必须要通过这个对象启动WPS，才可以实现对WPS 程序的控制；进而可以打开文档，实现对其他对象的控制。

- 遇到接口参数名为lcid的参数（ 此参数代表区域语言ID），默认传递2052（即简体中文）即可。

## 使用场景

该帮助文档主要用于C++开发基于WPS的解决方案，包括进程内调用（仅Windows）和跨进程调用的方案。进程内调用的解决方案也称为加载项，由WPS在启动的时候主动加载，可以通过自定义功能区、自定义工具栏和菜单栏、自定义任务窗格等方式来扩展 WPS 应用程序的功能，并与 WPS文档中的内容进行交互。进程外调用的解决方案可以通过调用二次开发C++接口来调起WPS并执行指定的任务，当然也可以将WPS嵌入到自己的程序，具体代码可以参考我们提供的示例文档。

更多核心能力，敬请期待！

有任何疑问，请加QQ：3253920855 获取帮助。

> WPS 开发文档： [WPS 集成模式/C++ 应用集成 WPS 指南/C++ 应用集成概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/C%2B%2B%20%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%20WPS%20%E6%8C%87%E5%8D%97/C%2B%2B%20%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# 创建首个C++ 集成应用示例

## 基本介绍

这份文档介绍了如何使用提供的Demo，在linux环境下使用C++开发基于 WPS 的解决方案。

## 源码下载

[访问下载页面](https://code.aliyun.com/zouyingfeng/wps/tree/master/cpp)

## 目录结构说明

将源码下载下来后，其目录结构如下：

\| -- include WPS头文件，包含公共以及WPS，ET，WPP各自对应的api接口。

\| -- src c++源代码文件。

\| -- README.md 帮助文件

## 环境准备

- 操作系统： Linux系统（中标麒麟，银河麒麟等国产操作系统也支持）
- 安装 Qt ( 兼容 Qt4或Qt5 )
- 安装 WPS Office 2019 For Linux

## 编译及运行

- 在终端中使用cd命令，进入代码的src目录，执行qmake命令。 ![](服务器端图像/qmake.png)
- 执行make，在当前目录下会生成wpsDemo。 ![](服务器端图像/make.png)
- 执行./wpsDemo，运行wpsDemo。 ![](服务器端图像/wpsDemo.png)
- 最终效果如图所示。 ![](服务器端图像/DemoUI.png)

> WPS 开发文档： [WPS 集成模式/C++ 应用集成 WPS 指南/创建首个C++ 集成应用示例](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/C%2B%2B%20%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%20WPS%20%E6%8C%87%E5%8D%97/%E5%88%9B%E5%BB%BA%E9%A6%96%E4%B8%AAC%2B%2B%20%E9%9B%86%E6%88%90%E5%BA%94%E7%94%A8%E7%A4%BA%E4%BE%8B.html)

------------------------------------------------------------------------
# 浏览器应用集成嵌入WPS概述

## 基本介绍

欢迎使用浏览器应用集成嵌入WPS指南。本文档包含浏览器应用集成接口和API事件的详细介绍、示例和参考资料，为您开发基于 WPS 应用程序提供解决方案。 本文档包含以下部分：

- 提供 按对象 分类并按字母顺序列出的新成员列表。
- 提供接口对象和枚举的列表
- 提供通过文件索引功能进行成员对象和枚举的查找操作。

## 使用场景

开发者可以编写JavaScript来调用插件功能，该插件可以被支持NPAPI机制的浏览器加载。该插件提供浏览器OA集成的二次开发技术方案，具备快速开发、轻量化的特性。针对企业用户，能快速集成到现有Web 系统，简化企业用户对文档的操作，实现文档自动化； 针对个人用户，可以提供更完整的个性化能力。

## 使用范例

这里提供一份最简单的 使用范例（Linux 版)，参见 创建首个浏览器集成插件示例， 并对涉及元素进行了说明 。

## 插件对浏览器的要求

因为该WPS插件使用NPAPI机制来和浏览器交互，故要求使用插件的浏览器必须支持NPAPI机制且必须开启NPAPI机制。

以下是支持的常见的浏览器及其版本：

- FireFox浏览器52及小于52的版本（高于52的版本不再支持NPAPI）
- Chrome浏览器45及小于45的版本（高于45的版本不再支持NPAPI）
- 360浏览器 （支持NPAPI的版本）
- UOS浏览器 （支持NPAPI的版本）

更多核心能力，敬请期待！

有任何疑问，请加QQ：3253920855 获取帮助。

> WPS 开发文档： [WPS 集成模式/浏览器应用集成嵌入WPS指南/浏览器应用集成嵌入WPS概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/%E6%B5%8F%E8%A7%88%E5%99%A8%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%E5%B5%8C%E5%85%A5WPS%E6%8C%87%E5%8D%97/%E6%B5%8F%E8%A7%88%E5%99%A8%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%E5%B5%8C%E5%85%A5WPS%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# 创建首个浏览器集成插件示例

## 环境准备

- 一台安装了LINUX系统的机器。
- 一个支持NPAPI的浏览器（参见 浏览器应用集成嵌入WPS概述 ） 。

## Demo代码

源码如下：

``` Html
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <title> wps二次开发演示 </title>
</head>

<body>
    <object type="application/x-wps" width="1024" height="768" id="wps"></object>
    <!--通过类型 “application/x-wps” 来加载WPS浏览器NPAPI插件，并传递宽高值-->
    <script type="text/javascript">
        //来获取wps的插件对象
        var obj = document.getElementById("wps");
        //获取WPS对象树的根节点，后续功能调用都是根据 Application 逐步得到的
        var Application = obj.Application;
        //创建一个新的空白文档
        Application.createDocument("wps");
        //获取文档内容区域
        var range = Application.ActiveDocument.Content;
        // 在文档内容区域插入一个3行5列的表格，设置表格显示边框
        Application.ActiveDocument.Tables.Add(range, 3, 5).Borders.Enable = 1;
     </script>
</body>
</html>
```

[在浏览器中打开此链接查看效果](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20集成模式/浏览器应用集成嵌入WPS指南/np_demo.htm)

## 效果如下图所示

![](服务器端图像/示例图.jpg)

## 浏览器集成插件Demo

[请点击查看更多demo链接](https://code.aliyun.com/zouyingfeng/wps/tree/master/np-example/browser-integration-wps)

> WPS 开发文档： [WPS 集成模式/浏览器应用集成嵌入WPS指南/创建首个浏览器集成插件示例](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/%E6%B5%8F%E8%A7%88%E5%99%A8%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%E5%B5%8C%E5%85%A5WPS%E6%8C%87%E5%8D%97/%E5%88%9B%E5%BB%BA%E9%A6%96%E4%B8%AA%E6%B5%8F%E8%A7%88%E5%99%A8%E9%9B%86%E6%88%90%E6%8F%92%E4%BB%B6%E7%A4%BA%E4%BE%8B.html)

------------------------------------------------------------------------
# 适用范围

移动端JSAPI 目前仅支持Android端，文字组件

> WPS 开发文档： [WPS 集成模式/Android 端应用集成 WPS 指南/Android应用基础概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/Android%20%E7%AB%AF%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%20WPS%20%E6%8C%87%E5%8D%97/Android%E5%BA%94%E7%94%A8%E5%9F%BA%E7%A1%80%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# 从Visual Basic转到 JavaScript

## JSAPI接口的差异

### 方法的差异

vb的方法可以不加括号，但jsapi中所有的方法都需要加括号，如果方法不加括号会被js语法判定为属性。

``` vb
Application.Workbooks(1).Close
```

``` JavaScript
Application.Workbooks.Item(1).Close()
```

vb的方法支持给部分参数赋值。但js对缺省的参数需要用undefined占位补齐，如下面例子中为Find方法的第一和第三个参数赋值，在js中Find方法第二个参数需要用undefined占位补齐

``` vb
Range("A1:N29").Find("香港特别", LookIn:=xlValues).Select
```

``` JavaScript
Application.ActiveSheet.Range("A1:N29").Find("我",undefined, xlValues).Select();
```

vb可通过数组方式取集合中的对象，jsapi必须通过Item方法获取集合中的对象，其中要注意二维数组取值时需要将value要改成Value2

``` vb
Application.Workbooks(1).Close
```

``` JavaScript
Application.Workbooks.Item(1).Close()
```

以下是二维数组的取值方法

``` vb
cells(2,3).value
```

``` JavaScript
Application.Cells.Item(2,3).Value2
```

### 属性的差异

vb中调用书写错误的属性会报错，js不会报错，这是一个bug，所以特别注意。

``` JavaScript
Application.ActiveDocument.Name111 // 不报错
```

vb支持thisdocument对象，jsapi暂时不支持该对象，如果遇到thisdocument对象，可以用ActiveDocument代替

``` vb
MsgBox (Application.ThisDocument.Name)
```

``` JavaScript
MsgBox (Application.ActiveDocument.Name)
```

### 事件的差异

jside的导航栏展示和自动补全的事件比较少，但js支持的事件基本和vb一致，下面给出一个vb事件转换为jside中事件的例子

``` vb
Private Sub Workbook_SheetSelectionChange(ByVal Target As Range)
Dim i As Integer
IF Target.Column>6 And Target.Column>37 && Target.Row>6 And Target.Row Mod 2 = 1 And Cells(Target.Row -1, 5).Value <> “”
Then
    With Target.Validation
        .Delete
        .Add Tpye:=xlValidateList,AlertStyle:=xlValidAlertStop,Operator:=xlBetween,Formular1:="=区域城市！$A2:$A30000")
        .IncellDropdown = true
```

``` JavaScript
function Workbook_SheetSelectionChange(Sh, Target) {
    if(Target.Column>6 && Target.Column>37 && Target.Row>6 && (Target.Row)%2n==1 && Cells.Item(Target.Row -1n, 5).Value2 !==””) {
        var temptarget = Target.Validation
        temptarget.Delete()
        temptarget.Add(xlValidateList,xlValidAlertStop,xlBetween,"=区域城市！$A2:$A30000")
        temptarget.IncellDropdown = true
    }
}
```

vb支持thisdocument对象，jsapi暂时不支持该对象，如果遇到thisdocument对象，可以用ActiveDocument代替

``` vb
MsgBox (Application.ThisDocument.Name)
```

``` JavaScript
MsgBox (Application.ActiveDocument.Name)
```

## 函数定义

vb用sub ... end sub关键字定义函数，js用funtion{}关键字定义函数

``` vb
Public(Private) Sub AddTemplate()
    MsgBox "d"
End Sub
```

``` JavaScript
function AddTemplate() {
    MsgBox（"d"）
}
```

## 数据类型

vb在数据定义时需要指明类型，但是js是动态类型，赋值后才有类型,js包括的基本类型：字符串（String）、数字(Number)、布尔(Boolean)、对空（Null）、未定义（Undefined）、Symbol，声明这些类型都用关键字var

``` vb
Dim i As Integer
```

``` JavaScript
var i
```

``` vb
Dim j As String
```

``` JavaScript
var j
```

**基本类型对比**

| js类型  | vb类型                                  |
|---------|-----------------------------------------|
| boolean | Boolean                                 |
| number  | Integer，Long，LongLong，Single，Double |
| string  | String                                  |

vb的bool类型首字母大写，js中都是小写

``` vb
True, False
```

``` JavaScript
true, false
```

**字符串**

js字符串使用反斜杠\\转义字符把特殊字符转换为字符串字符:

| 代码 | 结果 | 描述   |
|------|------|--------|
| \\   | '    | 单引号 |
| \\   | "    | 双引号 |
| \\   | \\   | 反斜杠 |

因此，在常见的文件路径的场景，需要采用“/”或者“\\”(Windows)来作为路径分隔符：

``` JavaScript
function test()
{
    Workbooks.Open("H:\\新建文件夹\\工作簿1.xlsx")
}
```

## 运算符

算术运算符的差异，如vb中的mod关键字应改为%

``` vb
Target.Row Mod 2 = 1
```

``` JavaScript
(Target.Row)%2==1
```

“+”可以用来拼接字符串

``` vb
AddIns.Add FileName:="C:\Program Files\Microsoft Office" _"\Templates\Letters Faxes\MyFax.dot", Install:=True
```

``` JavaScript
AddIns.Add("C:/Program Files/Microsoft Office" + "/Templates/Letters Faxes/MyFax.dot",true)
```

逻辑运算符

``` vb
1.逻辑与（AND）2.逻辑或（OR） 3.逻辑非（<>）
```

``` JavaScript
逻辑与（&&）2.逻辑或（||） 3.逻辑非（！）
```

注意js和vb判断空的逻辑不一样，vb内部会智能一点，遇到判断空的语句，尽量用"!"

``` vb
if Cells(1,2).value == ""  //空
```

``` JavaScript
if (Cells.Item(1,2).Value2)   //空
```

``` vb
if Cells(1,2).value <> ""  //非空
```

``` JavaScript
if (!Cells.Item(1,2).Value2)  //非空
```

## 枚举

vb和js枚举的使用上存在的差异如下，在vb中wdAlignParagraphLeft和WdParagraphAlignment.wdAlignParagraphLeft语法都正确，但在js中只有wdAlignParagraphLeft是正确用法

``` vb
//vb用以上两种方式的枚举都可以正常使用
Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
Application.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
```

``` JavaScript
//js只支持这一种
Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
```

## With关键字

vb有with关键字，js没有该关键字，转换时可以将with后面的对象赋值给一个变量，例子如下

``` vb
With Selection.Range.PageSetup.TextColumns
    .SetCount NumColumns:=2
    .EvenlySpaced = 0
End With
```

``` JavaScript
Var columens = Selection.Range.PageSetup.TextColumns
columens.SetCount(2)
columens.EvenlySpaced = 0
//注意TextColumns后面没有括号因为这是一个属性
```

## 条件语句

VB中的常用条件语句{IF 条件 Then 表达式 Endif}，在JS中应该转换为 if(条件) {表达式}

``` vb
IF Target.Column>6 And Target.Column>37 And Target.Row>6 And Target.Row Mod 2 = 1 And Cells(Target.Row -1, 5).Value <> ""
Then
    With Target.Validation
        .Delete
        .Add
Endif
```

``` JavaScript
if(Target.Column > 6 && Target.Column > 37 && Target.Row > 6 && (Target.Row) % 2 == 1 && Cells.Item(Target.Row -1n, 5).Value2 !== null) {
    var temptarget = Target.Validation
    temptarget.Delete()
    temptarget.Add(xlValidateList,xlValidAlertStop,xlBetween,"=区域城市！$A2:$A30000")
}
```

## 路径

VB中的常用条件语句{IF 条件 Then 表达式 Endif}，在JS中应该转换为 if(条件) {表达式}

``` vb
IF Target.Column>6 And Target.Column>37 And Target.Row>6 And Target.Row Mod 2 = 1 And Cells(Target.Row -1, 5).Value <> ""
Then
    With Target.Validation
        .Delete
        .Add
Endif
```

``` JavaScript
if(Target.Column > 6 && Target.Column > 37 && Target.Row > 6 && (Target.Row) % 2 == 1 && Cells.Item(Target.Row -1n, 5).Value2 !== null) {
    var temptarget = Target.Validation
    temptarget.Delete()
    temptarget.Add(xlValidateList,xlValidAlertStop,xlBetween,"=区域城市！$A2:$A30000")
}
```

> WPS 开发文档： [WPS 基础接口/从Visual Basic转到 JavaScrip](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E4%BB%8EVisual%20Basic%E8%BD%AC%E5%88%B0%20JavaScript.html)

------------------------------------------------------------------------
# WPS 加载项可用性

## 请确认安装的 WPS 安装包带有 WPS 加载项功能。

个人用户可以在 WPS 官网下载最新的安装包；企业用户请联系贵公司的 WPS 项目负责人。

## 启用 WPS 加载项

打开 WPS 安装目录，进入 wps.exe 程序同级目录，在cfgs文件夹下新建或打开oem.ini文件，在文档最后添加如下配置

``` ini
[Support]
##启用加载项
JsApiPlugin = true
##启用网页调试，详情请参考 WPS 基础接口说明
JsApiShowWebDebugger=true
[Server]
##指定加载项配置文件 jsplugins.xml 链接地址
JSPluginsServer = http://***/jsplugins.xml
```

## 加载项配置说明

加载项配置在文件 jsplugins.xml 中。jsplugins.xml 从服务器下载后会保存在本地，服务器更新后会重新下载。jsplugins.xml 格式如下：

``` xml
<jsplugins>
    <jsplugin name="demo_et"  version="3.6" type="et" url="http://localhost:8080/demo_et_3.6.7z"/>
    <jsplugin name="demo_wpp" version="3.6" type="wpp" url="http://localhost:8080/demo_wpp_3.6.7z"/>
</jsplugins>
```

\<jsplugins\> \*\*\* \</jsplugins\> 包含所有加载项配置

\<jsplugin \*\*\* /\> 代表一个加载项

name 为加载项名称

version 为该加载项版本号

type 为加载项类型

url 为加载项压缩包下载链接

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/WPS 加载项可用性](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%8F%AF%E7%94%A8%E6%80%A7.html)

------------------------------------------------------------------------
# WPS 加载项性能

WPS 加载项为了保证WPS 应用程序的安全性与稳定性，采用多进程架构，把网页渲染端和 WPS 应用进程分离； 同时通过共享内存、事件对象等技术，实现了一套进程间快速同步调用的机制。 在性能测试过程中，为10000个单元格赋值耗时2000毫秒，充分满足要求。

``` JavaScript
function TestSetCellFormula() 
{
    var cells = Application.ActiveWorkbook.ActiveSheet.Cells;
    let date = new Date();
    let start = date.getTime();
    for(var i = 1; i <= 100; ++i) 
    { 
        for(var j = 1; j <= 100; ++j) 
        {
            cells.Item(i, j).Formula = i + j; 
        } 
    } 
    date = new Date(); 
    let end = date.getTime(); 
    alert(end - start); return; 
}
```

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/WPS 加载项性能](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E6%80%A7%E8%83%BD.html)

------------------------------------------------------------------------
# 自定义功能区概述

自定义功能区采用通用的 [CustomUI标准](https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/edc80b05-9169-4ff7-95ee-03af067f35b1) 进行配置， 该标准定义了一整套标准的控件，比如按钮、下拉菜单、组合框；能够对控件的标签、图标、点击事件等属性进行配置。下面通过一个示例进行详细说明。

![](服务器端图像/ribbon.png)

图 1. Web 加载项自定义功能区

``` ini
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" onLoad="OnAddInLoad">
    <ribbon startFromScratch="false">
        <tabs>
            <tab id="WebAddinDemo" label="Web 加载项示例">
                <group id="btnGroup" label="示例分组">
                    <button onAction="OnClicked" label="示例按钮" getImage="GetImage"/>
                </group>
            </tab>
        </tabs>
    </ribbon>
</customUI>
```

**标记说明**

- onLoad 代表一个事件，仅在 WPS 应用装载该 WPS 加载项时触发一次。OnAddInLoad是开发者自定义的 JavaScript 函数，通常用来执行一些初始化操作。

- tabs 可以包含多个 tab，每一个 tab 对应一个自定义功能区。

- group 将多个控件划分成不同的分组，便于将相互关联的功能组织在一起。

- button 是一个按钮。onAction在用户点击后触发，OnClicked是开发者自定义的 JavaScript 函数。label 是按钮文字标签，getImage 用来自定义按钮图表 GetImage是开发者自定义的 JavaScript 函数，getImage 首先会在自定义功能区第一次显示的时候执行一次。当开发者调用刷新整个功能区或通过id刷新该控件时再次执行。

**更多说明**

请参考 [CustomUI标准](https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/edc80b05-9169-4ff7-95ee-03af067f35b1) 查看更多说明

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/自定义功能区/自定义功能区概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/%E8%87%AA%E5%AE%9A%E4%B9%89%E5%8A%9F%E8%83%BD%E5%8C%BA/%E8%87%AA%E5%AE%9A%E4%B9%89%E5%8A%9F%E8%83%BD%E5%8C%BA%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# RibbonControl对象

**RibbonControl** 对象是自定义功能区中的控件对象。

**属性**

| 名称    | 说明             |
|---------|------------------|
| Id  | 控件唯一ID       |
| Tag | 指定的任意字符串 |

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/自定义功能区/RibbonControl对象](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/%E8%87%AA%E5%AE%9A%E4%B9%89%E5%8A%9F%E8%83%BD%E5%8C%BA/RibbonControl%E5%AF%B9%E8%B1%A1.html)

------------------------------------------------------------------------
# RibbonUI对象

**RibbonUI** 对象是 onLoad 事件的参数。

**方法**

| 名称                     | 说明                                                                       |
|--------------------------|----------------------------------------------------------------------------|
| ActivateTab          | 使 WPS 应用程序激活显示当前自定义功能区。                                  |
| ActivateTabMso       | 激活内置的功能区，参数为指定的功能区名称。                                 |
| ActivateTabQ         | 激活自定义功能区，参数为自定义功能区名称及命名控件                         |
| Invalidate         | 将自定义功能区的数据缓存置为无效，WPS 应用程序会在下一次显示时再次获取更新 |
| InvalidateControl  | 指定更新一个控件，参数是控件Id                                             |
| InvalidateControlMso | 指定更新一个内置控件，参数是内置控件Id                                     |

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/自定义功能区/RibbonUI对象](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/%E8%87%AA%E5%AE%9A%E4%B9%89%E5%8A%9F%E8%83%BD%E5%8C%BA/RibbonUI%E5%AF%B9%E8%B1%A1.html)

------------------------------------------------------------------------
# 任务窗格概述

WPS 加载项的任务窗格是一个用来浏览网页的用户界面面板，通常停靠在 WPS 应用程序主窗口的一侧，开发者可以控制任务窗格停靠的位置及宽高。 但重要的是任务窗格中的这个网页可以和 WPS 直接完成交互，开发者可以提取 WPS 文档中的数据在网页中集中显示，也可以通过网页交互将数据直接写进文档。 下面的一段代码演示了如何打开一个任务窗格。

![](服务器端图像/taskpane.png)

图 1. Web 加载项任务窗格

``` JavaScript
let taskPane = Application.CreateTaskPane(GetUrlPath() + "/html_taskpane.html");
taskPane.DockPosition = Application.Enum.JSKsoEnum_msoCTPDockPositionLeft;
taskPane.Visible = true;
```

**代码说明**

- 调用wps上的方法CreateTaskPane创建任务窗格，参数是网页路径。返回值是一个 JavaScript 对象。

- DockPosition是任务窗格的停靠位置属性，这里将任务窗格停靠到左侧，默认是右侧。

- Visible控制任务窗格是否显示，这里设置为true后，将任务窗格显示出来

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/任务窗格/任务窗格概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/%E4%BB%BB%E5%8A%A1%E7%AA%97%E6%A0%BC/%E4%BB%BB%E5%8A%A1%E7%AA%97%E6%A0%BC%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# 事件概述

通过 WPS 加载项事件能够对 WPS 应用程序发出的事件添加 JavaScript 方法进行处理。 在通知型事件中，可以接收已经发生变化，比如通过WindowActivate事件，可以对文档的切换做一些功能性处理； 在询问型事件中，可以控制是否继续执行当前操作，比如通过WorkbookBeforeClose事件，可以取消文档的关闭。

``` JavaScript
Application.ApiEvent.AddApiEventListener("WorkbookBeforeClose", function(workBook){
    if (!workBook.Saved) {
        alert("请先保存文档！")
        Application.ApiEvent.Cancel = true;
    }
});
```

**代码说明**

注册WorkbookBeforeClose监听事件，在工作簿关闭时，判断该工作簿是否保存，如果未保存，弹出“请先保存文档！”提示，并取消关闭

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/事件/事件概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E4%BA%8B%E4%BB%B6%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# 表格事件

## 事件列表

| 名称                           | 级别        | 类型   | 触发时机|
|--------------------------------|-------------|--------|-----------|
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
| SheetBeforeRightClick          | Application | 询问型 | 当右击任一工作表之前触发此事件 。                                |
| SheetCalculate                 | Application | 通知型 | 在任一工作表进行计算时触发此事件。                               |
| WorkbookActivate               | Application | 通知型 | 任一工作簿被激活时，将触发此事件。                               |
| WorkbookDeactivate             | Application | 通知型 | 任一工作簿被切换到非激活状态时触发此事件。                       |
| WorkbookNewSheet               | Application | 通知型 | 在任一打开的工作簿中创建新工作表时触发此事件。                   |
| WindowResize                   | Application | 通知型 | 任一工作簿窗口调整大小时将触发此事件。                           |
| SheetFollowHyperlink           | Application | 通知型 | 单击任一工作表的超链接时触 发此事件。                            |
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

## 事件

### WorkbookOpen

当打开一个工作簿时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Wb   | 必选      | Workbook | 打开的工作簿对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( "WorkbookOpen" , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("打开工作簿")
}                                       
Application.ApiEvent.AddApiEventListener("WorkbookOpen", func) 
```

### NewWorkbook

当新建工作簿时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Wb   | 必选      | Workbook | 新建的工作簿对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " NewWorkbook " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("新建工作簿")
}                                       
Application.ApiEvent.AddApiEventListener("NewWorkbook", func) 
```

### WorkbookBeforeSave

任一工作簿被保存之前触发此事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明|
|----------|-----------|----------|--------------------------------------------------------------|
| Wb       | 必选      | Workbook | 工作簿对象。                                                 |
| SaveAsUI | 必选      | Boolean  | 如果为 true 则表示此次保存操作将会弹出保存或是另存为对话框。 |
| Cancel   | 必选      | Object   | 如果设置其属性Value为 true ，则不关闭工作簿。                |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookBeforeSave " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, cancal) {
    var msg = "您保存后要关闭吗？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};
Application.ApiEvent.AddApiEventListener("WorkbookBeforeSave", func) 
```

### WorkbookAfterSave

任一工作簿被保存之后触发此事件。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                      |
|---------|-----------|----------|-------------------------------------------|
| Wb      | 必选      | Workbook | 工作簿对象。                              |
| Success | 必选      | Boolean  | 如果保存操作成功，则为true; 否则为false。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookAfterSave " , callbackFunc )`

**示例**

``` JavaScript
function func(Success) {
    if(!Success)
        alert("保存成功")
};  
Application.ApiEvent.AddApiEventListener("WorkbookAfterSave", func) 
```

### WorkbookBeforeClose

任一打开的工作簿关闭之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Wb     | 必选      | Workbook | 工作簿对象。                                  |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不关闭工作簿。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookBeforeClose " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, cancal) {
    var msg = "您确定要关闭吗？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};  
Application.ApiEvent.AddApiEventListener("WorkbookBeforeClose", func) 
```

### WorkbookBeforePrint

在打印任一打开的工作簿之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Wb     | 必选      | Workbook | 工作簿对象。                                  |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不打印工作簿。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookBeforePrint " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, cancal) {
    var msg = "您打印后要关闭吗？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};
Application.ApiEvent.AddApiEventListener("WorkbookBeforePrint", func) 
```

### SheetSelectionChange

该工作簿任一工作表上的选定区域发生更改时，将触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                   |
|--------|-----------|----------|------------------------|
| Sh     | 必选      | Object   | 选区改变的工作表对象。 |
| Target | 必选      | Range    | 新选定的区域。         |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetSelectionChange " , callbackFunc )`

**示例**

``` JavaScript
function func(Sh, Target) {
    var msg = "工作表" + Sh + "选中的区域是：" + Target.Address();
    alert(msg);
}
Application.ApiEvent.AddApiEventListener("SheetSelectionChange", func) 
```

### SheetChange

当用户或外部链接更改了该工作簿任一工作表中的单元格时触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                   |
|--------|-----------|----------|------------------------|
| Sh     | 必选      | Object   | 选区改变的工作表对象。 |
| Target | 必选      | Range    | 新选定的区域。         |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetChange " , callbackFunc )`

**示例**

``` JavaScript
function func(Sh, Target) {
    var msg = "工作表" + Sh + "选中的区域是：" + Target.Address();
    alert(msg);
}
Application.ApiEvent.AddApiEventListener("SheetChange", func) 
```

### SheetActivate

当激活任一工作表时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明         |
|------|-----------|----------|--------------|
| Sh   | 必选      | Object   | 工作表对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetActivate " , callbackFunc )`

**示例**

``` JavaScript
let func = (Sh)=>{
    alert("工作表激活")
}                                       
Application.ApiEvent.AddApiEventListener("SheetActivate", func) 
```

### SheetDeactivate

当该工作簿任一工作表被切换到非激活状态时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明         |
|------|-----------|----------|--------------|
| Sh   | 必选      | Object   | 工作表对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetDeactivate " , callbackFunc )`

**示例**

``` JavaScript
let func = (Sh)=>{
    alert("工作表切换为未激活")
}                                       
Application.ApiEvent.AddApiEventListener("SheetDeactivate", func) 
```

### WindowActivate

任一工作簿窗口被激活时，将触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |
| Wn   | 必选      | Window   | 激活窗口对象。   |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowActivate " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, wn) {
    var msg = "您激活的窗口是: " + wn.Caption;
    alret(msg);
}           
Application.ApiEvent.AddApiEventListener("WindowActivate", func) 
```

### WindowDeactivate

任一工作簿窗口被切换到非激活状态时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |
| Wn   | 必选      | Window   | 激活窗口对象。   |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowDeactivate " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, wn) {
    var msg = "您取消激活的窗口是: " + wn.Caption;
    alret(msg);
}                                       
Application.ApiEvent.AddApiEventListener("WindowDeactivate", func) 
```

### SheetBeforeDoubleClick

当双击任一工作表之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sh     | 必选      | Object    | 双击的工作表对象。                            |
| Target | 必选      | Range对象 | 双击区域所在的单元格对象。                    |
| Cancel | 必选      | Object    | 如果设置其属性Value为 true ，则取消此次双击。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetBeforeDoubleClick " , callbackFunc )`

**示例**

``` JavaScript
function func(Sh, Target, Cancal) {
    var msg = "您双击的是：" + Target.Address() +"\n是否取消此次双击？"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("SheetBeforeDoubleClick", func) 
```

### SheetBeforeRightClick

当右击任一工作表之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sh     | 必选      | Object    | 单击的工作表对象。                            |
| Target | 必选      | Range对象 | 单击区域所在的单元格对象。                    |
| Cancel | 必选      | Object    | 如果设置其属性Value为 true ，则取消此次单击。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetBeforeRightClick " , callbackFunc )`

**示例**

``` JavaScript
function func(Sh, Target, Cancal) {
    var msg = "您单击的是：" + Target.Address() +"\n是否取消此次单击？"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("SheetBeforeRightClick", func) 
```

### SheetCalculate

在任一工作表进行计算时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Sh   | 必选      | Object   | 激活的工作表对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetCalculate " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("进行计算")
}                                       
Application.ApiEvent.AddApiEventListener("SheetCalculate", func) 
```

### WorkbookActivate

任一工作簿被激活时，将触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookActivate " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("工作簿激活")
}                                       
Application.ApiEvent.AddApiEventListener("WorkbookActivate", func) 
```

### WorkbookDeactivate

任一工作簿被切换到非激活状态时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookDeactivate " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("工作簿切换为未激活")
}                                       
Application.ApiEvent.AddApiEventListener("WorkbookDeactivate", func) 
```

### WorkbookNewSheet

在任一打开的工作簿中创建新工作表时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |
| Sh   | 必选      | Object   | 激活工作表对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WorkbookNewSheet " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("新建工作表")
}                                       
Application.ApiEvent.AddApiEventListener("WorkbookNewSheet", func) 
```

### WindowResize

任一工作簿窗口调整大小时将触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |
| Wn   | 必选      | Window   | 激活窗口对象。   |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowResize " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("窗口调整大小")
}                                       
Application.ApiEvent.AddApiEventListener("WindowResize", func) 
```

### SheetFollowHyperlink

单击任一工作表的超链接时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Sh   | 必选      | Object   | 激活工作表对象。 |
| Hl   | 必选      | Object   | 超链接对象。     |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetFollowHyperlink " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("跳转超链接")
}                                       
Application.ApiEvent.AddApiEventListener("SheetFollowHyperlink", func) 
```

### AfterCalculate

当计算完成后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterCalculate " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("计算完成")
}                                       
Application.ApiEvent.AddApiEventListener("AfterCalculate", func) 
```

### SheetBeforeDelete

当删除任一工作表之前触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Sh   | 必选      | Object   | 激活工作表对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SheetBeforeDelete " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("准备删除工作表")
}                                       
Application.ApiEvent.AddApiEventListener("SheetBeforeDelete", func) 
```

### DocumentRightsInfo

当操作文档权限时会触发此事件。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明         |
|------------|-----------|----------|--------------|
| Wb         | 必选      | Workbook | 工作簿对象。 |
| RightsInfo | 必选      | Int      | 权限信息。   |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentRightsInfo " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("操作文档权限")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentRightsInfo", func) 
```

### AfterLogin

当用户登陆之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterLogin " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("用户已登陆")
}                                       
Application.ApiEvent.AddApiEventListener("AfterLogin", func) 
```

### AfterLogout

当用户登出之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterLogout " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("用户已登出")
}                                       
Application.ApiEvent.AddApiEventListener("AfterLogout", func) 
```

### AfterVIPInfoChanged

当用户的会员信息变化之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterVIPInfoChanged " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("用户会员信息已变化")
}                                       
Application.ApiEvent.AddApiEventListener("AfterVIPInfoChanged", func) 
```

### AfterWebExtWatchingDataUpdated

当网页拓展监控的数据发生变化之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterWebExtWatchingDataUpdated " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("网页拓展监控的数据发生变化")
}                                       
Application.ApiEvent.AddApiEventListener("AfterWebExtWatchingDataUpdated", func) 
```

### AfterTaskPaneShow

当任务窗格显示之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterTaskPaneShow " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("任务窗格已显示")
}                                       
Application.ApiEvent.AddApiEventListener("AfterTaskPaneShow", func) 
```

### AfterTaskPaneHidden

当任务窗格隐藏之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterTaskPaneHidden " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("任务窗格已隐藏")
}                                       
Application.ApiEvent.AddApiEventListener("AfterTaskPaneHidden", func) 
```

### DocumentBeforeNew

在任一文档新建之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不新建工作簿。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeNew " , callbackFunc )`

**示例**

``` JavaScript
function func(Cancal) {
    var msg = "是否要新建工作簿"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                   
Application.ApiEvent.AddApiEventListener("DocumentBeforeNew", func) 
```

### DocumentBeforeOpen

在任一文档打开之前触发该事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|-----------------------------------------------|
| FilePath | 必选      | String   | 工作簿路径                                    |
| Cancel   | 必选      | Object   | 如果设置其属性Value为 true ，则不打开工作簿。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeOpen " , callbackFunc )`

**示例**

``` JavaScript
function func(FilePath, Cancal) {
    var msg = "是否要打开文档：" + FilePath
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeOpen", func) 
```

### DocumentBeforeCopy

在任一文档复制之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| Wb     | 必选      | Workbook | 工作簿对象。                            |
| Type   | 必选      | Long     | 复制类型。                              |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不复制。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeCopy " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("准备复制")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeCopy", func) 
```

### DocumentBeforePaste

在任一文档粘贴之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| Wb     | 必选      | Workbook | 工作簿对象。                            |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不粘贴。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforePaste " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("准备粘贴")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforePaste", func) 
```

### FileAfterSave

在任一文档保存之后触发该事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明         |
|----------|-----------|----------|--------------|
| Wb       | 必选      | Workbook | 工作簿对象。 |
| FilePath | 必选      | String   | 保存路径。   |
| Format   | 必选      | String   | 保存格式。   |

**语法**

`Application.ApiEvent.AddApiEventListener( " FileAfterSave " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("文档已保存")
}                                       
Application.ApiEvent.AddApiEventListener("FileAfterSave", func) 
```

### LinkedDataTypeConvert

在连接数据类型转变时触发该事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " LinkedDataTypeConvert " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("连接数据类型转变")
}                                       
Application.ApiEvent.AddApiEventListener("LinkedDataTypeConvert", func) 
```

### LinkedDataTypeChange

在连接数据类型改变时触发该事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " LinkedDataTypeChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("连接数据类型改变")
}                                       
Application.ApiEvent.AddApiEventListener("LinkedDataTypeChange", func) 
```

### LinkedDataTypeRefresh

在连接数据类型刷新时触发该事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " LinkedDataTypeRefresh " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("连接数据类型刷新")
}                                       
Application.ApiEvent.AddApiEventListener("LinkedDataTypeRefresh", func) 
```

### LinkedDataTypeCancel

在连接数据类型取消时触发该事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " LinkedDataTypeCancel " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("连接数据类型取消")
}                                       
Application.ApiEvent.AddApiEventListener("LinkedDataTypeCancel", func) 
```

### DocumentAfterPrint

该文档打印之后触发该事件。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明           |
|--------------|-----------|----------|----------------|
| Wb           | 必选      | Workbook | 工作簿对象。   |
| PrintOutPage | 必选      | Object   | 打印页面对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentAfterPrint " , callbackFunc )`

**示例**

``` JavaScript
let func = (wb)=>{
    alert("文档已打印")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentAfterPrint", func) 
```

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/事件/表格事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E8%A1%A8%E6%A0%BC%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# 文字事件

## 事件列表

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

## 事件

### DocumentOpen

文档打开时触发。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Doc  | 必选      | Document | 打开的文档对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentOpen " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("打开文档")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentOpen", func) 
```

### DocumentBeforeClose

任一打开的文档关闭之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                        |
|--------|-----------|----------|---------------------------------------------|
| Doc    | 必选      | Document | 文档对象。                                  |
| Cancel | 必选      | Boolean  | 如果设置其属性Value为 true ，则不关闭文档。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeClose " , callbackFunc )`

**示例**

``` JavaScript
function func(doc, cancal) {
    var msg = "是否要关闭文档？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};                                      
Application.ApiEvent.AddApiEventListener("DocumentBeforeClose", func) 
```

### DocumentBeforeSave

任一打开的文档保存之前触发此事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明|
|----------|-----------|----------|--------------------------------------------------------------|
| Doc      | 必选      | Document | 关闭的文档对象。                                             |
| SaveAsUI | 必选      | Boolean  | 如果为 true 则表示此次保存操作将会弹出保存或是另存为对话框。 |
| Cancel   | 必选      | Object   | 如果设置其属性Value为 true ，则不关闭文档。                  |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeSave " , callbackFunc )`

**示例**

``` JavaScript
function func(doc, cancal) {
    var msg = "您保存后要关闭吗？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};                                      
Application.ApiEvent.AddApiEventListener("DocumentBeforeSave", func) 
```

### DocumentAfterClose

任一打开的文档关闭之后触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明       |
|------|-----------|----------|------------|
| Doc  | 必选      | Document | 文档对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentAfterClose " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档已关闭")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentAfterClose", func) 
```

### DocumentChange

任一打开的文档关闭之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档内容发生变化")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentChange", func) 
```

### DocumentBeforePrint

任一打开的文档打印之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                        |
|--------|-----------|----------|---------------------------------------------|
| Doc    | 必选      | Document | 文档对象。                                  |
| Cancel | 必选      | Boolean  | 如果设置其属性Value为 true ，则不打印文档。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforePrint " , callbackFunc )`

**示例**

``` JavaScript
function func(doc, cancal) {
    var msg = "您要打印文档吗吗？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};                                      
Application.ApiEvent.AddApiEventListener("DocumentBeforePrint", func) 
```

### DocumentNew

新建任一文档时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Doc  | 必选      | Document | 新建的文档对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentNew " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("新建文档")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentNew", func) 
```

### DocumentRightsInfo

当操作文档权限时会触发此事件。

**参数**

| 名称       | 必选/可选 | 数据类型 | 说明             |
|------------|-----------|----------|------------------|
| Doc        | 必选      | Document | 新建的文档对象。 |
| RightsInfo | 必选      | Int      | 权限信息。       |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentRightsInfo " , callbackFunc )`

**示例**

``` JavaScript
function func(doc, info) {
    var msg = "文档的权限为：" + info;
    alert(msg);
}                                       
Application.ApiEvent.AddApiEventListener("DocumentRightsInfo", func) 
```

### DocumentBeforeCopy

在任一文档复制之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| Doc    | 必选      | Document | 文档对象。                              |
| Type   | 必选      | Long     | 复制类型。                              |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不复制。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeCopy " , callbackFunc )`

**示例**

``` JavaScript
function func(Doc, Type, Cancal) {
    var msg = "是否要复制文档内容"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeCopy", func) 
```

### DocumentBeforePaste

在任一文档粘贴之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| Doc    | 必选      | Document | 文档对象。                              |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不粘贴。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforePaste " , callbackFunc )`

**示例**

``` JavaScript
function func(Doc, Cancal) {
    var msg = "是否要粘贴文档内容"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforePaste", func) 
```

### DocumentBeforeNew

在任一文档新建之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                        |
|--------|-----------|----------|---------------------------------------------|
| Cancel | 必选      | Boolean  | 如果设置其属性Value为 true ，则不新建文档。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeNew " , callbackFunc )`

**示例**

``` JavaScript
function func(Cancal) {
    var msg = "是否要新建文档"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeNew", func) 
```

### DocumentBeforeOpen

在任一文档打开之前触发该事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                        |
|----------|-----------|----------|---------------------------------------------|
| FilePath | 必选      | String   | 文档路径                                    |
| Cancel   | 必选      | Object   | 如果设置其属性Value为 true ，则不打开文档。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeOpen " , callbackFunc )`

**示例**

``` JavaScript
function func(FilePath, Cancal) {
    var msg = "是否要打开文档" + FilePath
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeOpen", func) 
```

### DocumentAfterSave

在任一文档保存之后触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                      |
|--------|-----------|----------|-------------------------------------------|
| Doc    | 必选      | Document | 文档对象。                                |
| Result | 必选      | Boolean  | 如果保存操作成功，则为true; 否则为false。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentAfterSave " , callbackFunc )`

**示例**

``` JavaScript
function func(Result) {
    if(!Result)
        alert("保存成功")
};                                      
Application.ApiEvent.AddApiEventListener("DocumentAfterSave", func) 
```

### DocumentViewFocusIn

在任一文档获取到焦点之后触发该事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明       |
|------|-----------|----------|------------|
| Doc  | 必选      | Document | 文档对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentViewFocusIn " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档获取到焦点")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentViewFocusIn", func) 
```

### DocumentViewFocusOut

在任一文档失去焦点之后触发该事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明       |
|------|-----------|----------|------------|
| Doc  | 必选      | Document | 文档对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentViewFocusOut " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档失去焦点")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentViewFocusOut", func) 
```

### DocumentAfterPrint

在任一文档打印之后触发该事件。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明           |
|--------------|-----------|----------|----------------|
| Doc          | 必选      | Document | 文档对象。     |
| PrintOutPage | 必选      | Object   | 打印页面对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentAfterPrint " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档已打印")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentAfterPrint", func) 
```

### WindowActivate

当激活文档窗口时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明           |
|------|-----------|----------|----------------|
| Doc  | 必选      | Document | 激活文档对象。 |
| Wn   | 必选      | Window   | 激活窗口对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowActivate " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, wn) {
    var msg = "您激活的窗口是: " + wn.Caption;
    alret(msg);
}                                       
Application.ApiEvent.AddApiEventListener("WindowActivate", func) 
```

### WindowDeactivate

任一文档窗口被切换到非激活状态时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明           |
|------|-----------|----------|----------------|
| Doc  | 必选      | Document | 激活文档对象。 |
| Wn   | 必选      | Window   | 激活窗口对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowDeactivate " , callbackFunc )`

**示例**

``` JavaScript
function func(wb, wn) {
    var msg = "您取消激活的窗口是: " + wn.Caption;
    alret(msg);
}                                       
Application.ApiEvent.AddApiEventListener("WindowDeactivate", func) 
```

### WindowSelectionChange

活动窗口所选内容更改时将触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明       |
|------|-----------|----------|------------|
| Sel  | 必选      | Object   | 选择区域。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowSelectionChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("窗口所选内容更改")
}                                       
Application.ApiEvent.AddApiEventListener("WindowSelectionChange", func) 
```

### WindowBeforeRightClick

当右击任一文档窗口之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                      |
|--------|-----------|----------|-------------------------------------------|
| Sel    | 必选      | Object   | 选择区域。                                |
| Cancel | 必选      | Boolean  | 如果设置其属性Value为 true ，则取消单击。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowBeforeRightClick " , callbackFunc )`

**示例**

``` JavaScript
function func(Cancal) {
    var msg = "是否要取消单击"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("WindowBeforeRightClick", func) 
```

### WindowBeforeDoubleClick

当双击任一文档窗口之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                      |
|--------|-----------|----------|-------------------------------------------|
| Sel    | 必选      | Object   | 选择区域。                                |
| Cancel | 必选      | Boolean  | 如果设置其属性Value为 true ，则取消双击。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowBeforeDoubleClick " , callbackFunc )`

**示例**

``` JavaScript
function func(Cancal) {
    var msg = "是否要取消双击"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("WindowBeforeDoubleClick", func) 
```

### ApplicationQuit

当退出文字组件进程时触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ApplicationQuit " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文字组件退出")
}                                       
Application.ApiEvent.AddApiEventListener("ApplicationQuit", func) 
```

### AfterLogin

在用户登陆之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterLogin " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("用户已登陆")
}                                       
Application.ApiEvent.AddApiEventListener("AfterLogin", func) 
```

### AfterLogout

在用户登出之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterLogout " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("用户已登出")
}                                       
Application.ApiEvent.AddApiEventListener("AfterLogout", func) 
```

### AfterVIPInfoChanged

在用户会员信息变化之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterVIPInfoChanged " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("用户会员信息已变化")
}                                       
Application.ApiEvent.AddApiEventListener("AfterVIPInfoChanged", func) 
```

### AfterWebExtensionDataRangeChange

在网络扩展数据范围变化之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterWebExtensionDataRangeChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("网络扩展数据范围已变化")
}                                       
Application.ApiEvent.AddApiEventListener("AfterWebExtensionDataRangeChange", func) 
```

### AfterTaskPaneShow

当任务窗格显示之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterTaskPaneShow " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("任务窗格已显示")
}                                       
Application.ApiEvent.AddApiEventListener("AfterTaskPaneShow", func) 
```

### AfterTaskPaneHidden

当任务窗格隐藏之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterTaskPaneHidden " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("任务窗格已隐藏")
}                                       
Application.ApiEvent.AddApiEventListener("AfterTaskPaneHidden", func) 
```

### QueryDocumentPrintCopies

当打印复本排队时触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| Doc    | 必选      | Document | 文档对象。                              |
| Copies | 必选      | Int      | 要打印的份数。                          |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不打印。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " QueryDocumentPrintCopies " , callbackFunc )`

**示例**

``` JavaScript
function func(Doc, Copies, Cancal) {
    var msg = "是否要打印文档"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("QueryDocumentPrintCopies", func) 
```

### ContentChange

当文档内容发生变化时触发此事件。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明       |
|-------|-----------|----------|------------|
| Doc   | 必选      | Document | 文档对象。 |
| Range | 必选      | Object   | 范围对象。 |
| Type  | 必选      | Int      | 变化类型。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("文档内容发生变化")
}                                       
Application.ApiEvent.AddApiEventListener("ContentChange", func) 
```

### SelectionAfterStyleChange

当选择的内容字体变化之后触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明       |
|------|-----------|----------|------------|
| Sel  | 必选      | Object   | 选择区域。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SelectionAfterStyleChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("内容字体变化")
}                                       
Application.ApiEvent.AddApiEventListener("SelectionAfterStyleChange", func) 
```

### FileAfterSave

在文件保存之后触发该事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明       |
|----------|-----------|----------|------------|
| Doc      | 必选      | Document | 文档对象。 |
| FilePath | 必选      | String   | 保存路径。 |
| Format   | 必选      | String   | 保存格式。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " FileAfterSave " , callbackFunc )`

**示例**

``` JavaScript
function func(Doc, FilePath, Format) {
    var msg = "文档以" + Format + "格式保存在" + FilePath
    alret(msg);
}                                       
Application.ApiEvent.AddApiEventListener("FileAfterSave", func) 
```

### ContentControlAfterAdd

当内容控件被添加到文档之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlAfterAdd " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("内容控件被添加")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlAfterAdd", func) 
```

### ContentControlBeforeDelete

当内容控件被从文档删除之前触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlBeforeDelete " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("内容控件被从文档删除")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlBeforeDelete", func) 
```

### ContentControlOnExit

当退出内容控件时触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlOnExit " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("退出内容控件")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlOnExit", func) 
```

### ContentControlOnEnter

当进入内容控件时触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlOnEnter " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("进入内容控件")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlOnEnter", func) 
```

### ContentControlBeforeStoreUpdate

当内容控件的内容去更新XML映射内容时触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlBeforeStoreUpdate " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("内容控件的内容更新XML映射内容")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlBeforeStoreUpdate", func) 
```

### ContentControlBeforeContentUpdate

当内容控件的内容发生来自XML映射内容更新时触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " ContentControlBeforeContentUpdate " , callbackFunc )`

**示例**

``` JavaScript
let func = (doc)=>{
    alert("内容控件的内容发生来自XML映射内容更新")
}                                       
Application.ApiEvent.AddApiEventListener("ContentControlBeforeContentUpdate", func) 
```

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/事件/文字事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%96%87%E5%AD%97%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# 演示事件

## 事件列表

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

## 事件

### AfterLogin

用户登录之后触发该事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterLogin " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("用户登录")
}                                       
Application.ApiEvent.AddApiEventListener("AfterLogin", func)
```

### AfterLogout

用户登出之后触发该事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterLogout " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("用户登出")
}                                       
Application.ApiEvent.AddApiEventListener("AfterLogout", func)
```

### PresentationAfterOpen

任一演示文稿打开后触发该事件。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明       |
|------|-----------|--------------|------------|
| Pre  | 必选      | Presentation | 演示对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " PresentationAfterOpen " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("演示文稿打开")
}                                       
Application.ApiEvent.AddApiEventListener("PresentationAfterOpen", func)
```

### AfterTaskPaneHidden

当任务窗格隐藏之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterTaskPaneHidden " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("任务窗格隐藏")
}                                       
Application.ApiEvent.AddApiEventListener("AfterTaskPaneHidden", func)
```

### AfterTaskPaneShow

当任务窗格显示之后触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " AfterTaskPaneShow " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("任务窗格显示")
}                                       
Application.ApiEvent.AddApiEventListener("AfterTaskPaneShow", func)
```

### DocumentAfterPrint

当文档打印之后触发此事件。

**参数**

| 名称         | 必选/可选 | 数据类型     | 说明           |
|--------------|-----------|--------------|----------------|
| Pre          | 必选      | Presentation | 演示对象。     |
| PrintOutPage | 必选      | Object       | 打印页面对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentAfterPrint " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("文档打印")
}                                       
Application.ApiEvent.AddApiEventListener("DocumentAfterPrint", func)
```

### DocumentBeforeCopy

在任一文档复制之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型     | 说明                                    |
|--------|-----------|--------------|-----------------------------------------|
| Pre    | 必选      | Presentation | 演示对象。                              |
| Type   | 必选      | Long         | 复制类型。                              |
| Cancel | 必选      | Object       | 如果设置其属性Value为 true ，则不复制。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeCopy " , callbackFunc )`

**示例**

``` JavaScript
function func(Doc, Type, Cancal) {
    var msg = "是否要复制文档内容"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeCopy", func)
```

### DocumentBeforeNew

在任一文档新建之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                    |
|--------|-----------|----------|-----------------------------------------|
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则不新建。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeNew " , callbackFunc )`

**示例**

``` JavaScript
function func(Cancal) {
    var msg = "是否要新建文档"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeNew", func)
```

### DocumentBeforeOpen

在任一文档打开之前触发该事件。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                            |
|----------|-----------|----------|-------------------------------------------------|
| FilePath | 必选      | String   | 演示路径                                        |
| Cancel   | 必选      | Object   | 如果设置其属性Value为 true ，则不打开演示文档。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforeOpen " , callbackFunc )`

**示例**

``` JavaScript
function func(FilePath, Cancal) {
    var msg = "是否要打开文档" + FilePath
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforeOpen", func)
```

### DocumentBeforePaste

在任一文档粘贴之前触发该事件。

**参数**

| 名称   | 必选/可选 | 数据类型     | 说明                                    |
|--------|-----------|--------------|-----------------------------------------|
| Pre    | 必选      | Presentation | 演示对象。                              |
| Cancel | 必选      | Object       | 如果设置其属性Value为 true ，则不粘贴。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentBeforePaste " , callbackFunc )`

**示例**

``` JavaScript
function func(Pre, Cancal) {
    var msg = "是否要粘贴文档内容"
    if (confirm(msg) == true) {
        Cancal.Value = false;
    } else {
        Cancal.Value = true;
    }
}                                       
Application.ApiEvent.AddApiEventListener("DocumentBeforePaste", func)
```

### DocumentRightsInfo

当操作文档权限时会触发此事件。

**参数**

| 名称       | 必选/可选 | 数据类型     | 说明       |
|------------|-----------|--------------|------------|
| Pre        | 必选      | Presentation | 演示对象。 |
| RightsInfo | 必选      | Int          | 权限信息。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " DocumentRightsInfo " , callbackFunc )`

**示例**

``` JavaScript
function func(pre, info) {
    var msg = "文档的权限为：" + info;
    alert(msg);
}                                   
Application.ApiEvent.AddApiEventListener("DocumentRightsInfo", func)
```

### FileAfterSave

在文件保存之后触发该事件。

**参数**

| 名称     | 必选/可选 | 数据类型     | 说明       |
|----------|-----------|--------------|------------|
| Pre      | 必选      | Presentation | 演示对象。 |
| FilePath | 必选      | String       | 保存路径。 |
| Format   | 必选      | String       | 保存格式。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " FileAfterSave " , callbackFunc )`

**示例**

``` JavaScript
function func(Pre, FilePath, Format) {
    var msg = "文档以" + Format + "格式保存在" + FilePath
    alret(msg);
}                                       
Application.ApiEvent.AddApiEventListener("FileAfterSave", func)
```

### NewPresentation

当新建演示文稿时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明       |
|------|-----------|--------------|------------|
| Pre  | 必选      | Presentation | 演示对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " NewPresentation " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("新建演示文稿")
}                                       
Application.ApiEvent.AddApiEventListener("NewPresentation", func)
```

### PresentationBeforeClose

任一打开的演示文稿关闭之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型     | 说明                                        |
|--------|-----------|--------------|---------------------------------------------|
| Pre    | 必选      | Presentation | 演示对象。                                  |
| Cancel | 必选      | Boolean      | 如果设置其属性Value为 true ，则不关闭演示。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " PresentationBeforeClose " , callbackFunc )`

**示例**

``` JavaScript
function func(pre, cancal) {
    var msg = "是否要关闭文档？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};                                      
Application.ApiEvent.AddApiEventListener("PresentationBeforeClose", func)
```

### PresentationBeforePrint

任一演示文稿打印之前触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型     | 说明                                          |
|--------|-----------|--------------|-----------------------------------------------|
| Pre    | 必选      | Presentation | 演示对象。                                    |
| Cancel | 必选      | Object       | 如果设置其属性Value为 true ，则不打印工作簿。 |

**语法**

`Application.ApiEvent.AddApiEventListener( "PresentationBeforePrint" , callbackFunc )`

**示例**

``` JavaScript
function func(pre, cancal) {
    var msg = "是否要取消打印？";
    if (confirm(msg) == true) {
        cancal.Value = false;
    } else {
        cancal.Value = true;
    }
};                                      
Application.ApiEvent.AddApiEventListener("PresentationBeforePrint", func)
```

### PresentationClose

任一演示文稿关闭时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明       |
|------|-----------|--------------|------------|
| Pre  | 必选      | Presentation | 演示对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " PresentationClose " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("演示文稿关闭")
}                                       
Application.ApiEvent.AddApiEventListener("PresentationClose", func)
```

### PresentationCloseFinal

任一演示文稿关闭之后触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明       |
|------|-----------|--------------|------------|
| Pre  | 必选      | Presentation | 演示对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " PresentationCloseFinal " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("演示文稿关闭")
}                                       
Application.ApiEvent.AddApiEventListener("PresentationCloseFinal", func)
```

### PresentationNewSlide

任一演示文稿新建幻灯片时触发此事件。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明         |
|-------|-----------|----------|--------------|
| Slide | 必选      | Slide    | 幻灯片对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " PresentationNewSlide " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("新建幻灯片")
}                                       
Application.ApiEvent.AddApiEventListener("PresentationNewSlide", func)
```

### PresentationOpen

任一演示文稿打开时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明       |
|------|-----------|--------------|------------|
| Pre  | 必选      | Presentation | 演示对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " PresentationOpen " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("演示文稿打开")
}                                       
Application.ApiEvent.AddApiEventListener("PresentationOpen", func)
```

### PresentationPrint

任一演示文稿打印时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明       |
|------|-----------|--------------|------------|
| Pre  | 必选      | Presentation | 演示对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " PresentationPrint " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("演示文稿打印")
}                                       
Application.ApiEvent.AddApiEventListener("PresentationPrint", func)
```

### PresentationSave

任一演示文稿保存时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明       |
|------|-----------|--------------|------------|
| Pre  | 必选      | Presentation | 演示对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " PresentationSave " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("演示文稿保存")
}                                       
Application.ApiEvent.AddApiEventListener("PresentationSave", func)
```

### Quit

任一演示文稿退出时触发此事件。

**参数**

无

**语法**

`Application.ApiEvent.AddApiEventListener( " Quit " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("演示文稿退出")
}                                       
Application.ApiEvent.AddApiEventListener("Quit", func)
```

### SlideSelectionChanged

演示文稿幻灯片选择变化时触发此事件。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明       |
|-------|-----------|----------|------------|
| Range | 必选      | Object   | 选择对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SlideSelectionChanged " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("幻灯片选择变化")
}                                       
Application.ApiEvent.AddApiEventListener("SlideSelectionChanged", func)
```

### SlideShowBegin

演示文稿开始播放时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Shw  | 必选      | Object   | 幻灯片窗口对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SlideShowBegin " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("演示文稿开始播放")
}                                       
Application.ApiEvent.AddApiEventListener("SlideShowBegin", func)
```

### SlideShowEnd

演示文稿结束播放时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明       |
|------|-----------|--------------|------------|
| Pre  | 必选      | Presentation | 演示对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SlideShowEnd " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("演示文稿结束播放")
}                                       
Application.ApiEvent.AddApiEventListener("SlideShowEnd", func)
```

### SlideShowNextClick

演示文稿播放时下一次点击触发此事件。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明             |
|--------|-----------|----------|------------------|
| Shw    | 必选      | Object   | 幻灯片窗口对象。 |
| Effect | 必选      | Object   | 操作对象         |

**语法**

`Application.ApiEvent.AddApiEventListener( " SlideShowNextClick " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("点击")
}                                       
Application.ApiEvent.AddApiEventListener("SlideShowNextClick", func)
```

### SlideShowNextSlide

演示文稿播放下一个幻灯片时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Shw  | 必选      | Object   | 幻灯片窗口对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SlideShowNextSlide " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("播放下一个幻灯片")
}                                       
Application.ApiEvent.AddApiEventListener("SlideShowNextSlide", func)
```

### SlideShowOnNext

演示文稿播放点击下一页时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Shw  | 必选      | Object   | 幻灯片窗口对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SlideShowOnNext " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("点击下一页")
}                                       
Application.ApiEvent.AddApiEventListener("SlideShowOnNext", func)
```

### SlideShowOnPrevious

演示文稿播放点击上一页时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Shw  | 必选      | Object   | 幻灯片窗口对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " SlideShowOnPrevious " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("点击上一页")
}                                       
Application.ApiEvent.AddApiEventListener("SlideShowOnPrevious", func)
```

### WindowActivate

当激活演示窗口时触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明       |
|------|-----------|--------------|------------|
| Pre  | 必选      | Presentation | 演示对象。 |
| Wn   | 必选      | Object       | 窗口对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowActivate " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("激活演示窗口")
}                                       
Application.ApiEvent.AddApiEventListener("WindowActivate", func)
```

### WindowSelectionChange

活动窗口所选内容更改时将触发此事件。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明           |
|------|-----------|----------|----------------|
| Sel  | 必选      | Object   | 选择范围对象。 |

**语法**

`Application.ApiEvent.AddApiEventListener( " WindowSelectionChange " , callbackFunc )`

**示例**

``` JavaScript
let func = (slide)=>{
    alert("所选内容更改")
}                                       
Application.ApiEvent.AddApiEventListener("WindowSelectionChange", func)
```

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/事件/演示事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%BC%94%E7%A4%BA%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# JavaScript运行时说明

WPS宏编辑器集成了一个V8 引擎的 JavaScript 运行时，支持大部分ES6语法，因此宏编辑器支持JavaScript 标准内置对象， **注意** ，JS内置对象和浏览器的内置对象是不同的，WPS宏编辑器集成的是 JavaScript 运行时，而不是浏览器， 因此 WPS宏编辑器 不支持 浏览器的内置对象 。 具体API参见 <https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference/Global_Objects> 。如以下代码输出圆周率，可以在WPS宏编辑器中执行：

``` JavaScript
function abcd()
{
    Console.log(Math.PI);
}
```

使用过程中遇到的编译错误和语法错误具体原因和细节可以在 <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Errors> 上查看。

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/JavaScript运行时说明](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/JavaScript%E8%BF%90%E8%A1%8C%E6%97%B6%E8%AF%B4%E6%98%8E.html)

------------------------------------------------------------------------
# 事件概述

通过 WPS宏编辑器能够对 WPS 应用程序发出的事件添加 JavaScript 方法进行处理。 在通知型事件中，可以接收已经发生变化，比如通过WindowActivate事件，可以对文档的切换做一些功能性处理； 在询问型事件中，可以控制是否继续执行当前操作，比如通过WorkbookBeforeClose事件，可以取消文档的关闭。

``` JavaScript
function Application_WorkbookBeforeClose(Wb, Cancel)
{
    if (!Wb.Saved) {
        MsgBox("请先保存文档！")
        Cancel.Value = true;
    }
}
```

**代码说明**

注册WorkbookBeforeClose监听事件，在工作簿关闭时，判断该工作簿是否保存，如果未保存，弹出“请先保存文档！”提示，并取消关闭

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/事件概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E4%BA%8B%E4%BB%B6%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# Application 事件

## 事件列表

| 名称                   | 触发时机                                                    |
|------------------------|-------------------------------------------------------------|
| NewWorkbook            | 当新建工作簿时触发此事件。                                  |
| SheetActivate          | 当激活任一工作表时 触发 此事件。                            |
| SheetBeforeDelete      | 当删除任 一 工作表之前 触发 此事件。                        |
| SheetBeforeDoubleClick | 当双击任一工作表之前 触发 此事件。                          |
| SheetBeforeRightClick  | 当右击任一 工作表之前 触发 此事件 。                        |
| SheetCalculate         | 在任一工作表进行 计算时 触发此事件。                        |
| SheetChange            | 当用户或外部链接更改了任一 工作表中的单元格时 触发 此事件。 |
| SheetDeactivate        | 当任 一 工作表被切换到非激活 状态 时 触发 此事件。          |
| SheetFollowHyperlink   | 单击 任一 工作表的超链接 时 触 发 此事 件。                 |
| SheetSelectionChange   | 任一工作表上的选定区域发生更改时，将触发此事件。            |
| WindowActivate         | 任一工作簿窗口被激活时，将 触发 此事件。                    |
| WindowDeactivate       | 任一 工 作簿窗口 被切换到非激活状态 时 触发 此事件。        |
| WindowResize           | 任一工作簿窗口调整大小时将 触发 此事件。                    |
| WorkbookActivate       | 任一 工 作簿被激活时，将 触发 此事件。                      |
| WorkbookAfterSave      | 任一 工作簿被保存之后触发此事件。                           |
| WorkbookBeforeClose    | 任一打开的工作簿关闭之前触发此事件。                        |
| WorkbookBeforePrint    | 在打印任一打开的工作簿之前 触发 此事件。                    |
| WorkbookBeforeSave     | 任一 工作簿被保存 之前 触发 此事件。                        |
| WorkbookDeactivate     | 任一 工 作簿被切换到非激活状态时触发此事件。                |
| WorkbookNewSheet       | 在任一 打开的工作簿中创建新工作表时 触发此事件。            |
| WorkbookOpen           | 当打开一个工作簿时 触发 此事件。                            |

## 事件

### NewWorkbook

当新建一个工作簿时触发此事件。

**语法**

`function Application_NewWorkbook ( Wb ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 新建的工作簿对象 |

**示例**

当新建一个工作簿时，弹消息框提醒用户新建了工作簿。

``` JavaScript
function Application_NewWorkbook(Wb)
{
    MsgBox("您新建了工作簿："+Wb.Name)
}
```

### SheetActivate

当激活任一工作表时 触发 此事件。

**语法**

`function Application_SheetActivate ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Sh   | 必选      | Object   | 激活的工作表对象。 |

**示例**

当新建一个工作簿时，弹消息框提醒用户 激活 了 工作表 。

``` JavaScript
function Application_SheetActivate(Sh)
{
    MsgBox("您激活了工作表："+Sh.Name)
}
```

### SheetBeforeDelete

当删除任 一 工作表之前 触发 此事件。

**语法**

`function Application_SheetBeforeDelete ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                   |
|------|-----------|----------|------------------------|
| Sh   | 必选      | Object   | 即将删除的工作表对象。 |

**示例**

当删除任 一 工作表之前 ，弹消息框提醒用户。

``` JavaScript
function Application_SheetBeforeDelete(Sh)
{
    MsgBox("您即将删除工作表："+Sh.Name)
} 
```

### SheetBeforeDoubleClick

双击任一工作表之前触发此事件。

**语法**

`function Application_SheetBeforeDoubleClick ( Sh, Target, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sh     | 必选      | Object    | 双击的工作表对象。                            |
| Target | 必选      | Range对象 | 双击区域所在的单元格对象。                    |
| Cancel | 必选      | O bject   | 如果设置其属性Value为 true ，则取消此次双击。 |

**示例**

双击任一工作表之前，如果双击操作在A1单元格上，则不响应双击操作（默认是进入编辑），否则则继续响应。

``` JavaScript
function Application_SheetBeforeDoubleClick(Sh, rg, cancel)
{
    if (Target.Address() == "$A$1")
        Cancel.Value = true;
} 
```

### SheetBeforeRightClick

右击任一工作表之前触发此事件。

**语法**

`function Application_SheetBeforeRightClick ( Sh, Target, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sh     | 必选      | Object    | 右击 的工作表对象。                           |
| Target | 必选      | Range对象 | 右击区域所在的单元格对象。                    |
| Cancel | 必选      | Object    | 如果设置其属性Value为 true ，则取消此次右击。 |

**示例**

右击任一工作表之前，如果右击操作在A1单元格上，则不响应右击操作（默认是弹出右键菜单），否则则继续右击。

``` JavaScript
function Application_SheetBeforeRightClick(Sh, rg, cancel)
{
    if (Target.Address() == "$A$1")
        Cancel.Value = true;
}
```

### SheetCalculate

任一工作表进行计算时触发此事件。

**语法**

`function Application_SheetCalculate ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Sh   | 必选      | Object   | 计算的工作表对象。 |

**示例**

当工作表进行计算时，弹消息框提醒用户。

``` JavaScript
function Application_SheetCalculate(Sh)
{
    MsgBox("工作表正在计算："+Sh.Name)
} 
```

### SheetChange

当用户或外部链接更改了任一工作表中的单元格时触发此事件。

**语法**

`function Application_SheetChange ( Sh, Target ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                  |
|--------|-----------|-----------|-----------------------|
| Sh     | 必选      | Object    | 修改的工作表对象。    |
| Target | 必 选     | Range对象 | 修改的单元格Range对象 |

**示例**

工作表中的单元格被修改时，弹消息框提醒用户。

``` JavaScript
function Application_SheetChange(Sh, rg)
{
    MsgBox("工作表:" + Sh.Name + "。区域："+Target.Address() + "。发生了修改")
}
```

### SheetDeactivate

当任一工作表被切换到非激活状态时触发此事件。

**语法**

`function Application_SheetDeactivate ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                             |
|------|-----------|----------|----------------------------------|
| Sh   | 必选      | Object   | 被切换到非激活状态的工作表对象。 |

**示例**

任一工作表被切换到非激活状态时，弹消息框提醒用户。

``` JavaScript
function Application_SheetDeactivate(Sh)
{
    MsgBox("被切换到非激活的工作表："+Sh.Name)
} 
```

### SheetFollowHyperlink

单击任一工作表的超链接时触发此事件。

**语法**

`function Application_SheetFollowHyperlink ( Sh, Target ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型      | 说明                              |
|--------|-----------|---------------|-----------------------------------|
| Sh     | 必选      | Object        | 超链接所在的工作表对象。          |
| Target | 必选      | Hyperlink对象 | Hyperlink对象，代表超链接的目标。 |

**示例**

单击任一工作表的超链接时，弹消息框提醒用户。

``` JavaScript
function Application_SheetFollowHyperlink(Sh, Target)
{
    MsgBox("你点击了工作表\"" + Sh.Name + "\"上的超链接："+Target.Address)
} 
```

### SheetSelectionChange

该工作簿任一工作表上的选定区域发生更改时，将触发此事件。

**语法**

`function Application_SheetSelectionChange ( Sh, Target ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                   |
|--------|-----------|----------|------------------------|
| Sh     | 必选      | Object   | 选区改变的工作表对象。 |
| Target | 必选      | Range    | 新选定的区域。         |

**示例**

当工作表上的选定区域发生更改时，弹消息框提醒用户。

``` JavaScript
function Application_SheetSelectionChange(Sh, Target)
{
    MsgBox("工作表\"" + Sh.Name + "\"选中的区域是："+Target.Address() + "。")
} 
```

### WindowActivate

任一工作簿窗口被激活时，将 触发 此事件。

**语法**

`function Application_WindowActivate ( Wb , Wn ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明             |
|------|-----------|----------|------------------|
| Wb   | 必选      | Workbook | 激活工作簿对象。 |
| Wn   | 必选      | Window   | 激活窗口 对象。  |

**示例**

任 一工作簿窗口被 激活时 ， 输出激活的窗口标题。

``` JavaScript
function Application_WindowActivate(Wb, Wn)
        { Debug.
        Print("当前激活窗口是："+ Wn.Caption) }
        
        
```

### WindowDeactivate

任一 工 作簿窗口被切换到非激活状态时触发此事件。

**语法**

`function Application_WindowDeactivate ( Wb , Wn ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                              |
|------|-----------|----------|-----------------------------------|
| Wb   | 必选      | Workbook | 被切换到非激活状态的 工作簿对象。 |
| Wn   | 必选      | Window   | 被切换到非激活状态的 窗口 对象。  |

**示例**

任一 工 作簿窗口被切换到非激活状态时 ，输出激活的窗口标题。

``` JavaScript
function Application_WindowDeactivate(Wb, Wn)
{
    Debug.Print("当前被切换到非激活窗口是："+ Wn.Caption)
} 
```

### WindowResize

任一工作簿窗口调整大小时将 触发 此事件。

**语法**

`function Application_WindowResize ( Wb , Wn ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                         |
|------|-----------|----------|------------------------------|
| Wb   | 必选      | Workbook | 调整窗口大小 的 工作簿对象。 |
| Wn   | 必选      | Window   | 调整窗口大小 的 窗口 对象。  |

**示例**

当触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_WindowResize(Wb, Wn)
        { MsgBox(
        "窗口\""+ Wn.Caption + "\"大小发生了改变。") } 
        
```

### Workbook Activate

工作簿被激活时，将触发此事件。

**语法**

`function Application_Workbook Activate ( Wb ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明            |
|------|-----------|----------|-----------------|
| Wb   | 必选      | Workbook | 被切换的 工作簿 |

**示例**

工作簿被激活时，弹消息框提醒用户。

``` JavaScript
function Application_WorkbookActivate(Wb)
{
    MsgBox("工作簿:" + Wb.Name +"被激活了。");
}
```

### Workbook AfterSave

工作簿被保存之后触发此事件。

**语法**

`function Application_WorkbookAfterSave ( Wb, Success ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                      |
|---------|-----------|----------|-------------------------------------------|
| Wb      | 必选      | Workbook | 保存的工作簿                              |
| Success | 必选      | Boolean  | 如果保存操作成功，则为true; 否则为false。 |

**示例**

工作簿被保存后，弹消息框提醒用户。

``` JavaScript
function Application_WorkbookAfterSave(Wb, Success)
{
    if (Success)
    {
        MsgBox("文档被成功保存！")
    }
}
```

### Workbook BeforeClose

有工作簿被关闭之前触发此事件。

**语法**

`function Application_WorkbookBeforeClose ( Wb, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Wb     | 必选      | Workbook | 关闭的工作簿                                  |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则取消此次关闭。 |

**示例**

有工作簿被关闭时，弹消息框询问用户是否取消，用户点击“是”就会取消关闭工作簿。

``` JavaScript
function Application_WorkbookBeforeClose(Wb, Cancel)
{
    var ret = MsgBox("工作簿:"+ Wb.Name+"正在关闭。是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
} 
```

### Workbook BeforePrint

有工作簿被打印之前触发此事件

**语法**

`function Application_WorkbookBeforePrint ( Wb, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Wb     | 必选      | Workbook | 打印的工作簿                                  |
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则取消此次打印。 |

**示例**

有工作簿被印前时，弹消息框询问用户是否取消，用户点击“是”就会取消打印工作簿。

``` JavaScript
function Application_WorkbookBeforePrint(Wb, Cancel)
{
    var ret = MsgBox("工作簿:"+ Wb.Name+"即将打印。是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### Workbook BeforeSave

有 工 作 簿被 保存之前触发此事件。

**语法**

`function Application_Workbook BeforeSave ( Wb, S aveAsUI ， Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明|
|-----------|-----------|----------|--------------------------------------------------------------|
| Wb        | 必选      | Workbook | 保存的工作簿                                                 |
| S aveAsUI | 必选      | Boolean  | 如果为 true 则表示此次保存操作将会弹出保存或是另存为对话框。 |
| Cancel    | 必选      | Object   | 如果设置其属性Value为 true ，则取消此次保存。                |

**示例**

有工作簿被保存之前，如果会弹出保存或是另存为对话框的情况，弹消息框询问用户是否取消，用户点击“是”就会取消保存工作簿。

``` JavaScript
function Application_WorkbookBeforeSave(Wb, SaveAsUI, Cancel)
{
    if (SaveAsUI){
        var ret = MsgBox("工作簿:"+ Wb.Name+"即将保存。是否取消？", jsYesNo);
        if (ret == jsResultYes)
            Cancel.Value = true;
    }
}
```

### Workbook Deactivate

工作簿被切换到非激活（非活动）状态时触发此事件。

**语法**

`function Application_WorkbookDeactivate ( Wb ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明            |
|------|-----------|----------|-----------------|
| Wb   | 必选      | Workbook | 被切换的 工作簿 |

**示例**

工作簿被切换到非激活（非活动）状态时，弹消息框提醒用户。

``` JavaScript
function Application_WorkbookDeactivate(Wb)
{
    MsgBox("工作簿:" + Wb.Name +"被切换到非激活了。");
}
```

### Workbook NewSheet

该工作簿中创建新工作表时触发此事件。

**语法**

`function Application_WorkbookNewSheet ( Wb, Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                   |
|------|-----------|----------|------------------------|
| Wb   | 必选      | Workbook | 新建工作表所在的工作簿 |
| Sh   | 必选      | Object   | 新建的工作表           |

**示例**

创建新工作表时，弹消息框提醒用户。

``` JavaScript
function Application_WorkbookNewSheet(Wb, Sh)
{
    MsgBox("新建了工作表：" + Sh.Name)
}
```

### Workbook Open

工作簿打开时触发此事件

**语法**

`function Application_WorkbookOpe n ( Wb ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明         |
|------|-----------|----------|--------------|
| Wb   | 必选      | Workbook | 打开的工作簿 |

示例该工作簿被打开时，弹消息框提醒用户。

``` JavaScript
function Application_WorkbookOpen(Wb)
{
    MsgBox("打开了工作簿：" + Wb.Name);
} 
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/表格事件/Application 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E8%A1%A8%E6%A0%BC%E4%BA%8B%E4%BB%B6/Application%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# Workbook 事件

## 事件列表

| 名称                   | 触发时机      |
|------------------------|------------------------------------------------------------------------|
| Activate               | 该工 作簿被激活时，将触发此事件。                                      |
| AfterSave              | 该工作簿被保存之后触发此事件。                                         |
| BeforeClose            | 该工作簿关闭之前触发此事件。                                           |
| BeforePrint            | 该工作簿打印之前触发此事件。                                           |
| BeforeSave             | 该工作簿保存之前触发此事件。                                           |
| Deactivate             | 该 工 作簿被切换到非激活状态时触发此事件。                             |
| NewSheet               | 该 工作簿中创建新工作表时 触发此事件。                                 |
| Open                   | 该 工作簿打开时 触发 此事件。                                          |
| SheetActivate          | 当激活 该 工作簿 任一工作表时 触发 此事件。                            |
| SheetBeforeDelete      | 删除 该 工作簿 任 一 工作表之前 触发 此事件。                          |
| SheetBeforeDoubleClick | 双击 该 工作簿 任一工作表之前 触发 此事件。                            |
| SheetBeforeRightClick  | 右击 该 工作簿 任一 工作表之前 触发 此事件 。                          |
| SheetCalculate         | 在 该 工作簿 任一工作表进行 计算时 触发此事件。                        |
| SheetChange            | 当用户或外部链接更改了 该 工作簿 任一 工作表中的单元格时 触发 此事件。 |
| SheetDeactivate        | 当 该 工作簿 任 一 工作表被切换到非激活状态时 触发 此事件。            |
| SheetFollowHyperlink   | 单击 该 工作簿 任一 工作表的超链接 时 触 发 此事 件。                  |
| SheetSelectionChange   | 该 工作簿 任一工作表上的选定区域发生更改时，将触发此事件。             |

## 事件

### Activate

该工 作簿被激活时，将触发此事件。

**语法**

`function Workbook_Activate () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**示例**

该工 作簿被激活时 ，弹消息框提醒用户。

``` JavaScript
function Workbook_Activate()
{
    MsgBox("您激活了当前工作簿")
} 
```

### AfterSave

该工作簿被保存之后触发此事件。

**语法**

`function Workbook_AfterSave ( Success ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                      |
|---------|-----------|----------|-------------------------------------------|
| Success | 必选      | Boolean  | 如果保存操作成功，则为true; 否则为false。 |

**示例**

该工 作簿被保存后， 弹消息框提醒用户。

``` JavaScript
function Workbook_AfterSave(Success)
{
    if (Success)
    {
        MsgBox("文档被成功保存！")
    }
} 
```

### BeforeClose

该工作簿关闭之前触发此事件。

**语法**

`function Workbook_BeforeClose ( Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则取消此次关闭。 |

**示例**

该工作簿被关闭时，弹消息框询问用户是否取消，用户点击“是”就会取消关闭 工作簿 。

``` JavaScript
function Workbook_BeforeClose(Cancel)
{
    var ret = MsgBox("工作簿正在关闭。是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### BeforePrint

该工作簿打印之前触发此事件

**语法**

`function Workbook_BeforePrint ( Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                                          |
|--------|-----------|----------|-----------------------------------------------|
| Cancel | 必选      | Object   | 如果设置其属性Value为 true ，则取消此次打印。 |

**示例**

该工作簿打印前 时，弹消息框询问 用 户是否取消， 用户点击“是”就会取消打印 工作簿 。

``` JavaScript
function Workbook_BeforePrint(Cancel)
        {
        var ret = MsgBox("工作簿即将打印。是否取消？", jsYesNo);
        if (ret == jsResultYes) Cancel.Value = true; }
        
        
```

### BeforeSave

该工作簿保存之前触发此事件。

**语法**

`function Workbook_BeforeSave ( S aveAsUI ， Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称      | 必选/可选 | 数据类型 | 说明|
|-----------|-----------|----------|--------------------------------------------------------------|
| S aveAsUI | 必选      | Boolean  | 如果为 true 则表示此次保存操作将会弹出保存或是另存为对话框。 |
| Cancel    | 必选      | Object   | 如果设置其属性Value为 true ，则取消此次保存。                |

**示例**

对该工作簿 保存之前 ， 如果 会弹出 保存或是另存为 对话框 的情况， 弹消息框询问用户是否取消，用户点击“是”就会取消保存工作簿。

``` JavaScript
function Workbook_BeforeSave(SaveAsUI, Cancel)
{
    if (SaveAsUI){
        var ret = MsgBox("工作簿即将保存。是否取消？", jsYesNo);
        if (ret == jsResultYes)
            Cancel.Value = true;
    }
} 
```

### Deactivate

该 工 作簿被切换到非激活 （非活动 ） 状态时触发此事件。

**语法**

`function Workbook_Deactivate () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**示例**

该工作簿被 切换到非激活（非活动）状态时 ，弹消息框提醒用户。

``` JavaScript
function Workbook_Deactivate()
{
    MsgBox("工作簿被切换到非激活了。");
} 
```

### NewSheet

该 工作簿中创建新工作表时 触发此事件。

**语法**

`function Workbook_NewSheet ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明         |
|------|-----------|----------|--------------|
| Sh   | 必选      | Object   | 新建的工作表 |

**示例**

该工作簿 创建新工作表 时，弹消息框提醒用户。

``` JavaScript
function Workbook_NewSheet(Sh)
{
    MsgBox("新建了工作表：" + Sh.Name)
} 
```

### Open

该 工作簿打开时 触发 此事件

**语法**

`function Workbook_Open () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**示例**

该工作簿被打开时，弹消息框提醒用户。

``` JavaScript
function Workbook_Open()
{
    MsgBox("工作簿被打开了。");
}
```

### SheetActivate

当激活 该 工作簿 任一工作表时 触发 此事件。

**语法**

`function Workbook_SheetActivate ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Sh   | 必选      | Object   | 激活的工作表对象。 |

**示例**

当 激活 了 工作表 时，弹消息框提醒用户。

``` JavaScript
function Application_SheetActivate(Sh)
{
    MsgBox("您激活了工作表："+Sh.Name)
}
```

### SheetBeforeDelete

删除该工作簿任 一 工作表之前 触发 此事件。

**语法**

`function Workbook_SheetBeforeDelete ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Sh   | 必选      | Object   | 删除的工作表对象。 |

**示例**

删除该工作簿任 一 工作表之前 ，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetBeforeDelete(Sh)
{
    MsgBox("您删除了工作表："+Sh.Name)
}
```

### SheetBeforeDoubleClick

双击该工作簿任一工作表之前 触发 此事件。

**语法**

`function Workbook_SheetBeforeDoubleClick ( Sh, Target, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sh     | 必选      | Object    | 双击的工作表对象。                            |
| Target | 必选      | Range对象 | 双击区域所在的单元格对象。                    |
| Cancel | 必选      | O bject   | 如果设置其属性Value为 true ，则取消此次双击。 |

**示例**

双击该工作簿任一工作表之前 ，如果双击操作在A1单元格上，则不响应双击操作（默认是进入编辑），否则则继续响应。

``` JavaScript
function Workbook_SheetBeforeDoubleClick(Sh, Target, Cancel)
{
    if (Target.Address() == "$A$1")
        Cancel.Value = true;
} 
```

### SheetBeforeRightClick

右击该工作簿任一 工作表之前 触发 此事件 。

**语法**

`function Workbook_SheetBeforeRightClick ( Sh, Target, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                            |
|--------|-----------|-----------|-------------------------------------------------|
| Sh     | 必选      | Object    | 右击 的工作表对象。                             |
| Target | 必选      | Range对象 | 右击 区域所在的单元格对象。                     |
| Cancel | 必选      | Object    | 如果设置其属性Value为 true ，则取消此次 右击 。 |

**示例**

右击 该工作簿任一工作表之前 ，如果右击操作在A1单元格上，则不响应 右击 操作（默认是弹出右键菜单），否则则继续 右击 。

``` JavaScript
function Workbook_SheetBeforeRightClick(Sh, Target, Cancel)
        {
        if (Target.Address() == "$A$1") Cancel.Value = true; } 
        
```

### SheetCalculate

该工作簿任一工作表进行 计算时 触发此事件。

**语法**

`function Workbook_SheetCalculate ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明               |
|------|-----------|----------|--------------------|
| Sh   | 必选      | Object   | 计算的工作表对象。 |

**示例**

当 工作表进行 计算时 ，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetCalculate(Sh)
{
    MsgBox("工作表正在计算："+Sh.Name)
} 
```

### SheetChange

当用户或外部链接更改了该工作簿任一 工作表中的单元格时 触发 此事件。

**语法**

`function Workbook_SheetChange ( Sh, Target ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                  |
|--------|-----------|-----------|-----------------------|
| Sh     | 必选      | Object    | 修改的工作表对象。    |
| Target | 必 选     | Range对象 | 修改的单元格Range对象 |

**示例**

工作表中的单元格被修改 时，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetChange(Sh, Target)
{
    MsgBox("工作表:" + Sh.Name + "。区域："+Target.Address() + "。发生了修改")
}
```

### SheetDeactivate

当该工作簿任 一 工作表被切换到非激活状态时 触发 此事件。

**语法**

`function Workbook_SheetDeactivate ( Sh ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                             |
|------|-----------|----------|----------------------------------|
| Sh   | 必选      | Object   | 被切换到非激活状态的工作表对象。 |

**示例**

任 一 工作表被切换到非激活状态时 ，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetDeactivate(Sh)
{
    MsgBox("被切换到非激活的工作表："+Sh.Name)
} 
```

### SheetFollowHyperlink

单击该工作簿 任一 工作表的超链接 时 触 发 此事 件。

**语法**

`function Workbook_SheetFollowHyperlink ( Sh, Target ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型      | 说明                              |
|--------|-----------|---------------|-----------------------------------|
| Sh     | 必选      | Object        | 超链接所在的工作表对象。          |
| Target | 必选      | Hyperlink对象 | Hyperlink对象，代表超链接的目标。 |

**示例**

单击该工作簿 任一 工作表的超链接 时，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetFollowHyperlink(Sh, Target)
{
    MsgBox("你点击了工作表\"" + Sh.Name + "\"上的超链接："+Target.Address)
} 
```

### SheetSelectionChange

该 工作簿 任一工作表上的选定区域发生更改时，将触发此事件。

**语法**

`function Workbook_SheetSelectionChange ( Sh, Target ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                   |
|--------|-----------|----------|------------------------|
| Sh     | 必选      | Object   | 选区改变的工作表对象。 |
| Target | 必选      | Range    | 新选定的区域。         |

**示例**

当 工作表上的选定区域发生更改时 ，弹消息框提醒用户。

``` JavaScript
function Workbook_SheetSelectionChange(Sh, Target)
{
    MsgBox("工作表\"" + Sh.Name + "\"选中的区域是："+Target.Address() + "。")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/表格事件/Workbook 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E8%A1%A8%E6%A0%BC%E4%BA%8B%E4%BB%B6/Workbook%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# CheckBox 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| Click     | 当用户在对象上按下然后释放鼠标按钮时 触发此事件                                   |
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Change    | 当控件的内容发生变化时触发此事件                                                  |
| Enter     | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit      | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### Click

当用户在对象上按下然后释放鼠标按钮时触发此事件

**语法**

`function UserForm1_CheckBox1_Click () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_CheckBox1_Click()
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

**语法**

`function UserForm1_CheckBox1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_CheckBox1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_CheckBox1_MouseDown (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_CheckBox1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_CheckBox1_MouseUp (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_CheckBox1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_CheckBox1_MouseMove (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_CheckBox1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_CheckBox1_KeyUp (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_CheckBox1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}       
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_CheckBox1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_CheckBox1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm1_CheckBox1_KeyPress (keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_CheckBox1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Change

当控件的内容发生变化时触发此事件

**语法**

`function UserForm1_CheckBox1_Change () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ListBox1_Change()
{
    MsgBox("控件内容发生改变")
} 
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm1_CheckBox1_Enter () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_CheckBox1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm1_CheckBox 1_Exit (cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_CheckBox1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/CheckBox 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/CheckBox%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# ComboBox 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| Click     | 当用户在对象上按下然后释放鼠标按钮时 触发此事件                                   |
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Change    | 当控件的内容发生变化时触发此事件                                                  |
| Enter     | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit      | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### Click

当用户在对象上按下然后释放鼠标按钮时触发此事件

**语法**

`function UserForm1_ComboBox1_Click () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ComboBox1_Click()
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时触发此事件

**语法**

`function UserForm1_ComboBox1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。 参数

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_ComboBox1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_ComboBox1_MouseDown (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_ComboBox1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_ComboBox1_MouseUp (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_ComboBox1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_ComboBox1_MouseMove (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_ComboBox1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_ComboBox1_KeyUp (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_ComboBox1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}                   
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_ComboBox1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_ComboBox1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm1_ComboBox1_KeyPress (keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_ComboBox1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Change

当控件的内容发生变化时触发此事件

**语法**

`function UserForm1_ComboBox1_Change () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ComboBox1_Change()
{
    MsgBox("控件内容发生改变")
} 
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm1_ComboBox1_Enter () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ComboBox1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm1_ComboBox1_Exit (cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_ComboBox1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/ComboBox 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/ComboBox%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# CommandButton 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| Click     | 当用户在对象上按下然后释放鼠标按钮时 触发此事件                                   |
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Enter     | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit      | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### Click

当用户在对象上按下然后释放鼠标按钮时触发此事件

**语法**

`function UserForm1_CommandButton1_Click () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_CommandButton1_Click()
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

**语法**

`function UserForm1_CommandButton1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_CommandButton1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_CommandButton1_MouseDown(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_CommandButton1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_CommandButton1_MouseUp(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_CommandButton1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_CommandButton1_MouseMove(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_CommandButton1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_CommandButton1_KeyUp(keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_CommandButton1_KeyUp(keycode, shift) 
{
    MsgBox("您松开了" + keycode + "键")
}
                        
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_CommandButton1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_CommandButton1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm 1_CommandButton1_KeyPress(keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_CommandButton1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm 1_CommandButton1_Enter () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_CommandButton1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm 1_CommandButton1_Exit(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_CommandButton1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/CommandButton 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/CommandButton%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# Frame 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| Click     | 当用户在对象上按下然后释放鼠标按钮时 触发此事件                                   |
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Enter     | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit      | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### Click

当用户在对象上按下然后释放鼠标按钮时触发此事件

**语法**

`function UserForm1_Frame1_Click () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_Frame1_Click()
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

**语法**

`function UserForm1_Frame1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_Frame1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_Frame1_MouseDown (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_Frame1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_Frame1_MouseUp (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_Frame1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_Frame1_MouseMove (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_Frame1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_Frame1_KeyUp (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_Frame1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}                   
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_Frame1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_Frame1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm1_Frame1_KeyPress (keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_Frame1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm1_Frame1_Enter () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_Frame1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm1_Frame1_Exit (cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_Frame1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/Frame 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/Frame%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# HLayout 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| Click     | 当用户在对象上按下然后释放鼠标按钮时 触发此事件                                   |
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |

## 事件

### Click

当用户在对象上按下然后释放鼠标按钮时触发此事件

**语法**

`function UserForm1_HLayout1_Click () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_HLayout1_Click()
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

**语法**

`function UserForm1_HLayout1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_HLayout1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_HLayout1_MouseDown (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_HLayout1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_HLayout1_MouseUp (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_HLayout1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_HLayout1_MouseMove (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_HLayout1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_HLayout1_KeyUp (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_HLayout1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}                   
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_HLayout1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_HLayout1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm1_HLayout1_KeyPress (keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_HLayout1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/HLayout 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/HLayout%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# Label 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|-------------------------------------------------------------------------|
| Click     | 当用户在对象上按下然后释放鼠标按钮时 触发此事件                         |
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件 |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                         |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                         |
| MouseMove | 在用户移动鼠标时 触发此事件                                             |

## 事件

### Click

当用户在对象上按下然后释放鼠标按钮时触发此事件

**语法**

`function UserForm1_Label1_Click () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_Label1_Click()
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

**语法**

`function UserForm1_Label1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_Label1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_Label 1_MouseDown(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_Label1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserFo rm1_Label1_MouseUp(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_Label1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function User Form1_Label1_MouseMove(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_Label1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/Label 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/Label%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# ListBox 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| Click     | 当用户在对象上按下然后释放鼠标按钮时 触发此事件                                   |
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Change    | 当控件的内容发生变化时触发此事件                                                  |
| Enter     | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit      | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### Click

当用户在对象上按下然后释放鼠标按钮时触发此事件

**语法**

`function UserForm1_ListBox1_Click () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ListBox1_Click()
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

**语法**

`function UserForm1_ListBox1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_ListBox1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_ListBox1_MouseDown (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_ListBox1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_ListBox1_MouseUp (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_ListBox1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_ListBox1_MouseMove (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_ListBox1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_ListBox1_KeyUp (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_ListBox1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}                       
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_ListBox1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_ListBox1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm1_ListBox1_KeyPress (keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_ListBox1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Change

当控件的内容发生变化时触发此事件

**语法**

`function UserForm1_ListBox1_Change () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ListBox1_Change()
{
    MsgBox("控件内容发生改变")
} 
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm1_ListBox1_Enter () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ListBox1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm1_ListBox1_Exit (cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_ListBox1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/ListBox 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/ListBox%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# MultiPage 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| Click     | 当用户在对象上按下然后释放鼠标按钮时 触发此事件                                   |
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Change    | 当控件的内容发生变化时触发此事件                                                  |
| Enter     | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit      | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### Click

当用户在对象上按下然后释放鼠标按钮时触发此事件

**语法**

`function UserForm1_MultiPage1_Click (index) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明                 |
|-------|-----------|----------|----------------------|
| Index | 必选      | Number   | 该设置指定页面的下标 |

**示例**

``` JavaScript
function UserForm1_MultiPage1_Click(index)
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

**语法**

`function UserForm1_MultiPage1_DblClick(index, cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Index  | 必选      | Number   | 该设置指定页面的下标                                                |
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_MultiPage1_DblClick(index, cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_MultiPage1_MouseDown (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_MultiPage1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_MultiPage1_MouseUp (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_MultiPage1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_MultiPage1_MouseMove (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_MultiPage1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_MultiPage1_KeyUp (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_MultiPage1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}                       
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_MultiPage1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_MultiPage1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm1_MultiPage1_KeyPress (keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_MultiPage1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Change

当控件的内容发生变化时触发此事件

**语法**

`function UserForm1_MultiPage1_Change () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_MultiPage1_Change()
{
    MsgBox("控件内容发生改变")
} 
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm1_MultiPage1_Enter () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_MultiPage1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm1_MultiPage1_Ex it (cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_MultiPage1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/MultiPage 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/MultiPage%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# OptionButton 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| Click     | 当用户在对象上按下然后释放鼠标按钮时 触发此事件                                   |
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Change    | 当控件的内容发生变化时触发此事件                                                  |
| Enter     | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit      | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### Click

当用户在对象上按下然后释放鼠标按钮时触发此事件

**语法**

`function UserForm1_OptionButton1_Click () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_OptionButton1_Click()
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

**语法**

`function UserForm1_OptionButton1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_OptionButton1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_OptionButton1_MouseDown (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_OptionButton1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_OptionButton1_MouseUp (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_OptionButton1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_OptionButton1_MouseMove (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_OptionButton1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_OptionButton1_KeyUp (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_OptionButton1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}                       
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_OptionButton1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_OptionButton1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm1_OptionButton1_KeyPress (keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_OptionButton1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Change

当控件的内容发生变化时触发此事件

**语法**

`function UserForm1_OptionButton1_Change () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_OptionButton1_Change()
{
    MsgBox("控件内容发生改变")
} 
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm1_OptionButton1_Enter () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_OptionButton1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm1_OptionButton1_Exit (cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_OptionButton1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/OptionButton 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/OptionButton%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# ScrollBar 事件

## 事件列表

| 名称     | 触发时机       |
|----------|--|
| KeyUp    | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown  | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Change   | 当控件的内容发生变化时触发此事件                                                  |
| Enter    | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit     | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_ScrollBar1_KeyUp (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_ScrollBar1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}                       
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_ScrollBar1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_ScrollBar1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm1_ScrollBar1_KeyPress (keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_ScrollBar1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Change

当控件的内容发生变化时触发此事件

**语法**

`function UserForm1_ScrollBar1_Change () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ScrollBar1_Change()
{
    MsgBox("控件内容发生改变")
} 
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm1_ScrollBar1_Enter () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ScrollBar1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm1_ScrollBar1_Exit (cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_ScrollBar1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/ScrollBar 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/ScrollBar%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# SpinButton 事件

## 事件列表

| 名称     | 触发时机       |
|----------|--|
| KeyUp    | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown  | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Change   | 当控件的内容发生变化时触发此事件                                                  |
| Enter    | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit     | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_SpinButton1_KeyUp (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_SpinButton1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}                       
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_SpinButton1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_SpinButton1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm1_SpinButton1_KeyPress (keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_SpinButton1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Change

当控件的内容发生变化时触发此事件

**语法**

`function UserForm1_SpinButton1_Change () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_SpinButton1_Change()
{
    MsgBox("控件内容发生改变")
} 
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm1_SpinButton1_Enter () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_SpinButton1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm1_SpinButton1_Exit (cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_SpinButton1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/SpinButton 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/SpinButton%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# TextEdit 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Change    | 当控件的内容发生变化时触发此事件                                                  |
| Enter     | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit      | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### DblClick

**语法**

`function UserForm1_TextEdit1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_TextEdit1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_TextEdit1_MouseDown(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_TextEdit1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_TextEdit1_MouseUp(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_TextEdit1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_TextEdit1_MouseMove(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_TextEdit1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_TextEdit1_KeyUp(keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_TextEdit1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}               
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_TextEdit1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_TextEdit1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm 1_TextEdit1_KeyPress(keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_TextEdit1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Change

当控件的内容发生变化时触发此事件

**语法**

`function UserForm1_TextEdit1_Change () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_TextEdit1_Change()
{
    MsgBox("控件内容发生改变")
} 
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm 1_TextEdit1_Enter() { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_TextEdit1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm 1_TextEdi t1_Exit(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_TextEdit1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/TextEdit 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/TextEdit%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# ToggleButton 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| Click     | 当用户在对象上按下然后释放鼠标按钮时 触发此事件                                   |
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |
| Change    | 当控件的内容发生变化时触发此事件                                                  |
| Enter     | 在控件从同一窗体的控件上接收焦点之前触发此事件                                    |
| Exit      | 在一个控件上的焦点转移到同一窗体上的另一个控件之前触发此事件                      |

## 事件

### Click

当用户在对象上按下然后释放鼠标按钮时触发此事件

**语法**

`function UserForm1_ToggleButton1_Click () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ToggleButton1_Click()
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

**语法**

`function UserForm1_ToggleButton1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_ToggleButton1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_ToggleButton1_MouseDown (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_ToggleButton1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_ToggleButton1_MouseUp (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_ToggleButton1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_ToggleButton1_MouseMove (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_ToggleButton1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_ToggleButton1_KeyUp (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_ToggleButton1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}                   
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_ToggleButton1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_ToggleButton1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm1_ToggleButton1_KeyPress (keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_ToggleButton1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

### Change

当控件的内容发生变化时触发此事件

**语法**

`function UserForm1_ToggleButton1_Change () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ToggleButton1_Change()
{
    MsgBox("控件内容发生改变")
} 
```

### Enter

在控件从同一窗体的控件上接收焦点之前触发该事件

**语法**

`function UserForm1_ToggleButton1_Enter () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_ToggleButton1_Enter()
{
    MsgBox("焦点进入该控件")
}
```

### Exit

在一个控件上的焦点转移到同一窗体上的另一个控件之前触发该事件

**语法**

`function UserForm1_ToggleButton1_Exit (cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_ToggleButton1_Exit(cancel)
{
    MsgBox("焦点离开该控件")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/ToggleButton 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/ToggleButton%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# UserForm 事件

## 事件列表

| 名称       | 触发时机                                                      |
|------------|-------------------------------------------------|
| Initialize | 在加载对象之后但在显示之前触发此事件                          |
| Terminate  | 当通过将引用该对象的所有变量设置为 Nothing 或当对该对象的最后一个引用超出范围而从内存中删除对某个对象实例的所有引用时 触发此事件 |
| Click      | 当用户在对象上按下然后释放鼠标按钮时 触发此事件               |
| DblClick   | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件 |
| MouseDown  | 在用户按下鼠标按钮时 触发此事件                               |
| MouseUp    | 在用户释放鼠标按钮时 触发此事件                               |
| MouseMove  | 在用户移动鼠标时 触发此事件                                   |
| KeyUp      | 当用户在窗体或控件具有焦点时释放键时 触发此事件               |
| KeyDown    | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件           |
| KeyPress   | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件                                                |

## 事件

### Initialize

在加载对象之后但在显示之前 触发此事件

**语法**

`function UserForm1_Initiali ze () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_Initialize()
{
    MsgBox("您新建了用户窗体：")
}
```

### Terminate

当通过将引用该对象的所有变量设置为 Nothing 或当对该对象的最后一个引用超出范围而从内存中删除对某个对象实例的所有引用时 触发此事件

**语法**

`function UserForm1_Terminate () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_Terminate()
{
    MsgBox("您关闭了用户窗体：")
}
```

### Click

当用户在对象上按下然后释放鼠标按钮时 触发此事件

**语法**

`function UserForm1_Click () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_Click()
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

当 用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件

**语法**

`function UserForm1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时 触发此事件

**语法**

`function UserForm1_MouseDown(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时 触发此事件

**语法**

`function UserForm1_MouseUp (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_MouseMove(btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_KeyUp(keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}                       
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm 1_KeyPress(keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/UserForm 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/UserForm%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# VLayout 事件

## 事件列表

| 名称      | 触发时机       |
|-----------|--|
| Click     | 当用户在对象上按下然后释放鼠标按钮时 触发此事件                                   |
| DblClick  | 当用户在系统的双击时间限制内在对象上按下并释放鼠标左键两次时 触发此事件           |
| MouseDown | 在用户按下鼠标按钮时 触发此事件                                                   |
| MouseUp   | 在用户释放鼠标按钮时 触发此事件                                                   |
| MouseMove | 在用户移动鼠标时 触发此事件                                                       |
| KeyUp     | 当用户在窗体或控件具有焦点时释放键时 触发此事件                                   |
| KeyDown   | 当用户在窗体或控件具有焦点时按下某个键时 触发此事件                               |
| KeyPress  | 当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时 触发此事件 |

## 事件

### Click

当用户在对象上按下然后释放鼠标按钮时触发此事件

**语法**

`function UserForm1_VLayout1_Click () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

无

**示例**

``` JavaScript
function UserForm1_VLayout1_Click()
{
    MsgBox("您单击了一次：")
} 
```

### DblClick

**语法**

`function UserForm1_VLayout1_DblClick(cancel) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明       |
|--------|-----------|----------|---------------------------------------------------------------------|
| Cancel | 必选      | Integet  | 该设置确定是否触发事件。 将该参数设置为 True (1) 会取消触发该事件。 |

**示例**

``` JavaScript
function UserForm1_VLayout1_DblClick(cancel)
{
    MsgBox("您双击了一次：")
}
```

### MouseDown

在用户按下鼠标按钮时触发此事件

**语法**

`function UserForm1_VLayout1_MouseDown (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_VLayout1_MouseDown(btn, shift, x, y)
{
    MsgBox("您按下了鼠标：")
} 
```

### MouseUp

在用户释放鼠标按钮时触发此事件

**语法**

`function UserForm1_VLayout1_MouseUp (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_VLayout1_MouseUp(btn, shift, x, y)
{
    MsgBox("您松开了鼠标：")
} 
```

### MouseMove

在用户移动鼠标时触发此事件

**语法**

`function UserForm1_VLayout1_MouseMove (btn, shift, x, y) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明  |
|--------|-----------|----------|----------------------------------------------------------------|
| Button | 必选      | Integer  | 被按下以触发事件的鼠标按钮                                     |
| Shift  | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |
| X      | 必选      | Single   | 鼠标指针当前位置的 x 坐标                                      |
| Y      | 必选      | Single   | 鼠标指针当前位置的 y坐标                                       |

**示例**

``` JavaScript
function UserForm1_VLayout1_MouseMove(btn, shift, x, y)
{
    MsgBox("您移动了鼠标：")
}
```

### KeyUp

当用户在窗体或控件具有焦点时释放键时触发此事件

**语法**

`function UserForm1_VLayout1_KeyUp (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_VLayout1_KeyUp(keycode, shift)
{
    MsgBox("您松开了" + keycode + "键")
}                   
```

### KeyDown

当用户在窗体或控件具有焦点时按下某个键时触发此事件

**语法**

`function UserForm1_VLayout1_KeyDown (keycode, shift) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明  |
|---------|-----------|----------|----------------------------------------------------------------|
| keyCode | 必选      | Integer  | 键代码，您可以通过将参数设置为 0 来阻止对象接收按键            |
| Shift   | 必选      | Integer  | 按下或释放 Button 参数指定的按钮时 Shift、Ctrl 和 Alt 键的状态 |

**示例**

``` JavaScript
function UserForm1_VLayout1_KeyDown(keycode, shift)
{
    MsgBox("您按下了" + keycode + "键")
}
```

### KeyPress

当窗体或控件具有焦点时，当用户按下并释放对应于 ANSI 代码的键或组合键时触发此事件

**语法**

`function UserForm1_VLayout1_KeyPress (keyascii) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称     | 必选/可选 | 数据类型 | 说明                                          |
|----------|-----------|----------|---------------------------------|
| keyAscii | 必选      | Integer  | 返回一个数字 ANSI 键码。 KeyAscii 参数是通过引用传递的；更改它会向对象发送不同的字符。 将该参数设置为0会取消击键 |

**示例**

``` JavaScript
function UserForm1_VLayout1_KeyPress(keyascii)
{
    MsgBox("您点击了" + keycode + "键")
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/控件事件/VLayout 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%8E%A7%E4%BB%B6%E4%BA%8B%E4%BB%B6/VLayout%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# Application 事件

## 事件列表

| 名称                    | 触发时机                                                    |
|-------------------------|-------------------------------------------------------------|
| DocumentBeforeClose     | 当文档关闭前时触发此事件。                                  |
| DocumentBeforePrint     | 当文档打印之前时 触发 此事件。                              |
| DocumentChange          | 当创建新文档、打开已有文档或者切换激活文档 时 触发 此事件。 |
| DocumentOpen            | 当任一文档打开时 触发 此事件。                              |
| NewDocument             | 当 创建新文档 时 触发 此事件 。                             |
| Quit                    | 当关闭退出文字组件进程时触发此事件。                        |
| WindowActivate          | 当激活文档窗口 时 触发 此事件。                             |
| WindowBeforeDoubleClick | 当双击任一 文档窗口 之前 触发 此事件。                      |
| WindowBeforeRightClick  | 当右击任一 文档窗口 之前 触发 此事件 。                     |
| WindowDeactivate        | 任一 文档 窗口被切换到非激活状态时触发此事件。              |
| WindowSelectionChange   | 活动窗口所选内容更改时将 触发 此事件。                      |
| WindowSize              | 任一 文档 窗口调整大小时将 触发 此事件。                    |
| XMLSelectionChange      | 当前所选内容的 XML 父节点更改时将 触发 此事件。             |
| XMLValidationError      | 文档中存在验证错误时 时，将 触发 此事件。                   |

## 事件

### DocumentBef oreClose

当文档关闭前时触发此事件。

**语法**

`function Application_DocumentBef oreClose ( Doc, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型     | 说明                                        |
|--------|-----------|--------------|---------------------------------------------|
| Doc    | 必选      | Document对象 | 关闭的文档对象                              |
| Cancel | 必选      | Object       | 如果设置其属性Value为 true ，则不关闭文稿。 |

**示例**

当关闭一个 文档 时，弹消息框提醒用户即将关闭 文档 ， 询问 是否取消。用户点击“是”就会取消关闭 文档 。

``` JavaScript
function Application_DocumentBeforeClose(Doc, Cancel)
{
    var ret = MsgBox("文档\"" + Doc.Name + "\"" +"正在关闭。是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### DocumentBeforePrint

当文档打印之前时 触发 此事件。

**语法**

`function Application_DocumentBeforePrint ( Doc, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型     | 说明                                          |
|--------|-----------|--------------|-----------------------------------------------|
| Doc    | 必选      | Document对象 | 关闭的文档对象                                |
| Cancel | 必选      | Object       | 如果设置其属性Value为 true ，则不 打印文档 。 |

**示例**

当打印一个 文档 时，弹消息框提醒用户即将打印 文档 ， 询问 是否取消。用户点击“是”就会取消打印 文档。

``` JavaScript
function Application_DocumentBeforePrint(Doc, Cancel)
{
var ret = MsgBox("即将对文档\"" + Doc.Name + "\"" +"进行打印。是否取消？", jsYesNo);
if (ret == jsResultYes)
    Cancel.Value = true;
}
```

### DocumentChange

当创建新文档、打开已有文档或者切换激活文档 时 触发 此事件。

**语法**

`function Application_DocumentChange () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**示例**

当触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_DocumentChange()
{
    MsgBox("文档切换了！");
}
```

### DocumentOpen

当任一文档打开时 触发 此事件。

**语法**

`function Application_DocumentOpen ( Doc ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明           |
|------|-----------|--------------|----------------|
| Doc  | 必选      | Document对象 | 关闭的文档对象 |

**示例**

当触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_DocumentOpen(Doc)
{
    MsgBox("打开了文档:" + Doc.Name);
}
```

### NewDocument

当创建新文档 时 触发 此事件 。

**语法**

`function Application_NewDocument ( Doc ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明           |
|------|-----------|--------------|----------------|
| Doc  | 必选      | Document对象 | 关闭的文档对象 |

**示例**

当触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_NewDocument(Doc)
{
    MsgBox("新建了文档:" + Doc.Name);
}
```

### Quit

当关闭退出文字组件进程时触发此事件。

**语法**

`function Application_Quit () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

### WindowActivate

当激活文档窗口 时 触发 此事件。

**语法**

`function Application_WindowActivate ( Doc, Wn ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明            |
|------|-----------|--------------|-----------------|
| Doc  | 必选      | Document对象 | 激活 的文档对象 |
| Wn   | 必选      | Window 对象  | 激活 的窗口对象 |

**示例**

当切换文档窗口时，输出激活的窗口标题。

``` JavaScript
function Application_WindowActivate(Doc,  Wn)
{
    Debug.Print("当前激活窗口是："+ Wn.Caption)
}
```

### WindowBeforeDoubleClick

当双击任一文档窗口之前 触发 此事件。

**语法**

`function Application_WindowBeforeDoubleClick ( Sel , Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sel    | 必选      | Selection | 当前所选内容。                                |
| Cancel | 必选      | Object    | 如果设置其属性Value为 true ，则取消此次双击。 |

**示例**

当在一个窗口内双击 时，弹消息框提醒用户 ， 询问 是否取消。用户点击“是”就会取消该操作 。

``` JavaScript
function Application_WindowBeforeDoubleClick(Sel, Cancel)
{
    var ret = MsgBox("您进行了双击，当前选中内容是：" + Sel +"是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### WindowBeforeRightClick

当右击任一 文档窗口之前 触发 此事件 。

**语法**

`function Application_WindowBeforeRightClick ( Sel , Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型  | 说明                                          |
|--------|-----------|-----------|-----------------------------------------------|
| Sel    | 必选      | Selection | 当前所选内容。                                |
| Cancel | 必选      | Object    | 如果设置其属性Value为 true ，则取消此次右击。 |

**示例**

当在一个窗口内右击时，弹消息框提醒用户，询问是否取消。用户点击“是”就会取消该操作 。

``` JavaScript
function Application_WindowBeforeRightClick(Sel, Cancel)
{
    var ret = MsgBox("您进行了右击，当前选中内容是：" + Sel +"是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### WindowDeactivate

任一 文档 窗口被切换到非激活状态时触发此事件。

**语法**

`function Application_WindowDeactivate ( Doc, Wn ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明                          |
|------|-----------|--------------|-------------------------------|
| Doc  | 必选      | Document对象 | 被切换到非激活状态 的文档对象 |
| Wn   | 必选      | Window 对象  | 被切换到非激活状态 的窗口对象 |

**示例**

当切换文档窗口时，输出 被切换到非激活状态 的窗口 标题。

``` JavaScript
function Application_WindowDeactivate(Doc,  Wn)
{
    Debug.Print("当前被切换到非激活窗口是："+ Wn.Caption)
}
```

### WindowSelectionChange

活动窗口所选内容更改时将 触发 此事件。

**语法**

`function Application_WindowSelectionChange ( Sel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型      | 说明           |
|------|-----------|---------------|----------------|
| Sel  | 必选      | Selection对象 | 当前所选内容。 |

**示例**

当 活动窗口所选内容更改时 触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_WindowSelectionChange(Sel)
{
    MsgBox("当前选中内容是：" + Sel);
}
```

### WindowSize

任一文档窗口调整大小时将 触发 此事件。

**语法**

`function Application_WindowSize ( Doc, Wn ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型     | 说明                    |
|------|-----------|--------------|-------------------------|
| Doc  | 必选      | Document对象 | 调整窗口大小 的文档对象 |
| Wn   | 必选      | Window 对象  | 调整窗口大小 的窗口对象 |

**示例**

当触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_WindowSize(Doc, Wn)
{
    MsgBox("窗口\""+ Wn.Caption + "\"大小发生了改变。")
}
```

### XMLSelectionChange

当前所选内容的 XML 父节点更改时将 触发 此事件。

**语法**

`function Application_XMLSelectionChange ( Sel, OldXMLNode, NewXMLNode, Reason ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称       | 必选/可选 | 数据类型       | 说明                        |
|------------|-----------|----------------|-----------------------------|
| Sel        | 必选      | Selection 对象 | 当前所选内容。              |
| OldXMLNode | 必选      | XMLNode 对象   | 插入区域移动前的XML 节点。  |
| NewXMLNode | 必选      | XMLNode 对象   | 插入区域移动后的 XML 节点。 |
| Reason     | 必选      | long           | 更改的操作                  |

**示例**

以下示例在将新元素插入到文档中时，对新添加的 XML 元素进行验证。

``` JavaScript
function Application_XMLSelectionChange(Sel, OldXMLNode, NewXMLNode, Reason)
{
    if (Reason == wdXMLSelectionChangeReasonInsert)
    {
        if (OldXMLNode !== undefined && OldXMLNode !== null)
            NewXMLNode.Validate()
    }
}
```

### XMLValidationError

文档中存在验证错误时 时，将触发此事件。

**语法**

`function Application_XMLValidationError ( XMLNode ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称    | 必选/可选 | 数据类型     | 说明           |
|---------|-----------|--------------|----------------|
| XMLNode | 必选      | XMLNode 对象 | 无效的 XML节点 |

**示例**

当触发事件时，弹出消息框提醒用户 。

``` JavaScript
function Application_XMLValidationError(XMLNode)
{
    MsgBox("XML 节点"  + XMLNode.BaseName +" 无效:" + XMLNode.ValidationErrorText)
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/文字事件/Application 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%96%87%E5%AD%97%E4%BA%8B%E4%BB%B6/Application%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# Document 事件

## 事件列表

| 名称                              | 触发时机              |
|-----------------------------------|---------|
| Close                             | 当文档关闭时 触发 此事件。      |
| ContentControlAfterAdd            | 当内容控件被添加到文档之后 触发 此事件。                                                 |
| ContentControlBeforeContentUpdate | 当内容控件的内容发生来自XML映射内容更新时 触发 此事件。                                  |
| ContentControlBeforeDelete        | 当内容控件被从文档删除之前 触发 此事件。                                                 |
| ContentControlBeforeStoreUpdate   | 当内容控件的内容去更新XML映射内容时 触发 此事件。                                        |
| ContentControlOnEnter             | 当进入内容控件时 触发 此事件。  |
| ContentControlOnExit              | 当退出 内容控件时 触发 此事件。 |
| New                               | 当基于模板创建新文档时触发，此事件只发向基于的模板文档对象。                             |
| Open                              | 文档打开时触发。      |
| XMLAfterInsert                    | 向文档添加XML 元素时触发。如果同时向文档添加多个元素，则插入的每个元素都会触发这一事件。 |
| XMLBeforeDelete                   | 从文档删除XML 元素时触发。如果同时向文档删除多个元素，则删除的每个元素都会触发这一事件。 |

## 事件

### Close

当文档关闭时 触发 此事件。

**语法**

`function Document_Close ( Wb ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**示例**

当关闭文档时，弹消息框提醒用户。

``` JavaScript
function Document_Close()
{
    MsgBox("文档已关闭")
}
```

### ContentControlAfterAdd

当内容控件被添加到文档之后 触发 此事件。

**语法**

`function Document_ContentControlAfterAdd ( NewContentControl, InUndoRedo ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称              | 必选/可选 | 数据类型           | 说明                                     |
|:------------------|:----------|:-------------------|:-----------------------------------------|
| NewContentControl | 必选      | ContentControl对象 | 要添加的内容控件。                       |
| InUndoRedo        | 必选      | Boolean            | 此次添加操作是否是撤销或恢复操作所发生的 |

### ContentControlBeforeContentUpdate

当内容控件的内容发生来自XML映射内容更新时 触发 此事件。

**语法**

`function Document_ContentControlBeforeContentUpdate ( ContentControl, Content ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称           | 必选/可选 | 数据类型           | 说明        |
|----------------|-----------|--------------------|----------------------------------------------------------------------|
| ContentControl | 必选      | ContentControl对象 | 要更新的内容控件。                                                   |
| Content        | 必选      | String             | 控件的更新内容。 此参数用于更改 XML 数据的内容，以及设置其显示格式。 |

**备注**

对于重复内容，控件不触发此事件。

### ContentControlBeforeDelete

当内容控件被从文档删除之前 触发 此事件。

**语法**

`function Document_ContentControlBeforeDelete ( OldContentControl, InUndoRedo ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称              | 必选/可选 | 数据类型           | 说明                                     |
|-------------------|-----------|--------------------|------------------------------------------|
| OldContentControl | 必选      | ContentControl对象 | 要删除的内容控件。                       |
| InUndoRedo        | 必选      | Boolean            | 此次删除操作是否是撤销或恢复操作所发生的 |

### ContentControlBeforeStoreUpdate

当内容控件的内容去更新XML映射内容时 触发 此事件。

**语法**

`function Document_ContentControlBeforeStoreUpdate ( ContentControl, Content ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称           | 必选/可选 | 数据类型            | 说明   |
|----------------|-----------|---------------------|---------------------------------------------------------------------------|
| ContentControl | 必选      | ContentControl 对象 | 要更新的内容控件。                                                        |
| Content        | 必选      | String              | 去更新XML映射的内容。 在将内容设置到 XML映射 之前，可使用此参数更改该值。 |

**备注**

对于重复内容，控件不触发此事件。

### ContentControlOnEnter

当进入内容控件时 触发 此事件。

**语法**

`function Document_ContentControlOnEnter ( ContentControl ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称           | 必选/可选 | 数据类型           | 说明             |
|----------------|-----------|--------------------|------------------|
| ContentControl | 必选      | ContentControl对象 | 进入的内容控件。 |

**备注**

此事件仅对您要进入的内容控件而不对内容父控件触发。 例如，如果您的一个文本框内容控件嵌入内容组控件，您要将光标置于该文本框内容控件，则此事件仅对该文本框内容控件触发一次，而不对内容父组控件触发。

### ContentControlOnExit

当退出内容控件时 触发 此事件。

**语法**

`function Document_ContentControlOnExit ( ContentControl, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称           | 必选/可选 | 数据类型           | 说明                                      |
|----------------|-----------|--------------------|-------------------------------------------|
| ContentControl | 必选      | ContentControl对象 | 要离开的内容控件。                        |
| Cancel         | 必选      | Object             | 如果设置其属性Value为 true ，则取消退出。 |

**备注**

此事件仅对您要进入的内容控件而不对内容父控件触发。 例如，如果您的一个文本框内容控件嵌入内容组控件，您要将光标置于该文本框内容控件，则此事件仅对该文本框内容控件触发一次，而不对内容父组控件触发。

### New

当基于模板创建新文档时触发，此事件只发向基于的模板文档对象。

**语法**

`function Document_New () { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**示例**

当 模板创建新文档 时，弹消息框提醒用户。

``` JavaScript
function Document_New()
{
    MsgBox("模板新建了一个文档！")
}
```

### Open

文档打开时触发。

**语法**

`function Document_Open ( Wb ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**示例**

当 文档打开 时，弹消息框提醒用户。

``` JavaScript
function Document_Open()
{
    MsgBox("文档打开了！")
}
```

### XMLAfterInsert

向文档添加XML 元素时触发。如果同时向文档添加多个元素，则插入的每个元素都会触发这一事件。

**语法**

`function Document_XMLAfterInsert ( NewXMLNode, InUndoRedo ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称       | 必选/可选 | 数据类型    | 说明                                     |
|------------|-----------|-------------|------------------------------------------|
| NewXMLNode | 必选      | XMLNode对象 | 新添加的 XML 节点。                      |
| InUndoRedo | 必选      | Boolean     | 此次添加操作是否是撤销或恢复操作所发生的 |

**备注**

如果 InUndoRedo 参数为true，不要在 XMLAfterInsert 和 XMLBeforeDelete 中改变文档的XML。

如果 InUndoRedo 参数为 f alse，您 可 以 在 XMLAfterInsert 和 XMLBeforeDelete 中改变文档的XML， 但请注意 XMLAfterInsert 和 XMLBeforeDelete 事件会因 此存在嵌套 ，从而导致无限循环。 您 可 以在全局使用 Boolean 变量判断是否重入来防止 无限循环的出现。

**示例**

向文档添加XML 元素时，验证一个新添加的节点，如果节点无效，则显示一条描述验证错误的消息。

``` JavaScript
function Document_XMLAfterInsert(NewXMLNode, InUndoRedo)
{
    NewXMLNode.Validate()
    if (NewXMLNode.ValidationStatus != wdXMLValidationStatusOK)
    {
        MsgBox(NewXMLNode.ValidationErrorText)
    }
}
```

### XMLBeforeDelete

从文档删除XML 元素时触发。如果同时向文档删除多个元素，则删除的每个元素都会触发这一事件。

**语法**

`function Document_XMLBeforeDelete ( DeletedRange, OldXMLNode, InUndoRedo ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称         | 必选/可选 | 数据类型    | 说明                                      |
|--------------|-----------|-------------|-----------------------------|
| DeletedRange | 必选      | Range对象   | 被删除的 XML 元素的内容。 如果仅删除元素，并且未关联文本，DeletedRange 参数将不存在，因此将设置为 Nothing 。 |
| OldXMLNode   | 必选      | XMLNode对象 | 删除的 XML 节点。                         |
| InUndoRedo   | 必选      | Boolean     | 此次添加操作是否是撤销或恢复操作所发生的  |

**备注**

如果 InUndoRedo 参数为true，不要在 XMLAfterInsert 和 XMLBeforeDelete 中改变文档的XML。

如果 InUndoRedo 参数为 false，您可以在 XMLAfterInsert 和 XMLBeforeDelete 中改变文档的XML， 但请注意 XMLAfterInsert 和 XMLBeforeDelete 事件会因此存在嵌套，从而导致无限循环。 您可以在全局使用 Boolean 变量判断是否重入来防止 无限循环的出现。

**示例**

在 XML 元素被删除时。 如果元素中包含文本，则显示一条信息询问用户是否要删除元素所包含的文本。 如果用户通过单击"否"进行响应，则元素的内容将复制到剪贴板。

``` JavaScript
function Document_XMLBeforeDelete(DeletedRange, OldXMLNode, InUndoRedo)
{
    var ret = MsgBox("是否删除文本：" + DeletedRange.Text + "?", jsYesNo);
    if (ret == jsResultNo)
    {
        DeletedRange.Copy();
        MsgBox("文本内容已经复制到剪切板。")
    }   
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/文字事件/Document 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%96%87%E5%AD%97%E4%BA%8B%E4%BB%B6/Document%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# Application 事件

## 事件列表

| 名称                    | 触发时机                                     |
|-------------------------|----------------------------------------------|
| NewPresentation         | 当新建演示文稿时触发此事件。                 |
| PresentationBeforeClose | 任一打开的 演示文稿 关闭之前触发此事件。     |
| PresentationBeforeSave  | 任一 演示文稿 被保存 之前 触发 此事件。      |
| PresentationOpen        | 当打开任一演示文稿时 触发 此事件。           |
| PresentationSave        | 任一 演示文稿被 保存 时 触发 此事件。        |
| WindowActivate          | 任一 演示文稿 窗口被激活时，将 触发 此事件。 |

## 事件

### NewPresentation

当新建演示文稿时触发此事件 。

**语法**

`function Application_NewPresentation ( Pres ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型         | 说明                   |
|------|-----------|------------------|------------------------|
| Pres | 必选      | Presentation对象 | 新建的 演示文稿 对象。 |

**示例**

当新建一个 演示文稿 时，弹消息框提醒用户新建了 演示文稿 。

``` JavaScript
function Application_NewPresentation(Pres)
{
    MsgBox("您新建了演示文稿："+Pres.Name)
}
```

### PresentationBeforeClose

任一打开的 演示文稿 关闭之前触发此事件。

**语法**

`function Application_PresentationBeforeClose ( Pres, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型         | 说明                                            |
|--------|-----------|------------------|-------------------------------------------------|
| Pres   | 必选      | Presentation对象 | 关闭的 演示文稿 对象。                          |
| Cancel | 必选      | Object           | 如果设置其属性Value为 true ，则不关闭演示文稿。 |

**示例**

当关闭一个演示文稿时，弹消息框提醒用户即将关闭演示文稿，是否取消。用户点击“是”就会取消关闭 演示文稿。

``` JavaScript
function Application_PresentationBeforeClose(pres, Cancel)
{
    var ret = MsgBox("演示文稿\"" + pres.Name + "\"" +"正在关闭。是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### PresentationBeforeSave

任一 演示文稿被保存 之前 触发 此事件。

**语法**

`function Application_PresentationBeforeSave ( Pres, Cancel ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称   | 必选/可选 | 数据类型         | 说明                                            |
|--------|-----------|------------------|-------------------------------------------------|
| Pres   | 必选      | Presentation对象 | 保存的 演示文稿 对象。                          |
| Cancel | 必选      | Object           | 如果设置其属性Value为 true ，则不保存演示文稿。 |

**示例**

当保存一个演示文稿时，弹消息框提醒用户即将保存演示文稿，是否取消。用户点击“是”就会取消保存 演示文稿。

``` JavaScript
function Application_PresentationBeforeSave(pres, Cancel)
{
    var ret = MsgBox("演示文稿\"" + pres.Name + "\"" +"即将保存。是否取消？", jsYesNo);
    if (ret == jsResultYes)
        Cancel.Value = true;
}
```

### PresentationOpen

当打开任一演示文稿时 触发 此事件。

**语法**

`function Application_PresentationOpen ( Pres ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型         | 说明                   |
|------|-----------|------------------|------------------------|
| Pres | 必选      | Presentation对象 | 新建的 演示文稿 对象。 |

**示例**

当 打开任一演示文稿 时，弹消息框提醒用户打开了演示文稿。

``` JavaScript
function Application_PresentationOpen(Pres)
{
    MsgBox("您打开了演示文稿。")
}
```

### PresentationSave

任一 演示文稿被保存 时 触发 此事件。

**语法**

`function Application_PresentationSave ( Pres ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型         | 说明                   |
|------|-----------|------------------|------------------------|
| Pres | 必选      | Presentation对象 | 保存的 演示文稿 对象。 |

**示例**

当保存演示文稿时，弹消息框提醒用户 正在保存演示文稿 。

``` JavaScript
function Application_PresentationSave(pre)
{
    MsgBox("正在保存演示文稿："+ pre.Name)
}
```

### WindowActivate

任一 演示文稿 窗口被激活时，将 触发 此事件。

**语法**

`function Application_WindowActivate ( Pres, Wn ) { function_body_statements }`

*function_body_statements* 代表了响应函数的函数体的语句。

**参数**

| 名称 | 必选/可选 | 数据类型           | 说明                   |
|------|-----------|--------------------|------------------------|
| Pres | 必选      | Presentation对象   | 新建的 演示文稿 对象。 |
| Wn   | 必选      | DocumentWindow对象 | 激活的文档窗口对象。   |

**示例**

当切换演示文稿窗口时，输出激活的窗口标题。

``` JavaScript
function Application_WindowActivate(pre, win)
{
    Debug.Print("当前激活窗口是："+ win.Caption)
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/事件/演示事件/Application 事件](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E4%BA%8B%E4%BB%B6/%E6%BC%94%E7%A4%BA%E4%BA%8B%E4%BB%B6/Application%20%E4%BA%8B%E4%BB%B6.html)

------------------------------------------------------------------------
# DoEvents

处理进程的消息队列中的消息。

## 语法

**DoEvents** ()

## 说明

DoEvents 将暂时停止当前的宏执行，在 处理完进程的消息队列中的消息 后返回再继续宏的执行。

## 示例

此示例使用 DoEvents 函数来使得每迭代 1000 次循环就将 暂时停止当前的宏执行，处理完 进程的消息队列中的消息 后返回再继续宏的执行，以便迭代过程中在合适的时机更新文档视图或者响应鼠键消息来通过调试器中止迭代 。

``` JavaScript
function abcd()
{
    var beginTime = new Date();
    for (i = 0; i < 20000; i++) 
    {
        Selection.TypeText("金山办公软件");
        if (i % 1000 == 0)
            DoEvents()
    }
    var endTime = new Date();
    Debug.Print("用时共计"+(endTime-beginTime)+"ms");   
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/函数/DoEvents](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E5%87%BD%E6%95%B0/DoEvents.html)

------------------------------------------------------------------------
# InputBox

弹出 显示自定义提示信息 的输入对话框，等待用户在输入框中输入文本或单击按钮，然后返回输入框的内容。

## 语法

**InputBox** ( prompt, \[ *title* \], \[ *default* \], \[ *xpos* \], \[ *ypos* \] )

## 参数

| 名字      | 说明                                         |
|-----------|-------------------------------------------------------------------|
| *prompt*    | 必需项。 在对话框中显示的 消息的 字符串表达式 。                              |
| *title*   | 可选。 对话框标题栏中显示的字符串表达式。 如果省略 *title* ，则标题栏中将显示应用程序名称。                            |
| *default* | 可选。 文本框中显示的字符串表达式，在未提供其他输入时作为默认响应。 如果省略了 *default* ，文本框将显示为空。          |
| *xpos*    | 可选。 指定对话框的左边缘与屏幕的左边缘的水平距离（以缇为单位）的数值表达式。 如果省略了 *xpos* ，对话框将水平居中。   |
| *ypos*    | 可选。 指定对话框的上边缘与屏幕的顶部的垂直距离（以缇为单位）的数值表达式。 如果省略了 *ypos* ，对话框将 将垂直居中 。 |

## 返回值

如果用户选择"**确定**" 或按 **Enter**，**InputBox** 函数将返回文本框中的内容。 如果用户选择"**取消**"，函数将返回一个空字符串 。

## 示例

下面的代码是就是 **InputBox** 的一个示例。

``` JavaScript
function abcd()
{
    var strMessage = "这是一个输入对话框，请输入？";
    var strTitle = "我是标题";
    var strDefault = "我是标题";
    var ret = InputBox(strMessage, strTitle, strDefault, 3000, 4000)
    Debug.Print("输入对话框返回值是：" + ret)
}
```

![](Base64图像/Base64图像45来自_WPS%20基础接口_宏编辑器API参考_函数_InputBox.png)

![](Base64图像/Base64图像46来自_WPS%20基础接口_宏编辑器API参考_函数_InputBox.png)

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/函数/InputBox](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E5%87%BD%E6%95%B0/InputBox.html)

------------------------------------------------------------------------
# MsgBox

弹出消息对话框，等待用户单击按钮，并返回一个整数，指示用户单击的哪个按钮。

## 语法

**MsgBox** ( prompt, \[ *buttons* \], \[ *title* \] )

## 参数

| 名称      | 说明                                                                           |
|-----------|---------------------------------------------------------------------------------------------------|
| *prompt*    | 必需项。 在对话框中显示的 消息的 字符串表达式 。                            
| *buttons* | 可选。 数值表达式，用于指定要显示按钮的数量和类型、要使用的图标样式、默认按钮的标识和消息框的形式的值（ JSMsgBoxStyle ）之和。 如果省略，则 buttons 的默认值为 0。 |
| *title*   | 可选。 对话框 标题栏中显示的字符串表达式。 如果省略 title，则标题栏中将显示应用程序名称。             |

## 返回值

返回 JSMsgBoxResult 枚举 值。表示 用户点击MsgBox弹出消息框的按钮。 在 点击MsgBox弹出消息框 之前不会返回。 如果对话框中显示“ **取消** ”按钮，按 ESC 键与单击“ **取消** ”具有相同的作用。

## 示例

下面的代码是就是MsgBox的一个示例。

![](Base64图像/Base64图像47来自_WPS%20基础接口_宏编辑器API参考_函数_MsgBox.png)

![](Base64图像/Base64图像48来自_WPS%20基础接口_宏编辑器API参考_函数_MsgBox.png)

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/函数/MsgBox](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E5%87%BD%E6%95%B0/MsgBox.html)

------------------------------------------------------------------------
# 控件概述

WPS宏编辑器提供了命令按钮、 标签、文本框、复合框、列表框、复选框等控件，其跟api对象对应关系如下。

| 控件类型       | API对象       | 功能         |
|----------------|---------------|--------------------------------------------------------|
| 窗体       | UserForm      | 用于选中或是移动控件对象。              |
| 命令按钮   | CommandButton | 创建一个可以让用户选择，以完成一个命令的按钮                      |
| 标签       | Label         | 用于显示用户不能编辑的文本或图像，例如图形下的标题文本。          |
| 文本框     | TextEdit      | 用于输入或改变文本内容。                |
| 复合框     | ComboBox      | 复合框由一个文本输入控件和一个下拉菜单组成，用户可以从预先定义的列表中的选择一个选项，同时也可以直接通过文本框输入 |
| 列表框     | ListBox       | 用来显示用户可以选择的项目列表。如果不能一次显示全部项目的话，列表可以滚动。                 |
| 复选框     | CheckBox      | 可以显示出多重选择，用户可以选择一个或者多个选项，如果选中多个选项则显示出多重选择。         |
| 选项按钮   | OptionButton  | 可以显示出多重选择，用户只能选择一个。  |
| 滚动条     | ScrollBar   |提供在长列表工程或大量信息中快速浏览的图形工具，以比例方式指示出当前位置，或是做为一个输入设备，或速度及数量的指示器。 |
| 水平布局器 | HLayout       | 将内部的控件按照水平方向排布，一列一个。|
| 垂直布局器 | VLayout       | 将内部的控件按照垂直方向排布，一行一个。|
| 水平间隔   | HSpacer       | 填充水平方向上无用的空隙。              |
| 垂直间隔   | VSpacer       | 填充垂直方向上无用的空隙。              |
| 旋转按钮   | SpinButton    | 一种与其它控件并用微调控制的控件，也可以用来对一个范围的值或工程列表做向前或向后的遍历。     |
| 切换按钮   | ToggleButton  | 创建一个切换开关的按钮。                |
| 框架       | Frame         | 可以创建一个图形或控件的功能组。要为多个控件创建组，可先画框架，接着在框架中添加控件。       |
| 多页       | MultiPage   | 包含一个选项卡栏和一个页面区的容器类组件，通过切换选项卡来切换不同页面，从而达到在同一个窗口中自由切换不同的内容。  |

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/控件API参考/控件概述](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E6%8E%A7%E4%BB%B6API%E5%8F%82%E8%80%83/%E6%8E%A7%E4%BB%B6%E6%A6%82%E8%BF%B0.html)

------------------------------------------------------------------------
# 扩展接口使用说明

**基本介绍**

WPS的扩展接口是对WPS现有接口的进一步扩充，大大的丰富了WPS的接口，以便满足二次开发各种解决方案的需要。WPS所有的扩展接口都满足唯一性、对称性、自反性、传递性，也就是说通过扩展接口可以查询到对应的原始接口，同样的通过原始接口也可以查到对应的扩展接口。

> WPS 开发文档： [WPS 基础接口/WPS 扩展 API/扩展接口使用说明](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/WPS%20%E6%89%A9%E5%B1%95%20API/%E6%89%A9%E5%B1%95%E6%8E%A5%E5%8F%A3%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E.html)

------------------------------------------------------------------------
# NativeX扩展对象

## 背景

NativeX是提供给第三方开发者用native代码来扩展JS接口的技术方案。JSAPI提供的接口个数有限，当碰到不满足需求的情况下，第三方开发者可以用native代码(c++/ruby/python等)来扩展JS接口，加载第三方NativeX模块后，实现在JSAPI中使用自己扩展的接口。例如第三方开发者可以利用发布的SDK，将复杂的算法（如加密，解密算法）实现并生成dll或者so，提供给NativeX对象加载在JSAPI中使用。

## 使用SDK生成dll

NativeX对象中有三种调用方式Get,Set和Execute，在SDK中注册属性和函数名，指定回调函数，参数个数和参数类型，示例代码如下。

![](Base64图像/Base64图像49来自_WPS%20基础接口_WPS%20扩展%20API_NativeX扩展对象.png)

JSWpsRootObject为根对象，如果想返回NativeX子对象，创建JSWpsSubObject_1类，同样是继承父类NativeXObjectBase，在根对象处理返回值时，将NativeX子对象的指针返回即可。

![](Base64图像/Base64图像50来自_WPS%20基础接口_WPS%20扩展%20API_NativeX扩展对象.png)

## 注册工具

使用NativeX扩展对象之前，需要使用ksomisc工具将dll路径注册到文件中。

![](Base64图像/Base64图像51来自_WPS%20基础接口_WPS%20扩展%20API_NativeX扩展对象.png)

调用ksomisc程序，-regnativex参数设置dll路径， -unregnativex参数反注册， -i参数设置是否为进程内调用。在%AppData%\kingsoft\wps\jsaddons\binary\目录下生成WPSNativeX.conf文件，WPSNativeX.conf文件中包含了此dll中是否曾经发生过崩溃信息，dll路径信息，是否进程内调用信息。

![](Base64图像/Base64图像52来自_WPS%20基础接口_WPS%20扩展%20API_NativeX扩展对象.png)

## JSAPI中调用Nativex对象

使用根对象下的CreateObject方法创建NativeX对象，参数为利用SDK生成dll的名字。可以先将使用代码片段写入到加载项的按钮中，也可以直接在JS调试器中单步调用。NativeX扩展对象支持创建子对象，支持创建多实例，支持创建多对象。

![](Base64图像/Base64图像53来自_WPS%20基础接口_WPS%20扩展%20API_NativeX扩展对象.png)

> WPS 开发文档： [WPS 基础接口/WPS 扩展 API/NativeX扩展对象](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/WPS%20%E6%89%A9%E5%B1%95%20API/NativeX%E6%89%A9%E5%B1%95%E5%AF%B9%E8%B1%A1.html)

------------------------------------------------------------------------
# 表格idMso参考

| 控件名称                        | 控件类型     | 所属Tab                   | 功能描述                          |
|---------------------------------|--------------|---------------------------|-----------------------------------|
|                                 | comboBox     | TabChartTools             | 图表元素                          |
| PivotClearFilters               | button       | TabData                   | 全部显示                          |
| SortClear                       | button       | TabData                   | 全部显示                          |
| FilterReapply                   | button       | TabData                   | 重新应用                          |
| AdvancedFilterDialog            | toggleButton | TabData                   | 高级                              |
| SortAscendingExcel              | button       | TabData                   |                                   |
| SortDescendingExcel             | button       | TabData                   |                                   |
| AutoFilterClassic               | button       | TabData                   | 自动筛选                          |
| FillDown                        | button       | None                      | 向下填充                          |
| FillRight                       | button       | None                      | 右对齐                            |
| InputNumericSequence            | button       | TabData                   | 录入123序列                       |
| InputLetterSequence             | button       | TabData                   | 录入ABC序列                       |
| InputRomanNumericSequence       | button       | TabData                   | 录入罗马数字序列                  |
| FillBlankCells                  | button       | TabData                   | 空白单元格填充                    |
| DataValidation                  | splitButton  | TabData                   | 数据有效性                        |
| DataValidation                  | toggleButton | TabData                   | 数据有效性                        |
| DataValidationCircle            | toggleButton | TabData                   | 圈释无效数据                      |
| DataValidationCircleClear       | toggleButton | TabData                   | 清除验证标识圈                    |
|                                 | button       | TabData                   | 插入下拉列表                      |
| Consolidate                     | button       | TabData                   | 合并计算                          |
| OutlineUngroupMenu              | splitButton  | TabData                   | 取消组合                          |
| OutlineUngroup                  |              | None                      | 取消组合                          |
| OutlineSubtotals                | button       | TabData                   | 分类汇总                          |
| OutlineShowDetail               | button       | TabData                   | 展开明细                          |
| OutlineHideDetail               | button       | TabData                   | 折叠明细                          |
| OutlineGroup                    |              | None                      | 组合                              |
| RefreshAll                      | button       | TabData                   | 全部刷新                          |
|                                 | button       | TabData                   | 导入文本                          |
| RefreshAllMenu                  | splitButton  | TabData                   | 全部刷新                          |
| Refresh                         | button       | None                      | 刷新数据                          |
| ClearMenu                       | menu         | TabDesign                 | 清除表格样式                      |
| MacroPlay                       | button       | TabDevelopTools           | JS宏                              |
| MacroRecord                     | button       | TabDevelopTools           | 录制新宏                          |
| MacroRelativeReferences         | toggleButton | TabDevelopTools           | 使用相对引用                      |
| MacroSecurity                   | button       | TabDevelopTools           | 宏安全性                          |
| VisualBasic                     | button       | TabDevelopTools           | WPS 宏编辑器                      |
| ComAddInsDialog                 | button       | TabDevelopTools           | COM 加载项                        |
| DesignMode                      | toggleButton | TabDevelopTools           | 设计模式                          |
| ControlProperties               | button       | TabDevelopTools           | 控件属性                          |
| ViewCode                        | button       | TabDevelopTools           | 查看代码                          |
| ActiveXCheckBox                 | button       | TabDevelopTools           | 复选框                            |
| ActiveXTextBox                  | button       | TabDevelopTools           | 文本框                            |
| ActiveXButton                   | button       | TabDevelopTools           | 命令按钮                          |
| ActiveXListBox                  | button       | TabDevelopTools           | 列表框                            |
| ActiveXComboBox                 | button       | TabDevelopTools           | 组合框                            |
| ActiveXToggleButton             | button       | TabDevelopTools           | 切换按钮                          |
| ActiveXSpinButton               | button       | TabDevelopTools           | 数值调节按钮                      |
| ActiveXScrollBar                | button       | TabDevelopTools           | 滚动条                            |
| ActiveXLabel                    | button       | TabDevelopTools           | 标签                              |
| ActiveXImage                    | button       | TabDevelopTools           | 图像                              |
| MoreControlsDialog              | button       | TabDevelopTools           | 其他控件                          |
|                                 | button       | TabDevelopTools           | 命令按钮                          |
|                                 | button       | TabDevelopTools           | 复选框                            |
|                                 | button       | TabDevelopTools           | 列表框                            |
|                                 | button       | TabDevelopTools           | 组合框                            |
|                                 | button       | TabDevelopTools           | 文本框                            |
|                                 | button       | TabDevelopTools           | 标签                              |
|                                 | button       | TabDevelopTools           | 切换按钮                          |
|                                 | button       | TabDevelopTools           | 滚动条                            |
|                                 | button       | TabDevelopTools           | 数值调节按钮                      |
| TextBoxInsertHorizontal         | toggleButton | TabDrawingTools           | 横向文本框                        |
| TextBoxInsertVertical           | toggleButton | TabDrawingTools           | 竖向文本框                        |
| ObjectEditPoints                | toggleButton | TabDrawingTools           | 编辑顶点                          |
| ObjectPictureFill               | button       | TabDrawingTools           | 图片                              |
| OutlineLinePatternFill          | button       | TabDrawingTools           | 带图案线条                        |
| LineStylesDialog                | button       | TabDrawingTools           | 其他线条                          |
| ArrowStyleGallery               | gallery      | TabDrawingTools           | 箭头样式                          |
| ArrowsMore                      | button       | TabDrawingTools           | 其他箭头                          |
| ShapeStylesGallery              | gallery      | TabDrawingTools           | 形状样式                          |
| ShapeFillColorPicker            | gallery      | TabDrawingTools           | 填充                              |
| GradientGallery                 | gallery      | None                      | 渐变                              |
| MoreTextureOptions              | button       | None                      | 图片或纹理                        |
| ShapeOutlineColorPicker         | gallery      | TabDrawingTools           | 轮廓                              |
| ObjectBringForward              | button       | TabDrawingTools           | 上移一层                          |
| ObjectBringToFront              | button       | TabDrawingTools           | 置于顶层                          |
| ObjectSendBackward              | button       | TabDrawingTools           | 下移一层                          |
| ObjectSendToBack                | button       | TabDrawingTools           | 置于底层                          |
| ObjectsAlignLeft                | button       | TabDrawingTools           | 左对齐                            |
| ObjectsAlignCenterHorizontal    | button       | TabDrawingTools           | 水平居中                          |
| ObjectsAlignRight               | button       | TabDrawingTools           | 右对齐                            |
| ObjectsAlignTop                 | button       | TabDrawingTools           | 顶端对齐                          |
| ObjectsAlignMiddleVertical      | button       | TabDrawingTools           | 垂直居中                          |
| ObjectsAlignBottom              | button       | TabDrawingTools           | 底端对齐                          |
| AlignDistributeHorizontally     | button       | TabDrawingTools           | 横向分布                          |
| AlignDistributeVertically       | button       | TabDrawingTools           | 纵向分布                          |
| SnapToGrid                      | toggleButton | TabDrawingTools           | 对齐网格                          |
| ObjectsGroup                    | button       | TabDrawingTools           | 组合                              |
| ObjectsUngroup                  | button       | TabDrawingTools           | 取消组合                          |
| ObjectRotateFree                | button       | TabDrawingTools           | 自由旋转                          |
| ObjectRotateLeft90              | button       | TabDrawingTools           | 向左旋转 90°                      |
| ObjectRotateRight90             | button       | TabDrawingTools           | 向右旋转 90°                      |
| ObjectFlipHorizontal            | button       | TabDrawingTools           | 水平翻转                          |
| ObjectFlipVertical              | button       | TabDrawingTools           | 垂直翻转                          |
| SelectionPane                   | toggleButton | TabDrawingTools           | 选择窗格                          |
|                                 | editBox      | TabDrawingTools           | 大小                              |
|                                 | editBox      | TabDrawingTools           |                                   |
| ObjectFormatDialog              | button       | TabDrawingTools           | 设置对象格式                      |
| ObjectEditPoints                | toggleButton | TabDrawingTools_Vml       | 编辑顶点                          |
| ObjectPictureFill               | button       | TabDrawingTools_Vml       | 图片                              |
| GradientGallery                 | gallery      | None                      | 渐变                              |
| MoreTextureOptions              | button       | None                      | 图片或纹理                        |
| LineStylesDialog                | button       | TabDrawingTools_Vml       | 其他线条                          |
| ArrowStyleGallery               | gallery      | TabDrawingTools_Vml       | 箭头样式                          |
| ArrowsMore                      | button       | TabDrawingTools_Vml       | 其他箭头                          |
| ObjectsAlignLeft                | button       | TabDrawingTools_Vml       | 左对齐                            |
| ObjectsAlignCenterHorizontal    | button       | TabDrawingTools_Vml       | 水平居中                          |
| ObjectsAlignRight               | button       | TabDrawingTools_Vml       | 右对齐                            |
| ObjectsAlignTop                 | button       | TabDrawingTools_Vml       | 顶端对齐                          |
| ObjectsAlignMiddleVertical      | button       | TabDrawingTools_Vml       | 垂直居中                          |
| ObjectsAlignBottom              | button       | TabDrawingTools_Vml       | 底端对齐                          |
| SnapToGrid                      | toggleButton | TabDrawingTools_Vml       | 对齐网格                          |
| ObjectsGroup                    | button       | TabDrawingTools_Vml       | 组合                              |
| ObjectsUngroup                  | button       | TabDrawingTools_Vml       | 取消组合                          |
| ObjectRotateFree                | button       | TabDrawingTools_Vml       | 自由旋转                          |
| ObjectRotateLeft90              | button       | TabDrawingTools_Vml       | 向左旋转 90°                      |
| ObjectRotateRight90             | button       | TabDrawingTools_Vml       | 向右旋转 90°                      |
| ObjectFlipHorizontal            | button       | TabDrawingTools_Vml       | 水平翻转                          |
| ObjectFlipVertical              | button       | TabDrawingTools_Vml       | 垂直翻转                          |
|                                 | comboBox     | TabDrawingTools_Vml       | 字体                              |
| ClearContents                   | button       | None                      | 内容                              |
|                                 | editBox      | TabDrawingTools_Vml       | 高度:                             |
|                                 | editBox      | TabDrawingTools_Vml       | 宽度:                             |
| EquationOptions                 |              | TabEquationTools          | 公式选项                          |
| EquationInsertGallery           |              | TabEquationTools          | 公式                              |
| EquationProfessional            |              | TabEquationTools          | 专业型                            |
| EquationLinearFormat            |              | TabEquationTools          | 线性                              |
| EquationNormalText              |              | TabEquationTools          | 普通文本                          |
| EquationSymbolsInsertGallery    |              | TabEquationTools          | 公式符号                          |
| EquationFractionGallery         |              | TabEquationTools          | 分数                              |
| EquationScriptGallery           |              | TabEquationTools          | 上下标                            |
| EquationRadicalGallery          |              | TabEquationTools          | 根式                              |
| EquationIntegralGallery         |              | TabEquationTools          | 积分                              |
| EquationLargeOperatorGallery    |              | TabEquationTools          | 大型运算符                        |
| EquationDelimiterGallery        |              | TabEquationTools          | 括号                              |
| EquationFunctionGallery         |              | TabEquationTools          | 函数                              |
| EquationAccentGallery           |              | TabEquationTools          | 导数符号                          |
| EquationLimitGallery            |              | TabEquationTools          | 极限和对数                        |
| EquationOperatorGallery         |              | TabEquationTools          | 运算符                            |
| EquationMatrixGallery           |              | TabEquationTools          | 矩阵                              |
| InputNumericSequence            | button       | TabEtToolbox              | 录入123序列                       |
| InputLetterSequence             | button       | TabEtToolbox              | 录入ABC序列                       |
| InputRomanNumericSequence       | button       | TabEtToolbox              | 录入罗马数字序列                  |
| FillBlankCells                  | button       | TabEtToolbox              | 空白单元格填充                    |
| NumAsText                       | button       | TabEtToolbox              | 数字转为文本型数字                |
| TextAsNum                       | button       | TabEtToolbox              | 文本型数字转为数字                |
| KeepNumberical                  | button       | TabEtToolbox              | 仅保留数值                        |
| FormulaText                     | button       | TabEtToolbox              | 公式转文本                        |
| Round                           | button       | TabEtToolbox              | 四舍五入                          |
| InsertDiagonalHeader            | button       | TabEtToolbox              | 插入斜线表头                      |
| BatchAdd                        | button       | TabEtToolbox              | 加上                              |
| BatchSub                        | button       | TabEtToolbox              | 减去                              |
| BatchMul                        | button       | TabEtToolbox              | 乘以                              |
| BatchDiv                        | button       | TabEtToolbox              | 除以                              |
| DeleteBlankTable                | button       | TabEtToolbox              | 删除空白表                        |
| RemoveBlankRow                  | button       | TabEtToolbox              | 删除空行                          |
| RemoveLeadingSpace              | button       | TabEtToolbox              | 删除开头空格                      |
| RemoveTrailingSpace             | button       | TabEtToolbox              | 删除末尾空格                      |
| RemoveAllSpace                  | button       | TabEtToolbox              | 删除所有空格                      |
| FileSaveAsPicture               |              | TabFile                   | 输出为图片                        |
| FileNewMenu                     | menu         | TabFile                   | 新建                              |
| FileNewDefault                  | button       | TabFile                   | 新建                              |
| NewBlankXlsxFile                | button       | TabFile                   | 新建 Excel 2007/2010 文件         |
| FileNew                         |              | TabFile                   | 本机上的模板                      |
| NewFromDefaultTemplate          | button       | TabFile                   | 从默认模板新建                    |
| FilePrintMenu                   | splitButton  | TabFile                   | 打印                              |
| FilePrint                       | button       | TabFile                   | 打印                              |
| FilePrintPreview                | button       | TabFile                   | 打印预览                          |
|                                 | splitButton  | TabFile                   |                                   |
|                                 | button       | TabFile                   |                                   |
|                                 | button       | TabFile                   |                                   |
|                                 | splitButton  | TabFile                   | 备份与恢复                        |
|                                 | splitButton  | TabFile                   | 备份与恢复                        |
| FileInfoMenu                    | menu         | TabFile                   | 文档加密                          |
| FileMenuFileEncrypt             |              | TabFile                   | 文件加密                          |
| FileProperties                  | button       | TabFile                   | 属性                              |
| FileBackupManagement            | button       | TabFile                   | 备份管理                          |
| FileBackupManagement            | button       | TabFile                   | 备份管理                          |
| FileBackupHistory               | button       | TabFile                   | 历史版本                          |
| FileHelp                        | menu         | TabFile                   | 帮助                              |
| Help                            | button       | TabFile                   | WPS 表格 帮助                     |
| About                           | button       | TabFile                   | 关于 WPS 表格                     |
| FileSave                        | button       | TabFile                   | 保存                              |
| FileOpen                        | button       | TabFile                   | 打开                              |
| FileSave                        | button       | TabFile                   | 保存                              |
| FileSaveAsMenu                  | menu         | None                      | 另存为                            |
| FileSaveAsPdfOrXps              | button       | None                      | 输出为PDF                         |
| FileSlimming                    |              | None                      | 文件瘦身                          |
| DocumentSplitAndMerge           | menu         | None                      | 文档拆分合并                      |
| FileShare                       | button       | TabFile                   | 分享                              |
| FileMenuSendMail                | button       | TabFile                   | 发送邮件                          |
| ApplicationOptionsDialog        |              | TabFile                   | 选项                              |
| AutoSum                         | button       | TabFormulas               | 求和                              |
| AutoSumAverage                  | button       | TabFormulas               | 平均值                            |
| AutoSumCount                    | button       | TabFormulas               | 计数                              |
| AutoSumMax                      | button       | TabFormulas               | 最大值                            |
| AutoSumMin                      | button       | TabFormulas               | 最小值                            |
| AutoSumMoreFunctions            | button       | TabFormulas               | 其他函数                          |
| AutoSumMenu                     | menu         | TabFormulas               | 自动求和                          |
| FunctionWizard                  | button       | TabFormulas               | 插入函数                          |
| FormulaMoreFunctionsMenu        | menu         | TabFormulas               | 其他函数                          |
| NameManager                     | button       | TabFormulas               | 名称管理器                        |
| NameDefine                      | button       | TabFormulas               | 定义名称                          |
| NamePasteName                   | button       | TabFormulas               | 粘贴                              |
| NameCreateFromSelection         | button       | TabFormulas               | 指定                              |
| NameDefineMenu                  | menu         | TabFormulas               | 定义名称                          |
| FormulaEvaluate                 | button       | TabFormulas               | 公式求值                          |
| TracePrecedents                 | button       | TabFormulas               | 追踪引用单元格                    |
| TraceDependents                 | button       | TabFormulas               | 追踪从属单元格                    |
| TraceRemoveAllArrows            | button       | TabFormulas               | 移去箭头                          |
|                                 | button       | TabFormulas               | 移去引用单元格追踪箭头            |
|                                 | button       | TabFormulas               | 移去从属单元格追踪箭头            |
|                                 | splitButton  | TabFormulas               | 移去箭头                          |
|                                 | splitButton  | TabFormulas               | 移去箭头                          |
| ErrorCheckingMenu               | splitButton  | TabFormulas               |                                   |
| ShowFormulas                    | button       | None                      | 显示公式                          |
| CalculateSheet                  | button       | TabFormulas               | 计算工作表                        |
| EditLinks                       | button       | TabFormulas               | 编辑链接                          |
| Connections                     | group        | TabFormulas               | 连接                              |
| ObjectPictureFill               | button       | TabGraphicTool            | 图片                              |
| OutlineLinePatternFill          | button       | TabGraphicTool            | 带图案线条                        |
| LineStylesDialog                | button       | TabGraphicTool            | 其他线条                          |
| ShapeStylesGallery              | gallery      | TabGraphicTool            | 图形样式                          |
| ShapeFillColorPicker            | gallery      | TabGraphicTool            | 图形填充                          |
| GradientGallery                 | gallery      | None                      | 渐变                              |
| ShapeOutlineColorPicker         | gallery      | TabGraphicTool            | 图形轮廓                          |
| ObjectBringForward              | button       | TabGraphicTool            | 上移一层                          |
| ObjectBringToFront              | button       | TabGraphicTool            | 置于顶层                          |
| ObjectSendBackward              | button       | TabGraphicTool            | 下移一层                          |
| ObjectSendToBack                | button       | TabGraphicTool            | 置于底层                          |
| ObjectsAlignLeft                | button       | TabGraphicTool            | 左对齐                            |
| ObjectsAlignCenterHorizontal    | button       | TabGraphicTool            | 水平居中                          |
| ObjectsAlignRight               | button       | TabGraphicTool            | 右对齐                            |
| ObjectsAlignTop                 | button       | TabGraphicTool            | 顶端对齐                          |
| ObjectsAlignMiddleVertical      | button       | TabGraphicTool            | 垂直居中                          |
| ObjectsAlignBottom              | button       | TabGraphicTool            | 底端对齐                          |
| AlignDistributeHorizontally     | button       | TabGraphicTool            | 横向分布                          |
| AlignDistributeVertically       | button       | TabGraphicTool            | 纵向分布                          |
| SnapToGrid                      | toggleButton | TabGraphicTool            | 对齐网格                          |
| ObjectsGroup                    | button       | TabGraphicTool            | 组合                              |
| ObjectsUngroup                  | button       | TabGraphicTool            | 取消组合                          |
| ObjectRotateFree                | button       | TabGraphicTool            | 自由旋转                          |
| ObjectRotateLeft90              | button       | TabGraphicTool            | 向左旋转 90°                      |
| ObjectRotateRight90             | button       | TabGraphicTool            | 向右旋转 90°                      |
| ObjectFlipHorizontal            | button       | TabGraphicTool            | 水平翻转                          |
| ObjectFlipVertical              | button       | TabGraphicTool            | 垂直翻转                          |
| SelectionPane                   | toggleButton | TabGraphicTool            | 选择窗格                          |
|                                 | button       | None                      | 重设形状和大小                    |
|                                 | editBox      | TabGraphicTool            | 高度:                             |
|                                 | editBox      | TabGraphicTool            | 宽度:                             |
| ObjectFormatDialog              | button       | TabGraphicTool            | 设置对象格式                      |
| PasteAsPicture                  | button       | TabHome                   | 粘贴为图片                        |
| PasteFormatting                 | button       | TabHome                   | 带格式粘贴                        |
|                                 | button       | TabHome                   | 粘贴值到可见单元格                |
| PasteMenu                       | splitButton  | TabHome                   | 粘贴                              |
| PasteAllUsingSourceTheme        | button       | TabHome                   | 保留源格式                        |
| PasteValues                     | button       | TabHome                   | 值                                |
| PasteFormulas                   | button       | TabHome                   | 公式                              |
| PasteNoBorders                  | button       | TabHome                   | 无边框                            |
| PasteTranspose                  | button       | TabHome                   | 转置                              |
| PasteSpecialDialog              | button       | None                      | 选择性粘贴                        |
| Cut                             | button       | TabHome                   | 剪切                              |
| Copy                            | button       | TabHome                   | 复制                              |
| CopyPictureDialog               |              | TabHome                   | 复制为图片                        |
| ShowClipboard                   | button       | TabHome                   | 剪贴板                            |
| FormatPainter                   | toggleButton | None                      | 格式刷                            |
| Font                            | comboBox     | TabHome                   | 字体                              |
| FontSize                        | comboBox     | TabHome                   | 字号                              |
| FontSizeIncrease                | button       | TabHome                   | 增大字号                          |
| FontSizeDecrease                | button       | TabHome                   | 减小字号                          |
| Bold                            | toggleButton | TabHome                   | 加粗                              |
| Italic                          | toggleButton | TabHome                   | 倾斜                              |
| Underline                       | toggleButton | TabHome                   | 下划线                            |
|                                 | toggleButton | TabHome                   | 下划线                            |
|                                 | toggleButton | TabHome                   |                                   |
|                                 | toggleButton | TabHome                   |                                   |
| IndentDecreaseExcel             | button       | TabHome                   | 减少缩进量                        |
| IndentIncreaseExcel             | button       | TabHome                   | 增加缩进量                        |
| TextOrientationVertical         | toggleButton | TabHome                   | 竖排文字                          |
| TextOrientationRotateUp         | toggleButton | TabHome                   | 向上旋转文字                      |
| TextOrientationRotateDown       | toggleButton | TabHome                   | 向下旋转文字                      |
| BorderDrawLine                  | toggleButton | TabHome                   | 绘图边框                          |
| BorderDrawGrid                  | toggleButton | TabHome                   | 绘图边框网格                      |
| BorderErase                     | toggleButton | TabHome                   | 擦除边框                          |
| BorderStyle                     | dropDown     | TabHome                   | 线条样式                          |
| BorderColorPickerExcel          | gallery      | TabHome                   | 线条颜色                          |
| ClearAll                        | button       | TabHome                   | 全部                              |
| ClearFormats                    | button       | TabHome                   | 格式                              |
| ClearContents                   | button       | TabHome                   | 内容                              |
| ClearComments                   | button       | TabHome                   | 批注                              |
| BordersGallery                  | splitButton  | None                      |                                   |
| CellFillColorPicker             | gallery      | TabHome                   | 填充颜色                          |
| FontColorPicker                 | gallery      | TabHome                   | 字体颜色                          |
| GroupFont                       |              | TabHome                   | 字体                              |
|                                 | toggleButton | TabHome                   | 从左向右                          |
|                                 | toggleButton | TabHome                   | 从右向左                          |
|                                 | toggleButton | TabHome                   | 上下文                            |
| AlignTopExcel                   | toggleButton | TabHome                   | 顶端对齐                          |
| AlignBottomExcel                | toggleButton | TabHome                   | 底端对齐                          |
| AlignRight                      | toggleButton | TabHome                   | 右对齐                            |
| AlignJustify                    | toggleButton | TabHome                   | 两端对齐                          |
| AlignDistributed                | toggleButton | TabHome                   | 分散对齐                          |
| MergeCenter                     | toggleButton | TabHome                   | 合并居中                          |
| GroupAlignmentExcel             |              | TabHome                   | 对齐方式                          |
| AlignMiddleExcel                | toggleButton | None                      | 垂直居中                          |
| AlignLeft                       | toggleButton | None                      | 左对齐                            |
| AlignCenter                     | toggleButton | None                      | 水平居中                          |
| MergeCenterMenu                 | splitButton  | None                      | 合并单元格                        |
| NumberFormatGallery             | comboBox     | TabHome                   | 数字                              |
| PercentStyle                    | button       | TabHome                   | 百分比样式                        |
| CommaStyle                      | button       | TabHome                   | 千位分隔样式                      |
| DecimalsIncrease                | button       | TabHome                   | 增加小数位数                      |
| DecimalsDecrease                | button       | TabHome                   | 减少小数位数                      |
| GroupNumber                     |              | TabHome                   | 数字                              |
| AdvancedFilterDialog            | toggleButton | TabHome                   | 高级筛选                          |
| ConditionalFormattingMenu       | menu         | TabHome                   | 条件格式                          |
| Lock                            | toggleButton | TabHome                   | 锁定单元格                        |
| FormatCellsDialog               | button       | TabHome                   | 设置单元格格式                    |
| RowHeight                       | button       | TabHome                   | 行高                              |
| RowHeightAutoFit                | button       | TabHome                   | 最适合的行高                      |
| ColumnWidth                     | button       | TabHome                   | 列宽                              |
| ColumnWidthAutoFit              | button       | TabHome                   | 最适合的列宽                      |
| SheetRowsInsert                 |              | TabHome                   | 插入行                            |
| SheetColumnsInsert              | button       | TabHome                   | 插入列                            |
| CellsInsertDialog               | button       | TabHome                   | 插入单元格                        |
| CellsDeleteSmart                | button       | TabHome                   | 删除单元格                        |
| CellsDeleteSmart                | button       | TabHome                   | 删除单元格                        |
| CellStylesGallery               | splitButton  | TabHome                   |                                   |
| FillDown                        | button       | TabHome                   | 向下填充                          |
| FillRight                       | button       | TabHome                   | 向右填充                          |
| FillUp                          | button       | TabHome                   | 向上填充                          |
| FillLeft                        | button       | TabHome                   | 向左填充                          |
| InputNumericSequence            | button       | TabHome                   | 录入123序列                       |
| InputLetterSequence             | button       | TabHome                   | 录入ABC序列                       |
| InputRomanNumericSequence       | button       | TabHome                   | 录入罗马数字序列                  |
| FillBlankCells                  | button       | TabHome                   | 空白单元格填充                    |
| FillSeries                      | button       | TabHome                   | 序列                              |
| RowsHide                        |              | TabHome                   | 隐藏行                            |
| RowsUnhide                      | button       | TabHome                   | 取消隐藏行                        |
| SheetInsert                     | button       | TabHome                   | 插入工作表                        |
| SheetProtect                    | button       | TabHome                   | 保护工作表                        |
| SheetRename                     | button       | TabHome                   | 重命名                            |
| SheetHide                       | button       | TabHome                   | 隐藏工作表                        |
| SheetUnhide                     | button       | TabHome                   | 取消隐藏工作表                    |
| FindDialogExcel                 | button       | TabHome                   | 查找                              |
| GoTo                            | button       | TabHome                   | 定位                              |
| ObjectsSelect                   | toggleButton | TabHome                   | 选择对象                          |
|                                 | group        | TabHome                   | 智能工具箱                        |
| GroupEditing                    | group        | TabHome                   | 编辑                              |
| FillMenu                        | menu         | TabHome                   | 填充                              |
|                                 | group        | TabHome                   | 设置单元格格式                    |
| HideAndUnhideMenu               | menu         | TabHome                   | 隐藏与取消隐藏                    |
| SheetDelete                     | button       | TabHome                   | 删除工作表                        |
|                                 | group        | TabHome                   | 窗格                              |
|                                 | group        | TabHome                   | 工具箱                            |
|                                 | group        | TabHome                   | 其他                              |
| ConditionalFormattingMenu       | menu         | TabHome                   | 条件格式                          |
| InputNumericSequence            | button       | TabHome                   | 录入123序列                       |
| InputLetterSequence             | button       | TabHome                   | 录入ABC序列                       |
| InputRomanNumericSequence       | button       | TabHome                   | 录入罗马数字序列                  |
| FillBlankCells                  | button       | TabHome                   | 空白单元格填充                    |
| TextBoxInsertHorizontal         | toggleButton | TabInsert                 | 横向文本框                        |
| TextBoxInsertVertical           | toggleButton | TabInsert                 | 竖向文本框                        |
| TextBoxInsertMenu               | splitButton  | TabInsert                 | 文本框                            |
| PivotTableReport                | button       | TabInsert                 | 数据透视表                        |
| TableInsertExcel                | button       | None                      | 表格                              |
| PivotTableInsert                | button       | TabInsert                 | 数据透视表                        |
| HeaderFooterInsert              | button       | TabInsert                 | 页眉页脚                          |
| Camera                          | button       | TabInsert                 | 照相机                            |
| OleObjectctInsert               | button       | TabInsert                 | 对象                              |
| EquationInsertGallery           |              | TabInsert                 | 公式                              |
| EquationInsertNew               |              | TabInsert                 | 插入新公式                        |
| ChartTypeAllInsertDialog        | button       | TabInsert                 | 全部图表                          |
| OnlineDocerCommand_Chart        |              | None                      |                                   |
| PictureInsertFromFile           |              | TabInsert                 | 图片                              |
| PictureInsertFromFile           |              | TabInsert                 | 本地图片                          |
| ShapesInsertGallery             | gallery      | TabInsert                 | 形状                              |
| HyperlinkInsert                 | button       | TabInsert                 | 超链接                            |
| ActiveXLabel                    | button       | TabInsert                 | 标签                              |
| FormControlGroupBox             | button       | TabInsert                 | 分组框                            |
| ActiveXButton                   | button       | TabInsert                 | 按钮                              |
| FormControlCheckBox             | button       | TabInsert                 | 复选框                            |
| ActiveXListBox                  | button       | TabInsert                 | 列表框                            |
| ActiveXComboBox                 | button       | TabInsert                 | 组合框                            |
| ActiveXScrollBar                | button       | TabInsert                 | 滚动条                            |
| ActiveXSpinButton               | button       | TabInsert                 | 微调项                            |
| CodeEdit                        | button       | TabInsert                 | 编辑代码                          |
| ObjectsGroup                    | button       | TabInsert                 | 组合                              |
| ObjectsUngroup                  | button       | TabInsert                 | 取消组合                          |
| ObjectsGroupMenu                | menu         | TabInsert                 | 组合                              |
| ObjectRotateFree                | button       | TabInsert                 | 自由旋转                          |
| ObjectRotateLeft90              | button       | TabInsert                 | 向左旋转 90°                      |
| ObjectRotateRight90             | button       | TabInsert                 | 向右旋转 90°                      |
| ObjectFlipHorizontal            | button       | TabInsert                 | 水平翻转                          |
| ObjectFlipVertical              | button       | TabInsert                 | 垂直翻转                          |
| GroupArrange                    | group        | TabInsert                 | 排列                              |
| GroupArrange                    | group        | TabInsert                 | 排列                              |
| SheetBackground                 | button       | TabLayout                 | 背景图片                          |
| PageOrientationGallery          | gallery      | TabLayout                 | 纸张方向                          |
| PageSizeGallery                 | gallery      | TabLayout                 | 纸张大小                          |
| PrintAreaMenu                   | menu         | TabLayout                 | 打印区域                          |
| PrintTitles                     | button       | TabLayout                 | 打印标题                          |
| HeaderFooterInsert              | button       | TabLayout                 | 页眉页脚                          |
| PageBreakInsertExcel            | button       | TabLayout                 | 插入分页符                        |
| PageBreakRemove                 | button       | TabLayout                 | 删除分页符                        |
| PageBreaksResetAll              | button       | TabLayout                 | 重置所有分页符                    |
| PageBreakMenu                   | button       | TabLayout                 | 插入分页符                        |
| ObjectBringForward              | button       | TabLayout                 | 上移一层                          |
| ObjectBringToFront              | button       | TabLayout                 | 置于顶层                          |
| ObjectSendBackward              | button       | TabLayout                 | 下移一层                          |
| ObjectSendToBack                | button       | TabLayout                 | 置于底层                          |
| SelectionPane                   | toggleButton | TabLayout                 | 选择窗格                          |
| ObjectsAlignLeft                | button       | TabLayout                 | 左对齐                            |
| ObjectsAlignCenterHorizontal    | button       | TabLayout                 | 水平居中                          |
| ObjectsAlignRight               | button       | TabLayout                 | 右对齐                            |
| ObjectsAlignTop                 | button       | TabLayout                 | 顶端对齐                          |
| ObjectsAlignMiddleVertical      | button       | TabLayout                 | 垂直居中                          |
| ObjectsAlignBottom              | button       | TabLayout                 | 底端对齐                          |
| AlignDistributeHorizontally     | button       | TabLayout                 | 横向分布                          |
| AlignDistributeVertically       | button       | TabLayout                 | 纵向分布                          |
| SnapToGrid                      | toggleButton | TabLayout                 | 对齐网格                          |
| ObjectAlignMenu                 | menu         | TabLayout                 | 对齐                              |
| ObjectsGroup                    | button       | TabLayout                 | 组合                              |
| ObjectsUngroup                  | button       | TabLayout                 | 取消组合                          |
| ObjectRotateFree                | button       | TabLayout                 | 自由旋转                          |
| ObjectRotateLeft90              | button       | TabLayout                 | 向左旋转 90°                      |
| ObjectRotateRight90             | button       | TabLayout                 | 向右旋转 90°                      |
| ObjectFlipHorizontal            | button       | TabLayout                 | 水平翻转                          |
| ObjectFlipVertical              | button       | TabLayout                 | 垂直翻转                          |
| PageSetupDialog                 | button       | None                      | 页面设置                          |
| ContrastMore                    | button       | TabPictureTool            | 增加对比度                        |
| ContrastLess                    | button       | TabPictureTool            | 降低对比度                        |
| BrightnessMore                  | button       | TabPictureTool            | 增加亮度                          |
| BrightnessLess                  | button       | TabPictureTool            | 降低亮度                          |
| PictureSetTransparentColor      | button       | TabPictureTool            | 设置透明色                        |
| ObjectPictureFill               | button       | TabPictureTool            | 图片                              |
| GradientGallery                 | gallery      | None                      | 渐变                              |
| MoreTextureOptions              | button       | None                      | 图片或纹理                        |
| OutlineLinePatternFill          | button       | TabPictureTool            | 带图案线条                        |
| LineStylesDialog                | button       | TabPictureTool            | 其他线条                          |
| ObjectsAlignLeft                | button       | TabPictureTool            | 左对齐                            |
| ObjectsAlignCenterHorizontal    | button       | TabPictureTool            | 水平居中                          |
| ObjectsAlignRight               | button       | TabPictureTool            | 右对齐                            |
| ObjectsAlignTop                 | button       | TabPictureTool            | 顶端对齐                          |
| ObjectsAlignMiddleVertical      | button       | TabPictureTool            | 垂直居中                          |
| ObjectsAlignBottom              | button       | TabPictureTool            | 底端对齐                          |
| AlignDistributeHorizontally     | button       | TabPictureTool            | 横向分布                          |
| AlignDistributeVertically       | button       | TabPictureTool            | 纵向分布                          |
| SnapToGrid                      | toggleButton | TabPictureTool            | 对齐网格                          |
| ObjectsGroup                    | button       | TabPictureTool            | 组合                              |
| ObjectsUngroup                  | button       | TabPictureTool            | 取消组合                          |
| ObjectRotateFree                | button       | TabPictureTool            | 自由旋转                          |
| ObjectRotateLeft90              | button       | TabPictureTool            | 向左旋转 90°                      |
| ObjectRotateRight90             | button       | TabPictureTool            | 向右旋转 90°                      |
| ObjectFlipHorizontal            | button       | TabPictureTool            | 水平翻转                          |
| ObjectFlipVertical              | button       | TabPictureTool            | 垂直翻转                          |
| SelectionPane                   | toggleButton | TabPictureTool            | 选择窗格                          |
|                                 | editBox      | TabPictureTool            |                                   |
|                                 | editBox      | TabPictureTool            |                                   |
| PictureCrop                     | toggleButton | TabPictureTool            | 裁剪                              |
|                                 | button       | None                      | 重设形状和大小                    |
| ContrastMore                    | button       | TabPictureTool_Vml        | 增加对比度                        |
| ContrastLess                    | button       | TabPictureTool_Vml        | 降低对比度                        |
| BrightnessMore                  | button       | TabPictureTool_Vml        | 增加亮度                          |
| BrightnessLess                  | button       | TabPictureTool_Vml        | 降低亮度                          |
| PictureSetTransparentColor      | button       | TabPictureTool_Vml        | 设置透明色                        |
| ObjectPictureFill               | button       | TabPictureTool_Vml        | 图片                              |
| GradientGallery                 | gallery      | None                      | 渐变                              |
| MoreTextureOptions              | button       | None                      | 图片或纹理                        |
| LineStylesDialog                | button       | TabPictureTool_Vml        | 其他线条                          |
| ObjectsAlignLeft                | button       | TabPictureTool_Vml        | 左对齐                            |
| ObjectsAlignCenterHorizontal    | button       | TabPictureTool_Vml        | 水平居中                          |
| ObjectsAlignRight               | button       | TabPictureTool_Vml        | 右对齐                            |
| ObjectsAlignTop                 | button       | TabPictureTool_Vml        | 顶端对齐                          |
| ObjectsAlignMiddleVertical      | button       | TabPictureTool_Vml        | 垂直居中                          |
| ObjectsAlignBottom              | button       | TabPictureTool_Vml        | 底端对齐                          |
| SnapToGrid                      | toggleButton | TabPictureTool_Vml        | 对齐网格                          |
| ObjectsGroup                    | button       | TabPictureTool_Vml        | 组合                              |
| ObjectsUngroup                  | button       | TabPictureTool_Vml        | 取消组合                          |
| ObjectRotateFree                | button       | TabPictureTool_Vml        | 自由旋转                          |
| ObjectRotateLeft90              | button       | TabPictureTool_Vml        | 向左旋转 90°                      |
| ObjectRotateRight90             | button       | TabPictureTool_Vml        | 向右旋转 90°                      |
| ObjectFlipHorizontal            | button       | TabPictureTool_Vml        | 水平翻转                          |
| ObjectFlipVertical              | button       | TabPictureTool_Vml        | 垂直翻转                          |
| PictureCrop                     | toggleButton | TabPictureTool_Vml        | 裁剪                              |
|                                 | editBox      | TabPictureTool_Vml        | 高度:                             |
|                                 | editBox      | TabPictureTool_Vml        | 宽度:                             |
| Refresh                         | button       | None                      | 刷新数据                          |
| RefreshAll                      | button       | None                      | 全部刷新                          |
| FilePrint                       | button       | None                      | 打印                              |
| FilePrintQuick                  | button       | None                      | 直接打印                          |
| FilePrintEntireWorkbookQuick    | button       | TabPrintPreview           | 打印整个工作簿                    |
|                                 | comboBox     | TabPrintPreview           | 打印机                            |
|                                 | comboBox     | TabPrintPreview           | 纸张类型                          |
|                                 | comboBox     | TabPrintPreview           | 方式                              |
|                                 | comboBox     | TabPrintPreview           | 顺序                              |
|                                 | comboBox     | TabPrintPreview           | 缩放比例                          |
| ViewPageLayoutView              | toggleButton | None                      | 页面视图                          |
|                                 | button       | None                      | 同义词库                          |
| TranslateToSimplifiedChinese    | button       | TabReview                 | 繁转简                            |
| TranslateToTraditionalChinese   | button       | TabReview                 | 简转繁                            |
| ReviewNewComment                | button       | TabReview                 | 新建批注                          |
| ReviewDeleteComment             | button       | TabReview                 | 删除批注                          |
| ReviewPreviousComment           | button       | TabReview                 | 上一条                            |
| ReviewNextComment               | button       | TabReview                 | 下一条                            |
| ReviewShowOrHideComment         | toggleButton | TabReview                 | 显示/隐藏批注                     |
| ReviewShowAllComments           | toggleButton | TabReview                 | 显示所有批注                      |
|                                 | toggleButton | TabReview                 | 重置当前批注                      |
|                                 | toggleButton | TabReview                 | 重置所有批注                      |
| SheetProtect                    | button       | TabReview                 | 保护工作表                        |
| ReviewRestrictEditing           | toggleButton | TabReview                 | 保护工作簿                        |
| ReviewShareWorkbook             | button       | TabReview                 | 共享工作簿                        |
| ReviewProtectAndShareWorkbook   | button       | TabReview                 | 保护并共享工作簿                  |
| ReviewAllowUsersToEditRanges    | button       | TabReview                 | 允许用户编辑区域                  |
| ReviewHighlightChanges          | button       | TabReview                 | 突出显示修订                      |
| ReviewTrackChangesMenu          | menu         | TabReview                 | 修订                              |
| ShadowSemitransparentClassic    | button       | TabshadowDrawingTools     | 半透明阴影                        |
| ShadowOnOrOffClassic            | button       | TabshadowDrawingTools     | 设置/取消阴影                     |
| \_3DTiltUpClassic               | button       | TabshadowDrawingTools     | 上翘                              |
| \_3DEffectsOnOffClassic         | button       | TabshadowDrawingTools     | 设置三维效果                      |
| \_3DExtrusionPerspectiveClassic | button       | TabshadowDrawingTools     | 透视                              |
| \_3DExtrusionParallelClassic    | button       | TabshadowDrawingTools     | 平行                              |
| \_3DLightingFlatClassic         | button       | TabshadowDrawingTools     | 明亮                              |
| \_3DLightingNormalClassic       | button       | TabshadowDrawingTools     | 普通                              |
| \_3DLightingDimClassic          | button       | TabshadowDrawingTools     | 阴暗                              |
| \_3DExtrusionPerspectiveClassic | button       | TabshadowDrawingTools_Vml | 透视                              |
| ObjectBringForward              | button       | TabSlicerOptions          | 上移一层                          |
| ObjectBringToFront              | button       | TabSlicerOptions          | 置于顶层                          |
| ObjectSendBackward              | button       | TabSlicerOptions          | 下移一层                          |
| ObjectSendToBack                | button       | TabSlicerOptions          | 置于底层                          |
| ObjectsAlignLeft                | button       | TabSlicerOptions          | 左对齐                            |
| ObjectsAlignCenterHorizontal    | button       | TabSlicerOptions          | 水平居中                          |
| ObjectsAlignRight               | button       | TabSlicerOptions          | 右对齐                            |
| ObjectsAlignTop                 | button       | TabSlicerOptions          | 顶端对齐                          |
| ObjectsAlignMiddleVertical      | button       | TabSlicerOptions          | 垂直居中                          |
| ObjectsAlignBottom              | button       | TabSlicerOptions          | 底端对齐                          |
| AlignDistributeHorizontally     | button       | TabSlicerOptions          | 横向分布                          |
| AlignDistributeVertically       | button       | TabSlicerOptions          | 纵向分布                          |
| SnapToGrid                      | toggleButton | TabSlicerOptions          | 对齐网格                          |
| ObjectsGroup                    | button       | TabSlicerOptions          | 组合                              |
| ObjectsUngroup                  | button       | TabSlicerOptions          | 取消组合                          |
| ObjectRotateLeft90              | button       | TabSlicerOptions          | 向左旋转 90°                      |
| ObjectRotateRight90             | button       | TabSlicerOptions          | 向右旋转 90°                      |
| ObjectFlipHorizontal            | button       | TabSlicerOptions          | 水平翻转                          |
| ObjectFlipVertical              | button       | TabSlicerOptions          | 垂直翻转                          |
| SelectionPane                   | toggleButton | TabSlicerOptions          | 选择窗格                          |
| TextBoxInsertMenu               | splitButton  | TabTextTool               | 文本框                            |
| TextBoxInsertHorizontal         | toggleButton | None                      | 横向文本框                        |
| TextBoxInsertVertical           | toggleButton | None                      | 垂直平铺                          |
|                                 | button       | TabTextTool               | 清除艺术字                        |
| ShapeStylesGallery              | gallery      | TabTextTool               | 形状样式                          |
| ShapeFillColorPicker            | gallery      | TabTextTool               | 形状填充                          |
| GradientGallery                 | gallery      | None                      | 渐变                              |
| MoreTextureOptions              | button       | None                      | 图片或纹理                        |
| ShapeOutlineColorPicker         | gallery      | TabTextTool               | 形状轮廓                          |
|                                 | comboBox     | TabTextTool               | 字体                              |
|                                 | comboBox     | TabTextTool               | 字号                              |
| ClearAll                        | button       | TabTextTool               | 全部                              |
| ClearFormats                    | button       | TabTextTool               | 效果设置                          |
| ClearContents                   | button       | TabTextTool               | 内容                              |
| ClearComments                   | button       | TabTextTool               | 批注                              |
| AsianLayoutPhoneticGuide        | button       | TabTextTool               | 拼音指南                          |
| CharacterBorder                 | toggleButton | TabTextTool               | 字符边框                          |
|                                 | button       | TabTextTool               | 清除艺术字                        |
| FormatPainter                   | toggleButton | None                      | 格式刷                            |
| ViewNormalViewExcel             | toggleButton | TabView                   | 普通                              |
| ViewPageBreakPreviewView        | toggleButton | TabView                   | 分页预览                          |
| ViewPageLayoutView              | toggleButton | TabView                   | 页面布局                          |
| ViewCustomViews                 | button       | TabView                   | 自定义视图                        |
| ViewFullScreenView              | toggleButton | TabView                   | 全屏显示                          |
| ViewFullScreenView              | toggleButton | TabView                   | 全屏显示                          |
| ViewFormulaBar                  | checkBox     | TabView                   | 编辑栏                            |
| GridlinesExcel                  | checkBox     | TabView                   | 显示网格线                        |
| ViewHeadings                    | checkBox     | TabView                   | 显示行号列标                      |
| ZoomDialog                      | button       | TabView                   | 显示比例                          |
| WindowNew                       | button       | TabView                   | 新建窗口                          |
| WindowHide                      | button       | TabView                   | 隐藏                              |
| WindowUnhide                    | button       | TabView                   | 取消隐藏                          |
| WindowSplitToggle               | toggleButton | TabView                   | 拆分窗口                          |
| WindowSwitchWindowsMenuExcel    | button       | TabView                   | 切换窗口                          |
|                                 | button       | TabView                   | 取消冻结窗格                      |
|                                 | button       | TabView                   | 冻结窗格                          |
|                                 | button       | TabView                   | 冻结行                            |
|                                 | button       | TabView                   | 冻结列                            |
|                                 | button       | TabView                   | 冻结首行                          |
|                                 | button       | TabView                   | 冻结首列                          |
| MacroPlay                       | button       | TabView                   | JS 宏                             |
|                                 | splitButton  | TabView                   | 宏安全                            |
| VisualBasic                     | button       | TabView                   | WPS 宏编辑器                      |
| MenuMacros                      | splitButton  | TabView                   | JS 宏                             |
| ShadowSemitransparentClassic    | button       | TabWAshadowDrawingTools   | 半透明阴影                        |
| ShadowOnOrOffClassic            | button       | TabWAshadowDrawingTools   | 设置/取消阴影                     |
| \_3DTiltUpClassic               | button       | TabWAshadowDrawingTools   | 上翘                              |
| \_3DEffectsOnOffClassic         | button       | TabWAshadowDrawingTools   | 设置三维效果                      |
| \_3DExtrusionPerspectiveClassic | button       | TabWAshadowDrawingTools   | 透视                              |
| \_3DExtrusionParallelClassic    | button       | TabWAshadowDrawingTools   | 平行                              |
| \_3DLightingFlatClassic         | button       | TabWAshadowDrawingTools   | 明亮                              |
| \_3DLightingNormalClassic       | button       | TabWAshadowDrawingTools   | 普通                              |
| \_3DLightingDimClassic          | button       | TabWAshadowDrawingTools   | 阴暗                              |
| WordArtInsertDialogClassic      | button       | TabWordArt                | 艺术字库                          |
| WordArtSpacingVeryTight         | button       | TabWordArt                | 很紧                              |
| WordArtSpacingTight             | toggleButton | TabWordArt                | 紧密                              |
| WordArtSpacingNormal            | toggleButton | TabWordArt                | 常规                              |
| WordArtSpacingLoose             | toggleButton | TabWordArt                | 稀疏                              |
| WordArtSpacingVeryLoose         | toggleButton | TabWordArt                | 很松                              |
| WordArtVerticalText             | button       | TabWordArt                | 竖排                              |
| TextAlignLeft                   | button       | TabWordArt                | 左对齐                            |
| TextAlignCenter                 | button       | TabWordArt                | 居中                              |
| TextAlignRight                  | button       | TabWordArt                | 右对齐                            |
| TextAlignWordJustify            | button       | TabWordArt                | 单词调整                          |
| TextAlignLetterJustify          | button       | TabWordArt                | 字母调整                          |
| TextAlignStretchJustify         | button       | TabWordArt                | 延伸调整                          |
| ObjectPictureFill               | button       | TabWordArt                | 图片                              |
| GlowColorMoreColorsDialog       | button       | TabWordArt                | 其他填充颜色                      |
| GradientGallery                 | gallery      | None                      | 渐变                              |
| MoreTextureOptions              | button       | None                      | 图片或纹理                        |
| OutlineLinePatternFill          | button       | TabWordArt                | 带图案线条                        |
| LineStylesDialog                | button       | TabWordArt                | 其他线条                          |
| SelectionPane                   | toggleButton | TabWordArt                | 选择窗格                          |
| ObjectsAlignLeft                | button       | TabWordArt                | 左对齐                            |
| ObjectsAlignCenterHorizontal    | button       | TabWordArt                | 水平居中                          |
| ObjectsAlignRight               | button       | TabWordArt                | 右对齐                            |
| ObjectsAlignTop                 | button       | TabWordArt                | 顶端对齐                          |
| ObjectsAlignMiddleVertical      | button       | TabWordArt                | 垂直居中                          |
| ObjectsAlignBottom              | button       | TabWordArt                | 底端对齐                          |
| AlignDistributeHorizontally     | button       | TabWordArt                | 横向分布                          |
| AlignDistributeVertically       | button       | TabWordArt                | 纵向分布                          |
| SnapToGrid                      | toggleButton | TabWordArt                | 对齐网格                          |
| ObjectsGroup                    | button       | TabWordArt                | 组合                              |
| ObjectsUngroup                  | button       | TabWordArt                | 取消组合                          |
| ObjectRotateFree                | button       | TabWordArt                | 自由旋转                          |
| ObjectRotateLeft90              | button       | TabWordArt                | 向左旋转 90°                      |
| ObjectRotateRight90             | button       | TabWordArt                | 向右旋转 90°                      |
| ObjectFlipHorizontal            | button       | TabWordArt                | 水平翻转                          |
| ObjectFlipVertical              | button       | TabWordArt                | 垂直翻转                          |
| FileSaveAsPicture               |              | TabWorkSpace              | 输出为图片                        |
| ExportToPDF                     | button       | None                      | 输出为PDF                         |
| SetLanguage                     | button       | None                      | 更改语言                          |
| SwitchFace                      |              | None                      | 皮肤                              |
| TabbarQAndA                     |              | None                      |                                   |
| FileMenu                        |              | None                      |                                   |
| TabHome                         | tab          | None                      | 开始                              |
| TabInsert                       | tab          | None                      | 插入                              |
| TabPageLayoutExcel              | tab          | None                      | 页面视图                          |
| TabFormulas                     | tab          | None                      | 公式                              |
| TabData                         | tab          | None                      | 数据                              |
| TabReview                       | tab          | None                      | 审阅                              |
| TabView                         | tab          | None                      | 视图                              |
| TabPrintPreview                 | tab          | None                      | 打印预览                          |
| TabAddIns                       | tab          | None                      | 加载宏                            |
| TabSecurity                     | tab          | None                      | 安全性                            |
| TabDeveloper                    | tab          | None                      | 开发工具                          |
| TabWorkSpace                    | tab          | None                      | 云服务                            |
| TabPictureToolsFormat           | tab          | None                      | 图片工具                          |
| TabEtToolbox                    | tab          | None                      | 智能工具箱                        |
| ExportToPDF                     | button       | None                      | 输出为PDF                         |
| QAT_Menu                        |              | None                      |                                   |
| FileNewBlankDocument            | button       | None                      | 新建                              |
| FileOpen                        | button       | None                      | 打开                              |
| FileSave                        | button       | None                      |                                   |
| FilePrint                       | button       | None                      | 打印                              |
| FilePrintPreview                | button       | None                      | 打印                              |
| FilePrintQuick                  | button       | None                      | 直接打印                          |
| FilePrintPreview                | button       | None                      | 打印预览                          |
| Undo                            | button       | None                      | 撤消                              |
| Redo                            | button       | None                      | 恢复                              |
| MergeCells                      | button       | None                      | 合并单元格                        |
| MergeCenterMenu                 | splitButton  | None                      | 合并单元格                        |
| BorderNone                      | button       | None                      | 无边框                            |
| BordersAll                      | button       | None                      | 所有框线                          |
| BorderOutside                   | button       | None                      | 外侧框线                          |
| BorderBottom                    | button       | None                      | 下框线                            |
| BorderTop                       | button       | None                      | 上框线                            |
| BorderLeft                      | button       | None                      | 左框线                            |
| BorderRight                     | button       | None                      | 右框线                            |
| BorderDoubleBottom              | button       | None                      | 双底框线                          |
| BorderThickBottom               | button       | None                      | 粗底框线                          |
| BorderTopAndBottom              | button       | None                      | 上下框线                          |
| BorderTopAndThickBottom         | button       | None                      | 上框线和粗下框线                  |
| BorderTopAndDoubleBottom        | button       | None                      | 上框线和双下框线                  |
| BordersMoreDialog               | button       | None                      | 其他边框                          |
| BordersGallery                  | splitButton  | None                      |                                   |
| AlignCenter                     | toggleButton | None                      | 水平居中                          |
| AutoSum                         | button       | None                      | 求和                              |
| AutoSumAverage                  | button       | None                      | 平均值                            |
| AutoSumCount                    | button       | None                      | 计数                              |
| AutoSumMax                      | button       | None                      | 最大值                            |
| AutoSumMin                      | button       | None                      | 最小值                            |
| AutoSumMoreFunctions            | button       | None                      | 其他函数                          |
| ShadowOnOrOffClassic            | button       | None                      | 设置阴影                          |
| ShadowSemitransparentClassic    | button       | None                      | 半透明阴影                        |
| FormatPainter                   | toggleButton | None                      | 格式刷                            |
| AlignLeft                       | toggleButton | None                      | 左对齐                            |
| AlignMiddleExcel                | toggleButton | None                      | 垂直居中                          |
| InsertDiagonalHeader            | button       | None                      | 插入斜线表头                      |
| wps_SplitSheet                  |              | None                      | 拆分表格                          |
| wps_MergeSheet                  |              | None                      | 合并表格                          |
| FileMenuSendMail                |              | None                      | 发送邮件                          |
| OutlineGroup                    |              | None                      | 组合                              |
| OutlineUngroup                  |              | None                      | 取消组合                          |
| Refresh                         | button       | None                      | 刷新数据                          |
| RefreshAll                      | button       | None                      | 全部刷新                          |
| GoTo                            | button       | None                      | 定位                              |
| ReviewNewComment                | button       | None                      | 新建批注                          |
| ReviewNewComment                | button       | None                      |                                   |
| ReviewShowAllComments           | toggleButton | None                      | 显示所有批注                      |
| ContextMenuCell                 |              | None                      | 单元格                            |
| VisualBasic                     | button       | None                      | WPS 宏编辑器                      |
| MacroPlay                       | button       | None                      | JS 宏                             |
| HeaderFooterCurrentTimeInsert   | button       | None                      |                                   |
| FormatCellsDialog               | button       | None                      | 设置单元格格式                    |
| HeaderFooterCurrentDate         | button       | None                      |                                   |
| Bold                            | button       | None                      | 加粗                              |
| Copy                            | button       | None                      | 复制                              |
| TableInsertExcel                | button       | None                      | 表格                              |
| FillDown                        | button       | None                      | 向下填充                          |
| FindDialogExcel                 | button       | None                      | 查找                              |
| GoTo                            | button       | None                      | 定位                              |
| ReplaceDialog                   | button       | None                      | 替换                              |
| Italic                          | toggleButton | None                      | 倾斜                              |
| FileNewBlankDocument            | button       | None                      | 新建                              |
| FileNewBlankDocument            | button       | None                      | 新建                              |
| FileNewBlankDocument            | button       | None                      | 新建                              |
| FileNewBlankDocument            | button       | None                      | 新建                              |
| FileNewBlankDocument            | button       | None                      | 新建                              |
| FileNewBlankDocument            | button       | None                      | 新建                              |
| FileClose                       | button       | None                      | 关闭                              |
| PageSetupDialog                 | button       | None                      | 页面设置                          |
| Properties                      | button       | None                      | 属性                              |
| PasteSpecialDialog              | button       | None                      | 选择性粘贴                        |
| ClearContents                   | button       | None                      | 内容                              |
| FillRight                       | button       | None                      | 右对齐                            |
| FileSave                        | button       | None                      |                                   |
| FontSizeDecrease                | button       | None                      | 减小字号                          |
| FontSizeIncrease                | button       | None                      | 增大字号                          |
| CellsInsertDialog               | button       | None                      | 单元格                            |
| Underline                       | toggleButton | None                      | 下划线                            |
| Paste                           | button       | None                      | 粘贴                              |
| Cut                             | button       | None                      | 剪切                              |
| PasteValues                     | button       | None                      | 粘贴为数值                        |
| FileSaveAs                      | button       | None                      | 另存为                            |
| CalculateSheet                  | button       | None                      | 计算工作表                        |
| ApplicationExit                 | button       | None                      | 关闭                              |
| FileNewXlsxFile                 | button       | None                      | 新建 Excel 文件                   |
| FileNew                         | button       | None                      | 新建                              |
| FileNewFromTemplate             | button       | None                      | 本机上的模板                      |
| TextBoxInsertHorizontal         | toggleButton | None                      | 横向文本框                        |
| TextBoxInsertVertical           | toggleButton | None                      | 垂直平铺                          |
|                                 | splitButton  | None                      | 智能数据                          |
| ViewNormalViewExcel             | toggleButton | None                      | 普通视图                          |
| ViewPageLayoutView              | toggleButton | None                      | 页面视图                          |
| ViewPageBreakPreviewView        | toggleButton | None                      | 分页预览                          |
| OnlineDocerCommand_Chart        |              | None                      |                                   |
| GradientGallery                 | gallery      | None                      | 渐变                              |
| MoreTextureOptions              | button       | None                      | 图片或纹理                        |
|                                 | button       | None                      | 同义词库                          |
| FileSaveAsMenu                  | menu         | None                      | 另存为                            |
| FileSaveAsExcelEt               | button       | None                      | WPS 表格 文件（\*.et）            |
| FileSaveAsExcelEtt              | button       | None                      | WPS 表格 模板文件（\*.ett）       |
| FileSaveAsExcelEtx              | button       | None                      | WPS 表格 2007/2010 文件（\*.etx） |
| FileSaveAsExcelXls              | button       | None                      | Excel 97-2003 文件（\*.xls）      |
| FileSaveAsExcelXlt              | button       | None                      | Excel 97-2003 模板文件（\*.xlt）  |
| FileSaveAsExcelXlsx             | button       | None                      | Excel 文件（\*.xlsx）             |
| FileSaveAsOtherFormats          | button       | None                      | 其他格式                          |
| FileSaveAsPdfOrXps              | button       | None                      | 输出为PDF                         |
| FileSlimming                    |              | None                      | 文件瘦身                          |
|                                 | button       | None                      | 新建批注                          |
| ShowFormulas                    | button       | None                      | 显示公式                          |
| DisplayOutline                  | button       | None                      |                                   |
|                                 | button       | None                      | 取消冻结窗格                      |
|                                 | button       | None                      | 冻结窗格                          |
|                                 | button       | None                      | 冻结行                            |
|                                 | button       | None                      | 冻结列                            |
|                                 | button       | None                      | 冻结首行                          |
|                                 | button       | None                      | 冻结首列                          |
| DocumentSplit                   | button       | None                      | 拆分（取消拆分）                  |
| DocumentMerge                   | button       | None                      | 合并                              |
| DocumentSplitAndMerge           | menu         | None                      | 文档拆分合并                      |
|                                 | button       | None                      | 重设形状和大小                    |
| SaveAsPicture                   |              | None (Context Menu)       | 输出为图片                        |
| FileCloseAll                    | button       | None (Context Menu)       | 关闭所有文档                      |
| FileCloseOrCloseAll             | button       | None (Context Menu)       | 关闭所有文档                      |
| FileCloseOrCloseAll             | button       | None (Context Menu)       | 关闭所有文档                      |
| SaveAll                         | button       | None (Context Menu)       | 保存所有文档                      |
| SaveAll                         | button       | None (Context Menu)       | 保存所有文档                      |
| SaveAll                         | button       | None (Context Menu)       | 保存所有文档                      |
| FileExportToPDF                 | button       | None (Context Menu)       | 输出为PDF格式                     |
| exportToOFD                     | button       | None (Context Menu)       | 输出为OFD格式                     |
| FileMenuSendMail                | button       | None (Context Menu)       | 发送邮件                          |
| FileEncrypt                     | button       | None (Context Menu)       | 文件加密                          |
| PasteAsPicture                  | button       | None (Context Menu)       | 粘贴为图片                        |
| FillUp                          | button       | None (Context Menu)       | 向上填充                          |
| FillLeft                        | button       | None (Context Menu)       | 向左填充                          |
| FillSeries                      | button       | None (Context Menu)       | 序列                              |
| ClearAll                        | button       | None (Context Menu)       | 全部                              |
| ClearFormats                    | button       | None (Context Menu)       | 格式                              |
| ClearComments                   | button       | None (Context Menu)       | 批注                              |
| SheetDelete                     | button       | None (Context Menu)       | 删除工作表                        |
| EditLinks                       | button       | None (Context Menu)       | 编辑链接                          |
| ViewFormulaBar                  | checkBox     | None (Context Menu)       | 编辑栏                            |
| ViewFullScreenView              | toggleButton | None (Context Menu)       | 全屏显示                          |
| ZoomDialog                      | button       | None (Context Menu)       | 显示比例                          |
| InsertCellstMenu                | splitButton  | None (Context Menu)       | 单元格                            |
| SheetInsert                     | button       | None (Context Menu)       | 工作表                            |
| NameDefine                      | button       | None (Context Menu)       | 定义名称                          |
| NewName                         | button       | None (Context Menu)       | 定义名称                          |
| NameCreateFromSelection         | button       | None (Context Menu)       | 指定                              |
| ClipArtInsert                   | toggleButton | None (Context Menu)       | 剪贴画                            |
| EquationInsertGallery           |              | None (Context Menu)       | 公式                              |
| ColumnWidthAutoFit              | button       | None (Context Menu)       | 最适合的列宽                      |
| RowHeightAutoFit                | button       | None (Context Menu)       | 最适合的行高                      |
| RowHeightAutoFit                | button       | None (Context Menu)       | 最适合的行高                      |
| RowsHide                        |              | None (Context Menu)       | 隐藏                              |
| RowsUnhide                      | button       | None (Context Menu)       | 取消隐藏                          |
| TableColumnWidth                | control      | None (Context Menu)       | 列宽                              |
| ColumnsHide                     | button       | None (Context Menu)       | 隐藏                              |
| ColumnsUnhide                   | button       | None (Context Menu)       | 取消隐藏                          |
| SheetHide                       | button       | None (Context Menu)       | 隐藏工作表                        |
| SheetUnhide                     | button       | None (Context Menu)       | 取消隐藏工作表                    |
| SheetBackground                 | button       | None (Context Menu)       | 背景图片                          |
| ReviewShareWorkbook             | button       | None (Context Menu)       | 共享工作簿                        |
| ReviewHighlightChanges          | button       | None (Context Menu)       | 突出显示修订                      |
| SheetProtect                    | button       | None (Context Menu)       | 保护工作表                        |
| ReviewAllowUsersToEditRanges    | button       | None (Context Menu)       | 允许用户编辑区域                  |
| ReviewRestrictEditing           | toggleButton | None (Context Menu)       | 保护工作簿                        |
| ReviewProtectAndShareWorkbook   | button       | None (Context Menu)       | 保护并共享工作簿                  |
| GoalSeek                        | button       | None (Context Menu)       | 单变量求解                        |
| FormulaEvaluate                 | button       | None (Context Menu)       | 公式求值                          |
| ErrorCheckingMenu               | splitButton  | None (Context Menu)       | 错误检查                          |
| CircularReferences              | gallery      | None (Context Menu)       | 循环引用                          |
| Security                        | button       | None (Context Menu)       | 安全性                            |
| ComAddInsDialog                 | button       | None (Context Menu)       | COM 加载项                        |
| FileBackupManagement            |              | None (Context Menu)       | 备份管理                          |
| ApplicationOptionsDialog        |              | None (Context Menu)       | 选项                              |
| SetLanguage                     | button       | None (Context Menu)       | 更改语言                          |
| SortDialogClassic               | button       | None (Context Menu)       | 排序                              |
| AdvancedFilterDialog            | toggleButton | None (Context Menu)       | 高级筛选                          |
| Subtotals                       | button       | None (Context Menu)       | 分类汇总                          |
| DataValidation                  | button       | None (Context Menu)       | 有效性                            |
| Consolidate                     | button       | None (Context Menu)       | 合并计算                          |
| OutlineHideDetail               | button       | None (Context Menu)       | 隐藏明细数据                      |
| OutlineShowDetail               | button       | None (Context Menu)       | 显示明细数据                      |
| WindowSplitToggle               | toggleButton | None (Context Menu)       | 拆分                              |
| FreezePanes                     | button       | None (Context Menu)       | 冻结窗格                          |
| BlogHomePage                    | button       | None (Context Menu)       | WPS官方网站                       |
| About                           | button       | None (Context Menu)       | 关于 WPS 表格                     |
| ReviewDeleteComment             | button       | None (Context Menu)       | 删除批注                          |
| ReviewShowOrHideComment         | button       | None (Context Menu)       | 显示/隐藏批注                     |
| ReviewShowOrHideComment         | button       | None (Context Menu)       | 隐藏批注                          |
| RowHeight                       | button       | None (Context Menu)       | 行高                              |
| FillSeries                      | button       | None (Context Menu)       | 序列                              |
| FillSeries                      | button       | None (Context Menu)       | 序列                              |
| CalculationOptionsAutomatically | toggleButton | None (Context Menu)       | 无                                |
|                                 | button       | None (Context Menu)       | 组合                              |
|                                 | button       | None (Context Menu)       | 取消组合                          |
|                                 | button       | None (Context Menu)       | 取消组合                          |
| ObjectBringToFront              | button       | None (Context Menu)       | 置于顶层                          |
| ObjectSendToBack                | button       | None (Context Menu)       | 置于底层                          |
| ObjectBringForward              | button       | None (Context Menu)       | 上移一层                          |
| ObjectSendBackward              | button       | None (Context Menu)       | 下移一层                          |
| ObjectEditPoints                | toggleButton | None (Context Menu)       | 编辑顶点                          |
| ShapeStraightConnector          | toggleButton | None (Context Menu)       | 直线连接符                        |
| ShapeElbowConnectorArrow        | toggleButton | None (Context Menu)       | 肘形连接符                        |
| ControlProperties               | button       | None (Context Menu)       | 属性                              |
| ViewCode                        | button       | None (Context Menu)       | 查看代码                          |
| ErrorChecking                   | button       | None (Context Menu)       | 错误检查选项                      |
| SheetDelete                     | button       | None (Context Menu)       | 删除工作表                        |
| MacroRecord                     | button       | None (Context Menu)       | 录制宏                            |
| Cut                             | button       | None (Context Menu)       | 剪切                              |
| Copy                            | button       | None (Context Menu)       | 复制                              |
| PasteAsPicture                  | button       | None (Context Menu)       | 粘贴为图片                        |
| PasteSpecialDialog              | button       | None (Context Menu)       | 选择性粘贴                        |
|                                 | button       | None (Context Menu)       | 粘贴值到可见单元格                |
| PasteMenu                       | splitButton  | None (Context Menu)       | 粘贴                              |
| PasteAllUsingSourceTheme        | button       | None (Context Menu)       | 保留源格式                        |
| PasteValues                     | button       | None (Context Menu)       | 值                                |
| PasteFormulas                   | button       | None (Context Menu)       | 公式                              |
| PasteNoBorders                  | button       | None (Context Menu)       | 无边框                            |
| PasteTranspose                  | button       | None (Context Menu)       | 转置                              |
| DataTypeShowCard                | button       | None (Context Menu)       | 显示数据类型卡                    |
| DataTypeRefresh                 | button       | None (Context Menu)       | 刷新                              |
| DataTypeChange                  | button       | None (Context Menu)       | 更改                              |
| DataTypeConvertToText           | button       | None (Context Menu)       | 转换为文本                        |
|                                 | button       | None (Context Menu)       | 科学计数                          |
| AlignTopExcel                   | toggleButton | None (Context Menu)       | 顶端对齐                          |
| AlignLeft                       | toggleButton | None (Context Menu)       | 左对齐                            |
| AlignMiddleExcel                | toggleButton | None (Context Menu)       | 垂直居中                          |
| AlignCenter                     | toggleButton | None (Context Menu)       | 水平居中                          |
| AlignBottomExcel                | toggleButton | None (Context Menu)       | 底端对齐                          |
| AlignRight                      | toggleButton | None (Context Menu)       | 右对齐                            |
| MergeCenterMenu                 | splitButton  | None (Context Menu)       | 合并                              |
| AutoSum                         | button       | None (Context Menu)       | 求和                              |
| AutoSumAverage                  | button       | None (Context Menu)       | 平均值                            |
| AutoSumCount                    | button       | None (Context Menu)       | 计数                              |
| AutoSumMax                      | button       | None (Context Menu)       | 最大值                            |
| AutoSumMin                      | button       | None (Context Menu)       | 最小值                            |
| AutoSumMoreFunctions            | button       | None (Context Menu)       | 其他函数                          |
|                                 | comboBox     | None (Context Menu)       | 字体                              |
|                                 | comboBox     | None (Context Menu)       | 字号                              |
|                                 | button       | None (Context Menu)       | 粘贴                              |
|                                 | button       | None (Context Menu)       | 粘贴                              |
| PasteSpecialDialog              | button       | None (Context Menu)       | 选择性粘贴                        |
| PasteSpecialDialog              | button       | None (Context Menu)       | 选择性粘贴                        |
| ClearAll                        | button       | None (Context Menu)       | 全部                              |
| ClearFormats                    | button       | None (Context Menu)       | 格式                              |
| ClearContents                   | button       | None (Context Menu)       | 内容                              |
| ClearComments                   | button       | None (Context Menu)       | 批注                              |
| PivotTableOptions               | button       | None (Context Menu)       | 数据透视表选项                    |
|                                 | button       | None (Context Menu)       | 数据透视图选项                    |
| ObjectBringToFrontMenu          | button       | None (Context Menu)       | 置于顶层                          |
| ObjectSendToBack                | button       | None (Context Menu)       | 置于底层                          |
| InputNumericSequence            | button       | None (Context Menu)       | 录入123序列                       |
| InputLetterSequence             | button       | None (Context Menu)       | 录入ABC序列                       |
| InputRomanNumericSequence       | button       | None (Context Menu)       | 录入罗马数字序列                  |
| FillBlankCells                  | button       | None (Context Menu)       | 空白单元格填充                    |
| NewBlankXlsxFile                | button       | None (Context Menu)       | 新建                              |
| Camera                          | button       | None (Context Menu)       | 照相机                            |
| SheetRename                     | button       | None (Context Menu)       | 重命名                            |
| ReviewTrackChangesMenu          | menu         | None (Context Menu)       | 修订                              |
|                                 | button       | None (Context Menu)       | 取消组合                          |
| FontSizeIncrease                | button       | None (Context Menu)       | 增大字号                          |
| FontSizeDecrease                | button       | None (Context Menu)       | 减小字号                          |

> WPS 开发文档： [WPS 基础接口/idMso列表/表格idMso参考](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/idMso%E5%88%97%E8%A1%A8/%E8%A1%A8%E6%A0%BCidMso%E5%8F%82%E8%80%83.html)

------------------------------------------------------------------------
# 文字idMso参考

| 控件名称                               | 控件类型     | 所属Tab               | 功能描述                           |
|----------------------------------------|--------------|-----------------------|------------------------------------|
|                                        | comboBox     | TabChartTools         | 图表元素                           |
| BorderBottomNoToggle                   | button       | TabDesign             | 下框线                             |
| BorderTop                              | toggleButton | TabDesign             | 上框线                             |
| BorderLeft                             | toggleButton | TabDesign             | 左框线                             |
| BorderRight                            | toggleButton | TabDesign             | 右框线                             |
| BordersAll                             | button       | TabDesign             | 所有框线                           |
| BorderOutside                          | button       | TabDesign             | 外侧框线                           |
|                                        | dropDown     | TabDesign             | 线型                               |
|                                        | dropDown     | TabDesign             | 粗细                               |
| TableEraser                            | toggleButton | TabDesign             | 擦除                               |
| TableDrawTable                         | toggleButton | TabDesign             | 绘制表格                           |
| TableDrawTable                         | toggleButton | TabDesign             | 绘制表格                           |
| GroupMacroPlay                         | group        | TabDevelopTools       | 宏                                 |
| MacroPlay                              | button       | TabDevelopTools       | JS宏                               |
| MacroRecordOrStop                      | button       | TabDevelopTools       | 录制新宏                           |
| MacroSecurity                          | button       | TabDevelopTools       | 宏安全性                           |
| VisualBasic                            | button       | TabDevelopTools       | WPS 宏编辑器                       |
| GroupAddins                            | group        | TabDevelopTools       | 加载项                             |
| AddInManager                           | button       | TabDevelopTools       | 加载项                             |
| ComAddInsDialog                        | button       | TabDevelopTools       | COM 加载项                         |
| ActiveXCheckBox                        | button       | TabDevelopTools       | 复选框                             |
| ActiveXTextBox                         | button       | TabDevelopTools       | 文本框                             |
| ListCommands                           | button       | TabDevelopTools       | 命令按钮                           |
| ActiveXListBox                         | button       | TabDevelopTools       | 列表框                             |
| ActiveXComboBox                        | button       | TabDevelopTools       | 组合框                             |
| ActiveXToggleButton                    | button       | TabDevelopTools       | 切换按钮                           |
| ActiveXSpinButton                      | button       | TabDevelopTools       | 数值调节按钮                       |
| ActiveXScrollBar                       | button       | TabDevelopTools       | 滚动条                             |
| ActiveXLabel                           | button       | TabDevelopTools       | 标签                               |
| ActiveXImage                           | button       | TabDevelopTools       | 图像                               |
| MoreControlsDialog                     | button       | TabDevelopTools       | 其他控件                           |
| GroupControls                          | group        | TabDevelopTools       | 控件                               |
| DesignMode                             | toggleButton | TabDevelopTools       | 设计模式                           |
| ControlProperties                      | button       | TabDevelopTools       | 控件属性                           |
| ViewVisualBasicCode                    | button       | TabDevelopTools       | 查看代码                           |
| ContentControlsGroupMenu               | menu         | TabDevelopTools       | 组合                               |
| ContentControlsGroup                   | button       | TabDevelopTools       | 组合                               |
| ContentControlsUngroup                 | button       | TabDevelopTools       | 取消组合                           |
| ContentControlRichText                 | button       | TabDevelopTools       | 格式文本内容控件                   |
| ContentControlText                     | button       | TabDevelopTools       | 纯文本内容控件                     |
| ContentControlPicture                  | button       | TabDevelopTools       | 图片内容控件                       |
| ContentControlBuildingBlockGallery     | button       | TabDevelopTools       | 构建基块库内容控件                 |
| ContentControlCheckBox                 | button       | TabDevelopTools       | 复选框内容控件                     |
| ContentControlComboBox                 | button       | TabDevelopTools       | 组合框内容控件                     |
| ContentControlDropDownList             | button       | TabDevelopTools       | 下拉列表内容控件                   |
| ContentControlData                     | button       | TabDevelopTools       | 日期选取器内容控件                 |
| ContentControlRepeating                | button       | TabDevelopTools       | 重复节内容控件                     |
| ControlsGalleryClassic                 | gallery      | TabDevelopTools       | 旧式工具                           |
| XmlStructure                           | button       | TabDevelopTools       | 结构                               |
| XmlSchema                              | button       | TabDevelopTools       | 架构                               |
| XmlExpansionPacksWord                  | button       | TabDevelopTools       | 扩展包                             |
| GroupXmlMappingPane                    | group        | TabDevelopTools       | 映射                               |
| ToggleXmlMappingPane                   | button       | TabDevelopTools       | XML 映射窗格                       |
| ObjectEditPoints                       | toggleButton | TabDrawingTools       | 编辑顶点                           |
| ObjectFillMoreColorsDialog             | button       | TabDrawingTools       | 其他填充颜色                       |
| ArrowsMore                             | button       | TabDrawingTools       | 其他箭头                           |
| FontColorMoreColorsDialogExcel         | button       | TabDrawingTools       | 其他字体颜色                       |
| ShapeStylesGallery                     | gallery      | TabDrawingTools       | 形状样式                           |
| ShapeFillMoreGradientsDialog           | button       | TabDrawingTools       | 填充                               |
| Bold                                   | toggleButton | TabDrawingTools       | 加粗                               |
| Italic                                 | toggleButton | TabDrawingTools       | 倾斜                               |
| Underline                              | toggleButton | TabDrawingTools       | 下划线                             |
| AlignLeft                              | toggleButton | TabDrawingTools       | 左对齐                             |
| AlignCenter                            | toggleButton | TabDrawingTools       | 居中对齐                           |
| AlignRight                             | toggleButton | TabDrawingTools       | 右对齐                             |
| AlignJustify                           | toggleButton | TabDrawingTools       | 两端对齐                           |
| TextWrappingInLineWithText             | toggleButton | TabDrawingTools       | 嵌入型                             |
| TextWrappingSquare                     | toggleButton | TabDrawingTools       | 四周型环绕                         |
| TextWrappingTight                      | toggleButton | TabDrawingTools       | 紧密型环绕                         |
| TextWrappingBehindText                 | toggleButton | TabDrawingTools       | 衬于文字下方                       |
| TextWrappingInFrontOfText              | toggleButton | TabDrawingTools       | 浮于文字上方                       |
| TextWrappingTopAndBottom               | toggleButton | TabDrawingTools       | 上下型环绕                         |
| TextWrappingThrough                    | toggleButton | TabDrawingTools       | 穿越型环绕                         |
| ObjectBringToFront                     | button       | TabDrawingTools       | 置于顶层                           |
| ObjectBringForward                     | button       | TabDrawingTools       | 上移一层                           |
| ObjectSendToBack                       |              | TabDrawingTools       | 置于底层                           |
| ObjectSendBehindText                   | button       | TabDrawingTools       | 衬于文字下方                       |
| ObjectsAlignLeft                       | button       | TabDrawingTools       | 左对齐                             |
| AlignDistributeHorizontally            | button       | TabDrawingTools       | 横向分布                           |
| AlignDistributeVertically              | button       | TabDrawingTools       | 纵向分布                           |
| ViewGridlinesWord                      | checkBox     | TabDrawingTools       | 网格线                             |
| AlignJustifyMenu                       |              | TabDrawingTools       | 对齐                               |
| ObjectRotateLeft90                     | button       | TabDrawingTools       | 向左旋转 90°                       |
|                                        | editBox      | TabDrawingTools       | 大小                               |
|                                        | editBox      | TabDrawingTools       |                                    |
| FontColorPicker                        | gallery      | TabDrawingTools       | 文本颜色                           |
|                                        | comboBox     | TabDrawingTools       | 字体名                             |
| ObjectFillMoreColorsDialog             |              | TabDrawingTools_Vml   | 其他填充颜色                       |
| ArrowsMore                             |              | TabDrawingTools_Vml   | 其他箭头                           |
| ObjectBringToFront                     | button       | TabDrawingTools_Vml   | 置于顶层                           |
| ObjectSendBehindText                   |              | TabDrawingTools_Vml   | 衬于文字下方                       |
| AlignDistributeHorizontally            | button       | TabDrawingTools_Vml   | 横向分布                           |
| AlignDistributeVertically              | button       | TabDrawingTools_Vml   | 纵向分布                           |
| AlignJustifyMenu                       |              | TabDrawingTools_Vml   | 对齐                               |
| Font                                   | comboBox     | TabDrawingTools_Vml   | 字体名                             |
|                                        | editBox      | TabDrawingTools_Vml   | 高度                               |
|                                        | editBox      | TabDrawingTools_Vml   | 宽度                               |
| EquationOptions                        |              | TabEquationTools      | 公式选项                           |
| EquationInsertGallery                  |              | TabEquationTools      | 公式                               |
| EquationProfessional                   |              | TabEquationTools      | 专业型                             |
| EquationLinearFormat                   |              | TabEquationTools      | 线性                               |
| EquationNormalText                     |              | TabEquationTools      | 普通文本                           |
| EquationSymbolsInsertGallery           |              | TabEquationTools      | 公式符号                           |
| EquationFractionGallery                |              | TabEquationTools      | 分数                               |
| EquationScriptGallery                  |              | TabEquationTools      | 上下标                             |
| EquationRadicalGallery                 |              | TabEquationTools      | 根式                               |
| EquationIntegralGallery                |              | TabEquationTools      | 积分                               |
| EquationLargeOperatorGallery           |              | TabEquationTools      | 大型运算符                         |
| EquationDelimiterGallery               |              | TabEquationTools      | 括号                               |
| EquationFunctionGallery                |              | TabEquationTools      | 函数                               |
| EquationAccentGallery                  |              | TabEquationTools      | 导数符号                           |
| EquationLimitGallery                   |              | TabEquationTools      | 极限和对数                         |
| EquationOperatorGallery                |              | TabEquationTools      | 运算符                             |
| EquationMatrixGallery                  |              | TabEquationTools      | 矩阵                               |
| FileSaveAsPicture                      |              | TabFile               | 输出为图片                         |
| FileNewMenu                            | menu         | TabFile               | 新建                               |
| FileNew                                | button       | TabFile               | 新建                               |
| FileNewFromTemplate                    | button       | TabFile               | 本机上的模板                       |
| FileNewFromDefaultTemplate             | button       | TabFile               | 从默认模板新建                     |
| FilePrintMenu                          | splitButton  | TabFile               | 打印                               |
| FilePrint                              | button       | TabFile               | 打印                               |
| FilePrintPreview                       | button       | TabFile               | 打印预览                           |
|                                        | splitButton  | TabFile               |                                    |
|                                        | button       | TabFile               |                                    |
|                                        | button       | TabFile               |                                    |
| FileInfoMenu                           |              | TabFile               | 文档加密                           |
| FileProperties                         | button       | TabFile               | 属性                               |
|                                        | splitButton  | TabFile               | 备份与恢复                         |
|                                        | splitButton  | TabFile               | 备份与恢复                         |
| FileBackupManagement                   | menu         | TabFile               | 备份管理                           |
| FileBackupManagement                   | button       | TabFile               | 备份管理                           |
| FileBackupHistory                      | button       | TabFile               | 历史版本                           |
| FileHelp                               | menu         | TabFile               | 帮助                               |
| Help                                   | button       | TabFile               | WPS 文字 帮助                      |
| About                                  | button       | TabFile               | 关于 WPS 文字                      |
| FileSave                               | button       | TabFile               | 保存                               |
| FileOpen                               | button       | TabFile               | 打开                               |
| FileSaveAsMenu                         | menu         | None                  | 另存为                             |
| FileSaveAsPdfOrXps                     | button       | None                  | 输出为PDF                          |
| FileOfdPrintMenu                       | button       | None                  | 输出为OFD                          |
| FileSaveAsPicture                      | button       | None                  | 输出为图片                         |
| DocumentSplitAndMerge                  |              | None                  | 文档拆分合并                       |
| FileShare                              | button       | TabFile               | 分享文档                           |
| FileMenuSendMail                       | button       | TabFile               | 发送邮件                           |
| ApplicationOptionsDialog               |              | TabFile               | 选项                               |
| ObjectFillMoreColorsDialog             | button       | TabGraphicTool        | 其他填充颜色                       |
| FontColorMoreColorsDialogExcel         | button       | TabGraphicTool        | 其他字体颜色                       |
| ShapeStylesGallery                     | gallery      | TabGraphicTool        | 图形样式                           |
| ShapeFillMoreGradientsDialog           | button       | TabGraphicTool        | 图形填充                           |
| Bold                                   | toggleButton | TabGraphicTool        | 加粗                               |
| Italic                                 | toggleButton | TabGraphicTool        | 倾斜                               |
| Underline                              | toggleButton | TabGraphicTool        | 下划线                             |
| AlignLeft                              | toggleButton | TabGraphicTool        | 左对齐                             |
| AlignCenter                            | toggleButton | TabGraphicTool        | 居中                               |
| AlignRight                             | toggleButton | TabGraphicTool        | 右对齐                             |
| AlignJustify                           | toggleButton | TabGraphicTool        | 两端对齐                           |
| TextWrappingInLineWithText             | toggleButton | TabGraphicTool        | 嵌入型                             |
| TextWrappingSquare                     | toggleButton | TabGraphicTool        | 四周型环绕                         |
| TextWrappingTight                      | toggleButton | TabGraphicTool        | 紧密型环绕                         |
| TextWrappingBehindText                 | toggleButton | TabGraphicTool        | 衬于文字下方                       |
| TextWrappingInFrontOfText              | toggleButton | TabGraphicTool        | 浮于文字上方                       |
| TextWrappingTopAndBottom               | toggleButton | TabGraphicTool        | 上下型环绕                         |
| TextWrappingThrough                    | toggleButton | TabGraphicTool        | 穿越型环绕                         |
| ObjectBringToFront                     | button       | TabGraphicTool        | 置于顶层                           |
| ObjectBringForward                     | button       | TabGraphicTool        | 上移一层                           |
| ObjectSendToBackMenu                   |              | TabGraphicTool        | 置于底层                           |
| ObjectSendBehindText                   | button       | TabGraphicTool        | 衬于文字下方                       |
| ObjectsAlignLeft                       | button       | TabGraphicTool        | 左对齐                             |
| AlignDistributeHorizontally            | button       | TabGraphicTool        | 横向分布                           |
| AlignDistributeVertically              | button       | TabGraphicTool        | 纵向分布                           |
| ViewGridlinesWord                      | checkBox     | TabGraphicTool        | 网格线                             |
| AlignJustifyMenu                       |              | TabGraphicTool        | 对齐                               |
| ObjectRotateLeft90                     | button       | TabGraphicTool        | 向左旋转 90°                       |
|                                        | button       | None                  | 重设形状和大小                     |
|                                        | editBox      | TabGraphicTool        | 高度                               |
|                                        | editBox      | TabGraphicTool        | 宽度                               |
| FontColorPicker                        | gallery      | TabGraphicTool        | 文本颜色                           |
|                                        | comboBox     | TabGraphicTool        | 字体名                             |
|                                        | gallery      | TabHeaderFooter       | 配套组合                           |
|                                        | gallery      | TabHeaderFooter       | 页眉                               |
|                                        | gallery      | TabHeaderFooter       | 页脚                               |
| DateAndTimeInsert                      | button       | TabHeaderFooter       | 日期和时间                         |
| PictureInsertFromFile                  | button       | TabHeaderFooter       | 本地图片                           |
| FieldInsert                            | button       | TabHeaderFooter       | 域                                 |
| HeaderFooterInsert                     | button       | TabHeaderFooter       | 插入                               |
| HeaderFooterPositionHeaderFromTop      | editBox      | TabHeaderFooter       | 页眉顶端距离                       |
| HeaderFooterPositionFooterFromBottom   | editBox      | TabHeaderFooter       | 页脚底端距离                       |
| PasteKeepSourceFormat                  | button       | TabHome               | 保留源格式                         |
| PasteMatchCurrentFormat                | button       | TabHome               | 匹配当前格式                       |
| PasteSetDefault                        | button       | TabHome               | 设置默认粘贴                       |
| PasteMenu                              | splitButton  | TabHome               | 粘贴                               |
| PasteTextOnly                          |              | None                  | 只粘贴文本                         |
| PasteSpecialDialog                     | button       | None                  | 选择性粘贴                         |
| Cut                                    | button       | TabHome               | 剪切                               |
| Copy                                   | button       | TabHome               | 复制                               |
| ShowClipboard                          | button       | TabHome               | 剪贴板                             |
| FormatPainter                          | toggleButton | None                  | 格式刷                             |
| Bold                                   | toggleButton | TabHome               | 加粗                               |
| Italic                                 | toggleButton | TabHome               | 倾斜                               |
| UnderlineGallery                       | gallery      | TabHome               | 下划线                             |
| UnderlineMoreUnderlinesDialog          | button       | TabHome               | 其他下划线                         |
| UnderlineColorGallery                  | gallery      | TabHome               | 下划线颜色                         |
| UnderlineColorMoreColorsDialog         | button       | TabHome               | 其他下划线颜色                     |
| UnderlineGallery                       | toggleButton | TabHome               | 下划线                             |
| Strikethrough                          | toggleButton | TabHome               | 删除线                             |
| EmphasisMark                           | toggleButton | TabHome               | 着重号                             |
| Superscript                            | toggleButton | TabHome               | 上标                               |
| Subscript                              | toggleButton | TabHome               | 下标                               |
| TextEffectsGallery                     | gallery      | TabHome               | 文本效果                           |
| WordArtInsertGallery                   | gallery      | TabHome               | 艺术字                             |
| TextEffectShadowGallery                | gallery      | TabHome               | 阴影                               |
| TextReflectionGallery                  | gallery      | TabHome               | 倒影                               |
| TextEffectGlowGallery                  | gallery      | TabHome               | 发光                               |
| TextEffectThreeDRotationGallery        | gallery      | TabHome               | 三维旋转                           |
| TextEffectTransformGallery             | gallery      | TabHome               | 转换                               |
| ClearFormatting                        | button       | TabHome               | 清除格式                           |
| CharacterShading                       | toggleButton | TabHome               | 字符底纹                           |
| ChangeCaseGallery                      | gallery      | TabHome               | 更改大小写                         |
| CharacterBorder                        | toggleButton | TabHome               | 字符边框                           |
| AsianLayoutPhoneticGuide               | button       | TabHome               | 拼音指南                           |
| Font                                   | comboBox     | TabHome               | 字体                               |
|                                        | comboBox     | TabHome               | 字体                               |
| StrikethroughAndEmphasisMark           | splitButton  | TabHome               | 删除线                             |
| ChangeCaseGallery                      | gallery      | TabHome               | 更改大小写                         |
| AsianLayoutCharactersEnclose           | button       | TabHome               | 带圈字符                           |
| AlignLeft                              | toggleButton | TabHome               | 左对齐                             |
| AlignCenter                            | toggleButton | TabHome               | 居中对齐                           |
| AlignRight                             | toggleButton | TabHome               | 右对齐                             |
| AlignJustify                           | toggleButton | TabHome               | 两端对齐                           |
| ParagraphDistributed                   | toggleButton | TabHome               | 分散对齐                           |
| GroupFont                              | group        | TabHome               | 字体                               |
| FontSize                               | comboBox     | None                  | 字号                               |
| FontSizeIncreaseWord                   | button       | None                  | 增大字体                           |
| FontSizeDecreaseWord                   | button       | None                  | 减小字体                           |
| TextHighlightColorPicker               | gallery      | None                  | 突出显示                           |
| FontColorPicker                        | gallery      | None                  | 文本颜色                           |
|                                        | menu         | TabHome               | 横纵向字号                         |
|                                        | comboBox     | None                  | 横向字号                           |
|                                        | comboBox     | None                  | 纵向字号                           |
|                                        | gallery      | None                  | 加粗                               |
| Bullets                                | toggleButton | TabHome               | 项目符号                           |
| Bullets                                | toggleButton | TabHome               | 项目符号                           |
| BulletsGalleryWord                     | gallery      | TabHome               | 项目符号                           |
| BulletDefineNew                        | button       | TabHome               | 自定义项目符号                     |
| NumberingGalleryWord                   | gallery      | TabHome               | 编号                               |
| DefineNewNumberFormat                  | button       | TabHome               | 自定义编号                         |
| ListLevelGallery                       | gallery      | TabHome               | 更改编号级别                       |
| IndentDecreaseWord                     | button       | TabHome               | 减少缩进量                         |
| TextDirectionLeftToRight               | button       | TabHome               | 从左向右                           |
| TextDirectionRightToLeft               | button       | TabHome               | 从右向左                           |
| ParagraphDistributed                   | toggleButton | TabHome               |                                    |
| LineSpacingMoreLineSpacingDialog       | button       | TabHome               | 其他                               |
| LineSpacingGallery                     | gallery      | TabHome               | 行距                               |
| FillColorsMoreFillColorsDialog         | button       | TabHome               | 其他填充颜色                       |
| ShadingColorPicker                     | gallery      | TabHome               | 底纹颜色                           |
| BorderBottomNoToggle                   | button       | TabHome               | 下框线                             |
| BorderTopWord                          | button       | TabHome               | 上框线                             |
| BorderLeftWord                         | button       | TabHome               | 左框线                             |
| BorderRightWord                        | button       | TabHome               | 右框线                             |
| BorderNone                             | button       | TabHome               | 无框线                             |
| BordersAll                             | button       | TabHome               | 所有框线                           |
| BorderOutside                          | button       | TabHome               | 外侧框线                           |
| BorderInside                           | button       | TabHome               | 内部框线                           |
| BorderInsideHorizontal                 | button       | TabHome               | 内部横框线                         |
| BorderInsideVertical                   | button       | TabHome               | 内部竖框线                         |
| BordersShadingDialogWord               | button       | TabHome               | 边框和底纹                         |
| TableBordersMenu                       | splitButton  | TabHome               | 外侧框线                           |
|                                        | splitButton  | TabHome               | 边框                               |
| EastAsianEditingMarks                  | menu         | TabHome               | 显示/隐藏编辑标记                  |
| EastAsianParagraphMarks                | toggleButton | TabHome               | 显示/隐藏段落标记                  |
| EastAsianParagraphDistributed          | toggleButton | TabHome               | 显示/隐藏段落布局按钮              |
| Tabs                                   | button       | TabHome               | 制表位                             |
|                                        | menu         | TabHome               | 工具                               |
| GroupParagraph                         | group        | TabHome               |                                    |
| IndentIncreaseWord                     | button       | None                  | 增加缩进量                         |
| GroupParagraph                         | group        | TabHome               | 段落                               |
| QuickStylesGallery                     | gallery      | TabHome               | 样式                               |
| QuickStylesSaveSelectionAsNew          | button       | TabHome               | 新建样式                           |
| ClearFormatting                        | button       | TabHome               | 清除格式                           |
|                                        | button       | TabHome               | 显示更多样式                       |
| QuickStylesSaveClearSelection          | splitButton  | TabHome               | 新建样式                           |
| GroupStyles                            | group        | TabHome               | 样式                               |
| FindDialog                             | button       | TabHome               | 查找                               |
| ReplaceDialog                          | button       | TabHome               | 替换                               |
| GoTo                                   | button       | TabHome               | 定位                               |
| FindAndReplaceDialog                   | button       | TabHome               | 查找替换                           |
| ObjectsSelect                          | toggleButton | TabHome               | 选择对象                           |
| SelectTableWithDashedBorders           | toggleButton | TabHome               | 虚框选择表格                       |
| SelectionPane                          | button       | TabHome               | 选择窗格                           |
| SelectMenu                             | menu         | TabHome               | 选择                               |
| SelectAll                              | button       | TabHome               | 全选                               |
| GroupEditing                           | group        | TabHome               | 编辑                               |
| WordTools                              | menu         | TabHome               | 文字排版                           |
| GroupParagraphTools                    | group        | TabHome               | 工具                               |
| FontDialog                             | button       | TabHome               | 字体                               |
| ParagraphDialog                        | button       | TabHome               | 段落                               |
| StylesPane                             | button       | TabHome               | 样式和格式                         |
| InkBallpointPen                        | button       | TabInkTools           | 圆珠笔                             |
| InkWatercolorPen                       | button       | TabInkTools           | 水彩笔                             |
| InkHighlighterPen                      | button       | TabInkTools           | 荧光笔                             |
| InkDrawPen                             | splitButton  | TabInkTools           | 笔                                 |
| InkShowLine                            | button       | TabInkTools           | 直线                               |
| InkShowWave                            | button       | TabInkTools           | 波浪线                             |
| InkShowRect                            | button       | TabInkTools           | 矩形                               |
| InkDrawShape                           | splitButton  | TabInkTools           | 形状                               |
| InkEraser                              | button       | TabInkTools           | 橡皮擦                             |
|                                        | gallery      | TabInkTools           | 墨迹样式                           |
| InkColor                               | menu         | TabInkTools           | 颜色                               |
| InkLineStyle                           | gallery      | TabInkTools           | 线型                               |
| ExitInkDraw                            | button       | TabInkTools           | 关闭                               |
|                                        | toggleButton | TabInkTools_mac       | 笔                                 |
|                                        | toggleButton | TabInkTools_mac       | 类型                               |
|                                        | gallery      | TabInkTools_mac       | 墨迹样式                           |
| PageBreakInsertWord                    | button       | TabInsert             | 分页符                             |
| TextWrappingBreakInsertWord            | button       | TabInsert             | 换行符                             |
| NextPageSectionBreakInsertWord         | button       | TabInsert             | 下一页分节符                       |
| ContinuousSectionBreakInsertWord       | button       | TabInsert             | 连续分节符                         |
| EvenPageSectionBreakInsertWord         | button       | TabInsert             | 偶数页分节符                       |
| OddPageSectionBreakInsertWord          | button       | TabInsert             | 奇数页分节符                       |
| PageBreakGallery                       | gallery      | TabInsert             | 分页                               |
| ColumnBreakInsertWord                  | button       | TabInsert             | 分栏符                             |
| DeletePageNumInsertWord                | button       | TabInsert             | 删除页码                           |
| BlankPageVertical                      | button       | TabInsert             | 竖向                               |
| BlankPageHorizontal                    | button       | TabInsert             | 横向                               |
| BlankPageInsertGallery                 | gallery      | TabInsert             | 空白页                             |
| TableInsertDialogWord                  |              | TabInsert             | 插入表格                           |
| TableDrawTable                         | toggleButton | TabInsert             | 绘制表格                           |
| ConvertTextToTable                     | button       | TabInsert             | 文本转换成表格                     |
| ConvertTableToText                     | button       | TabInsert             | 表格转换成文本                     |
| TableInsertGallery                     | gallery      | TabInsert             | 表格                               |
| TableInsertGallery                     | gallery      | TabInsert             | 表格                               |
| TableInsertGallery                     | gallery      | TabInsert             | 表格                               |
| GroupInsertTables                      | group        | TabInsert             | 表格                               |
| PictureInsertGallery                   | gallery      | TabInsert             | 图片                               |
| PictureInsertFromFile                  | button       | TabInsert             | 来自文件                           |
| PictureToText                          | button       | TabInsert             | 图片转文字                         |
| PictureInsertFromFile                  | button       | TabInsert             | 图片                               |
| DocumentBarcode                        | button       | TabInsert             | 公文二维条码                       |
| BarcodeInsert                          | splitButton  | TabInsert             | 条形码                             |
| ClipArtInsert                          | toggleButton | TabInsert             | 剪贴画                             |
| ShapesInsertGallery                    | gallery      | TabInsert             | 形状                               |
| ShapesInsertGallery                    | gallery      | TabInsert             | 形状                               |
| InsertCanvas                           | button       | TabInsert             | 新建绘图画布                       |
| ChartInsert                            | button       | TabInsert             | 图表                               |
| TextMulti-line                         | toggleButton | TabInsert             | 多行文字                           |
| TextBoxDrawMenu                        | menu         | TabInsert             | 文本框                             |
| TextBoxInsertHorizontalWord            | button       | TabInsert             | 横向                               |
| TextBoxInsertVerticalWord              | button       | TabInsert             | 竖向                               |
| WordArtInsertGallery                   | gallery      | TabInsert             | 艺术字                             |
| DropCapOptionsDialog                   | button       | TabInsert             | 首字下沉                           |
| OleObjectInsertMenu                    | splitButton  | TabInsert             | 对象                               |
| TextFromFileInsert                     | button       | TabInsert             | 文件中的文字                       |
| OleObjectInsertMenu                    | splitButton  | TabInsert             | 对象                               |
| ReviewNewComment                       | button       | None                  | 批注                               |
| HeaderFooterGallery                    | gallery      | TabInsert             | 页眉页脚                           |
| PageNumber                             | button       | TabInsert             | 页码                               |
| HeaderFooterPageNumberInsertGallery    | gallery      | TabInsert             | 页码                               |
| PageNumberFormatGallery                | gallery      | TabInsert             | 页码                               |
| GroupHeaderFooter                      | group        | TabInsert             | 页眉页脚                           |
| WatermarkGallery                       | gallery      | TabInsert             | 水印                               |
| WatermarkInsert                        | button       | TabInsert             | 插入水印                           |
| WatermarkRemove                        | button       | TabInsert             | 删除文档中的水印                   |
| GroupInsertLinks                       | group        | TabInsert             | 链接                               |
| HyperlinkInsert                        | button       | TabInsert             | 超链接                             |
| CrossReferenceInsert                   | button       | TabInsert             | 交叉引用                           |
| BookmarkInsert                         | button       | TabInsert             | 书签                               |
| HyperlinkInsert                        | button       | TabInsert             | 超链接                             |
| CoverPageInsertGallery                 | gallery      | TabInsert             | 封面页                             |
| CoverPageInsertGallery                 | gallery      | TabInsert             | 封面页                             |
| GroupInsertTables                      | group        | TabInsert             | 表格                               |
| SmartArtInsert                         | button       | TabInsert             | 智能图形                           |
| DateAndTimeInsert                      | button       | TabInsert             | 日期                               |
| QuickPartsInsertGallery                | gallery      | TabInsert             | 文档部件                           |
| AutoTextGallery                        | gallery      | TabInsert             | 自动图文集                         |
| PageX                                  | button       | TabInsert             | 第 X 页                            |
| TotalofYPage                           | button       | TabInsert             | 共 Y 页                            |
| PageXofY                               | button       | TabInsert             | 第 X 页 共 Y 页                    |
| SaveSelectionToAutoTextGallery         | button       | TabInsert             | 将所选内容保存到自动图文集库       |
| FieldInsert                            | button       | TabInsert             | 域                                 |
| DocumentFieldInsert                    | button       | TabInsert             | 公文域                             |
| GroupInsertSymbols                     | group        | TabInsert             | 符号                               |
| SymbolInsertGallery                    | gallery      | None                  | 符号                               |
| EquationInsertGallery                  |              | TabInsert             | 公式                               |
| EquationInsertNew                      |              | TabInsert             | 插入新公式                         |
| NumberInsert                           | button       | TabInsert             | 数字                               |
| FieldText                              | button       | TabInsert             | 文字型窗体域                       |
| FieldCheckBox                          | button       | TabInsert             | 复选框型窗体域                     |
| FieldDrop-Down                         | button       | TabInsert             | 下拉型窗体域                       |
| FieldOptions                           | button       | TabInsert             | 属性                               |
| FieldShading                           | toggleButton | TabInsert             | 窗体域底纹                         |
| FieldsReset                            | button       | TabInsert             | 重置窗体域                         |
| FieldsProtect                          | toggleButton | TabInsert             | 保护窗体                           |
| CoverPageInsertGallery                 | gallery      | TabInsert_Vml         | 封面页                             |
| TableInsertDialogWord                  |              | TabInsert_Vml         | 插入表格                           |
| TableDrawTable                         | toggleButton | TabInsert_Vml         | 绘制表格                           |
| ConvertTextToTable                     | button       | TabInsert_Vml         | 文本转换成表格                     |
| TableInsertGallery                     | gallery      | TabInsert_Vml         | 表格                               |
| TableInsertGallery                     | gallery      | TabInsert_Vml         | 表格                               |
| PictureInsertFromFile                  | button       | TabInsert_Vml         | 来自文件                           |
| ShapesInsertGallery                    | gallery      | TabInsert_Vml         | 形状                               |
| ShapesInsertGallery                    | gallery      | TabInsert_Vml         | 形状                               |
| InsertCanvas                           | button       | TabInsert_Vml         | 新建绘图画布                       |
| ChartInsert                            | button       | TabInsert_Vml         | 图表                               |
| TextBoxDrawMenu                        | menu         | TabInsert_Vml         | 文本框                             |
| DropCapOptionsDialog                   | button       | TabInsert_Vml         | 首字下沉                           |
| OleObjectInsertMenu                    | splitButton  | TabInsert_Vml         | 对象                               |
| ReviewNewComment                       | button       | None                  | 批注                               |
| PageNumber                             | button       | TabInsert_Vml         | 页码                               |
| WatermarkGallery                       | gallery      | TabInsert_Vml         | 水印                               |
| HyperlinkInsert                        | button       | TabInsert_Vml         | 超链接                             |
| CrossReferenceInsert                   | button       | TabInsert_Vml         | 交叉引用                           |
| BookmarkInsert                         | button       | TabInsert_Vml         | 书签                               |
| SymbolInsertGallery                    | gallery      | TabInsert_Vml         | 符号                               |
| SymbolInsertGallery                    | gallery      | TabInsert_Vml         | 符号                               |
| FieldShading                           | toggleButton | TabInsert_Vml         | 窗体域底纹                         |
| InsertScreenGrab                       |              | None                  | 截图取字                           |
| DateAndTimeInsert                      | button       | TabInsert_Vml         | 日期                               |
| QuickPartsInsertGallery                | gallery      | TabInsert_Vml         | 文档部件                           |
| AutoTextGallery                        | gallery      | TabInsert_Vml         | 自动图文集                         |
| SaveSelectionToAutoTextGallery         | button       | TabInsert_Vml         | 将所选内容保存到自动图文集库       |
| FieldInsert                            | button       | TabInsert_Vml         | 域                                 |
| DocumentFieldInsert                    | button       | TabInsert_Vml         | 公文域                             |
| EquationInsertGallery                  |              | TabInsert_Vml         | 公式                               |
| EquationInsertNew                      |              | TabInsert_Vml         | 插入新公式                         |
| CoverPageInsertGallery                 | gallery      | TabInsert_Vml         | 封面页                             |
| MailMergeFindRecipient                 | button       | TabMailings           | 收件人                             |
| MessagePrevious                        | button       | TabMailings           | 上一条                             |
|                                        | button       | TabMailings           |                                    |
| MailMergeGoToRecord                    | editBox      | TabMailings           | 跳转记录                           |
| MessageNext                            | button       | TabMailings           | 下一条                             |
|                                        | button       | TabMailings           |                                    |
| MailMergeMergeToPrinter                | button       | TabMailings           | 合并到打印机                       |
|                                        | button       | TabMailings           | 合并发送                           |
|                                        | button       | TabOfficial_Elements  | 页码                               |
|                                        | gallery      | TabOfficial_Elements  | 水印                               |
|                                        | toggleButton | TabOfficial_Elements  | 绘制表格                           |
|                                        | button       | TabOfficial_Elements  | 文本转换成表格                     |
|                                        | gallery      | TabOfficial_Elements  | 表格                               |
| PictureInsertFromFile                  | button       | TabOfficial_Elements  | 本地图片                           |
|                                        | gallery      | TabOfficial_Elements  | 符号                               |
|                                        | gallery      | TabOfficial_Elements  | 符号                               |
|                                        | button       | TabOfficial_Elements  | 图表                               |
|                                        | menu         | TabOfficial_Elements  | 文本框                             |
|                                        | gallery      | TabOfficial_Elements  | 形状                               |
|                                        | gallery      | TabOfficial_Elements  | 形状                               |
|                                        | button       | TabOfficial_Elements  | 新建绘图画布                       |
|                                        | button       | TabOfficial_Elements  | 公文域                             |
|                                        | button       | TabOfficial_Elements  | 超链接                             |
|                                        | button       | TabOfficial_Elements  | 书签                               |
|                                        | button       | TabOfficial_Print     | 打印                               |
|                                        | dropDown     | TabOutline            | 大纲级别                           |
|                                        | dropDown     | TabOutline            | 大纲级别                           |
|                                        | dropDown     | TabOutline            | 大纲级别                           |
|                                        | dropDown     | TabOutline            | 显示级别                           |
| OutlineViewClose                       | button       | TabOutline            | 关闭                               |
| GroupThemeEdit                         | group        | TabPageLayout         | 编辑主题                           |
| ThemeSearchOfficeOnline                | button       | TabPageLayout         | 主题                               |
| ThemesGallery                          | gallery      | TabPageLayout         | 主题                               |
| ThemeFontsMenu                         | menu         | TabPageLayout         | 字体                               |
| ThemeEffectsSearchOfficeOnline         | button       | TabPageLayout         | 效果                               |
| ThemeEffectsGallery                    | gallery      | TabPageLayout         | 效果                               |
| TextDirectionGalleryWord               | gallery      | TabPageLayout         | 文字方向                           |
| TextDirectionOptionsDialog             | button       | TabPageLayout         | 文字方向选项                       |
| PageMarginsGallery                     | gallery      | TabPageLayout         | 页边距                             |
| MarginsCustomMargins                   | button       | TabPageLayout         | 自定义页边距                       |
| PageOrientationGallery                 | gallery      | TabPageLayout         | 纸张方向                           |
| PageSizeGallery                        | gallery      | TabPageLayout         | 纸张大小                           |
| PageSizeMorePaperSizesDialog           | button       | TabPageLayout         | 其它页面大小                       |
| ColumnsOne                             | button       | TabPageLayout         | 一栏                               |
| ColumnsTwo                             | button       | TabPageLayout         | 两栏                               |
| ColumnsThree                           | button       | TabPageLayout         | 三栏                               |
| ColumnsDialog                          | button       | TabPageLayout         | 更多分栏                           |
| TableColumnsGallery                    | gallery      | TabPageLayout         | 分栏                               |
| BreakPage                              | button       | TabPageLayout         | 分页符                             |
| BreakTextWrapping                      | button       | TabPageLayout         | 换行符                             |
| BreakNextPageSection                   | button       | TabPageLayout         | 下一页分节符                       |
| BreakContinuousSection                 | button       | TabPageLayout         | 连续分节符                         |
| BreakEvenPageSection                   | button       | TabPageLayout         | 偶数页分节符                       |
| BreakOddPageSection                    | button       | TabPageLayout         | 奇数页分节符                       |
| BreaksMenu                             | menu         | TabPageLayout         | 分隔符                             |
| BreakColumn                            | button       | TabPageLayout         | 分栏符                             |
| LineNumbersOff                         | toggleButton | TabPageLayout         | 无                                 |
| LineNumbersContinuous                  | toggleButton | TabPageLayout         | 连续                               |
| LineNumbersResetPage                   | toggleButton | TabPageLayout         | 每页重编行号                       |
| LineNumbersResetSection                | toggleButton | TabPageLayout         | 每节重编行号                       |
| LineNumbersSuppress                    | toggleButton | TabPageLayout         | 禁止用于当前段落                   |
| LineNumbersDoNotShowLineNumbeforBlank  | toggleButton | TabPageLayout         | 空行不显示行号                     |
| LineNumbersOptionsDialog               | toggleButton | TabPageLayout         | 行编号选项                         |
| LineNumbersMenu                        | menu         | TabPageLayout         | 行号                               |
| GroupPageLayoutSetup                   | group        | TabPageLayout         | 页面设置                           |
| PageColorMoreColorsDialog              | button       | TabPageLayout         | 其他填充颜色                       |
| PageColorPicker                        | gallery      | TabPageLayout         | 取色器                             |
| PageColorGallery                       | gallery      | TabPageLayout         | 背景                               |
| BackgroundPicture                      | button       | TabPageLayout         | 图片背景                           |
| BackgroundMoreBackgroundDialog         | button       | TabPageLayout         | 其他背景                           |
| Gradient                               | button       | TabPageLayout         | 渐变                               |
| Texture                                | button       | TabPageLayout         | 纹理                               |
| Pattern                                | button       | TabPageLayout         | 图案                               |
| WatermarkGallery                       | gallery      | TabPageLayout         | 水印                               |
| WatermarkCustomDialog                  | button       | TabPageLayout         | 插入水印                           |
| WatermarkRemove                        | button       | TabPageLayout         | 删除文档中的水印                   |
| PageBorderAndShadingDialog             | button       | TabPageLayout         | 页面边框                           |
| GroupPageBackground                    | group        | TabPageLayout         | 页面背景                           |
| GenkoSettingDialog                     | button       | TabPageLayout         | 稿纸设置                           |
| GroupGenkoSetting                      | group        | TabPageLayout         | 稿纸设置                           |
| TextWrappingInLineWithText             | toggleButton | TabPageLayout         | 嵌入型                             |
| TextWrappingSquare                     | toggleButton | TabPageLayout         | 四周型环绕                         |
| TextWrappingTight                      | toggleButton | TabPageLayout         | 紧密型环绕                         |
| TextWrappingBehindText                 | toggleButton | TabPageLayout         | 衬于文字下方                       |
| TextWrappingInFrontOfText              | toggleButton | TabPageLayout         | 浮于文字上方                       |
| TextWrappingTopAndBottom               | toggleButton | TabPageLayout         | 上下型环绕                         |
| TextWrappingThrough                    | toggleButton | TabPageLayout         | 穿越型环绕                         |
| TextWrappingMenu                       | menu         | TabPageLayout         | 文字环绕                           |
| ObjectBringForwardMenu                 | splitButton  | TabPageLayout         | 上移一层                           |
| ObjectBringToFront                     | button       | TabPageLayout         | 置于顶层                           |
| ObjectBringInFrontOfText               | button       | TabPageLayout         | 浮于文字上方                       |
| ObjectBringForward                     | button       | TabPageLayout         | 上移一层                           |
| ObjectSendBackwardMenu                 | splitButton  | TabPageLayout         | 下移一层                           |
| ObjectSendToBack                       | button       | TabPageLayout         | 置于底层                           |
| ObjectSendBehindText                   | button       | TabPageLayout         | 衬于文字下方                       |
| ObjectSendBackward                     | button       | TabPageLayout         | 下移一层                           |
| SelectionPane                          | toggleButton | TabPageLayout         | 选择窗格                           |
| ObjectsAlignLeft                       | button       | TabPageLayout         | 左对齐                             |
| ObjectsAlignCenterHorizontalSmart      | button       | TabPageLayout         | 水平居中                           |
| ObjectsAlignRightSmart                 | button       | TabPageLayout         | 右对齐                             |
| ObjectsAlignTopSmart                   | button       | TabPageLayout         | 顶端对齐                           |
| ObjectsAlignMiddleVerticalSmart        | button       | TabPageLayout         | 垂直居中                           |
| ObjectsAlignBottomSmart                | button       | TabPageLayout         | 底端对齐                           |
| AlignDistributeHorizontally            | button       | TabPageLayout         | 横向分布                           |
| AlignDistributeVertically              | button       | TabPageLayout         | 纵向分布                           |
| EqualHeight                            | button       | TabPageLayout         | 等高                               |
| EqualWidth                             | button       | TabPageLayout         | 等宽                               |
| EqualSize                              | button       | TabPageLayout         | 等尺寸                             |
| ObjectsAlignRelativeToContainerSmart   | toggleButton | TabPageLayout         | 相对于页                           |
| ViewGridlinesWord                      | checkBox     | TabPageLayout         | 网格线                             |
| GridSettings                           | button       | TabPageLayout         | 绘图网格                           |
| AlignJustifyMenu                       | toggleButton | TabPageLayout         | 对齐                               |
| ObjectsGroup                           | button       | TabPageLayout         | 组合                               |
| ObjectsUngroup                         | button       | TabPageLayout         | 取消组合                           |
| ObjectsGroupMenu                       | menu         | TabPageLayout         | 组合                               |
| ObjectsGroupMenu                       | menu         | TabPageLayout         | 组合                               |
| ObjectRotateFree                       | button       | TabPageLayout         | 自由旋转                           |
| ObjectRotateLeft90                     | button       | TabPageLayout         | 向左旋转 90°                       |
| ObjectRotateRight90                    | button       | TabPageLayout         | 向右旋转 90°                       |
| FlipHorizontal                         | button       | TabPageLayout         | 水平翻转                           |
| FlipVertical                           | button       | TabPageLayout         | 垂直翻转                           |
| GroupArrange                           | group        | TabPageLayout         | 排列                               |
| ObjectRotateMenu                       | menu         | TabPageLayout         | 旋转                               |
| GroupParagraphLayout                   | group        | TabPageLayout         | 段落                               |
| PageSetupDialog                        | button       | None                  | 页面设置                           |
|                                        | dropDown     | TabParagraph          | 目录级别                           |
|                                        | editBox      | TabParagraph          | 行距                               |
| StylesPane                             | button       | TabParagraph          | 样式和格式                         |
| ParagraphDialog                        | button       | TabParagraph          | 段落                               |
| PageSetupDialog                        | button       | None                  | 页面设置                           |
| Bullets                                | toggleButton | TabParagraph          | 项目符号                           |
| FontSize                               | comboBox     | TabParagraph          | 字号                               |
| FontColorMoreColorsDialogExcel         | button       | TabParagraph          | 其他字体颜色                       |
| FontColorPicker                        | gallery      | TabParagraph          | 字体颜色                           |
| Bold                                   | toggleButton | TabParagraph          | 加粗                               |
| Underline                              | toggleButton | TabParagraph          | 下划线                             |
| AlignLeft                              | toggleButton | TabParagraph          | 左对齐                             |
| AlignCenter                            | toggleButton | TabParagraph          | 居中对齐                           |
| AlignRight                             | toggleButton | TabParagraph          | 右对齐                             |
| FontSizeIncreaseWord                   | button       | None                  | 增大字体                           |
| FontSizeDecreaseWord                   | button       | None                  | 减小字体                           |
|                                        | comboBox     | TabParagraph          | 字体名                             |
| PictureInsertFromFile                  | button       | TabPictureTools       | 本地图片                           |
| GroupPictureTools                      | group        | TabPictureTools       | 插入                               |
| SnapToGrid                             | toggleButton | TabPictureTools       | 增加亮度                           |
| ObjectFillMoreColorsDialog             | button       | TabPictureTools       | 其他填充颜色                       |
| GradientGallery                        | gallery      | TabPictureTools       | 渐变                               |
| MoreTextureOptions                     | button       | TabPictureTools       | 图片或纹理                         |
| ObjectBringToFront                     | button       | TabPictureTools       | 置于顶层                           |
| ObjectSendToBack                       |              | TabPictureTools       | 置于底层                           |
| ObjectSendBehindText                   | button       | TabPictureTools       | 衬于文字下方                       |
| AlignDistributeHorizontally            | button       | TabPictureTools       | 横向分布                           |
| AlignDistributeVertically              | button       | TabPictureTools       | 纵向分布                           |
| ObjectsRegroup                         | button       | TabPictureTools       | 组合                               |
| ObjectsUnGroup                         | button       | TabPictureTools       | 取消组合                           |
| ObjectRotateLeft90                     | button       | TabPictureTools       | 向左旋转 90°                       |
| FontSizeIncrease                       | button       | TabPictureTools       | 略向左移                           |
| FontSizeDecrease                       | button       | TabPictureTools       | 略向右移                           |
| TextWrappingInLineWithText             | toggleButton | TabPictureTools       | 嵌入型                             |
| TextWrappingSquare                     | toggleButton | TabPictureTools       | 四周型环绕                         |
| TextWrappingTight                      | toggleButton | TabPictureTools       | 紧密型环绕                         |
| TextWrappingBehindText                 | toggleButton | TabPictureTools       | 衬于文字下方                       |
| TextWrappingInFrontOfText              | toggleButton | TabPictureTools       | 浮于文字上方                       |
| TextWrappingTopAndBottom               | toggleButton | TabPictureTools       | 上下型环绕                         |
| TextWrappingThrough                    | toggleButton | TabPictureTools       | 穿越型环绕                         |
|                                        | button       | None                  | 重设形状和大小                     |
|                                        | editBox      | TabPictureTools       |                                    |
|                                        | editBox      | TabPictureTools       |                                    |
| ObjectFillMoreColorsDialog             |              | TabPictureTools_Vml   | 其他填充颜色                       |
| ObjectBringToFront                     | button       | TabPictureTools_Vml   | 置于顶层                           |
| ObjectSendBehindText                   |              | TabPictureTools_Vml   | 衬于文字下方                       |
| AlignDistributeHorizontally            | button       | TabPictureTools_Vml   | 横向分布                           |
| AlignDistributeVertically              | button       | TabPictureTools_Vml   | 纵向分布                           |
|                                        | editBox      | TabPictureTools_Vml   | 高度                               |
|                                        | editBox      | TabPictureTools_Vml   | 宽度                               |
| FilePrint                              | button       | None                  | 打印                               |
| FilePrintQuick                         | button       | None                  | 直接打印                           |
|                                        | comboBox     | TabPrintPreview       | 打印机                             |
|                                        | comboBox     | TabPrintPreview       | 纸张类型                           |
| PageOrientationGallery                 | gallery      | TabPrintPreview       | 纸张方向                           |
| PageMarginsGallery                     | gallery      | TabPrintPreview       | 页边距                             |
|                                        | comboBox     | TabPrintPreview       | 方式                               |
| TableOfContentsDialog                  | button       | TabReferences         | 插入目录                           |
| TOCLevelGallery                        | splitButton  | TabReferences         | 目录级别                           |
| TOCLevel1                              | button       | TabReferences         | 1 级目录                           |
| TOCLevel2                              | button       | TabReferences         | 2 级目录                           |
| TOCLevel3                              | button       | TabReferences         | 3 级目录                           |
| TOCLevel4                              | button       | TabReferences         | 4 级目录                           |
| TOCLevel5                              | button       | TabReferences         | 5 级目录                           |
| TOCLevel6                              | button       | TabReferences         | 6 级目录                           |
| TOCLevel7                              | button       | TabReferences         | 7 级目录                           |
| TOCLevel8                              | button       | TabReferences         | 8 级目录                           |
| TOCLevel9                              | button       | TabReferences         | 9 级目录                           |
| TOCBodyText                            | button       | TabReferences         | 普通文本                           |
| TableOfContentsGallery                 | gallery      | TabReferences         | 目录                               |
| TableOfContentsCustom                  | button       | TabReferences         | 自定义目录                         |
| TableOfContentsRemove                  | button       | TabReferences         | 删除目录                           |
| GroupTableOfContents                   | group        | TabReferences         | 目录                               |
| TableOfContentsUpdateClassic           | button       | None                  | 更新目录                           |
| GroupFootnotesAndEndnotes              | group        | TabReferences         | 脚注和尾注                         |
| FootnoteInsert                         | button       | TabReferences         | 插入脚注                           |
| FootnotePreviousWord                   | button       | TabReferences         | 上一条脚注                         |
| FootnoteNextWord                       | button       | TabReferences         | 下一条脚注                         |
| EndnoteInsertWord                      | button       | TabReferences         | 插入尾注                           |
| EndnotePreviousWord                    | button       | TabReferences         | 上一条尾注                         |
| EndnoteNextWord                        | button       | TabReferences         | 下一条尾注                         |
| FootnoteEndnote Separator              | toggleButton | TabReferences         | 脚注/尾注分隔线                    |
| CaptionInsert                          | button       | TabReferences         | 题注                               |
| TableOfFiguresInsert                   | button       | TabReferences         | 插入表目录                         |
| CrossReferenceInsert                   | button       | TabReferences         | 交叉引用                           |
| GroupCaptions                          | group        | TabReferences         | 题注                               |
| GroupMailMerge                         | group        | TabReferences         | 邮件合并                           |
| ShowMailingsContext                    | toggleButton | TabReferences         | 邮件                               |
| IndexMarkEntry                         | button       | TabReferences         | 标记索引项                         |
| IndexInsert                            | button       | TabReferences         | 插入索引                           |
| IndexUpdate                            | button       | TabReferences         | 更新索引                           |
| GroupIndex                             | group        | TabReferences         | 索引                               |
| SpellingMenu                           | splitButton  | TabReview             | 拼写检查                           |
| SpellingAndGrammar                     | button       | None                  | 拼写检查                           |
| WordCount                              | button       | TabReview             | 字数统计                           |
| GroupProofing                          | group        | TabReview             | 校对                               |
| OpenThesaurus                          | button       | None                  | 同义词库                           |
| GroupChineseTranslation                | group        | TabReview             | 中文简繁转换                       |
| TranslateToSimplifiedChinese           | button       | TabReview             | 繁转简                             |
| TranslateToTraditionalChinese          | button       | TabReview             | 简转繁                             |
| ReviewHandwritingComment               | button       | TabReview             | 手写批注                           |
| ReviewDeleteCommentsMenu               | splitButton  | TabReview             | 删除                               |
| ReviewDeleteComment                    | splitButton  | TabReview             | 删除批注                           |
| ReviewDeleteAllCommentsInDocument      | button       | TabReview             | 删除文档中的所有批注               |
| GroupComments                          | group        | TabReview             | 批注                               |
| ReviewNewComment                       | button       | None                  | 插入批注                           |
| ReviewPreviousComment                  | button       | TabReview             | 上一条                             |
| ReviewNextComment                      | button       | TabReview             | 下一条                             |
| ReviewTrackChangesMenu                 | splitButton  | TabReview             | 修订                               |
| ReviewTrackChanges                     | toggleButton | None                  | 修订                               |
| ReviewChangeTrackingOptions            | button       | TabReview             | 修订选项...                        |
| ReviewChangeUserName                   | button       | TabReview             | 更改用户名...                      |
| ReviewShowAllReviewers                 | checkBox     | TabReview             | 所有审阅人                         |
| ReviewShowReviewersMenu                | menu         | TabReview             | 显示标记                           |
| ReviewShowComments                     | toggleButton | TabReview             | 批注                               |
| ReviewShowInsertionsAndDeletions       | toggleButton | TabReview             | 插入和删除                         |
| ReviewShowFormatting                   | toggleButton | TabReview             | 格式设置                           |
| ReviewShowReviewersMenu                | menu         | TabReview             | 审阅人                             |
| ReviewBalloonsMenu                     | menu         | TabReview             | 使用批注框                         |
| ReviewShowRevisionsInBalloons          | toggleButton | TabReview             | 在批注框中显示修订内容             |
| ReviewShowRevisionsInline              | toggleButton | TabReview             | 以嵌入方式显示所有修订             |
| ReviewShowRevisorInformationInBalloons | toggleButton | TabReview             | 在批注框中显示修订者信息           |
| ReviewShowRevisor                      | toggleButton | TabReview             | 显示审阅者                         |
| ReviewShowRevisorImage                 | toggleButton | TabReview             | 显示审阅者头像                     |
| ReviewShowDate                         | toggleButton | TabReview             | 显示日期                           |
| ReviewShowTime                         | toggleButton | TabReview             | 显示时间                           |
| ReviewShowCommentShading               | toggleButton | TabReview             | 显示批注底纹                       |
| ReviewAcceptAllChangesInDocument       | button       | TabReview             | 接受对文档所做的所有修订           |
| ReviewAcceptChangeMenu                 | splitButton  | TabReview             | 接受                               |
| ReviewAcceptChange                     | button       | TabReview             | 接受修订                           |
| ReviewAcceptAllFormatChanges           | button       | TabReview             | 接受所有的格式修订                 |
| ReviewAcceptAllChangesShown            |              | TabReview             | 接受所有显示的修订                 |
| ReviewRejectAllChangesInDocument       | button       | TabReview             | 拒绝对文档所做的所有修订           |
| ReviewRejectChangeMenu                 | splitButton  | TabReview             | 拒绝                               |
| ReviewRejectChange                     | button       | TabReview             | 拒绝所选修订                       |
| ReviewRejectAllFormatChanges           | button       | TabReview             | 拒绝所有的格式修订                 |
| ReviewRejectAllChangesShown            |              | TabReview             | 拒绝所有显示的修订                 |
| ReviewMenu                             | splitButton  | TabReview             | 审阅                               |
| ReviewShowReviewersMenu                | splitButton  | TabReview             | 审阅人                             |
| ReviewShowReviewers                    | button       | TabReview             | 审阅人                             |
| ReviewShowReviewTimeMenu               | splitButton  | TabReview             | 审阅时间                           |
| ReviewShowReviewTime                   | button       | TabReview             | 审阅时间                           |
| ReviewReviewingPaneMenu                | splitButton  | TabReview             | 审阅窗格                           |
| ReviewReviewingPaneVertical            | toggleButton | TabReview             | 垂直审阅窗格                       |
| ReviewReviewingPaneHorizontal          | toggleButton | TabReview             | 水平审阅窗格                       |
| GroupChangesTracking                   | group        | TabReview             | 修订                               |
| ReviewDisplayForReview                 | dropDown     | TabReview             | 显示以审阅                         |
| ReviewTrackChangesMenu                 | splitButton  | TabReview             | 修订                               |
| ReviewPreviousChange                   | button       | TabReview             | 上一条                             |
| ReviewNextChange                       | button       | TabReview             | 下一条                             |
|                                        | checkBox     | TabReview             | 标尺                               |
|                                        | checkBox     | TabReview             | 网格线                             |
|                                        | button       | TabReview             | 单页                               |
|                                        | button       | TabReview             | 多页                               |
| ReviewCompareMenu                      | menu         | TabReview             | 比较                               |
| ReviewCompareTwoVersions               | button       | TabReview             | 比较                               |
| ReviewShowSourceDocumentsMenu          | gallery      | TabReview             | 显示源文档                         |
| ReviewHideSourceDocumentsMenu          | gallery      | TabReview             | 隐藏源文档                         |
| ReviewShowSourceDocumentsOriginal      | button       | TabReview             | 显示原始文档                       |
| ReviewShowSourceDocumentsRevised       | button       | TabReview             | 显示修订的文档                     |
| ReviewShowSourceDocumentsBoth          | button       | TabReview             | 显示原始及修订的文档               |
| GroupCompare                           | group        | TabReview             | 比较                               |
| ReviewRestrictFormatting               | button       | TabReview             | 限制编辑                           |
| ReviewRestrictFormatting               | toggleButton | TabReview             | 文档安全                           |
| PageOrientationGallery                 | gallery      | TabSection            | 纸张方向                           |
| NextPageSectionBreak                   | button       | TabSection            | 下一页分节符                       |
| ContinuousSectionBreak                 | button       | TabSection            | 连续分节符                         |
| EvenPageSectionBreak                   | button       | TabSection            | 偶数页分节符                       |
| OddPageSectionBreak                    | button       | TabSection            | 奇数页分节符                       |
| HeaderAndFooter                        | button       | TabSection            | 页眉页脚                           |
| GoToPreviousSection                    | button       | TabSection            | 上一节                             |
| GoToNextSection                        | button       | TabSection            | 下一节                             |
| GroupCoverAndTOC                       | group        | TabSection            | 封面/目录                          |
| CoverPageInsertGallery                 | gallery      | TabSection            | 封面页                             |
| CoverPageInsertGallery                 | gallery      | TabSection            | 封面页                             |
| TableOfContentsGallery                 | gallery      | TabSection            | 目录页                             |
| TableOfContentsDialog                  | button       | TabSection            | 自定义目录                         |
| TableOfContentsRemove                  | button       | TabSection            | 删除目录                           |
| PageNumber                             | toggleButton | TabSection            | 页码                               |
| NavigationPaneShowHide                 | splitButton  | TabSection            | 导航窗格                           |
| NavigationPaneFind                     | button       | TabSection            | 章节导航                           |
| GroupPageSetupDialog                   | group        | TabSection            | 页面设置                           |
| PageMarginsGallery                     | gallery      | TabSection            | 页边距                             |
| MarginsCustomMargins                   | button       | TabSection            | 自定义页边距                       |
| PageSizeGallery                        | gallery      | TabSection            | 纸张大小                           |
| PageSizeMorePaperSizesDialog           | button       | TabSection            | 其他页面大小                       |
| SectionBreakInsert                     | button       | TabSection            | 新增节                             |
| SectionDelete                          | button       | TabSection            | 删除本节                           |
| GroupHeaderFooterPageNumberInsert      | group        | TabSection            | 页码                               |
| PageNumbers                            | button       | TabSection            | 页码                               |
| PageNumbersRemove                      | button       | TabSection            | 删除页码                           |
| HeaderFooterDifferentFirstPageWord     | checkBox     | TabSection            | 首页不同                           |
|                                        | button       | TabSection            | 奇偶页不同                         |
| GroupHeaderFooter                      | group        | TabSection            | 页眉页脚                           |
| HeaderLinkToPrevious                   | checkBox     | TabSection            | 页眉同前节                         |
| FooterLinkToPrevious                   | checkBox     | TabSection            | 页脚同前节                         |
| PageSetupDialog                        | button       | None                  | 页面设置                           |
| ReviewRestrictFormatting               | button       | TabSecurity           | 限制编辑                           |
| ObjectBringToFront                     | button       | TabSmartArtDesign     | 置于顶层                           |
| ObjectBringForward                     | button       | TabSmartArtDesign     | 上移一层                           |
| ObjectSendToBack                       |              | TabSmartArtDesign     | 置于底层                           |
| ObjectSendBehindText                   | button       | TabSmartArtDesign     | 浮于文字下方                       |
| TextWrappingInLineWithText             | toggleButton | TabSmartArtDesign     | 嵌入型                             |
| TextWrappingSquare                     | toggleButton | TabSmartArtDesign     | 四周型环绕                         |
| TextWrappingTight                      | toggleButton | TabSmartArtDesign     | 紧密型环绕                         |
| TextWrappingBehindText                 | toggleButton | TabSmartArtDesign     | 浮于文字下方                       |
| TextWrappingInFrontOfText              | toggleButton | TabSmartArtDesign     | 浮于文字上方                       |
| TextWrappingTopAndBottom               | toggleButton | TabSmartArtDesign     | 上下型环绕                         |
| TextWrappingThrough                    | toggleButton | TabSmartArtDesign     | 穿越型环绕                         |
| ObjectsAlignLeft                       | button       | TabSmartArtDesign     | 居左对齐                           |
| AlignDistributeHorizontally            | button       | TabSmartArtDesign     | 横向分布                           |
| AlignDistributeVertically              | button       | TabSmartArtDesign     | 纵向分布                           |
| ViewGridlinesWord                      | checkBox     | TabSmartArtDesign     | 网络线                             |
|                                        | editBox      | TabSmartArtDesign     | 高度                               |
|                                        | editBox      | TabSmartArtDesign     | 宽度                               |
|                                        | comboBox     | TabSmartArtFormatTool | 字号                               |
| Bullets                                | toggleButton | TabSmartArtFormatTool | 项目符号                           |
| ObjectFillMoreColorsDialog             |              | TabSmartArtFormatTool | 其他填充颜色                       |
| Bold                                   |              | TabSmartArtFormatTool | 加粗                               |
| Underline                              | toggleButton | TabSmartArtFormatTool | 下划线                             |
| AlignDistributeHorizontally            | button       | TabSmartArtFormatTool | 横向分布                           |
| AlignDistributeVertically              | button       | TabSmartArtFormatTool | 纵向分布                           |
| AlignJustifyMenu                       |              | TabSmartArtFormatTool | 对齐                               |
| ClearFormatting                        | button       | TabSmartArtFormatTool | 清除格式                           |
| Superscript                            | toggleButton | None                  | 上标                               |
| Subscript                              | toggleButton | None                  | 下标                               |
| ShapeStylesGallery                     | gallery      | TabSmartArtFormatTool | 形状风格                           |
| FontColorPicker                        | gallery      | TabStudentTools       |                                    |
| FontColorMoreColorsDialogExcel         | button       | None                  | 其他字体颜色                       |
| PictureInsertFromFile                  | button       | TabStudentTools       |                                    |
| SpellingMenu                           | splitButton  | TabStudentTools       |                                    |
| SpellingAndGrammar                     | button       | None                  | 拼写检查                           |
| ShapesInsertGallery                    | gallery      | TabStudentTools       |                                    |
| ShapesInsertGallery                    | gallery      | TabStudentTools       |                                    |
| InsertCanvas                           | button       | TabStudentTools       |                                    |
| ReviewTrackChangesMenu                 | splitButton  | TabStudentTools       |                                    |
| ReviewTrackChanges                     | toggleButton | None                  | 修订                               |
| WordCount                              | button       | TabStudentTools       |                                    |
|                                        | toggleButton | None                  | 阅读版式                           |
| TextBoxDrawMenu                        | menu         | TabStudentTools       |                                    |
| OpenThesaurus                          | button       | None                  | 同义词库                           |
| InsertScreenGrab                       |              | None                  | 截图取字                           |
| EquationInsertGallery                  |              | None                  | 公式                               |
| SymbolInsertGallery                    | gallery      | None                  | 符号                               |
| TableShowGridlines                     | toggleButton | TabTableTools         | 显示虚框                           |
| TableDrawTable                         | toggleButton | TabTableTools         | 绘制表格                           |
| TableEraser                            | toggleButton | TabTableTools         | 擦除                               |
| CellsDelete                            | button       | TabTableTools         | 单元格                             |
| SplitCells                             | button       | TabTableTools         | 拆分单元格                         |
| TableRowsDistribute                    | button       | TabTableTools         | 平均分布各行                       |
| TableColumnsDistribute                 | button       | TabTableTools         | 平均分布各列                       |
| TablePropertiesDialog                  | button       | TabTableTools         | 表格属性                           |
|                                        | menu         | TabTableTools         | 选择                               |
| FontColorMoreColorsDialogExcel         | button       | TabTableTools         | 其他字体颜色                       |
| FontColorPicker                        | gallery      | TabTableTools         | 字体颜色                           |
| Bold                                   | button       | TabTableTools         | 加粗                               |
| Underline                              | toggleButton | TabTableTools         | 下划线                             |
|                                        | comboBox     | TabTableTools         | 字体名                             |
|                                        | gallery      | None                  | 加粗                               |
|                                        | comboBox     | TabTextTool           | 字号                               |
| ObjectFillMoreColorsDialog             |              | TabTextTool           | 其他填充颜色                       |
| Bold                                   |              | TabTextTool           | 加粗                               |
| AlignDistributeHorizontally            | button       | TabTextTool           | 横向分布                           |
| AlignDistributeVertically              | button       | TabTextTool           | 纵向分布                           |
| AlignJustifyMenu                       |              | TabTextTool           | 对齐                               |
| ClearFormatting                        | button       | TabTextTool           | 清除格式                           |
| FormatPainter                          | toggleButton | None                  | 格式刷                             |
|                                        | comboBox     | TabTextTool           | 字体名                             |
|                                        | gallery      | None                  | 加粗                               |
| Superscript                            | toggleButton | None                  | 上标                               |
| Subscript                              | toggleButton | None                  | 下标                               |
| FontColorPicker                        | gallery      | None                  | 文本颜色                           |
| Bullets                                | toggleButton | TabTextTool           | 项目符号                           |
| ShapeStylesGallery                     | gallery      | TabTextTool           | 形状样式                           |
| ShapeFillMoreGradientsDialog           | button       | TabTextTool           | 形状填充                           |
| ViewPrintLayoutView                    | toggleButton | TabView               | 页面                               |
| ViewFullScreenView                     | button       | TabView               | 全屏显示                           |
| ViewOutlineView                        | toggleButton | TabView               | 大纲                               |
| ViewOutlineView                        | toggleButton | TabView               | Web版式                            |
| ViewFullScreenReadingView              | toggleButton | TabView               | 阅读版式                           |
|                                        | toggleButton | TabView               | 公文版式                           |
| GroupDocumentViews                     | group        | TabView               | 文档视图                           |
| ViewFullScreenView                     | button       | TabView               | 全屏显示                           |
| ViewPrintLayoutView                    | toggleButton | TabView               | 页面                               |
| ProtectEyes                            | toggleButton | TabView               | 护眼模式                           |
| NavigationPanePlaceOnLeft              | toggleButton | TabView               | 靠左                               |
| NavigationPanePlaceOnRight             | toggleButton | TabView               | 靠右                               |
| NavigationPaneInvisible                | toggleButton | TabView               | 隐藏                               |
| NavigationPaneShowHide                 | splitButton  | TabView               | 导航窗格                           |
| GroupViewShowHide                      | group        | TabView               | 显示                               |
| ViewRulerWord                          | checkBox     | TabView               | 标尺                               |
| ViewMarkup                             | checkBox     | TabView               | 标记                               |
| ViewGridlinesWord                      | checkBox     | TabView               | 网格线                             |
| ViewTaskWindow                         | checkBox     | TabView               | 任务窗格                           |
| ViewTableGridlines                     | checkBox     | TabView               | 表格虚框                           |
| ZoomDialog                             | button       | TabView               | 显示比例                           |
| ZoomCurrent100                         | button       | TabView               | 100%                               |
| GroupZoom                              | group        | TabView               | 显示比例                           |
| ZoomPageWidth                          | button       | TabView               | 页宽                               |
| ZoomOnePage                            | button       | TabView               | 单页                               |
| ZoomTwoPages                           | button       | TabView               | 多页                               |
| WindowNew                              | button       | TabView               | 新建窗口                           |
| WindowSideBySideSynchronousScrolling   | toggleButton | TabView               | 同步滚动                           |
| WindowResetPosition                    | toggleButton | TabView               | 重设位置                           |
| WindowSideBySide                       | toggleButton | TabView               | 并排比较                           |
| GroupWindow                            | group        | TabView               | 窗口                               |
| WindowsArrangeAll                      | splitButton  | TabView               | 重排窗口                           |
| WindowsHorizontal                      | toggleButton | TabView               | 水平平铺                           |
| WindowsVertical                        | toggleButton | TabView               | 垂直平铺                           |
| WindowsCascade                         | toggleButton | TabView               | 层叠                               |
| MenuMacros                             | toggleButton | TabView               | JS 宏                              |
| PlayMacro                              | button       | TabView               | JS 宏                              |
| MacroSecurity                          | button       | TabView               | 宏安全性                           |
| VisualBasic                            | button       | TabView               | WPS 宏编辑器                       |
| GroupMacros                            | group        | TabView               | 宏                                 |
| WordArtSpacingVeryTight                | toggleButton | TabWordArt            | 很紧                               |
| WordArtSpacingTight                    | toggleButton | TabWordArt            | 紧密                               |
| WordArtSpacingNormal                   | toggleButton | TabWordArt            | 常规                               |
| WordArtSpacingLoose                    | toggleButton | TabWordArt            | 稀疏                               |
| WordArtSpacingVeryLoose                | toggleButton | TabWordArt            | 很松                               |
| GlowColorMoreColorsDialog              | button       | TabWordArt            | 其他填充颜色                       |
| TextWrappingInLineWithText             | toggleButton | TabWordArt            | 嵌入型                             |
| TextWrappingSquare                     | toggleButton | TabWordArt            | 四周型环绕                         |
| TextWrappingTight                      | toggleButton | TabWordArt            | 紧密型环绕                         |
| TextWrappingBehindText                 | toggleButton | TabWordArt            | 衬于文字下方                       |
| TextWrappingInFrontOfText              | toggleButton | TabWordArt            | 浮于文字上方                       |
| TextWrappingTopAndBottom               | toggleButton | TabWordArt            | 上下型环绕                         |
| TextWrappingThrough                    | toggleButton | TabWordArt            | 穿越型环绕                         |
| ObjectBringToFront                     | button       | TabWordArt            | 置于顶层                           |
| ObjectSendToBack                       |              | TabWordArt            | 置于底层                           |
| ObjectSendBehindText                   | button       | TabWordArt            | 衬于文字下方                       |
| AlignDistributeHorizontally            | button       | TabWordArt            | 横向分布                           |
| AlignDistributeVertically              | button       | TabWordArt            | 纵向分布                           |
| ViewGridlinesWord                      | checkBox     | TabWordArt            | 网格线                             |
| AlignJustifyMenu                       |              | TabWordArt            | 对齐                               |
| ObjectRotateLeft90                     | button       | TabWordArt            | 向左旋转 90°                       |
| FileSaveAsPicture                      |              | TabWorkSpace          | 输出为图片                         |
| ReviewRestrictFormatting               | button       | TabWorkSpace          | 限制编辑                           |
| FileSaveAsPdfOrXps                     | button       | TabWorkSpace          | 输出为PDF                          |
| FileSaveAsPdfOrXps                     | button       | TabWorkSpace          | 输出为PDF                          |
| FileSaveAsPicture                      |              | TabWorkSpace          | 输出为图片                         |
| SwitchFace                             |              | None                  | 皮肤                               |
| TabbarQAndA                            |              | None                  |                                    |
| FileMenu                               |              | None                  |                                    |
| ReviewNewComment                       | button       | None                  | 插入批注                           |
| ReviewTrackChanges                     | toggleButton | None                  | 修订                               |
| AsianLayoutTwoLinesInOne               | button       | None                  | 双行合一                           |
| FitText                                | button       | None                  | 调整宽度                           |
| TabHome                                | tab          | None                  | 开始                               |
| TabInsert                              | tab          | None                  | 插入                               |
|                                        | tab          | None                  | 要素                               |
| TabPageLayoutWord                      | tab          | None                  | 页面布局                           |
| TabReferences                          | tab          | None                  | 引用                               |
| TabMailings                            | tab          | None                  | 邮件                               |
| TabReviewWord                          | tab          | None                  | 审阅                               |
| TabView                                | tab          | None                  | 视图                               |
| TabSection                             | tab          | None                  | 章节                               |
|                                        | tab          | None                  | 文印                               |
| TabOfdPrintPreview                     | tab          | None                  | 预览OFD效果                        |
| TabAddIns                              | tab          | None                  | 加载项                             |
| TabSecurity                            | tab          | None                  | 安全                               |
| TabDeveloper                           | tab          | None                  | 开发工具                           |
| TabWorkSpace                           | tab          | None                  | 云服务                             |
| TabPrintPreview                        | tab          | None                  | 打印预览                           |
| GroupParagraph                         | group        | None                  | 段落布局                           |
| TabTableToolsLayout                    | tab          | None                  | 表格工具                           |
| TabStudentToolsLayout                  | tab          | None                  |                                    |
| TabTableToolsDesign                    | tab          | None                  | 表格样式                           |
| TabOutlining                           | tab          | None                  | 大纲                               |
| TabHeaderAndFooterToolsDesign          | tab          | None                  | 页眉页脚                           |
| TabPictureToolsFormat                  | tab          | None                  | 图片工具                           |
| TabWordArtToolsFormat                  | tab          | None                  | 艺术字                             |
|                                        | tab          | None                  | 签批工具                           |
| FileSaveAsPdfOrXps                     | button       | None                  | 输出为PDF                          |
| QAT_Menu                               |              | None                  | 快速访问工具栏                     |
| FileNewBlankDocument                   | button       | None                  | 新建                               |
| FileOpen                               | button       | None                  | 打开                               |
| FileSave                               | button       | None                  | 保存                               |
| FilePrint                              | button       | None                  | 打印                               |
| FilePrintPreview                       | button       | None                  | 打印                               |
| FilePrintQuick                         | button       | None                  | 直接打印                           |
| FilePrintPreview                       | button       | None                  | 打印预览                           |
| Undo                                   | button       | None                  | 撤消                               |
| Redo                                   | button       | None                  | 恢复                               |
| FormatPainter                          | toggleButton | None                  | 格式刷                             |
| FontSizeIncreaseWord                   | button       | None                  | 增大字体                           |
| FontColorMoreColorsDialogExcel         | button       | None                  | 其他字体颜色                       |
| FontColorPicker                        | gallery      | None                  | 文本颜色                           |
| FontSize                               | comboBox     | None                  | 字号                               |
|                                        | comboBox     | None                  | 横向字号                           |
|                                        | comboBox     | None                  | 纵向字号                           |
| FontSizeDecreaseWord                   | button       | None                  | 减小字体                           |
| FileMenuSendMail                       |              | None                  | 发送邮件                           |
| FileFind                               | button       | None                  | 查找                               |
| ReplaceDialog                          |              | None                  | 替换                               |
| ReviewNewComment                       | button       | None                  | 批注                               |
| ContextMenuText                        | contextMenu  | None                  | 文本                               |
| ContextMenuHyperlink                   | contextMenu  | None                  | 超链接上下文菜单                   |
| ContextMenuComment                     | contextMenu  | None                  | 备注                               |
| ContextMenuList                        | contextMenu  | None                  | 编号上下文菜单                     |
| ContextMenuRevision                    | contextMenu  | None                  | 修订上下文菜单                     |
| ContextMenuSpell                       | contextMenu  | None                  | 拼写检查建议上下文菜单             |
|                                        | contextMenu  | None                  | 文本                               |
|                                        | contextMenu  | None                  | 编号上下文菜单                     |
| SaveSelectionToAutoTextGallery         | button       | None                  | 创建"自动图文集"                   |
| PicturesCompress                       | button       | None                  | 压缩图片                           |
| VisualBasic                            | button       | None                  | WPS 宏编辑器                       |
| MacroPlay                              | button       | None                  | JS 宏                              |
| FieldCodes                             | toggleButton | None                  | 查看域代码                         |
| Subscript                              | toggleButton | None                  | 下标                               |
| SelectAll                              | button       | None                  | 全选                               |
|                                        | toggleButton | None                  | 阅读版式                           |
| ViewFullScreenView                     | button       | None                  | 全屏显示                           |
| Bold                                   | toggleButton | None                  | 加粗                               |
| Copy                                   | button       | None                  | 复制                               |
| FontDialog                             | button       | None                  | 字体                               |
| AlignCenter                            | toggleButton | None                  | 居中对齐                           |
| PageBreakInsertWord                    | button       | None                  | 分页符                             |
| FindDialogExcel                        | button       | None                  | 查找                               |
| GoTo                                   | button       | None                  | 定位                               |
| ReplaceDialog                          | button       | None                  | 替换                               |
| Italic                                 | toggleButton | None                  | 倾斜                               |
| AlignJustify                           | toggleButton | None                  | 两端对齐                           |
| HyperlinkInsert                        | button       | None                  | 超链接                             |
| AlignLeft                              | toggleButton | None                  | 左对齐                             |
| FileNewBlankDocument                   | button       | None                  | 新建                               |
| FileNewBlankDocument                   | button       | None                  | 新建                               |
| FileNewBlankDocument                   | button       | None                  | 新建                               |
| FileNewBlankDocument                   | button       | None                  | 新建                               |
| FileCloseOrCloseAll                    | button       | None                  | 关闭                               |
| PageSetupDialog                        | button       | None                  | 页面设置                           |
| Properties                             | button       | None                  | 属性                               |
| PasteSpecialDialog                     | button       | None                  | 选择性粘贴                         |
| PasteTextOnly                          |              | None                  | 只粘贴文本                         |
| ClearContents                          | button       | None                  | 内容                               |
| AlignRight                             | toggleButton | None                  | 右对齐                             |
| FileSave                               | button       | None                  | 保存                               |
| FontSizeDecrease                       | button       | None                  | 减小字体                           |
| FontSizeIncrease                       | button       | None                  | 增大字体                           |
| Superscript                            | toggleButton | None                  | 上标                               |
| ReviewTrackChanges                     | toggleButton | None                  | 修订                               |
| WordCount                              | button       | None                  | 字数统计                           |
|                                        | button       | None                  | 从左向右文字方向                   |
|                                        | button       | None                  | 从右向左文字方向                   |
| UnderlineGallery                       | gallery      | None                  | 下划线                             |
|                                        | gallery      | None                  | 加粗                               |
| Paste                                  | button       | None                  | 粘贴                               |
| Cut                                    | button       | None                  | 剪切                               |
| Undo                                   | button       | None                  | 撤消                               |
| Help                                   | button       | None                  | WPS 文字 帮助                      |
| FileSaveAs                             | button       | None                  | 另存为                             |
| SpellingAndGrammar                     | button       | None                  | 拼写检查                           |
| CharacterFormattingReset               | button       | None                  | 重设字符格式                       |
| OpenThesaurus                          | button       | None                  | 同义词库                           |
| IndentDecreaseWord                     | button       | None                  | 减少缩进量                         |
| IndentIncreaseWord                     | button       | None                  | 增加缩进量                         |
| FileOfdPrintPreview                    | button       | None                  | 预览OFD效果                        |
| FitText                                | button       | None                  | 调整宽度                           |
| ApplicationExit                        | button       | None                  | 退出                               |
| FileNewFromTemplate                    | button       | None                  | 从更多模板新建                     |
| FileNew                                | button       | None                  | 新建                               |
| TaskPaneRestrict                       | button       | None                  | 限制编辑                           |
| InsertScreenGrab                       |              | None                  | 截图取字                           |
| TableOfContentsUpdateClassic           | button       | None                  | 更新目录                           |
| FileSaveAsMenu                         | menu         | None                  | 另存为                             |
|                                        | button       | None                  | WPS 公文 (\*.wps)                  |
|                                        | button       | None                  | WPS 公文模板 (\*.wpt)              |
| FileSaveAsWps                          | button       | None                  | WPS 文字 文件（\*.wps）            |
| FileSaveAsWpt                          | button       | None                  | WPS 文字 模板文件（\*.wpt）        |
| FileSaveAsWpsx                         | button       | None                  | WPS 文字 2007/2010 文件（\*.wpsx） |
| FileSaveAsDoc                          | button       | None                  | Word 97-2003 文件（\*.doc）        |
| FileSaveAsDot                          | button       | None                  | Word 97-2003 模板文件（\*.dot）    |
| FileSaveAsWordDocx                     | button       | None                  | Word 文件（\*.docx）               |
| FileSaveAsOtherFormats                 | button       | None                  | 其他格式                           |
| FileOfdPrintMenu                       | button       | None                  | 输出为OFD                          |
| FileSaveAsOfd                          | button       | None                  | 输出为OFD                          |
| FileOfdPrintPreview                    | button       | None                  | 预览OFD效果                        |
| FileSaveAsPdfOrXps                     | button       | None                  | 输出为PDF                          |
| FileSaveAsPicture                      | button       | None                  | 输出为图片                         |
| MergeCells                             | button       | None                  | 合并单元格                         |
| NumberingRestart                       | button       | None                  | 重新开始编号                       |
|                                        | button       | None                  | 查找                               |
| SymbolInsertGallery                    | gallery      | None                  | 符号                               |
| SymbolInsertGallery                    | gallery      | None                  | 符号                               |
| SymbolsDialog                          | button       | None                  | 其他符号                           |
| EquationInsertGallery                  |              | None                  | 公式                               |
| EquationInsertNew                      |              | None                  | 插入新公式                         |
|                                        | button       | None                  | 重设形状和大小                     |
| DocumentSplit                          |              | None                  | 文档拆分                           |
| DocumentMerge                          |              | None                  | 文档合并                           |
| DocumentSplitAndMerge                  |              | None                  | 文档拆分合并                       |
| ReviewPreviousComment                  | button       | None                  |                                    |
| ReviewNextComment                      | button       | None                  |                                    |
| Strikethrough                          | toggleButton | None                  | 删除线                             |
| Heading1                               |              | None                  |                                    |
| Heading2                               |              | None                  |                                    |
| Heading3                               |              | None                  |                                    |
| EndnoteInsertWord                      | button       | None                  |                                    |
| FileSaveAsPicture                      |              | None (Context Menu)   | 输出为图片                         |
| FileCloseAll                           | button       | None (Context Menu)   | 关闭所有文档                       |
| FileCloseAll                           | button       | None (Context Menu)   | 关闭所有文档                       |
| FileCloseAll                           | button       | None (Context Menu)   | 关闭所有文档                       |
| SaveAll                                | button       | None (Context Menu)   | 保存所有文档                       |
| FileSaveAsPdfOrXps                     | button       | None (Context Menu)   | 输出为PDF格式                      |
| SaveAsOFD                              | button       | None (Context Menu)   | 输出为OFD格式                      |
| FileMenuSendMail                       | button       | None (Context Menu)   | 发送邮件                           |
| SaveAll                                | button       | None (Context Menu)   | 保存所有文档                       |
| SaveAll                                | button       | None (Context Menu)   | 保存所有文档                       |
| FileConvert                            | button       | None (Context Menu)   | 转换                               |
| FileEncrypt                            | button       | None (Context Menu)   | 文件加密                           |
| ClearFormats                           | button       | None (Context Menu)   | 格式                               |
| ViewPrintLayoutView                    | toggleButton | None (Context Menu)   | 页面                               |
| GroupHeaderFooter                      | group        | None (Context Menu)   | 页眉页脚                           |
| ViewFullScreenView                     | button       | None (Context Menu)   | 全屏显示                           |
| DateAndTimeInsert                      | button       | None (Context Menu)   | 日期和时间                         |
| AutoTextGallery                        | gallery      | None (Context Menu)   | 自动图文集                         |
| PageX                                  | button       | None (Context Menu)   | 第 X 页                            |
| TotalofYPage                           | button       | None (Context Menu)   | 共 Y 页                            |
| PageXofY                               | button       | None (Context Menu)   | 第 X 页 共 Y 页                    |
| SaveSelectionToAutoTextGallery         | button       | None (Context Menu)   | 将所选内容保存到自动图文集库       |
| FieldInsert                            | button       | None (Context Menu)   | 域                                 |
| ReviewNewComment                       | button       | None (Context Menu)   | 批注                               |
| FootnoteEndnoteDialog                  | button       | None (Context Menu)   | 脚注和尾注                         |
| CaptionInsert                          | button       | None (Context Menu)   | 题注                               |
| CrossReferenceInsert                   | button       | None (Context Menu)   | 交叉引用                           |
| IndexInsert                            | button       | None (Context Menu)   | 目录                               |
| PictureInsertFromFile                  | button       | None (Context Menu)   | 来自文件                           |
| InsertCanvas                           | button       | None (Context Menu)   | 插入画布                           |
| ClipArtInsert                          | toggleButton | None (Context Menu)   | 剪贴画                             |
| EquationInsertGallery                  |              | None (Context Menu)   | 公式                               |
| DropCapOptionsDialog                   | button       | None (Context Menu)   | 首字下沉                           |
| TextDirectionOptionsDialog             | button       | None (Context Menu)   | 文字方向                           |
|                                        | button       | None (Context Menu)   | 更改大小写                         |
| AsianLayoutPhoneticGuide               | button       | None (Context Menu)   | 拼音指南                           |
| ReviewRestrictFormatting               | button       | None (Context Menu)   | 限制编辑                           |
| MacroSecurity                          | splitButton  | None (Context Menu)   | 安全性                             |
| AddInManager                           | button       | None (Context Menu)   | 加载项                             |
| ComAddInsDialog                        | button       | None (Context Menu)   | COM 加载项                         |
| FileBackupManagement                   |              | None (Context Menu)   | 备份管理                           |
| ApplicationOptionsDialog               |              | None (Context Menu)   | 选项                               |
| TableDrawTable                         | toggleButton | None (Context Menu)   | 绘制表格                           |
| TableRowsInsertAboveExcel              | button       | None (Context Menu)   | 在上方插入行                       |
| TableRowsInsertBelowExcel              | button       | None (Context Menu)   | 在下方插入行                       |
| InsertCellstMenu                       | splitButton  | None (Context Menu)   | 单元格                             |
| CellsDelete                            | button       | None (Context Menu)   | 单元格                             |
| TableSelect                            | button       | None (Context Menu)   | 表格                               |
| SplitCells                             | button       | None (Context Menu)   | 拆分单元格                         |
| TableAutoFitWindow                     | button       | None (Context Menu)   | 根据窗口调整表格                   |
| TableAutoFitContent                    | button       | None (Context Menu)   | 根据内容调整表格                   |
| TableToolTransposeTable                | button       | None (Context Menu)   | 行列互换                           |
| TableRowsDistribute                    | button       | None (Context Menu)   | 平均分布各行                       |
| TableColumnsDistribute                 | button       | None (Context Menu)   | 平均分布各列                       |
| TableInsertMultidiagonalCell           | button       | None (Context Menu)   | 绘制斜线表头                       |
| TableShowGridlines                     | toggleButton | None (Context Menu)   | 显示虚框                           |
|                                        | button       | None (Context Menu)   | 文本转换成表格                     |
| ConvertTableToText                     | button       | None (Context Menu)   | 表格转换成文本                     |
| TablePropertiesDialog                  | button       | None (Context Menu)   | 表格属性                           |
| WindowNew                              | button       | None (Context Menu)   | 新建窗口                           |
| BlogHomePage                           | button       | None (Context Menu)   | WPS官方网站                        |
| About                                  | button       | None (Context Menu)   | 关于 WPS 文字                      |
| PasteTextOnly                          |              | None (Context Menu)   | 只粘贴文本                         |
| PasteTextOnly                          |              | None (Context Menu)   | 只粘贴文本                         |
| ReviewDeleteCommentsMenu               | splitButton  | None (Context Menu)   | 删除批注                           |
| EditField                              | button       | None (Context Menu)   | 编辑域                             |
| GoToFootnote                           | button       | None (Context Menu)   | 定位至脚注                         |
| CellFillColorPicker                    | button       | None (Context Menu)   | 单元格对齐方式                     |
| GoToEndnote                            | button       | None (Context Menu)   | 定位至尾注                         |
| HangingIndent                          | button       | None (Context Menu)   | 悬挂缩进                           |
| Cancel                                 | button       | None (Context Menu)   | 取消                               |
| ReviewAcceptChange                     | button       | None (Context Menu)   | 接受修订                           |
| ReviewAcceptChange                     | button       | None (Context Menu)   | 拒绝修订                           |
| ReviewAcceptChange                     | button       | None (Context Menu)   | 接受插入                           |
| ReviewRejectChange                     | button       | None (Context Menu)   | 拒绝插入                           |
| ReviewAcceptChange                     | button       | None (Context Menu)   | 接受删除                           |
| ReviewRejectChange                     | button       | None (Context Menu)   | 拒绝删除                           |
| ReviewAcceptChange                     | button       | None (Context Menu)   | 接受格式更改                       |
| ReviewRejectChange                     | button       | None (Context Menu)   | 拒绝格式更改                       |
| ReviewAcceptChange                     | button       | None (Context Menu)   | 接受格式更改                       |
| ReviewRejectChange                     | button       | None (Context Menu)   | 拒绝格式更改                       |
| ReviewAcceptChange                     | button       | None (Context Menu)   | 接受格式更改                       |
| ReviewRejectChange                     | button       | None (Context Menu)   | 拒绝格式更改                       |
| ReviewAcceptChange                     | button       | None (Context Menu)   | 接受格式更改                       |
| ReviewRejectChange                     | button       | None (Context Menu)   | 拒绝格式更改                       |
| ReviewAcceptChange                     | button       | None (Context Menu)   | 接受修订                           |
| ReviewRejectChange                     | button       | None (Context Menu)   | 拒绝修订                           |
| ObjectsUnGroup                         | button       | None (Context Menu)   | 取消组合                           |
| ObjectBringToFront                     | button       | None (Context Menu)   | 置于顶层                           |
| ObjectSendToBack                       | button       | None (Context Menu)   | 置于底层                           |
| ObjectBringForward                     | button       | None (Context Menu)   | 上移一层                           |
| ObjectSendBackward                     | button       | None (Context Menu)   | 下移一层                           |
| ObjectSendBehindText                   | button       | None (Context Menu)   | 衬于文字下方                       |
| ObjectEditPoints                       | toggleButton | None (Context Menu)   | 编辑顶点                           |
| ShapeStraightConnector                 | toggleButton | None (Context Menu)   | 直线连接符                         |
| ShapeElbowConnector                    | toggleButton | None (Context Menu)   | 肘形连接符                         |
| ControlProperties                      | button       | None (Context Menu)   | 属性                               |
| ViewVisualBasicCode                    | button       | None (Context Menu)   | 查看代码                           |
| MacroRecord                            | button       | None (Context Menu)   | 录制新宏                           |
| RealTimeBackupStatus                   |              | None (Context Menu)   | 实时备份                           |
| TextWrappingMenu                       | menu         | None (Context Menu)   | 文字环绕                           |
| Copy                                   | button       | None (Context Menu)   | 复制                               |
| Cut                                    | button       | None (Context Menu)   | 剪切                               |
| PasteTextOnly                          |              | None (Context Menu)   | 只粘贴文本                         |
| ApplicationOptionsDialog               |              | None (Context Menu)   | 设置默认粘贴                       |
| PasteSpecialDialog                     | button       | None (Context Menu)   | 选择性粘贴                         |
| PasteMenu                              | splitButton  | None (Context Menu)   | 粘贴                               |
| AlignLeft                              | toggleButton | None (Context Menu)   | 左对齐                             |
| AlignCenter                            | toggleButton | None (Context Menu)   | 居中                               |
| AlignRight                             | toggleButton | None (Context Menu)   | 右对齐                             |
| AlignJustify                           | toggleButton | None (Context Menu)   | 两端对齐                           |
| AlignCenter                            | toggleButton | None (Context Menu)   | 居中                               |
| Font                                   | comboBox     | None (Context Menu)   | 字体名                             |
| Font                                   | comboBox     | None (Context Menu)   | 字体名                             |
|                                        | comboBox     | None (Context Menu)   | 字体名                             |
|                                        | button       | None (Context Menu)   | 横向字号                           |
|                                        | button       | None (Context Menu)   | 纵向字号                           |
|                                        | button       | None (Context Menu)   | 粘贴                               |
|                                        | button       | None (Context Menu)   | 粘贴                               |
|                                        | button       | None (Context Menu)   | 索引                               |
|                                        | button       | None (Context Menu)   | 减少缩进量                         |
|                                        | button       | None (Context Menu)   | 增加缩进量                         |
| ObjectBringToFrontMenu                 | splitButton  | None (Context Menu)   | 置于顶层                           |
| ObjectSendToBackMenu                   | splitButton  | None (Context Menu)   | 置于底层                           |
| LineNumbersMenu                        | menu         | None (Context Menu)   | 行号                               |
|                                        | menu         | None (Context Menu)   | 文本框                             |
| ObjectsUnGroup                         | button       | None (Context Menu)   | 取消组合                           |
| FontSizeIncreaseWord                   | button       | None (Context Menu)   | 增大字体                           |
| FontSizeDecreaseWord                   | button       | None (Context Menu)   | 縮小字型                           |

> WPS 开发文档： [WPS 基础接口/idMso列表/文字idMso参考](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/idMso%E5%88%97%E8%A1%A8/%E6%96%87%E5%AD%97idMso%E5%8F%82%E8%80%83.html)

------------------------------------------------------------------------
# 演示idMso参考

| 控件名称                              | 控件类型     | 所属Tab                 | 功能描述                              |
|---------------------------------------|--------------|-------------------------|---------------------------------------|
| AnimationGallery                      | gallery      | TabAnimation            | 动画                                  |
| AnimationCustom                       | toggleButton | TabAnimation            | 动画窗格                              |
|                                       | comboBox     | TabAnimation            | 开始播放:                             |
|                                       | editBox      | TabAudioTool            | 淡入                                  |
|                                       | editBox      | TabAudioTool            | 淡出                                  |
|                                       | editBox      | TabAudioTool            | 页停止                                |
|                                       | dropDown     | TabAudioTool            | 自动                                  |
| PowerPointPageSetup                   | button       | TabDesign               | 页面设置                              |
| GroupPageSetup                        | group        | TabDesign               | 幻灯片大小                            |
| Magic                                 | button       | TabDesign               | 魔法                                  |
| ImportTemplates                       | button       | TabDesign               | 导入模板                              |
| InvolvedTemplate                      | button       | TabDesign               | 本文模板                              |
| DesignSlideMaster                     | button       | TabDesign               | 编辑母版                              |
| SlideThemesGallery                    | gallery      | TabDesign               | 设计模板                              |
| OnlineTemplate                        | button       | TabDesign               | 更多设计                              |
| SlideThemesGallery                    | gallery      | TabDesign               | 设计模板                              |
| ResetSlide                            | button       | TabDesign               | 重置                                  |
| ObjectFillMoreColorsDialog            | button       | TabDesignTable          | 其他填充颜色                          |
| GradientGallery                       | gallery      | TabDesignTable          | 渐变                                  |
| MoreTextureOptions                    | button       | TabDesignTable          | 纹理                                  |
|                                       | dropDown     | TabDesignTable          | 笔样式                                |
|                                       | dropDown     | TabDesignTable          | 笔划粗细                              |
| ShapeFillMoreGradientsDialog          | button       | TabDesignTable          | 填充                                  |
| ClearMenu                             | menu         | TabDesignTable          | 清除表格样式                          |
| VisualBasic                           | button       | TabDevelopTools         | WPS 宏编辑器                          |
| AddInsDialog                          | button       | TabDevelopTools         | 加载项                                |
| ComAddInsDialog                       | button       | TabDevelopTools         | COM 加载项                            |
| Cancel                                | button       | TabDevelopTools         | 其他控件                              |
| FindNext                              | button       | TabDevelopTools         | 控件列表                              |
| ControlProperties                     | button       | TabDevelopTools         | 控件属性                              |
| ObjectEditPoints                      | toggleButton | TabDrawingTool          | 编辑顶点                              |
| ObjectEditPoints                      | toggleButton | TabDrawingTool          | 编辑顶点                              |
| ObjectFillMoreColorsDialog            | button       | TabDrawingTool          | 其他填充颜色                          |
| FontColorMoreColorsDialogExcel        | button       | TabDrawingTool          | 其他字体颜色                          |
| GradientGallery                       | gallery      | None                    | 渐变                                  |
| ObjectBringForward                    | button       | TabDrawingTool          | 上移一层                              |
| ObjectSendBackward                    | button       | TabDrawingTool          | 下移一层                              |
| AlignDistributeHorizontally           | button       | TabDrawingTool          | 横向分布                              |
| AlignDistributeVertically             | button       | TabDrawingTool          | 纵向分布                              |
| ObjectsRegroup                        | button       | TabDrawingTool          | 组合                                  |
| ObjectsUnGroup                        | button       | TabDrawingTool          | 取消组合                              |
| ObjectRotateLeft90                    | button       | TabDrawingTool          | 向左旋转 90°                          |
|                                       | editBox      | TabDrawingTool          | 大小                                  |
|                                       | editBox      | TabDrawingTool          |                                       |
| Bold                                  | toggleButton | TabDrawingTool          | 加粗                                  |
| Underline                             | toggleButton | TabDrawingTool          | 下划线                                |
| ObjectEffectShadowGallery             | gallery      | TabDrawingTool          | 阴影                                  |
| OutlineExpand                         | button       | TabDrawingTool          | 上标                                  |
| OutlineCollapse                       | button       | TabDrawingTool          | 下标                                  |
| FontSizeIncrease                      | button       | TabDrawingTool          | 增大字号                              |
| FontSizeDecrease                      | button       | TabDrawingTool          | 减小字号                              |
| Bullets                               | toggleButton | TabDrawingTool          | 项目符号                              |
|                                       | comboBox     | TabDrawingTool          | 字体                                  |
| FontSize                              | comboBox     | TabDrawingTool          | 字号                                  |
| Numbering                             | toggleButton | TabDrawingTool          | 编号                                  |
| IndentDecreaseExcel                   | button       | TabDrawingTool          | 减少缩进量                            |
| IndentIncreaseExcel                   | button       | TabDrawingTool          | 增加缩进量                            |
| ParagraphDistributed                  | toggleButton | TabDrawingTool          | 分散对齐                              |
| WindowMoreWindowsDialog               | button       | TabDrawingTool          | 增大段落行距                          |
| ParagraphMarks                        | button       | TabDrawingTool          | 1.0                                   |
| AlignLeft                             | button       | TabDrawingTool          | 1.5                                   |
| AlignRight                            | button       | TabDrawingTool          | 2.0                                   |
| AlignCenter                           | button       | TabDrawingTool          | 2.5                                   |
| AlignJustify                          | button       | TabDrawingTool          | 3.0                                   |
|                                       | comboBox     | TabDrawingTool_Vml      | 字体                                  |
|                                       | editBox      | TabDrawingTool_Vml      | 高度                                  |
|                                       | editBox      | TabDrawingTool_Vml      | 宽度                                  |
| EquationOptions                       |              | TabEquationTools        | 公式选项                              |
| EquationInsertGallery                 |              | TabEquationTools        | 公式                                  |
| EquationProfessional                  |              | TabEquationTools        | 专业型                                |
| EquationLinearFormat                  |              | TabEquationTools        | 线性                                  |
| EquationNormalText                    |              | TabEquationTools        | 普通文本                              |
| EquationSymbolsInsertGallery          |              | TabEquationTools        | 公式符号                              |
| EquationFractionGallery               |              | TabEquationTools        | 分数                                  |
| EquationScriptGallery                 |              | TabEquationTools        | 上下标                                |
| EquationRadicalGallery                |              | TabEquationTools        | 根式                                  |
| EquationIntegralGallery               |              | TabEquationTools        | 积分                                  |
| EquationLargeOperatorGallery          |              | TabEquationTools        | 大型运算符                            |
| EquationDelimiterGallery              |              | TabEquationTools        | 括号                                  |
| EquationFunctionGallery               |              | TabEquationTools        | 函数                                  |
| EquationAccentGallery                 |              | TabEquationTools        | 导数符号                              |
| EquationLimitGallery                  |              | TabEquationTools        | 极限和对数                            |
| EquationOperatorGallery               |              | TabEquationTools        | 运算符                                |
| EquationMatrixGallery                 |              | TabEquationTools        | 矩阵                                  |
| FileSaveAsPicture                     |              | TabFile                 | 输出为图片                            |
| FileNewMenu                           | menu         | TabFile                 | 新建                                  |
| FileNew                               | button       | TabFile                 | 新建                                  |
| FileNewFromTemplate                   | button       | TabFile                 | 本机上的模板                          |
| FileNewFromDefaultTemplate            | button       | TabFile                 | 从默认模板新建                        |
| FileSaveAsMenu                        | menu         | TabFile                 | 另存为                                |
| FileSaveAsDps                         | button       | TabFile                 | WPS 演示 文件（\*.dps）               |
| FileSaveAsDpt                         | button       | TabFile                 | WPS 演示 模板文件（\*.dpt）           |
| FileSaveAsPpt                         | button       | TabFile                 | PowerPoint 97-2003 文件（\*.ppt）     |
| FileSaveAsPot                         | button       | TabFile                 | PowerPoint 97-2003 模板文件（\*.pot） |
| FileSaveAsPps                         | button       | TabFile                 | PowerPoint 97-2003 放映文件（\*.pps） |
| FileSaveAsOtherFormats                | button       | TabFile                 | 其他格式                              |
| FileMenuPackageMenu                   | menu         | TabFile                 | 文件打包                              |
| FilePackageIntoFolder                 | menu         | TabFile                 | 将演示文档打包成文件夹                |
| FilePackageIntoZip                    | menu         | TabFile                 | 将演示文档打包成压缩文件              |
| FilePrintMenu                         | splitButton  | TabFile                 | 打印                                  |
| FilePrint                             | button       | TabFile                 | 打印                                  |
| FilePrintPreview                      | button       | TabFile                 | 打印预览                              |
|                                       | splitButton  | TabFile                 |                                       |
|                                       | button       | TabFile                 |                                       |
|                                       | button       | TabFile                 |                                       |
| FileInfoMenu                          | menu         | TabFile                 | 文档加密                              |
| FileProperties                        | button       | TabFile                 | 属性                                  |
|                                       | splitButton  | TabFile                 | 备份与恢复                            |
|                                       | splitButton  | TabFile                 | 备份与恢复                            |
| FileBackupManagement                  | menu         | TabFile                 | 备份管理                              |
| FileBackupManagement                  | button       | TabFile                 | 备份管理                              |
| FileBackupHistory                     | button       | TabFile                 | 历史版本                              |
| FileHelp                              | menu         | TabFile                 | 帮助                                  |
| Help                                  | button       | TabFile                 | WPS 演示 帮助                         |
| About                                 | button       | TabFile                 | 关于 WPS 演示                         |
| FileSave                              | button       | TabFile                 | 保存                                  |
| FileOpen                              | button       | TabFile                 | 打开                                  |
| FileSaveAsPdfOrXps                    | button       | None                    | 输出为PDF                             |
| FileShare                             | button       | TabFile                 | 分享文档                              |
| FileMenuSendMail                      | button       | TabFile                 | 发送邮件                              |
| ApplicationOptionsDialog              |              | TabFile                 | 选项                                  |
| ObjectFillMoreColorsDialog            | button       | TabGraphicTool          | 其他填充颜色                          |
| GradientGallery                       | gallery      | None                    | 渐变                                  |
| ObjectBringForward                    | button       | TabGraphicTool          | 上移一层                              |
| ObjectSendBackward                    | button       | TabGraphicTool          | 下移一层                              |
| AlignDistributeHorizontally           | button       | TabGraphicTool          | 横向分布                              |
| AlignDistributeVertically             | button       | TabGraphicTool          | 纵向分布                              |
| ObjectsRegroup                        | button       | TabGraphicTool          | 组合                                  |
| ObjectsUnGroup                        | button       | TabGraphicTool          | 取消组合                              |
| ObjectRotateLeft90                    | button       | TabGraphicTool          | 向左旋转 90°                          |
|                                       | button       | None                    | 重设形状和大小                        |
|                                       | editBox      | TabGraphicTool          | 高度                                  |
|                                       | editBox      | TabGraphicTool          | 宽度                                  |
| Bold                                  | toggleButton | TabGraphicTool          | 加粗                                  |
| Underline                             | toggleButton | TabGraphicTool          | 下划线                                |
| ObjectEffectShadowGallery             | gallery      | TabGraphicTool          | 阴影                                  |
| FontColorMoreColorsDialogExcel        | button       | TabGraphicTool          | 其他字体颜色                          |
| OutlineExpand                         | button       | TabGraphicTool          | 上标                                  |
| OutlineCollapse                       | button       | TabGraphicTool          | 下标                                  |
| FontSizeIncrease                      | button       | TabGraphicTool          | 增大字号                              |
| FontSizeDecrease                      | button       | TabGraphicTool          | 减小字号                              |
| Bullets                               | toggleButton | TabGraphicTool          | 项目符号                              |
|                                       | comboBox     | TabGraphicTool          | 字体                                  |
| FontSize                              | comboBox     | TabGraphicTool          | 字号                                  |
| Numbering                             | toggleButton | TabGraphicTool          | 编号                                  |
| IndentDecreaseExcel                   | button       | TabGraphicTool          | 减少缩进量                            |
| IndentIncreaseExcel                   | button       | TabGraphicTool          | 增加缩进量                            |
| ParagraphDistributed                  | toggleButton | TabGraphicTool          | 分散对齐                              |
| WindowMoreWindowsDialog               | button       | TabGraphicTool          | 增大段落行距                          |
| ParagraphMarks                        | button       | TabGraphicTool          | 1.0                                   |
| AlignLeft                             | button       | TabGraphicTool          | 1.5                                   |
| AlignRight                            | button       | TabGraphicTool          | 2.0                                   |
| AlignCenter                           | button       | TabGraphicTool          | 2.5                                   |
| AlignJustify                          | button       | TabGraphicTool          | 3.0                                   |
| PasteSourceFormatting                 |              | TabHome                 | 带格式粘贴                            |
| PasteTextOnly                         |              | TabHome                 | 只粘贴文本                            |
| PasteSpecialDialog                    |              | TabHome                 | 选择性粘贴                            |
| Cut                                   |              | TabHome                 | 剪切                                  |
| Copy                                  |              | TabHome                 | 复制                                  |
| FormatPainter                         |              | TabHome                 | 格式刷                                |
| ShowClipboard                         | button       | TabHome                 | 剪贴板                                |
| Paste                                 | button       | TabHome                 | 粘贴                                  |
| ViewSlideShowView                     |              | TabHome                 | 从头开始                              |
| SlideShowFromBeginning                |              | TabHome                 | 从头开始                              |
| SlideShowFromCurrent                  |              | TabHome                 | 当页开始                              |
| SlideNewGallery                       | gallery      | TabHome                 | 新建幻灯片                            |
| Bold                                  | toggleButton | TabHome                 | 加粗                                  |
| Italic                                | toggleButton | TabHome                 | 倾斜                                  |
| Underline                             | toggleButton | TabHome                 | 下划线                                |
| ObjectEffectShadowGallery             | gallery      | TabHome                 | 阴影                                  |
| TextHighlightColorPicker              | gallery      | TabHome                 | 突出显示                              |
| FontColorMoreColorsDialogExcel        | button       | TabHome                 | 其他字体颜色                          |
| OutlineExpand                         | button       | TabHome                 | 上标                                  |
| OutlineCollapse                       | button       | TabHome                 | 下标                                  |
| FontSizeIncrease                      | button       | TabHome                 | 增大字号                              |
| FontSizeDecrease                      | button       | TabHome                 | 减小字号                              |
| ClearFormats                          | button       | TabHome                 | 清除所有格式                          |
|                                       | comboBox     | None                    | 字体                                  |
| FontSize                              | comboBox     | TabHome                 | 字号                                  |
| Bullets                               | toggleButton | TabHome                 | 项目符号                              |
| IndentDecreaseExcel                   | button       | TabHome                 | 减少缩进量                            |
| IndentIncreaseExcel                   | button       | TabHome                 | 增加缩进量                            |
| ParagraphMarks                        | button       | TabHome                 | 110                                   |
| AlignLeft                             | button       | TabHome                 |                                       |
| AlignRight                            | button       | TabHome                 |                                       |
| AlignCenter                           | button       | TabHome                 |                                       |
| AlignJustify                          | button       | TabHome                 |                                       |
| WindowMoreWindowsDialog               | button       | None                    | 增大段落行距                          |
| FilePackageIntoFolder                 | button       | TabHome                 | 打包成文件夹                          |
| ObjectBringToFront                    | button       | TabHome                 | 置于顶层                              |
| ObjectSendToBack                      | button       | TabHome                 | 置于底层                              |
| ObjectBringForward                    | button       | TabHome                 | 上移一层                              |
| ObjectSendBackward                    | button       | TabHome                 | 下移一层                              |
| ObjectUnGroup                         | button       | TabHome                 | 取消组合                              |
| AlignDistributeHorizontally           | button       | TabHome                 | 横向分布                              |
| AlignDistributeVertically             | button       | TabHome                 | 纵向分布                              |
| ObjectRotateLeft90                    | button       | TabHome                 | 向左旋转 90°                          |
| ObjectFillMoreColorsDialog            | button       | TabHome                 | 其他填充颜色                          |
| ArrowStyleGallery                     | gallery      | TabHome                 | 箭头样式                              |
| ShapeScribble                         | button       | TabHome                 | 无三维效果                            |
| ShapesInsertGallery                   |              | TabHome                 | 形状                                  |
| PictureInsertFromFilePowerPoint       |              | TabHome                 | 图片                                  |
| ObjectsArrangeMenu                    | menu         | TabHome                 | 排列                                  |
| ObjectsRegroup                        | button       | None                    | 组合                                  |
| ShapeFillColorPicker                  | gallery      | TabHome                 | 填充                                  |
| GradientGallery                       | gallery      | None                    | 渐变                                  |
| OutlineWeightGallery                  | gallery      | TabHome                 | 轮廓                                  |
| ReplaceMenu                           | splitButton  | TabHome                 | 替换                                  |
| PresentationTool                      | menu         | TabHome                 | 演示工具                              |
| TableInsert                           | button       | TabInsert               | 插入表格                              |
| SlideNewGallery                       | gallery      | TabInsert               | 新建幻灯片                            |
| PictureInsertFromFilePowerPoint       |              | TabInsert               | 本地图片                              |
| TableInsertGallery                    | gallery      | TabInsert               | 表格                                  |
| ShapesInsertGallery                   |              | TabInsert               | 形状                                  |
| ShapesInsertGallery                   |              | TabInsert               | 形状                                  |
| HeaderFooterInsert                    | button       | TabInsert               | 页眉页脚                              |
| NumberInsert                          | button       | TabInsert               | 艺术字                                |
| OleObjectctInsert                     | button       | TabInsert               | 对象                                  |
| TextAlignLeft                         | button       | TabInsert               | 幻灯片编号                            |
| TextAlignLeft                         | button       | TabInsert               | 幻灯片编号                            |
| DateAndTimeInsert                     | button       | TabInsert               | 日期和时间                            |
| EquationInsertGallery                 |              | TabInsert               | 公式                                  |
| EquationInsertNew                     |              | TabInsert               | 插入新公式                            |
| FilePackageIntoFolder                 | button       | TabInsert               | 打包成文件夹                          |
| SoundInsertMenu                       | splitButton  | TabInsert               | 音频                                  |
| MovieInsert                           | splitButton  | TabInsert               | 视频                                  |
| ObjectFillMoreColorsDialog            | button       | TabOrgChart             | 其他填充颜色                          |
| GradientGallery                       | gallery      | TabOrgChart             | 渐变                                  |
| MoreTextureOptions                    | button       | TabOrgChart             | 纹理                                  |
| ArrowStyleGallery                     | gallery      | TabOrgChart             | 箭头样式                              |
| FontColorMoreColorsDialogExcel        | button       | TabOrgChart             | 其他字体颜色                          |
| ShapeFillColorPicker                  | gallery      | TabOrgChart             | 形状填充                              |
| GroupPictureTools                     | group        | TabPictureTool          | 插入                                  |
| SnapToGrid                            | toggleButton | TabPictureTool          | 增加亮度                              |
| ObjectFillMoreColorsDialog            | button       | TabPictureTool          | 其他填充颜色                          |
| GradientGallery                       | gallery      | None                    | 渐变                                  |
| ObjectBringForward                    | button       | TabPictureTool          | 上移一层                              |
| ObjectSendBackward                    | button       | TabPictureTool          | 下移一层                              |
| AlignDistributeHorizontally           | button       | TabPictureTool          | 横向分布                              |
| AlignDistributeVertically             | button       | TabPictureTool          | 纵向分布                              |
| ObjectsRegroup                        | button       | TabPictureTool          | 组合                                  |
| ObjectsUnGroup                        | button       | TabPictureTool          | 取消组合                              |
| ObjectRotateLeft90                    | button       | TabPictureTool          | 向左旋转 90°                          |
| FontSizeIncrease                      | button       | TabPictureTool          | 略向左移                              |
| FontSizeDecrease                      | button       | TabPictureTool          | 略向右移                              |
|                                       | button       | None                    | 重设形状和大小                        |
|                                       | editBox      | TabPictureTool          |                                       |
|                                       | editBox      | TabPictureTool          |                                       |
| FilePrint                             | button       | None                    | 打印                                  |
|                                       | comboBox     | TabPrintPreview         | 打印机                                |
|                                       | comboBox     | TabPrintPreview         | 纸张类型                              |
|                                       | comboBox     | TabPrintPreview         | 份数                                  |
|                                       | comboBox     | TabPrintPreview         | 顺序                                  |
|                                       | comboBox     | TabPrintPreview         | 方式                                  |
|                                       | comboBox     | TabPrintPreview         | 缩放比例:                             |
|                                       | button       | None                    | 同义词库                              |
| ReviewNewComment                      | button       | TabReview               | 插入批注                              |
| ReviewPreviousComment                 | button       | TabReview               | 上一条                                |
| ReviewNextComment                     | button       | TabReview               | 下一条                                |
| TranslateToSimplifiedChinese          | button       | TabReview               | 繁转简                                |
| ShapeScribble                         | button       | TabshadowDrawingTools   | 无三维效果                            |
| Bold                                  | toggleButton | TabSlideMaster          | 加粗                                  |
| Underline                             | toggleButton | TabSlideMaster          | 下划线                                |
| ObjectEffectShadowGallery             | gallery      | TabSlideMaster          | 阴影                                  |
| FontColorMoreColorsDialogExcel        | button       | TabSlideMaster          | 其他字体颜色                          |
| Bullets                               | toggleButton | TabSlideMaster          | 项目符号                              |
| BulletsAndNumberingBulletsDialog      |              | TabSlideMaster          | 其他项目符号                          |
|                                       | comboBox     | TabSlideMaster          | 字体                                  |
|                                       | comboBox     | TabSlideMaster          | 字号                                  |
| ViewSlideShowView                     |              | TabSlideShow            | 从头开始                              |
| SlideShowFromBeginning                |              | TabSlideShow            | 从头开始                              |
| SlideShowFromCurrent                  |              | None                    | 当页开始                              |
| SlideShowCustom                       |              | TabSlideShow            | 自定义放映                            |
| SlideShowRehearseTimings              |              | TabSlideShow            | 排练全部                              |
|                                       | comboBox     | TabSlideShow            | 放映位置                              |
| AnimationTransitionGallery            | gallery      | TabSlideTrans           | 切换                                  |
| GroupTransitionStyles                 | group        | TabSlideTrans           | 效果选项                              |
| ObjectBringForward                    | button       | TabSmartArtDesign       | 上移一层                              |
| ObjectSendBackward                    | button       | TabSmartArtDesign       | 下移一层                              |
| Bold                                  | toggleButton | TabSmartArtFormatTool   | 加粗                                  |
| Italic                                | toggleButton | TabSmartArtFormatTool   | 倾斜                                  |
| Underline                             | toggleButton | TabSmartArtFormatTool   | 下划线                                |
| ObjectEffectShadowGallery             | gallery      | TabSmartArtFormatTool   | 阴影                                  |
| FontColorMoreColorsDialogExcel        | button       | TabSmartArtFormatTool   | 其他字体颜色                          |
| OutlineExpand                         | button       | TabSmartArtFormatTool   | 上标                                  |
| OutlineCollapse                       | button       | TabSmartArtFormatTool   | 下标                                  |
| FontSizeIncrease                      | button       | TabSmartArtFormatTool   | 增大字号                              |
| FontSizeDecrease                      | button       | TabSmartArtFormatTool   | 减小字号                              |
| ClearFormats                          | button       | TabSmartArtFormatTool   | 清除所有格式                          |
|                                       | comboBox     | None                    | 字体                                  |
| FontSize                              | comboBox     | TabSmartArtFormatTool   | 字号                                  |
| ObjectFillMoreColorsDialog            | button       | TabSmartArtFormatTool   | 其他填充颜色                          |
| PictureInsertFromFilePowerPoint       |              | TabStudentTools         |                                       |
| ShapesInsertGallery                   |              | TabStudentTools         |                                       |
| ShapesInsertGallery                   |              | TabStudentTools         |                                       |
| EquationInsertGallery                 |              | None                    | 公式                                  |
| ReviewNewComment                      | button       | TabStudentTools         |                                       |
| MergeCells                            | button       | TabTableTool            | 合并单元格                            |
| SplitCells                            | button       | TabTableTool            | 拆分单元格                            |
| GroupTableAlignment                   | group        | TabTableTool            | 对齐方式                              |
| ObjectBringForward                    | button       | TabTableTool            | 上移一层                              |
| ObjectBringToFront                    | button       | TabTableTool            | 置于顶层                              |
| ObjectSendBackward                    | button       | TabTableTool            | 下移一层                              |
| ObjectSendToBack                      | button       | TabTableTool            | 置于底层                              |
| AlignDistributeHorizontally           | button       | TabTableTool            | 横向分布                              |
| AlignDistributeVertically             | button       | TabTableTool            | 纵向分布                              |
| Underline                             | toggleButton | TabTableTool            | 下划线                                |
| ObjectEffectShadowGallery             | gallery      | TabTableTool            | 阴影                                  |
| TextHighlightColorPicker              | gallery      | TabTableTool            | 突出显示                              |
| FontColorMoreColorsDialogExcel        | button       | TabTableTool            | 其他字体颜色                          |
| OutlineExpand                         | button       | TabTableTool            | 上标                                  |
| OutlineCollapse                       | button       | TabTableTool            | 下标                                  |
| FontSizeIncrease                      | button       | TabTableTool            | 增大字号                              |
| FontSizeDecrease                      | button       | TabTableTool            | 减小字号                              |
| WindowMoreWindowsDialog               | button       | TabTableTool            | 增大段落行距                          |
|                                       | comboBox     | TabTableTool            | 字体                                  |
|                                       | comboBox     | TabTableTool            | 字号                                  |
| Bold                                  | toggleButton | TabTextTool             | 加粗                                  |
| Italic                                | toggleButton | TabTextTool             | 倾斜                                  |
| Underline                             | toggleButton | TabTextTool             | 下划线                                |
| ObjectEffectShadowGallery             | gallery      | TabTextTool             | 阴影                                  |
| TextHighlightColorPicker              | gallery      | TabTextTool             | 突出显示                              |
| FontSizeIncrease                      | button       | TabTextTool             | 增大字号                              |
| FontSizeDecrease                      | button       | TabTextTool             | 减小字号                              |
| ClearFormats                          | button       | TabTextTool             | 清除所有格式                          |
|                                       | comboBox     | TabTextTool             | 字体                                  |
|                                       | comboBox     | TabTextTool             | 字号                                  |
| Bullets                               | toggleButton | TabTextTool             | 项目符号                              |
|                                       | button       | TabTextTool             | 清除艺术字                            |
| GradientGallery                       | gallery      | None                    | 渐变                                  |
|                                       | dropDown     | TabVideoTool            | 自动                                  |
| WindowNew                             | button       | TabView                 | 新建窗口                              |
| GroupMacros                           | group        | TabView                 | JS 宏                                 |
| VisualBasic                           | button       | TabView                 | WPS 宏编辑器                          |
| FontSizeIncrease                      | button       | TabWAshadowDrawingTools | 略向左移                              |
| FontSizeDecrease                      | button       | TabWAshadowDrawingTools | 略向右移                              |
| ShapeScribble                         | button       | TabWAshadowDrawingTools | 无三维效果                            |
| GlowColorMoreColorsDialog             | button       | TabWordArt              | 其他填充颜色                          |
| GradientGallery                       | gallery      | TabWordArt              | 渐变                                  |
| MoreTextureOptions                    | button       | TabWordArt              | 纹理                                  |
| ObjectBringForward                    | button       | TabWordArt              | 上移一层                              |
| ObjectSendBackward                    | button       | TabWordArt              | 下移一层                              |
| AlignDistributeHorizontally           | button       | TabWordArt              | 横向分布                              |
| AlignDistributeVertically             | button       | TabWordArt              | 纵向分布                              |
| ObjectsRegroup                        | button       | TabWordArt              | 组合                                  |
| ObjectsUnGroup                        | button       | TabWordArt              | 取消组合                              |
| ObjectRotateLeft90                    | button       | TabWordArt              | 向左旋转 90°                          |
| FileSaveAsPicture                     |              | TabWorkSpace            | 输出为图片                            |
| ExportToPDF                           | button       | TabWorkSpace            | 输出为PDF                             |
| FileSaveAsPicture                     |              | TabWorkSpace            | 输出为图片                            |
| ExportToPDF                           | button       | None                    | 输出为PDF                             |
| SetLanguage                           | button       | None                    | 更改语言                              |
| SwitchFace                            |              | None                    | 皮肤                                  |
| TabbarQAndA                           |              | None                    |                                       |
| FileMenu                              |              | None                    |                                       |
| TabHome                               | tab          | None                    | 开始                                  |
| TabInsert                             | tab          | None                    | 插入                                  |
| TabDesign                             |              | None                    | 设计                                  |
| TabSecurity                           | tab          | None                    | 安全                                  |
| TabWorkSpace                          | tab          | None                    | 云服务                                |
| TabStudentToolsLayout                 | tab          | None                    |                                       |
| TabAddIns                             |              | None                    | 加载项                                |
| ExportToPDF                           | button       | None                    | 输出为PDF                             |
| QAT_Menu                              |              | None                    | 快速访问菜单                          |
| FileNewBlankDocument                  | button       | None                    | 新建空白文档                          |
| FileOpen                              | button       | None                    | 打开                                  |
| FileSave                              | button       | None                    | 保存                                  |
| FilePrint                             | button       | None                    | 打印                                  |
| FilePrintPreview                      | button       | None                    | 打印                                  |
| FilePrintQuick                        | button       | None                    | 直接打印                              |
| FilePrintPreview                      | button       | None                    | 打印预览                              |
| Redo                                  | gallery      | None                    | 恢复                                  |
| FormatPainter                         |              | None                    | 格式刷                                |
|                                       | comboBox     | None                    | 字体                                  |
| WindowMoreWindowsDialog               | button       | None                    | 增大段落行距                          |
| ObjectsRegroup                        | button       | None                    | 组合                                  |
| FileMenuSendMail                      |              | None                    | 发送邮件                              |
| ContextMenuShape                      | contextMenu  | None                    | 形状                                  |
| ContextMenuTextEdit                   | contextMenu  | None                    | 文字编辑                              |
| PicturesCompress                      | button       | None                    | 压缩图片                              |
| VisualBasic                           | button       | None                    | WPS 宏编辑器                          |
| SelectAll                             | button       | None                    | 全选                                  |
| Bold                                  | toggleButton | None                    | 加粗                                  |
| Copy                                  | button       | None                    | 复制                                  |
| FindDialogExcel                       | button       | None                    | 查找                                  |
| ReplaceDialog                         | button       | None                    | 替换                                  |
| Italic                                | toggleButton | None                    | 倾斜                                  |
| FileNewBlankDocument                  | button       | None                    | 新建空白文档                          |
| FileNewBlankDocument                  | button       | None                    | 新建空白文档                          |
|                                       | button       | None                    | 新建空白文档                          |
| FileNew                               | button       | None                    | 新建空白文档                          |
| PageSetupDialog                       | button       | None                    | 页面设置                              |
| Properties                            | button       | None                    | 属性                                  |
| FileSave                              | button       | None                    | 保存                                  |
| FontSizeDecrease                      | button       | None                    | 减小字号                              |
| FontSizeIncrease                      | button       | None                    | 增大字号                              |
| Underline                             | toggleButton | None                    | 下划线                                |
| Paste                                 | button       | None                    | 粘贴                                  |
| Cut                                   | button       | None                    | 剪切                                  |
| FileSaveAs                            | button       | None                    | 另存为                                |
| ViewSlideShowView                     |              | None                    | 幻灯片放映                            |
|                                       | button       | None                    | 同义词库                              |
| IndentDecreaseExcel                   | button       | None                    | 减少缩进量                            |
| IndentIncreaseExcel                   | button       | None                    | 增加缩进量                            |
| SlideShowFromCurrent                  |              | None                    | 当页开始                              |
| ApplicationExit                       | button       | None                    | 退出                                  |
| BorderBottomNoToggle                  | button       | None                    | 下框线                                |
| BorderTop                             | toggleButton | None                    | 上框线                                |
| BorderLeft                            | toggleButton | None                    | 左框线                                |
| BorderRight                           | toggleButton | None                    | 右框线                                |
| BordersAll                            | button       | None                    | 所有框线                              |
| BorderOutside                         | button       | None                    | 外侧框线                              |
| ThemeBrowseForThemesPowerPoint        |              | None                    | 导入模板                              |
| SlideShowFromBeginning                |              | None                    | 从头开始                              |
| FileNewFromTemplate                   | button       | None                    | 本机上的模板                          |
| GradientGallery                       | gallery      | None                    | 渐变                                  |
| FileSaveAsMenu                        | menu         | None                    | 另存为                                |
| FileSaveAsDps                         | button       | None                    | WPS 演示 文件（\*.dps）               |
| FileSaveAsPpt                         | button       | None                    | PowerPoint 97-2003 文件（\*.ppt）     |
| FileSaveAsPot                         | button       | None                    | PowerPoint 97-2003 模板文件（\*.pot） |
| FileSaveAsPps                         | button       | None                    | PowerPoint 97-2003 放映文件（\*.pps） |
| FileSaveAsOtherFormats                | button       | None                    | 其他格式                              |
| FileSaveAsPdfOrXps                    | button       | None                    | 输出为PDF                             |
| EquationInsertGallery                 |              | None                    | 公式                                  |
| EquationInsertNew                     |              | None                    | 插入新公式                            |
|                                       | button       | None                    | 重设形状和大小                        |
| PresentationToolRibbon                |              | None                    |                                       |
| FileSaveAsPicture                     |              | None (Context Menu)     | 输出为图片                            |
| FileCloseOrCloseAll                   | button       | None (Context Menu)     | 关闭                                  |
| FileCloseAll                          | button       | None (Context Menu)     | 关闭所有文档                          |
| FileCloseAll                          | button       | None (Context Menu)     | 关闭所有文档                          |
| FileCloseAll                          | button       | None (Context Menu)     | 关闭所有文档                          |
| SaveAll                               | button       | None (Context Menu)     | 保存所有文档                          |
| FileExportToPDF                       | button       | None (Context Menu)     | 输出为PDF格式                         |
| FileSaveAsOFD                         | button       | None (Context Menu)     | 输出为OFD格式                         |
| SaveAll                               | button       | None (Context Menu)     | 保存所有文档                          |
| SaveAll                               | button       | None (Context Menu)     | 保存所有文档                          |
| FilePackageIntoFolder                 | button       | None (Context Menu)     | 打包成文件夹                          |
| FilePackageIntoZip                    | button       | None (Context Menu)     | 将演示文档打包成压缩文件              |
| FileMenuSendMail                      | button       | None (Context Menu)     | 发送邮件                              |
| FileEncrypt                           | button       | None (Context Menu)     | 文件加密                              |
| PasteTextOnly                         |              | None (Context Menu)     | 只粘贴文本                            |
| ClipArtInsert                         | toggleButton | None (Context Menu)     | 剪贴画                                |
| DataTable                             | group        | None (Context Menu)     | 表格                                  |
| MacroSecurity                         | splitButton  | None (Context Menu)     | 安全性                                |
| FileBackupManagement                  |              | None (Context Menu)     | 备份管理                              |
| ApplicationOptionsDialog              |              | None (Context Menu)     | 选项                                  |
| SetLanguage                           | button       | None (Context Menu)     | 更改语言                              |
| ViewSlideShowView                     |              | None (Context Menu)     | 观看放映                              |
| SlideShowRehearseTimings              |              | None (Context Menu)     | 排练计时                              |
| SlideShowCustom                       |              | None (Context Menu)     | 自定义放映                            |
| BlogHomePage                          | button       | None (Context Menu)     | WPS官方网站                           |
| ObjectsRegroup                        | button       | None (Context Menu)     | 组合                                  |
| ObjectsUngroup                        | button       | None (Context Menu)     | 取消组合                              |
| ObjectBringToFront                    | button       | None (Context Menu)     | 置于顶层                              |
| ObjectSendToBack                      | button       | None (Context Menu)     | 置于底层                              |
| ObjectBringForward                    | button       | None (Context Menu)     | 上移一层                              |
| ObjectSendBackward                    | button       | None (Context Menu)     | 下移一层                              |
| ObjectSetShapeDefaults                | button       | None (Context Menu)     | 设为默认形状样式                      |
| ObjectEditPoints                      | toggleButton | None (Context Menu)     | 编辑顶点                              |
| SheetColumnsDelete                    | button       | None (Context Menu)     | 删除列                                |
| SplitCells                            | button       | None (Context Menu)     | 拆分单元格                            |
| Magnifier                             | checkBox     | None (Context Menu)     | 放大                                  |
| InkColorPicker                        | gallery      | None (Context Menu)     | 墨迹颜色                              |
| InkColorPicker                        | gallery      | None (Context Menu)     | 墨迹颜色                              |
| OrganizationChartInsertSubordinate    | button       | None (Context Menu)     | 下属                                  |
| OrganizationChartInsertCoworker       | button       | None (Context Menu)     | 同事                                  |
| OrganizationChartInsertAssistant      | button       | None (Context Menu)     | 助手                                  |
| SmartArtOrganizationChartStandard     | button       | None (Context Menu)     | 标准                                  |
| SmartArtOrganizationChartBoth         | button       | None (Context Menu)     | 两边悬挂                              |
| SmartArtOrganizationChartLeftHanging  | button       | None (Context Menu)     | 左悬挂                                |
| SmartArtOrganizationChartRightHanging | button       | None (Context Menu)     | 右悬挂                                |
| OrganizationChartSelectLevel          | button       | None (Context Menu)     | 级别                                  |
| OrganizationChartSelectBranch         | button       | None (Context Menu)     | 分支                                  |
| ShapeStraightConnectorArrow           | toggleButton | None (Context Menu)     | 直线连接符                            |
| ControlProperties                     | button       | None (Context Menu)     | 属性                                  |
| DiagramChangeToCycleClassic           | button       | None (Context Menu)     | 循环型                                |
| DiagramChangeToRadialClassic          | button       | None (Context Menu)     | 射线型                                |
| DiagramChangeToPyramidClassic         | button       | None (Context Menu)     | 棱锥型                                |
| DiagramChangeToVennDiagramClassic     | button       | None (Context Menu)     | 维恩型                                |
| DiagramChangeToTargetClassic          | button       | None (Context Menu)     | 目标图                                |
| ReviewNewComment                      | button       | None (Context Menu)     | 插入批注                              |
| Cut                                   | button       | None (Context Menu)     | 剪切                                  |
| Copy                                  | button       | None (Context Menu)     | 复制                                  |
| PasteTextOnly                         |              | None (Context Menu)     | 无格式文本                            |
| PasteSpecialDialog                    | button       | None (Context Menu)     | 选择性粘贴                            |
| Paste                                 | button       | None (Context Menu)     | 粘贴                                  |
|                                       | comboBox     | None (Context Menu)     | 字号                                  |
|                                       | comboBox     | None (Context Menu)     | 字体                                  |
| TextHighlightColorPicker              | gallery      | None (Context Menu)     |                                       |
|                                       | button       | None (Context Menu)     | 粘贴                                  |
| ObjectBringToFront                    | button       | None (Context Menu)     | 置于顶层                              |
| ObjectSendToBack                      | button       | None (Context Menu)     | 置于底层                              |
| FileMenuPackageMenu                   | menu         | None (Context Menu)     | 文件打包                              |
| ObjectsUngroup                        | button       | None (Context Menu)     | 取消组合                              |
| FontSizeIncrease                      | button       | None (Context Menu)     | 增大字号                              |
| FontSizeDecrease                      | button       | None (Context Menu)     | 减小字号                              |

> WPS 开发文档： [WPS 基础接口/idMso列表/演示idMso参考](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/idMso%E5%88%97%E8%A1%A8/%E6%BC%94%E7%A4%BAidMso%E5%8F%82%E8%80%83.html)

------------------------------------------------------------------------
# 文档说明

**根对象**

帮助文档中所有的代码示例均是从Application根对象开始获取下面的对象，对于老版本的WPS应用程序，根对象是wps.Application，需要注意区别。

**枚举类型**

在WPS宏编辑器环境（JDE）是可以直接访问的，但是在加载项环境中需要从枚举对象（Application.Enum)中获取，以下示例为给当前文档第一个段落左边框设置线宽

``` JavaScript
Application.ActiveDocument.Paragraphs.Item(1)
  .Borders.Item(Application.Enum.wdBorderLeft).LineWidth = Application.Enum.wdLineWidth025pt
 
```

> WPS 开发文档： [文档说明](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/%E6%96%87%E6%A1%A3%E8%AF%B4%E6%98%8E.html)

------------------------------------------------------------------------
