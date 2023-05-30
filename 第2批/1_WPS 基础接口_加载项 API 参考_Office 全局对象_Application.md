# Application 对象

## 简介

Application 是浏览器 Window 对象的子对象。 Application 对象在浏览器创建时自动注入到浏览器执行上下文环境

## 说明

Wps 对象是访问 Application 对象模型的根对象，它在浏览器创建后动态注入到浏览器执行上下文中，下列代码展示如何从 Wps 对象拿到 Application 对象，然后从 Application 对象拿到当前文档的名字。

``` JavaScript
let app = Application
let doc = app.ActiveDocument
if(doc){
    let docName = doc.Name
}
```

除了原office标准对象外的扩展对象也都是 Application 对象的子对象，具体 Application 对象扩展了哪些对象，可以参阅 Application 对象成员相关说明，下列代码展示了如果从 Application 对象拿到 FileSystem 对象。

``` JavaScript
let fs =  Application.FileSystem
```

Application 对象还包含一些函数，具体 Application 对象支持哪些函数，可以参阅 Application 对像成员说明，下列代码展示了如何从 Application 对象创建一个网页对话框。

``` JavaScript
let width =  300 * window.devicePixelRatio
let hight= 200 * window.devicePixelRatio
let bModal = true
Application.ShowDialog(""http://www.wps.cn"", "wps首页", width, hight, bModal)
```

## 方法

| 序号 | 名称                                          | 说明                                                                         |
|:----:|:----------------------------------------------|:-----------------------------------------------------------------------------|
|  1   | [CreateTaskpane](#Application.CreateTaskpane) | 用给定的url创建一个 Taskpane 。                                              |
|  2   | [GetTaskpane](#Application.GetTaskpane)       | 根据给定的ID值，返回对应的 TaskPane 对象，这个ID值与 TaskPane.ID 对应。      |
|  3   | [ShowDialog](#Application.ShowDialog)         | 根据给定的url、标题、宽高等信息创建一个对话框，对话框中的内容是一个web网页。 |
|  4   | [alert](#Application.alert)                   | 弹出alert警告。                                                              |
|  5   | [confirm](#Application.confirm)               | 弹出一个确认框。返回值为一个bool变量，代表对确认框的操作是确认还是取消。     |

## 属性

| 序号 | 名称                                        | 说明                                                                                                                                                                                                                                                      |
|:----:|:--------------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ApiEvent](#Application.ApiEvent)           | 返回一个 ApiEvent 对象，有关 ApiEvent 对象的相关说明，请参阅 ApiEvent 对象。                                                                                                                                                                              |
|  2   | [Env](#Application.Env)                     | 返回一个 Env 对象，该对象代表获取系统环境相关的操作对象，有关Env对象的详细属性和方法，可以参考 Env 对象。                                                                                                                                                 |
|  3   | [FileSystem](#Application.FileSystem)       | 返回一个 FileSystem 对象，该对象包含了对文件和文件夹的常见操作，有关FileSystem对象的详细属性和方法，可以参考 FileSystem 对象。                                                                                                                            |
|  4   | [PluginStorage](#Application.PluginStorage) | 返回一个 PluginStorage 对象，该对象与 LocalStorage 的用法类似，但该对象的数据不会被持久化，在wps的jsapi插件中，该对象在一个插件中对应着同一个实例，因此该对象可以在一个jsapi插件的不同网页间共享数据，有关该对象的详细介绍，可以参考 PluginStorage 对象。 |

## 成员方法

### Application.CreateTaskpane

用给定的url创建一个 **Taskpane** 。

**语法**

**express.CreateTaskpane ( *url* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明        |
|-------|-----------|----------|-------------|
| *url* | 必选      | String   | 给定url路径 |

### Application.GetTaskpane

根据给定的ID值，返回对应的 TaskPane 对象，这个ID值与 TaskPane.ID 对应。

**语法**

**express.GetTaskpane ( *ID* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称 | 必选/可选 | 数据类型 | 说明                |
|------|-----------|----------|---------------------|
| *ID* | 必选      | Number   | 表示TaskPane.ID属性 |

### Application.ShowDialog

根据给定的url、标题、宽高等信息创建一个对话框，对话框中的内容是一个web网页。

**语法**

**express.ShowDialog ( *url* , *caption* , *width* , *height* , *bModal* , *hasCaption* , *resizeEdge* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型 | 说明                                                                                                                                            |
|--------------|-----------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------|
| *url*        | 必选      | String   | 此参数代表对话框网页的url, 这个url可以是一个http/https的网页，也可以是一个本地资源网页。                                                        |
| *caption*    | 必选      | String   | 对话框的标题。                                                                                                                                  |
| *width*      | 必选      | Long     | 表示对话框的宽度，这个宽度是物理像素值，尤其要注意在普通屏和高分辨率屏下是不一样的，常见的做法是借助于 window.devicePixelRatio 来消除这个影响。 |
| *height*     | 必选      | Long     | 表示对话框的高度，这个高度是物理像素值，尤其要注意在普通屏和高分辨率屏下是不一样的，常见的做法是借助于 window.devicePixelRatio 来消除这个影响。 |
| *bModal*     | 必选      | Bool     | 表示对话框是否模态。                                                                                                                            |
| *hasCaption* | 可选      | Bool     | 默认为true，表示对话框是否有标题栏和边框。为false时为无边框窗口，允许开发者对标题栏，窗口阴影等进行高级自定义                                   |
| *resizeEdge* | 可选      | Bool     | 窗口为无边框窗口时有效。默认为2，表示操作对话框缩放的宽度，单位是像素。                                                                         |

**示例**

``` JavaScript
以下代码示例创建了一个对话框，并让这个对话框模态显示。
let width = 400 * window.devicePixelRatio
let height = 300 * window.devicePixelRatio
Application.ShowDialog("https://www.wps.cn", "wps网站", width, height, true)
```

### Application.alert

弹出alert警告。

**语法**

**express.alert ( *text* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                          |
|--------|-----------|----------|-------------------------------|
| *text* | 必选      | String   | 表示警告框要显示的字符串内容. |

**说明**

alert实际上是浏览器 **window** 对象内置的一个方法，在 **wps** 的js环境中，wps对这个对象进行了拦截，主要原因是由于chrome的技术架构决定的。 wps的js环境是内置了一个chrome的内核来展示网页、执行js脚本，在chrome的技术架构中，chrome的js执行、css及html渲染、网页uil用户操作等都分别是它自己的线程中， 同样的，如果alert不作拦截，当网页前端代码调用alert时，这个alert框仅仅只是阻塞住了chrome的ui线程，而对wps主线程没有影响，这样就会产生alert框弹出来了，但wps仍然能 响应用户操作的奇怪现象，因此，wps对alert框进行了拦截，让alert框在wps主线程中弹出来，当alert弹出时，阻塞的会是wps主线程。

### Application.confirm

弹出一个确认框。返回值为一个bool变量，代表对确认框的操作是确认还是取消。

**语法**

**express.confirm ( *text* )**

\*express是一个代表 Application 对象的变量。

**参数**

| 名称   | 必选/可选 | 数据类型 | 说明                          |
|--------|-----------|----------|-------------------------------|
| *text* | 必选      | String   | 表示确认框要显示的字符串内容. |

**说明**

confirm与alert一样，实际上也是浏览器 **window** 对象内置的一个方法，在wps的js环境中，wps对这个对象进行了拦截，主要原因是由于chrome的技术架构决定的。 wps的js环境是内置了一个chrome的内核来展示网页、执行js脚本，在chrome的技术架构中，chrome的js执行、css及html渲染、网页uil用户操作等都分别是它自己的线程中， 同样的，如果不作拦截，当网页前端代码调用confirm时，这个confirm框仅仅只是阻塞住了chrome的ui线程，而对wps主线程没有影响，这样就会产生confirm框弹出来了，但wps仍然能 响应用户操作的奇怪现象，因此，wps对confirm框进行了拦截，让confirm框在wps主线程中弹出来，当confirm弹出时，阻塞的会是wps主线程。

## 成员属性

### Application.ApiEvent

返回一个 ApiEvent 对象，有关 **ApiEvent** 对象的相关说明，请参阅 ApiEvent 对象。

**语法**

**express.ApiEvent**

\*express是一个代表 Application 对象的变量。

### Application.Env

返回一个 Env 对象，该对象代表获取系统环境相关的操作对象，有关Env对象的详细属性和方法，可以参考 Env 对象。

**语法**

**express.Env**

\*express是一个代表 Application 对象的变量。

### Application.FileSystem

返回一个 FileSystem 对象，该对象包含了对文件和文件夹的常见操作，有关FileSystem对象的详细属性和方法，可以参考 FileSystem 对象。

**语法**

**express.FileSystem**

\*express是一个代表 Application 对象的变量。

### Application.PluginStorage

返回一个 PluginStorage 对象，该对象与 **LocalStorage** 的用法类似，但该对象的数据不会被持久化，在wps的jsapi插件中，该对象在一个插件中对应着同一个实例，因此该对象可以在一个jsapi插件的不同网页间共享数据，有关该对象的详细介绍，可以参考 PluginStorage 对象。

**语法**

**express.PluginStorage**

\*express是一个代表 Application 对象的变量。

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/Office 全局对象/Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
