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
