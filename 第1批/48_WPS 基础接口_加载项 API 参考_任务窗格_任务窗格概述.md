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
