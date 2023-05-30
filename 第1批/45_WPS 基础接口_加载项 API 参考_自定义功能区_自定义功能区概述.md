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
