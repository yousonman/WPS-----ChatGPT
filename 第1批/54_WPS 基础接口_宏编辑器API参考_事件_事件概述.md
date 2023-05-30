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
