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
