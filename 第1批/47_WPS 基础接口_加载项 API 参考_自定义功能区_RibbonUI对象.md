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
