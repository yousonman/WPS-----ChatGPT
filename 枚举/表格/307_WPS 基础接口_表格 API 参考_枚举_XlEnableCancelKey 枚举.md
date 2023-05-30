# XlEnableCancelKey 枚举

指定 WPS Office ET 2007 如何处理 Ctrl+Break（或 Esc、Command+Period）用户中断以用于运行程序。

| 名称           | 值  | 说明                                                                                               |
|----------------|-----|----------------------------------------------------------------------------------------------------|
| xlDisabled     | 0   | 完全禁用“取消”键捕获功能。                                                                         |
| xlErrorHandler | 2   | 将中断作为错误发送给运行程序，由 On Error GoTo 语句设置的错误处理程序捕获。可捕获的错误代码为 18。 |
| xlInterrupt    | 1   | 中断当前运行程序，用户可进行调试或结束程序的运行。                                                 |

> WPS 开发文档： [WPS 基础接口/表格 API 参考/枚举/XlEnableCancelKey 枚举](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlEnableCancelKey%20%E6%9E%9A%E4%B8%BE.html)

------------------------------------------------------------------------
