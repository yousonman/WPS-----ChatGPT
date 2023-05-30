# XlRobustConnect 枚举

指定数据透视表缓存与其数据源连接的方式。

| 名称         | 值  | 说明                                                                                       |
|--------------|-----|--------------------------------------------------------------------------------------------|
| xlAlways     | 1   | 缓存始终使用外部源信息（由 SourceConnectionFile 或 SourceDataFile 属性定义）进行重新连接。 |
| xlAsRequired | 0   | 缓存通过 Connection 属性使用外部源信息进行重新连接。                                       |
| xlNever      | 2   | 缓存从不使用源信息进行重新连接。                                                           |

> WPS 开发文档： [WPS 基础接口/表格 API 参考/枚举/XlRobustConnect 枚举](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlRobustConnect%20%E6%9E%9A%E4%B8%BE.html)

------------------------------------------------------------------------
