# WdFieldKind 枚举

指定 **Field** 对象的域类型。

| 名称            | 值  | 说明                                                                                                                     |
|-----------------|-----|--------------------------------------------------------------------------------------------------------------------------|
| wdFieldKindCold | 3   | 没有结果的域，如索引项 (XE) 域、目录项 (TC) 域或 Private 域。                                                            |
| wdFieldKindHot  | 1   | 每次显示或页面重新设置格式时自动更新的域，但也可进行手动更新（如 INCLUDEPICTURE 或 FORMDROPDOWN）。                      |
| wdFieldKindNone | 0   | 无效的域（如其中没有任何内容的一对域字符）。                                                                             |
| wdFieldKindWarm | 2   | 能进行更新并有结果的域。该类型包括随数据源变化而自动更新的域，也包括那些可以手动进行更新的域（如 DATE 或 INCLUDETEXT）。 |

> WPS 开发文档： [WPS 基础接口/文字 API 参考/枚举/WdFieldKind 枚举](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdFieldKind%20%E6%9E%9A%E4%B8%BE.html)

------------------------------------------------------------------------
