# MsoSyncConflictResolutionType 枚举

指定在同步共享文档时如何解决冲突。与 **Sync** 对象的 **ResolveConflict** 方法一起使用。

| 名称                      | 值  | 说明                                                                                                                                                                       |
|---------------------------|-----|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| msoSyncConflictClientWins | 0   | 用本地副本代替服务器副本。                                                                                                                                                 |
| msoSyncConflictMerge      | 2   | 将对服务器副本所做的更改合并到本地副本中。要解决冲突并成功地合并更改，必须在合并更改后保存活动文档，然后再次使用 msoSyncConflictClientWins 选项调用 ResolveConflict 方法。 |
| msoSyncConflictServerWins | 1   | 用服务器副本代替本地副本。                                                                                                                                                 |

> WPS 开发文档： [WPS 基础接口/表格 API 参考/枚举/MsoSyncConflictResolutionType 枚举](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoSyncConflictResolutionType%20%E6%9E%9A%E4%B8%BE.html)

------------------------------------------------------------------------
