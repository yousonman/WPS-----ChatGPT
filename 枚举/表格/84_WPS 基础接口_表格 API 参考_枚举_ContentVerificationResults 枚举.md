# ContentVerificationResults 枚举

提供验证文档内容是否已更改的状态。

| 名称                 | 值  | 说明                                 |
|----------------------|-----|--------------------------------------|
| contverresError      | 0   | 验证导致了错误。                     |
| contverresModified   | 4   | 文档内容自从经过数字签名之后已修改。 |
| contverresUnverified | 2   | 文档尚未验证。                       |
| contverresValid      | 3   | 文档内容已验证且有效。               |
| contverresVerifying  | 1   | 文档内容当前正在验证。               |

> WPS 开发文档： [WPS 基础接口/表格 API 参考/枚举/ContentVerificationResults 枚举](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/ContentVerificationResults%20%E6%9E%9A%E4%B8%BE.html)

------------------------------------------------------------------------
