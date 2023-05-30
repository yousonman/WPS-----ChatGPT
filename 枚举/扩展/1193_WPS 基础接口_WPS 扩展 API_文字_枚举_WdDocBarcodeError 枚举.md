# WdDocBarcodeError 枚举

表示插入公文二维码的错误结果提示。

| 名称                  | 值  | 说明                                                             |
|-----------------------|-----|------------------------------------------------------------------|
| WdBarcodeNoError      | 0   | 成功插入公文二维码。                                             |
| WdBarcodeFieldFormat  | 1   | 插入二维码存在字段格式错误，比如条码编号，日期格式，枚举值错误。 |
| WdBarcodeFieldBytes   | 2   | 插入二维码存在字段的字节数不在合法范围内。                       |
| WdBarcodeNoDocField   | 3   | 插入二维码时该文档没有公文域“条码”。                             |
| WdRowsInsufficient    | 4   | 固定高度时，给定的高度太小导致无法容纳二维码。                   |
| WdRowsExceed          | 5   | 固定高度时，指定的行太多，导致二维码高度超出已有限制。           |
| WdColumnsInsufficient | 6   | 固定宽度时，给定的宽度太小导致无法容纳二维码。                   |
| WdColumnsExceed       | 7   | 固定宽度时，指定的列太多，导致二维码宽度超出已有限制。           |

> WPS 开发文档： [WPS 基础接口/WPS 扩展 API/文字/枚举/WdDocBarcodeError 枚举](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/WPS%20%E6%89%A9%E5%B1%95%20API/%E6%96%87%E5%AD%97/%E6%9E%9A%E4%B8%BE/WdDocBarcodeError%20%E6%9E%9A%E4%B8%BE.html)

------------------------------------------------------------------------
