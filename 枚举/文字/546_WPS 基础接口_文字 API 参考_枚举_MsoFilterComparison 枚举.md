# MsoFilterComparison 枚举

指定如何比较 **ODSOFilter** 对象的 **Column** 和 **CompareTo** 属性。

| 名称                                | 值  | 说明                                                                     |
|-------------------------------------|-----|--------------------------------------------------------------------------|
| msoFilterComparisonContains         | 8   | 如果 CompareTo 字符串的任意部分包含在列值中，则列与 CompareTo 匹配。     |
| msoFilterComparisonEqual            | 0   | 如果 CompareTo 值与列值相同，则列与 CompareTo 匹配。                     |
| msoFilterComparisonGreaterThan      | 3   | 如果列值大于 CompareTo 值，则列与 CompareTo 匹配。                       |
| msoFilterComparisonGreaterThanEqual | 5   | 如果列值大于或等于 CompareTo 值，则列与 CompareTo 匹配。                 |
| msoFilterComparisonIsBlank          | 6   | 如果列为空，则列通过筛选。                                               |
| msoFilterComparisonIsNotBlank       | 7   | 如果列为空，则列通过筛选。                                               |
| msoFilterComparisonLessThan         | 2   | 如果列值小于 CompareTo 值，则列与 CompareTo 匹配。                       |
| msoFilterComparisonLessThanEqual    | 4   | 如果列值小于或等于 CompareTo 值，则列与 CompareTo 匹配。                 |
| msoFilterComparisonNotContains      | 9   | 如果 CompareTo 字符串的任何部分都未包含在列值中，则列与 CompareTo 匹配。 |
| msoFilterComparisonNotEqual         | 1   | 如果 CompareTo 值不等于列值，则列与 CompareTo 匹配。                     |

> WPS 开发文档： [WPS 基础接口/文字 API 参考/枚举/MsoFilterComparison 枚举](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoFilterComparison%20%E6%9E%9A%E4%B8%BE.html)

------------------------------------------------------------------------
