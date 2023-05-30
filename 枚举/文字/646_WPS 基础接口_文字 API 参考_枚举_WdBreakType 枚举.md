# WdBreakType 枚举

指定分隔符的类型。

| 名称                     | 值  | 说明                                                                                                                           |
|--------------------------|-----|--------------------------------------------------------------------------------------------------------------------------------|
| wdColumnBreak            | 8   | 插入点处的分栏符。                                                                                                             |
| wdLineBreak              | 6   | 换行符。                                                                                                                       |
| wdLineBreakClearLeft     | 9   | 换行符。                                                                                                                       |
| wdLineBreakClearRight    | 10  | 换行符。                                                                                                                       |
| wdPageBreak              | 7   | 插入点处的分页符。                                                                                                             |
| wdSectionBreakContinuous | 3   | 新节不包含相应分页符。                                                                                                         |
| wdSectionBreakEvenPage   | 4   | 使下一节从下一偶数页开始的分节符。如果分节符落入偶数页，则 WPS 将下一奇数页留为空白。                                          |
| wdSectionBreakNextPage   | 2   | 分节符在下一页。                                                                                                               |
| wdSectionBreakOddPage    | 5   | 使下一节从下一奇数页开始的分节符。如果分节符落入奇数页，则 WPS 将下一偶数页留为空白。                                          |
| wdTextWrappingBreak      | 11  | 结束当前行，并强制文字在图片、表格或其他项目的下方继续。文字将在下一个空行（且该空行不包含与左边距或右边距对齐的表格）上继续。 |

> WPS 开发文档： [WPS 基础接口/文字 API 参考/枚举/WdBreakType 枚举](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdBreakType%20%E6%9E%9A%E4%B8%BE.html)

------------------------------------------------------------------------
