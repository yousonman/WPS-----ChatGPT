# WPS 加载项性能

WPS 加载项为了保证WPS 应用程序的安全性与稳定性，采用多进程架构，把网页渲染端和 WPS 应用进程分离； 同时通过共享内存、事件对象等技术，实现了一套进程间快速同步调用的机制。 在性能测试过程中，为10000个单元格赋值耗时2000毫秒，充分满足要求。

``` JavaScript
function TestSetCellFormula() 
{
    var cells = Application.ActiveWorkbook.ActiveSheet.Cells;
    let date = new Date();
    let start = date.getTime();
    for(var i = 1; i <= 100; ++i) 
    { 
        for(var j = 1; j <= 100; ++j) 
        {
            cells.Item(i, j).Formula = i + j; 
        } 
    } 
    date = new Date(); 
    let end = date.getTime(); 
    alert(end - start); return; 
}
```

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/WPS 加载项性能](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E6%80%A7%E8%83%BD.html)

------------------------------------------------------------------------
