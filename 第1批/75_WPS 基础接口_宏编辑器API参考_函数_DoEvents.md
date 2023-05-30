# DoEvents

处理进程的消息队列中的消息。

## 语法

**DoEvents** ()

## 说明

DoEvents 将暂时停止当前的宏执行，在 处理完进程的消息队列中的消息 后返回再继续宏的执行。

## 示例

此示例使用 DoEvents 函数来使得每迭代 1000 次循环就将 暂时停止当前的宏执行，处理完 进程的消息队列中的消息 后返回再继续宏的执行，以便迭代过程中在合适的时机更新文档视图或者响应鼠键消息来通过调试器中止迭代 。

``` JavaScript
function abcd()
{
    var beginTime = new Date();
    for (i = 0; i < 20000; i++) 
    {
        Selection.TypeText("金山办公软件");
        if (i % 1000 == 0)
            DoEvents()
    }
    var endTime = new Date();
    Debug.Print("用时共计"+(endTime-beginTime)+"ms");   
}
```

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/函数/DoEvents](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/%E5%87%BD%E6%95%B0/DoEvents.html)

------------------------------------------------------------------------
