# JavaScript运行时说明

WPS宏编辑器集成了一个V8 引擎的 JavaScript 运行时，支持大部分ES6语法，因此宏编辑器支持JavaScript 标准内置对象， **注意** ，JS内置对象和浏览器的内置对象是不同的，WPS宏编辑器集成的是 JavaScript 运行时，而不是浏览器， 因此 WPS宏编辑器 不支持 浏览器的内置对象 。 具体API参见 <https://developer.mozilla.org/zh-CN/docs/Web/JavaScript/Reference/Global_Objects> 。如以下代码输出圆周率，可以在WPS宏编辑器中执行：

``` JavaScript
function abcd()
{
    Console.log(Math.PI);
}
```

使用过程中遇到的编译错误和语法错误具体原因和细节可以在 <https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Errors> 上查看。

> WPS 开发文档： [WPS 基础接口/宏编辑器API参考/JavaScript运行时说明](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%AE%8F%E7%BC%96%E8%BE%91%E5%99%A8API%E5%8F%82%E8%80%83/JavaScript%E8%BF%90%E8%A1%8C%E6%97%B6%E8%AF%B4%E6%98%8E.html)

------------------------------------------------------------------------
