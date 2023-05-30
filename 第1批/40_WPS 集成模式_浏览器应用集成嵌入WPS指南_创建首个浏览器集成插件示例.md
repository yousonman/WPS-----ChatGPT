# 创建首个浏览器集成插件示例

## 环境准备

- 一台安装了LINUX系统的机器。
- 一个支持NPAPI的浏览器（参见 浏览器应用集成嵌入WPS概述 ） 。

## Demo代码

源码如下：

``` Html
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <title> wps二次开发演示 </title>
</head>

<body>
    <object type="application/x-wps" width="1024" height="768" id="wps"></object>
    <!--通过类型 “application/x-wps” 来加载WPS浏览器NPAPI插件，并传递宽高值-->
    <script type="text/javascript">
        //来获取wps的插件对象
        var obj = document.getElementById("wps");
        //获取WPS对象树的根节点，后续功能调用都是根据 Application 逐步得到的
        var Application = obj.Application;
        //创建一个新的空白文档
        Application.createDocument("wps");
        //获取文档内容区域
        var range = Application.ActiveDocument.Content;
        // 在文档内容区域插入一个3行5列的表格，设置表格显示边框
        Application.ActiveDocument.Tables.Add(range, 3, 5).Borders.Enable = 1;
     </script>
</body>
</html>
```

[在浏览器中打开此链接查看效果](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20集成模式/浏览器应用集成嵌入WPS指南/np_demo.htm)

## 效果如下图所示

![](服务器端图像/示例图.jpg)

## 浏览器集成插件Demo

[请点击查看更多demo链接](https://code.aliyun.com/zouyingfeng/wps/tree/master/np-example/browser-integration-wps)

> WPS 开发文档： [WPS 集成模式/浏览器应用集成嵌入WPS指南/创建首个浏览器集成插件示例](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/%E6%B5%8F%E8%A7%88%E5%99%A8%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%E5%B5%8C%E5%85%A5WPS%E6%8C%87%E5%8D%97/%E5%88%9B%E5%BB%BA%E9%A6%96%E4%B8%AA%E6%B5%8F%E8%A7%88%E5%99%A8%E9%9B%86%E6%88%90%E6%8F%92%E4%BB%B6%E7%A4%BA%E4%BE%8B.html)

------------------------------------------------------------------------
