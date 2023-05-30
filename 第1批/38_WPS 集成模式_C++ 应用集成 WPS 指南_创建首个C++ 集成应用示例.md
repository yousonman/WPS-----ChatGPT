# 创建首个C++ 集成应用示例

## 基本介绍

这份文档介绍了如何使用提供的Demo，在linux环境下使用C++开发基于 WPS 的解决方案。

## 源码下载

[访问下载页面](https://code.aliyun.com/zouyingfeng/wps/tree/master/cpp)

## 目录结构说明

将源码下载下来后，其目录结构如下：

\| -- include WPS头文件，包含公共以及WPS，ET，WPP各自对应的api接口。

\| -- src c++源代码文件。

\| -- README.md 帮助文件

## 环境准备

- 操作系统： Linux系统（中标麒麟，银河麒麟等国产操作系统也支持）
- 安装 Qt ( 兼容 Qt4或Qt5 )
- 安装 WPS Office 2019 For Linux

## 编译及运行

- 在终端中使用cd命令，进入代码的src目录，执行qmake命令。 ![](服务器端图像/qmake.png)
- 执行make，在当前目录下会生成wpsDemo。 ![](服务器端图像/make.png)
- 执行./wpsDemo，运行wpsDemo。 ![](服务器端图像/wpsDemo.png)
- 最终效果如图所示。 ![](服务器端图像/DemoUI.png)

> WPS 开发文档： [WPS 集成模式/C++ 应用集成 WPS 指南/创建首个C++ 集成应用示例](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/C%2B%2B%20%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%20WPS%20%E6%8C%87%E5%8D%97/%E5%88%9B%E5%BB%BA%E9%A6%96%E4%B8%AAC%2B%2B%20%E9%9B%86%E6%88%90%E5%BA%94%E7%94%A8%E7%A4%BA%E4%BE%8B.html)

------------------------------------------------------------------------
