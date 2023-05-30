# WPS 加载项可用性

## 请确认安装的 WPS 安装包带有 WPS 加载项功能。

个人用户可以在 WPS 官网下载最新的安装包；企业用户请联系贵公司的 WPS 项目负责人。

## 启用 WPS 加载项

打开 WPS 安装目录，进入 wps.exe 程序同级目录，在cfgs文件夹下新建或打开oem.ini文件，在文档最后添加如下配置

``` ini
[Support]
##启用加载项
JsApiPlugin = true
##启用网页调试，详情请参考 WPS 基础接口说明
JsApiShowWebDebugger=true
[Server]
##指定加载项配置文件 jsplugins.xml 链接地址
JSPluginsServer = http://***/jsplugins.xml
```

## 加载项配置说明

加载项配置在文件 jsplugins.xml 中。jsplugins.xml 从服务器下载后会保存在本地，服务器更新后会重新下载。jsplugins.xml 格式如下：

``` xml
<jsplugins>
    <jsplugin name="demo_et"  version="3.6" type="et" url="http://localhost:8080/demo_et_3.6.7z"/>
    <jsplugin name="demo_wpp" version="3.6" type="wpp" url="http://localhost:8080/demo_wpp_3.6.7z"/>
</jsplugins>
```

\<jsplugins\> \*\*\* \</jsplugins\> 包含所有加载项配置

\<jsplugin \*\*\* /\> 代表一个加载项

name 为加载项名称

version 为该加载项版本号

type 为加载项类型

url 为加载项压缩包下载链接

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/WPS 加载项可用性](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E5%8A%A0%E8%BD%BD%E9%A1%B9%20API%20%E5%8F%82%E8%80%83/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%8F%AF%E7%94%A8%E6%80%A7.html)

------------------------------------------------------------------------
