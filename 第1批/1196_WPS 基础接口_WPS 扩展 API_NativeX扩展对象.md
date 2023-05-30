# NativeX扩展对象

## 背景

NativeX是提供给第三方开发者用native代码来扩展JS接口的技术方案。JSAPI提供的接口个数有限，当碰到不满足需求的情况下，第三方开发者可以用native代码(c++/ruby/python等)来扩展JS接口，加载第三方NativeX模块后，实现在JSAPI中使用自己扩展的接口。例如第三方开发者可以利用发布的SDK，将复杂的算法（如加密，解密算法）实现并生成dll或者so，提供给NativeX对象加载在JSAPI中使用。

## 使用SDK生成dll

NativeX对象中有三种调用方式Get,Set和Execute，在SDK中注册属性和函数名，指定回调函数，参数个数和参数类型，示例代码如下。

![](Base64图像/Base64图像49来自_WPS%20基础接口_WPS%20扩展%20API_NativeX扩展对象.png)

JSWpsRootObject为根对象，如果想返回NativeX子对象，创建JSWpsSubObject_1类，同样是继承父类NativeXObjectBase，在根对象处理返回值时，将NativeX子对象的指针返回即可。

![](Base64图像/Base64图像50来自_WPS%20基础接口_WPS%20扩展%20API_NativeX扩展对象.png)

## 注册工具

使用NativeX扩展对象之前，需要使用ksomisc工具将dll路径注册到文件中。

![](Base64图像/Base64图像51来自_WPS%20基础接口_WPS%20扩展%20API_NativeX扩展对象.png)

调用ksomisc程序，-regnativex参数设置dll路径， -unregnativex参数反注册， -i参数设置是否为进程内调用。在%AppData%\kingsoft\wps\jsaddons\binary\目录下生成WPSNativeX.conf文件，WPSNativeX.conf文件中包含了此dll中是否曾经发生过崩溃信息，dll路径信息，是否进程内调用信息。

![](Base64图像/Base64图像52来自_WPS%20基础接口_WPS%20扩展%20API_NativeX扩展对象.png)

## JSAPI中调用Nativex对象

使用根对象下的CreateObject方法创建NativeX对象，参数为利用SDK生成dll的名字。可以先将使用代码片段写入到加载项的按钮中，也可以直接在JS调试器中单步调用。NativeX扩展对象支持创建子对象，支持创建多实例，支持创建多对象。

![](Base64图像/Base64图像53来自_WPS%20基础接口_WPS%20扩展%20API_NativeX扩展对象.png)

> WPS 开发文档： [WPS 基础接口/WPS 扩展 API/NativeX扩展对象](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/WPS%20%E6%89%A9%E5%B1%95%20API/NativeX%E6%89%A9%E5%B1%95%E5%AF%B9%E8%B1%A1.html)

------------------------------------------------------------------------
