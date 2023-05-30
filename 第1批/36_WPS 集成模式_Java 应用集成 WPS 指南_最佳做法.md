# 最佳做法

## 默认参数

Java中不提供函数默认参数，因而调用接口时必须给定义的所有参数都传参，如果某个对应位置的参数需要使用默认值，请传递 **Variant.getMissing()** ，例如：

``` Java
app.get_ActiveDocument().Protect(WdProtectionType.wdAllowOnlyReading,
　　Variant.getMissing(),
　　Variant.getMissing(),
　　Variant.getMissing(),
　　Variant.getMissing()
);
```

## 函数返回值

返回值是api对象类型的接口可以直接判断返回值是否为null进行容错校验，如 **Application.get_ActiveDocument()** 无返回值或返回值是数字，字符串等类型的接口，需捕获Exception异常，以此判断操作是否执行成功。

## 属性获取和设置

所有对象的属性在设置和获取时，在java中均需使用加上get_或put_前缀的对应方法，无法通过“对象名.属性”直接调用，例如：

``` Java
CommandBarControl controlOfd = commandBar.get_Controls().get_Item(26);
controlOfd.put_Visible(false);
// 不能使用 controlOfd.visible = false;
```

## 关于lcid参数

遇到接口参数名为 lcid 的参数，默认传递2052即可。

## Canves句柄说明

1.获取用于装载WPS的Canves的句柄(pid和peer)的语句，必须在 mainFrame.setVisible(true) 之后调用，否则拿到空指针。

2.获取用于装载WPS的Canves的句柄(pid和peer)的语句，不能写在JPanel的构造函数里，必须在JPanel对象创建完成后再获取，否则拿到空指针。

## 注册事件回调流程

1.注册关闭事件。

``` Java
EventCookie cookie = 
    app.advise(ApplicationEvents4.class, new ApplicationEvents4() {
        @Override
        public void DocumentBeforeClose(Document doc, Holder<Boolean> cancel) {
            JOptionPane.showMessageDialog(null,"关闭事件被拦截");
            cancel.value = true; //cancel为true，代表取消关闭
        }
    });
```

2.取消注册关闭事件。

3.添加一个按钮，给按钮注册事件

``` Java
cookie.close();
cookie = null; 
```

``` Java
//添加一个按钮
CommandBar bar = app.get_CommandBars().Add(
    "test", 1, Variant.getMissing(), Variant.getMissing());
CommandBarControl control = 
    bar.get_Controls().Add(1, 1, "test", 1, "test");
bar.put_Visible(true);
control.put_Visible(true);
control.put_Caption("test");

//给按钮注册事件
EventCookie cookie = control.advise(_CommandBarButtonEvents.class,
    new _CommandBarButtonEvents() {
        @Override
        public void Click(CommandBarButton ctrl, Holder<Boolean> cancelDefault) {
            JOptionPane.showMessageDialog(null,"你点击了按钮 test");
            cancelDefault.value = true;
        }
    });
```

4.取消注册按钮时间

``` Java
cookie.close(); cookie = null;
```

## 调用WPS 2016程序的说明

JavaAPI Demo中默认开启了加密，在WPS2016下会导致新建文档阻塞。使用WPS2016的用户需要将代码中这一行前面的注释去掉：

``` Java
args.setCrypted(false); //wps2016需要关闭加密
```

> WPS 开发文档： [WPS 集成模式/Java 应用集成 WPS 指南/最佳做法](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/Java%20%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%20WPS%20%E6%8C%87%E5%8D%97/%E6%9C%80%E4%BD%B3%E5%81%9A%E6%B3%95.html)

------------------------------------------------------------------------
