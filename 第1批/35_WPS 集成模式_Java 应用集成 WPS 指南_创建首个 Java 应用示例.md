# 创建首个 Java 应用示例

## 环境搭建

1）新建空的Java项目。 2）导入依赖jar包，以idea为例： File\>\>Project Structure\>\>Libraries\>\>Add(+)\>\>Java\>\>选择所有jar包\>\>apply\>\>ok 。

## 编写代码测试主类创建空白swing窗体

``` Java
import javax.swing.*;
public class TestMain {
     public static void main(String[] args) {
    //Linux下必须加这一句
    java.lang.System.setProperty("sun.awt.xembedserver", "true");

    SwingUtilities.invokeLater(new Runnable() {
        @Override
        public void run() {
            JFrame mainFrame = new JFrame();
                //设置显示窗口标题
                mainFrame.setTitle("WPS JAVA接口调用演示");
                //设置窗口显示尺寸
                mainFrame.setSize(1524, 768);
                    //置窗口是否可以关闭
                    mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
                //禁止缩放
                mainFrame.setResizable(false);
                //设置窗口是否可见
                mainFrame.setVisible(true);
            }
        });
    }
}
```

## 添加按钮事件

``` Java
Panel wpsPanel = new JPanel();
wpsPanel.setLayout(new BorderLayout());

//初始化一个空白的画布用于装载WPS
Canvas client = new Canvas() {
        public void paint(Graphics g) {
        super.paint(g);
    }
};
//添加画布
wpsPanel.add(client, BorderLayout.CENTER);
//添加JPane
mainFrame.add(wpsPanel);
```

## 将WPS嵌入到窗体内

``` Java
//获取用于装载WPS的Canves的句柄
WindowIDProvider pid = (WindowIDProvider)client.getPeer();
XEmbedCanvasPeer peer = (XEmbedCanvasPeer)pid;
peer.removeXEmbedDropTarget();
//将画布句柄和长宽等信息作为参数初始化WPS窗体
WpsArgs args = WpsArgs.ARGS_MAP.get(WpsArgs.App.WPS);
args.setWinid(pid.getWindow());
args.setHeight(client.getHeight());
args.setWidth(client.getWidth());
Application app = ClassFactory.createApplication();// 创建WPS Application对象
```

## 调用api接口

``` Java
//新建文档
app.get_Documents().Add(Variant.getMissing(),
    Variant.getMissing(), Variant.getMissing(), Variant.getMissing());
//获取版本号
String version = app.get_Build();
//光标位置插入文本，此处写入版本号
app.get_Selection().TypeText(version);
//关闭文档
app.get_ActiveDocument().Close(false, 
Variant.getMissing(), Variant.getMissing());
```

## 完整范例

1.[一个简单范例](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20集成模式/Java%20应用集成%20WPS%20指南/smalldemo.htm)

2.[获取更多范例](https://code.aliyun.com/zouyingfeng/wps/tree/master/java)

> WPS 开发文档： [WPS 集成模式/Java 应用集成 WPS 指南/创建首个 Java 应用示例](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/Java%20%E5%BA%94%E7%94%A8%E9%9B%86%E6%88%90%20WPS%20%E6%8C%87%E5%8D%97/%E5%88%9B%E5%BB%BA%E9%A6%96%E4%B8%AA%20Java%20%E5%BA%94%E7%94%A8%E7%A4%BA%E4%BE%8B.html)

------------------------------------------------------------------------
