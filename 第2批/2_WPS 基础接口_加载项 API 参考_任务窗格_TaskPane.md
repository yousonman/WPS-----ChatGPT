# TaskPane 对象

## 简介

TaskPane 对象是一个用来浏览网页的用户界面面板

## 方法

| 序号 | 名称                           | 说明                                                  |
|:----:|:-------------------------------|:------------------------------------------------------|
|  1   | [Delete](#TaskPane.Delete)     | 此方法用于删除当前的 Taskpane 。                      |
|  2   | [Navigate](#TaskPane.Navigate) | 此方法用于在当前的 Taskpane 上重新载入url对应的网页。 |

## 属性

| 序号 | 名称                                   | 说明                                                     |
|:----:|:---------------------------------------|:---------------------------------------------------------|
|  1   | [DockPosition](#TaskPane.DockPosition) | 可读写，设置或返回一个整型，代表 Taskpane 的停靠位置。   |
|  2   | [Height](#TaskPane.Height)             | 返回一个整数，此整数代码 Taskpane 的高度。               |
|  3   | [ID](#TaskPane.ID)                     | 返回一个整数，此整数代码 Taskpane 的ID。                 |
|  4   | [Visible](#TaskPane.Visible)           | 可读写，设置或返回一个bool类型，代表 Taskpane 的可见性。 |
|  5   | [Width](#TaskPane.Width)               | 返回一个整数，此整数代码 Taskpane 的宽度。               |

## 成员方法

### TaskPane.Delete

此方法用于删除当前的 **Taskpane** 。

**语法**

**express.Delete ()**

\*express是一个代表 TaskPane 对象的变量。

### TaskPane.Navigate

此方法用于在当前的 **Taskpane** 上重新载入url对应的网页。

**语法**

**express.Navigate ( *url* )**

\*express是一个代表 TaskPane 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型 | 说明                                                     |
|-------|-----------|----------|----------------------------------------------------------|
| *url* | 必选      | String   | 这个url可以是一个在线的网址，也可以是一个本地的html资源. |

**示例**

``` JavaScript
let tp = Application.CreateTaskpane("https://www.kingsoft.com") if (tp){ tp.Visible = true tp.Navigate("https://www.wps.cn") }
```

## 成员属性

### TaskPane.DockPosition

可读写，设置或返回一个整型，代表 **Taskpane** 的停靠位置。

**语法**

**express.DockPosition**

\*express是一个代表 TaskPane 对象的变量。

**说明**

{ msoCTPDockPositionLeft : 0, msoCTPDockPositionTop : 1, msoCTPDockPositionRight : 2, msoCTPDockPositionBottom : 3, msoCTPDockPositionFloating : 4 }

### TaskPane.Height

返回一个整数，此整数代码 **Taskpane** 的高度。

**语法**

**express.Height**

\*express是一个代表 TaskPane 对象的变量。

### TaskPane.ID

返回一个整数，此整数代码 **Taskpane** 的ID。

**语法**

**express.ID**

\*express是一个代表 TaskPane 对象的变量。

### TaskPane.Visible

可读写，设置或返回一个bool类型，代表 **Taskpane** 的可见性。

**语法**

**express.Visible**

\*express是一个代表 TaskPane 对象的变量。

### TaskPane.Width

返回一个整数，此整数代码 **Taskpane** 的宽度。

**语法**

**express.Width**

\*express是一个代表 TaskPane 对象的变量。

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/任务窗格/TaskPane](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
