# PlaySettings 对象

## 简介

包含关于指定的媒体剪辑在幻灯片放映中的播放方式的信息

## 属性

| 序号 | 名称                                                     | 说明                                                                                                                                                                                                                                         |
|:----:|:---------------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [ActionVerb](#PlaySettings.ActionVerb)                   | 返回或设置包含 OLE 动作的字符串，该动作将在幻灯片放映过程中当指定的 OLE 对象进行动画显示时运行。默认动作指定 OLE 对象在上一个动画或幻灯片切换之后运行的动作（例如，播放一波形文件，或显示数据使得用户可以进行修改）。 String 类型，可读/写。 |
|  2   | [Application](#PlaySettings.Application)                 | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                                                                                                                      |
|  3   | [HideWhileNotPlaying](#PlaySettings.HideWhileNotPlaying) | 决定在幻灯片放映期间指定的媒体剪辑在不播放时是否隐藏。可读/写 MsoTriState 类型。                                                                                                                                                             |
|  4   | [LoopUntilStopped](#PlaySettings.LoopUntilStopped)       | 决定指定影片或声音是否持续循环播放，直到出现下列情况：开始下一个影片或声音、用户单击幻灯片或发生幻灯片切换。 MsoTriState 类型，可读/写。                                                                                                     |
|  5   | [Parent](#PlaySettings.Parent)                           | 返回指定对象的父对象。                                                                                                                                                                                                                       |
|  6   | [PauseAnimation](#PlaySettings.PauseAnimation)           | 决定在指定的媒体剪辑播放结束前是否暂停幻灯片放映。可读/写 MsoTriState 类型。                                                                                                                                                                 |
|  7   | [PlayOnEntry](#PlaySettings.PlayOnEntry)                 | 决定指定的影片或声音后是否在激活后自动播放。可读/写 MsoTriState 类型。                                                                                                                                                                       |
|  8   | [RewindMovie](#PlaySettings.RewindMovie)                 | 决定在指定的影片播放完毕后是否自动重新显示该影片的第一帧。可读/写 MsoTriState 类型。                                                                                                                                                         |
|  9   | [StopAfterSlides](#PlaySettings.StopAfterSlides)         | 返回或设置在媒体剪辑播放完毕前要放映的幻灯片数。可读/写 Long 类型。                                                                                                                                                                          |

## 成员属性

### PlaySettings.ActionVerb

返回或设置包含 OLE 动作的字符串，该动作将在幻灯片放映过程中当指定的 OLE 对象进行动画显示时运行。默认动作指定 OLE 对象在上一个动画或幻灯片切换之后运行的动作（例如，播放一波形文件，或显示数据使得用户可以进行修改）。 **String** 类型，可读/写。

**语法**

**express.ActionVerb**

\*express是一个代表 PlaySettings 对象的变量。

**示例**

``` JavaScript
/*本示例指定当前演示文稿第一张幻灯片的第三个形状被激活时将自动打开以供编辑。但第三个形状必须是包含声音或视频对象并支持“Edit”动词的 OLE 对象。*/
function test(){
　　　　let OLEobj = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
　　　　let oleobj = OLEobj.AnimationSettings.PlaySettings
  　　　　  oleobj.PlayOnEntry = true
  　　　　  oleobj.ActionVerb = "Edit"
}
```

### PlaySettings.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 PlaySettings 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

``` JavaScript
function AddAndSave(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

``` JavaScript
function test() {
    let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes
    for (let i = 1; i <= shpOle.Count; i++) {
        if (shpOle.Item(i).Type == msoLinkedOLEObject) {
            alert(shpOle.Item(i).OLEFormat.Application.Name)
        }
    }
}
```

### PlaySettings.HideWhileNotPlaying

决定在幻灯片放映期间指定的媒体剪辑在不播放时是否隐藏。可读/写 **MsoTriState** 类型。

**语法**

**express.HideWhileNotPlaying**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。          |
|--------------------------------------------------------|
| msoCTrue                                               |
| msoFalse                                               |
| msoTriStateMixed                                       |
| msoTriStateToggle                                      |
| msoTrue 幻灯片放映过程中指定的媒体剪辑在不播放时隐藏。 |

**示例**

``` JavaScript
/*以下示例在当前演示文稿的第一张幻灯片中插入名为“Clock.avi”的影片，设置它在幻灯片切换后自动播放，并指定该影片对象在不播放时隐藏。*/
function test(){
　　　　let mydocument = Application.ActivePresentation.Slides.Item(1).Shapes.AddOLEObject(10,10,250,250,undefined,"C:\\WINNT\\clock.avi")
　　　　let mydocument1 = mydocument.AnimationSettings.PlaySettings
 　　　　   mydocument1.PlayOnEntry = true
 　　　　   mydocument1.HideWhileNotPlaying = msoTrue
}
```

### PlaySettings.LoopUntilStopped

决定指定影片或声音是否持续循环播放，直到出现下列情况：开始下一个影片或声音、用户单击幻灯片或发生幻灯片切换。 **MsoTriState** 类型，可读/写。

**语法**

**express.LoopUntilStopped**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                                              |
|--------------------------------------------------------------------------------------------|
| msoCTrue                                                                                   |
| msoFalse                                                                                   |
| msoTriStateMixed                                                                           |
| msoTriStateToggle                                                                          |
| msoTrue 指定影片或声音持续循环，直到开始下一个影片或声音、用户单击幻灯片或发生幻灯片切换。 |

**示例**

本示例将当前演示文稿第一张幻灯片第三个形状设为以动画顺序开始播放，并持续播放直到开始下一个媒体剪辑。第三个形状必须是声音或影片对象。

``` JavaScript
Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).AnimationSettings.PlaySettings.LoopUntilStopped = msoTrue
```

### PlaySettings.Parent

返回指定对象的父对象。

**语法**

**express.Parent**

\*express是一个代表 PlaySettings 对象的变量。

**示例**

``` JavaScript
/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/
function test(){
　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes
　　　　let mydocument = myShapes.AddShape(msoShapeOval,50,50,300,150).TextFrame
   　　　　 mydocument.TextRange.Text = "Test text"
   　　　　 mydocument.Parent.Rotation = 45
}
```

### PlaySettings.PauseAnimation

决定在指定的媒体剪辑播放结束前是否暂停幻灯片放映。可读/写 **MsoTriState** 类型。

**语法**

**express.PauseAnimation**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

要使 **PauseAnimation** 属性设置生效，指定形状的 **PlayOnEntry** 属性必须设置为 **msoTrue** 。

| MsoTriState 可以是下列 MsoTriState 常量之一。      |
|----------------------------------------------------|
| msoCTrue                                           |
| msoFalse 在背景中播放媒体剪辑时继续幻灯片放映。    |
| msoTriStateMixed                                   |
| msoTriStateToggle                                  |
| msoTrue 在指定的媒体剪辑播放结束前暂停幻灯片放映。 |

**示例**

``` JavaScript
/*本示例指定当前演示文稿第一张幻灯片的第三个形状被激活后自动播放，并且当背景中播放该影片时暂停播放幻灯片放映。第三个形状必须是声音或影片对象。*/
function test(){
　　　　let OLEobj = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
　　　　let OLEobj1 = OLEobj.AnimationSettings.PlaySettings
   　　　　 OLEobj1.PlayOnEntry = msoTrue
    　　　　OLEobj1.PauseAnimation = msoTrue
}
```

### PlaySettings.PlayOnEntry

决定指定的影片或声音后是否在激活后自动播放。可读/写 **MsoTriState** 类型。

**语法**

**express.PlayOnEntry**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

将此属性设置为 **msoTrue** ，从而把 **AnimationSettings** 对象的 **Animate** 属性设为 **msoTrue** 。将 **Animate** 属性设置为 **msoFalse** 可自动将 **PlayOnEntry** 属性设为 **msoFalse** 。

使用 **ActionVerb** 属性设置媒体剪辑被激活时调用的动作。

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
|-----------------------------------------------|
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue 指定的影片或声音被激活后自动播放。    |

**示例**

``` JavaScript
/*本示例指定当前演示文稿第一张幻灯片的第三个形状在激活后自动播放。第三个形状必须是声音或影片对象。*/
function test(){
　　　　let OLEobj = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
　　　　OLEobj.AnimationSettings.PlaySettings.PlayOnEntry = msoTrue
}
```

### PlaySettings.RewindMovie

决定在指定的影片播放完毕后是否自动重新显示该影片的第一帧。可读/写 **MsoTriState** 类型。

**语法**

**express.RewindMovie**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。              |
|------------------------------------------------------------|
| msoCTrue                                                   |
| msoFalse                                                   |
| msoTriStateMixed                                           |
| msoTriStateToggle                                          |
| msoTrue 在指定的影片放映完毕后自动重新显示该影片的第一帧。 |

**示例**

``` JavaScript
/*本示例指定影片的第一帧（由当前演示文稿第一张幻灯片的第三个形状代表）在影片播放完毕后重新显示。第三个形状必须是影片对象。*/
function test(){
　　　　let OLEobj = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
　　　　OLEobj.AnimationSettings.PlaySettings.RewindMovie = msoTrue
}
```

### PlaySettings.StopAfterSlides

返回或设置在媒体剪辑播放完毕前要放映的幻灯片数。可读/写 **Long** 类型。

**语法**

**express.StopAfterSlides**

\*express是一个代表 PlaySettings 对象的变量。

**说明**

要使 **StopAfterSlides** 属性设置生效，必须将指定幻灯片的 **PauseAnimation** 属性设为 **False** ，并且必须将 **PlayOnEntry** 属性设为 **True** 。

当指定数目的幻灯片放映结束或剪辑播放完毕时（无论哪种情况先发生），媒体剪辑将停止播放。0（零）值指定在放映完当前幻灯片后剪辑停止播放。

**示例**

``` JavaScript
/*本例指定当媒体剪辑（由活动演示文稿中第一张幻灯片上的第三个形状代表）被激活时自动播放，并且当背景中播放该媒体剪辑的同时幻灯片将继续放映，还指定当放映完三张幻灯片或该剪辑播放完毕时（无论哪种情况先发生），媒体剪辑将停止播放。第三个形状必须是声音或影片对象。*/
function test(){
　　　　let OLEobj = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)
　　　　let OLEobj1 = OLEobj.AnimationSettings.PlaySettings
   　　　　 OLEobj1.PlayOnEntry = true
   　　　　 OLEobj1.PauseAnimation = false
   　　　　 OLEobj1.StopAfterSlides = 3
}
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/PlaySettings](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
