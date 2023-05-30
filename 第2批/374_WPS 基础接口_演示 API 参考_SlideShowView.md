# SlideShowView 对象

## 简介

代表幻灯片放映窗口中的视图。

## 方法

<table>
<colgroup>
<col style="width: 33%" />
<col style="width: 33%" />
<col style="width: 33%" />
</colgroup>
<thead>
<tr class="header" style="text-align:center;vertical-align:middle;">
<th style="text-align: center;">序号</th>
<th style="text-align: left;">名称</th>
<th style="text-align: left;">说明</th>
</tr>
</thead>
<tbody>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">1</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.Exit">Exit</a></td>
<td style="text-align: left;" data-valian="middle"><p>结束指定的幻灯片放映。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">2</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.First">First</a></td>
<td style="text-align: left;" data-valian="middle"><p>设置指定的幻灯片放映视图，显示演示文稿的第一张幻灯片。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">3</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.GetClickIndex">GetClickIndex</a></td>
<td style="text-align: left;" data-valian="middle"><p>返回启动当前正在幻灯片上播放或者刚刚已经播放完的动画的当前鼠标单击的索引号。</p>
<p>返回值类型为Long</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">4</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.GotoSlide">GotoSlide</a></td>
<td style="text-align: left;" data-valian="middle"><p>在幻灯片放映期间切换到指定幻灯片。可指定是否要重运行动画效果。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">5</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.Last">Last</a></td>
<td style="text-align: left;" data-valian="middle"><p>将指定的幻灯片放映视图设置为显示演示文稿中的上一张幻灯片。</p></td>
</tr>
<tr class="even" data-editname="functions">
<td style="text-align: center;" data-valian="middle">6</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.Next">Next</a></td>
<td style="text-align: left;" data-valian="middle"><p>显示紧接当前显示的幻灯片之后的幻灯片。如果显示的是最后一张幻灯片，则 Next 方法会关闭演讲者模式的幻灯片放映并以展台模式返回到第一张幻灯片。使用 SlideShowWindow 对象的 View 属性可返回 SlideShowView 对象。</p></td>
</tr>
<tr class="odd" data-editname="functions">
<td style="text-align: center;" data-valian="middle">7</td>
<td style="text-align: left;" data-valian="middle"><a href="#SlideShowView.Previous">Previous</a></td>
<td style="text-align: left;" data-valian="middle"><p>显示紧邻当前显示的幻灯片之前的幻灯片。如果当前以展台幻灯片放映显示第一张幻灯片，使用 Previous 方法将在幻灯片放映中切换到最后一张幻灯片；如果当前显示的是演示文稿中的第一张幻灯片，则此方法无效。使用 SlideShowWindow 对象的 View 属性可返回 SlideShowView 对象。</p></td>
</tr>
</tbody>
</table>

## 属性

| 序号 | 名称                                                              | 说明                                                                                                                                        |
|:----:|:------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------|
|  1   | [AdvanceMode](#SlideShowView.AdvanceMode)                         | 返回一个值，该值指示如何在指定视图中切换幻灯片放映。只读 PpSlideShowAdvanceMode 类型。                                                      |
|  2   | [Application](#SlideShowView.Application)                         | 返回一个 Application 对象，该对象表示指定对象的创建者。                                                                                     |
|  3   | [CurrentShowPosition](#SlideShowView.CurrentShowPosition)         | 返回当前幻灯片在指定视图下显示的幻灯片放映中的位置。只读 Long 类型。                                                                        |
|  4   | [LastSlideViewed](#SlideShowView.LastSlideViewed)                 | 返回一个 Slide 对象，该对象代表指定的幻灯片放映视图中在当前幻灯片之间观看的幻灯片。                                                         |
|  5   | [PointerColor](#SlideShowView.PointerColor)                       | 返回一个 ColorFormat 对象，该对象代表在幻灯片放映过程中指定演示文稿的指针颜色。一旦幻灯片放映结束，该颜色将转变为演示文稿的默认颜色。只读。 |
|  6   | [PointerType](#SlideShowView.PointerType)                         | 返回或设置在幻灯片放映中使用的指针类型。可读/写 PpSlideShowPointerType 类型。                                                               |
|  7   | [PresentationElapsedTime](#SlideShowView.PresentationElapsedTime) | 返回指定的幻灯片放映开始后经过的秒数。只读 Long 类型。                                                                                      |
|  8   | [Slide](#SlideShowView.Slide)                                     | 返回一个 Slide 对象，该对象代表在指定的幻灯片放映窗口视图中当前显示的幻灯片。只读。                                                         |
|  9   | [SlideElapsedTime](#SlideShowView.SlideElapsedTime)               | 返回当前幻灯片已放映的秒数。可读/写 Long 类型。                                                                                             |
|  10  | [State](#SlideShowView.State)                                     | 返回或设置幻灯片放映的状态。可读/写 PpSlideShowState 类型。                                                                                 |
|  11  | [Zoom](#SlideShowView.Zoom)                                       | 以正常大小的百分比形式返回指定幻灯片放映窗口视图的缩放设置。可以为 10% 到 400% 之间的值。只读 Integer 类型。                                |

## 成员方法

### SlideShowView.Exit

结束指定的幻灯片放映。

**语法**

**express.Exit ()**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*以下示例将结束第一个幻灯片放映窗口中的幻灯片放映。*/
Application.SlideShowWindows.Item(1).View.Exit()
```

### SlideShowView.First

设置指定的幻灯片放映视图，显示演示文稿的第一张幻灯片。

**语法**

**express.First ()**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

如果使用 **First** 方法在幻灯片放映过程中从一张幻灯片切换到另一张幻灯片，返回原幻灯片时，它的动画继续从被中断处开始放映。

**示例**

``` JavaScript
本示例设置第一个幻灯片放映窗口以显示演示文稿的第一张幻灯片。
Application.SlideShowWindows.Item(1).View.First()
```

### SlideShowView.GetClickIndex

返回启动当前正在幻灯片上播放或者刚刚已经播放完的动画的当前鼠标单击的索引号。

返回值类型为Long

**语法**

**express.GetClickIndex ()**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

使用 **GetClickCount** 方法可返回为幻灯片定义的鼠标单击次数。

如果幻灯片不包含任何动画或者用户尚未换到某个动画，则 **GetClickIndex** 方法返回 0。如果幻灯片包含自动播放的动画并且该用户移到上一页，则 **GetClickIndex** 方法返回 **msoClickStateBeforeAutomaticAnimcations** 。

**示例**

``` JavaScript
Application.SlideShowWindows.Item(1).View.GetClickIndex()
```

### SlideShowView.GotoSlide

在幻灯片放映期间切换到指定幻灯片。可指定是否要重运行动画效果。

**语法**

**express.GotoSlide ( *Index* , *ResetSlide* )**

\*express是一个代表 SlideShowView 对象的变量。

**参数**

| 名称         | 必选/可选 | 数据类型    | 说明                                                                                                                                                                                                                                                                              |
|--------------|-----------|-------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| *Index*      | 必选      | Integer     | 要切换到的幻灯片的编号。                                                                                                                                                                                                                                                          |
| *ResetSlide* | 可选      | MsoTriState | 如果将 ResetSlide 设置为 msoFalse，那么在幻灯片放映期间，当您从一张幻灯片切换到另一张，再返回到第一张幻灯片时，动画将从刚刚中断的位置继续播放。如果将 ResetSlide 设置为 msoTrue，当从一张幻灯片切换到另一张幻灯片，再返回到第一张幻灯片时，将重新播放整个动画。默认值为 msoTrue。 |

**说明**

|                                               |
|-----------------------------------------------|
| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 默认值。                          |

**示例**

``` JavaScript
/*本示例在第一个幻灯片放映窗口中从当前幻灯片切换到第三张幻灯片。如果在幻灯片放映过程中切换回当前幻灯片，将重新播放整个动画。*/
Application.SlideShowWindows.Item(1).View.GotoSlide(3)

/*本示例在第一个幻灯片放映窗口中从当前幻灯片切换到第三张幻灯片。如果在幻灯片放映过程中切换回当前幻灯片，动画将从中断位置重新播放。*/
Application.SlideShowWindows.Item(1).View.GotoSlide(3, msoFalse)
```

### SlideShowView.Last

将指定的幻灯片放映视图设置为显示演示文稿中的上一张幻灯片。

**语法**

**express.Last ()**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

如果使用 **Last** 方法在幻灯片放映过程中从一张幻灯片切换到另一张幻灯片，则在返回原幻灯片时，其中的动画继续从被中断处开始放映。

**示例**

``` JavaScript
/*本例将第一个幻灯片放映窗口设置为显示演示文稿中的上一张幻灯片。*/
Application.SlideShowWindows.Item(1).View.Last()
```

### SlideShowView.Next

显示紧接当前显示的幻灯片之后的幻灯片。如果显示的是最后一张幻灯片，则 **Next** 方法会关闭演讲者模式的幻灯片放映并以展台模式返回到第一张幻灯片。使用 **SlideShowWindow** 对象的 **View** 属性可返回 **SlideShowView** 对象。

**语法**

**express.Next ()**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*本示例显示第一个幻灯片放映窗口中紧随在当前显示的幻灯片之后的下一张幻灯片。*/
Application.SlideShowWindows.Item(1).View.Next()
```

### SlideShowView.Previous

显示紧邻当前显示的幻灯片之前的幻灯片。如果当前以展台幻灯片放映显示第一张幻灯片，使用 **Previous** 方法将在幻灯片放映中切换到最后一张幻灯片；如果当前显示的是演示文稿中的第一张幻灯片，则此方法无效。使用 **SlideShowWindow** 对象的 **View** 属性可返回 **SlideShowView** 对象。

**语法**

**express.Previous ()**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*本示例在第一个幻灯片放映窗口中放映紧挨当前显示的幻灯片之前的幻灯片。*/
Application.SlideShowWindows.Item(1).View.Previous()
```

## 成员属性

### SlideShowView.AdvanceMode

返回一个值，该值指示如何在指定视图中切换幻灯片放映。只读 **PpSlideShowAdvanceMode** 类型。

**语法**

**express.AdvanceMode**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

|                                                                     |
|---------------------------------------------------------------------|
| PpSlideShowAdvanceMode 可以是下列 PpSlideShowAdvanceMode 常量之一。 |
| **ppSlideShowManualAdvance**                                        |
| **ppSlideShowRehearseNewTimings**                                   |
| **ppSlideShowUseSlideTimings**                                      |

### SlideShowView.Application

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/
function test(pptPres) {
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}

/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/
function test() {
  let shpOle = Application.ActivePresentation.Slides.Item(3).Shapes
  for(let i = 1; i <= shpOle.Count; i++) {
    if(shpOle.Item(i).Type == msoLinkedOLEObject) {
        MsgBox(shpOle.Item(i).OLEFormat.Application.Name)
    }
  }
}
```

### SlideShowView.CurrentShowPosition

返回当前幻灯片在指定视图下显示的幻灯片放映中的位置。只读 **Long** 类型。

**语法**

**express.CurrentShowPosition**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

如果指定视图包含自定义放映，则 **CurrentShowPosition** 属性会返回当前幻灯片在自定义放映中的位置，而不是在整个演示文稿中的位置。

**示例**

``` JavaScript
/*本示例将一个变量设置为当前幻灯片在第一个幻灯片放映窗口中放映的幻灯片放映中的位置。*/
let lastSlideSeen = Application.SlideShowWindows.Item(1).View.CurrentShowPosition
```

### SlideShowView.LastSlideViewed

返回一个 **Slide** 对象，该对象代表指定的幻灯片放映视图中在当前幻灯片之间观看的幻灯片。

**语法**

**express.LastSlideViewed**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*本示例切换到第一个幻灯片放映窗口中当前幻灯片之前的幻灯片。*/
Application.SlideShowWindows.Item(1).View.GotoSlide(SlideShowWindows.Item(1).View.LastSlideViewed.SlideIndex)
```

### SlideShowView.PointerColor

返回一个 **ColorFormat** 对象，该对象代表在幻灯片放映过程中指定演示文稿的指针颜色。一旦幻灯片放映结束，该颜色将转变为演示文稿的默认颜色。只读。

**语法**

**express.PointerColor**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

要将指针更改为绘图笔，可将 **PointerType** 属性设置为 **ppSlideShowPointerPen** 。

### SlideShowView.PointerType

返回或设置在幻灯片放映中使用的指针类型。可读/写 **PpSlideShowPointerType** 类型。

**语法**

**express.PointerType**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

|                                                                     |
|---------------------------------------------------------------------|
| PpSlideShowPointerType 可以是下列 PpSlideShowPointerType 常量之一。 |
| **ppSlideShowPointerAlwaysHidden**                                  |
| **ppSlideShowPointerArrow**                                         |
| **ppSlideShowPointerAutoArrow**                                     |
| **ppSlideShowPointerNone**                                          |
| **ppSlideShowPointerPen**                                           |

**示例**

``` JavaScript
/*本示例执行当前演示文稿的一个幻灯片放映，将指针改变为绘图笔。*/
function test() {
  let currView = Application.ActivePresentation.SlideShowSettings.Run().View
      currView.PointerType = ppSlideShowPointerPen
}
```

### SlideShowView.PresentationElapsedTime

返回指定的幻灯片放映开始后经过的秒数。只读 **Long** 类型。

**语法**

**express.PresentationElapsedTime**

\*express是一个代表 SlideShowView 对象的变量。

**示例**

``` JavaScript
/*如果幻灯片放映开始已有五分钟，本示例转向第一个幻灯片放映窗口中第七张幻灯片。*/
function test() {
  if(Application.SlideShowWindows.Item(1).View.PresentationElapsedTime > 3) {
      Application.SlideShowWindows.Item(1).View.GotoSlide(3)
  }
}
```

### SlideShowView.Slide

返回一个 **Slide** 对象，该对象代表在指定的幻灯片放映窗口视图中当前显示的幻灯片。只读。

**语法**

**express.Slide**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

如果当前显示的幻灯片来源于一个嵌入演示文稿，则可以使用 Slide 对象（从 **Slide** 属性返回）的 **Parent** 属性返回包含该幻灯片的嵌入演示文稿。（ SlideShowWindow 对象或 DocumentWindow 对象的 Presentation 属性返回创建该窗口的演示文稿，而不是该嵌入演示文稿。）

### SlideShowView.SlideElapsedTime

返回当前幻灯片已放映的秒数。可读/写 **Long** 类型。

**语法**

**express.SlideElapsedTime**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

使用 **ResetSlideTime** 方法重置当前放映的幻灯片的播放时间。

**示例**

``` JavaScript
/*本示例为正在第一个幻灯片放映窗口中放映的幻灯片的播放时间设置一个变量，并显示该变量的值。*/
function test() {
  let currTime = Application.SlideShowWindows.Item(1).View.SlideElapsedTime
  alert(currTime)
}
```

### SlideShowView.State

返回或设置幻灯片放映的状态。可读/写 **PpSlideShowState** 类型。

**语法**

**express.State**

\*express是一个代表 SlideShowView 对象的变量。

**说明**

|                                                         |
|---------------------------------------------------------|
| PpSlideShowState 可以是下列 PpSlideShowState 常量之一。 |
| **ppSlideShowBlackScreen**                              |
| **ppSlideShowDone**                                     |
| **ppSlideShowPaused**                                   |
| **ppSlideShowRunning**                                  |
| **ppSlideShowWhiteScreen**                              |

**示例**

``` JavaScript
/*本示例将第一个幻灯片放映窗口的视图状态设为黑屏。*/
Application.SlideShowWindows.Item(1).View.State = ppSlideShowBlackScreen
```

### SlideShowView.Zoom

以正常大小的百分比形式返回指定幻灯片放映窗口视图的缩放设置。可以为 10% 到 400% 之间的值。只读 **Integer** 类型。

**语法**

**express.Zoom**

\*express是一个代表 SlideShowView 对象的变量。

> WPS 开发文档： [WPS 基础接口/演示 API 参考/SlideShowView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
