# View 对象

## 简介

代表指定文档窗口中当前编辑的视图。

## 说明

View 对象可以代表任何文档窗口视图：普通视图、幻灯片视图、大纲视图、幻灯片浏览视图、备注页视图、幻灯片母版视图、讲义母版视图或备注母版视图。 View 对象的某些属性和方法仅在特定视图中使用。如果试图使用不适合 View 对象的属性或方法，将产生错误。

## 属性

| 序号 | 名称                               | 说明                                                                                             |
|:----:|:-----------------------------------|:-------------------------------------------------------------------------------------------------|
|  1   | [Application](#View.Application)   | 返回一个 Application 对象，该对象表示指定对象的创建者。                                          |
|  2   | [PrintOptions](#View.PrintOptions) | 返回一个 PrintOptions 对象，该对象代表与指定演示文稿共同保存的打印选项。只读。                   |
|  3   | [Slide](#View.Slide)               | 返回或设置一个 Slide 对象，该对象代表在指定文档窗口视图中当前显示的幻灯片。可读/写。             |
|  4   | [Type](#View.Type)                 | 返回一个 PpViewType 常量，该常量代表视图的类型。只读。                                           |
|  5   | [Zoom](#View.Zoom)                 | 以百分比形式返回或设置指定视图的缩放设置。可以为 10% 到 400% 之间的值。 Integer 类型，可读/写 。 |

## 成员属性

### View.Application

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/
function test() {
    let pptPres = Application.ActivePresentation
    pptPres.Slides.Add(1, 1)
    pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")
}
```

### View.PrintOptions

返回一个 PrintOptions 对象，该对象代表与指定演示文稿共同保存的打印选项。只读。

**语法**

**express.PrintOptions**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*以下示例打印当前演示文稿中的隐藏幻灯片，并调整幻灯片的尺寸以适应纸张大小*/
function test(){
    let pres = Application.ActivePresentation
    let opt = pres.PrintOptions
    opt.PrintHiddenSlides = true
    opt.FitToPage = true
    pres.PrintOut()
}
```

### View.Slide

返回或设置一个 Slide 对象，该对象代表在指定文档窗口视图中当前显示的幻灯片。可读/写。

**语法**

**express.Slide**

\*express是一个代表 View 对象的变量。

**说明**

如果当前显示的幻灯片来源于一个嵌入演示文稿，则可以使用 **Slide** 对象（从 **Slide** 属性返回）的 **Parent** 属性返回包含该幻灯片的嵌入演示文稿。（ **SlideShowWindow** 对象或 **DocumentWindow** 对象的 **Presentation** 属性返回创建该窗口的演示文稿，而不是该嵌入演示文稿。）

**示例**

``` JavaScript
/*本示例显示当前窗口的第二个幻灯片的名称*/
Application.ActivePresentation.Slides.Item(2).Name
```

### View.Type

返回一个 PpViewType 常量，该常量代表视图的类型。只读。

**语法**

**express.Type**

\*express是一个代表 View 对象的变量。

**说明**

PpViewType 可以是下列 PpViewType 常量之一。 ppViewHandoutMaster ppViewMasterThumbnails ppViewNormal ppViewNotesMaster ppViewNotesPage ppViewOutline ppViewPrintPreview ppViewSlide ppViewSlideMaster ppViewSlideSorter ppViewThumbnails ppViewTitleMaster

### View.Zoom

以百分比形式返回或设置指定视图的缩放设置。可以为 10% 到 400% 之间的值。 **Integer** 类型，可读/写 。

**语法**

**express.Zoom**

\*express是一个代表 View 对象的变量。

**示例**

``` JavaScript
/*以下示例将第一个文档窗口的视图缩小为 30%*/
Application.ActivePresentation.Windows.Item(1).View.Zoom = 30
```

> WPS 开发文档： [WPS 基础接口/演示 API 参考/View](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
