# ColorStops 对象

## 简介

指定的数据系列中所有 ColorStop 对象的集合。

## 说明

``` JavaScript
/*以下示例显示具有 LinearGradients 的 ColorStops。*/
function test(){
let myInterior = Selection.Interior
    myInterior.Pattern = xlPatternLinearGradient
    myInterior.Gradient.Degree = 90
    myInterior.Gradient.ColorStops.Clear()

    //adds stops after any have been cleared
let myColorStops1 = Selection.Interior.Gradient.ColorStops.Add(0)
    myColorStops1.ThemeColor = xlThemeColorDark1
    myColorStops1.TintAndShade = 0

let myColorStops2 = Selection.Interior.Gradient.ColorStops.Add(1)
    myColorStops2.ThemeColor = xlThemeColorAccent1
    myColorStops2.TintAndShade = 0
}
```

每个 ColorStop 对象代表一个区域或选定内容中渐变填充的一个颜色光圈。

## 方法

| 序号 | 名称                           | 说明 |
|:----:|:-------------------------------|:-----|
|  1   | [Reserve](#ColorStops.Reserve) |      |

## 属性

| 序号 | 名称                           | 说明 |
|:----:|:-------------------------------|:-----|
|  1   | [Reserve](#ColorStops.Reserve) |      |

## 成员方法

### ColorStops.Reserve

**语法**

**express.Reserve ()**

\*express是一个代表 ColorStops 对象的变量。

## 成员属性

### ColorStops.Reserve

**语法**

**express.Reserve**

\*express是一个代表 ColorStops 对象的变量。

> WPS 开发文档： [WPS 基础接口/表格 API 参考/ColorStops](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
