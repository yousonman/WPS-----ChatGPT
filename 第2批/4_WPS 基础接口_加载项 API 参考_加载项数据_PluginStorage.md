# PluginStorage 对象

## 简介

通过 PluginStorage 对象可以实现一个 WPS 加载项中的多个网页之间的数据共享。 注意通过 PluginStorage 对象只能传递简单的数据类型，JSON 数据建议转成字符串后再进行共享。

## 说明

PluginStorage 对象是 Application 对象的子对象，每一个 WPS 加载项都只有一个 PluginStorage 实例。 注意 PluginStorage 对象不能持久化，所有数据只在关闭 WPS 加载项之前有效。 数据持久化可以将数据写入本地文件中，参考 FileSystem 对象 。

``` JavaScript
let ps = Application.PluginStorage
```

下列代码示例往 PluginStorage 中存储一个key为"count", 值为5的代码，如果 PluginStorage 中key为count的代码已存在，则会用新值进行覆盖。

``` JavaScript
Application.PluginStorage.setItem("count", 5)
```

以下示例枚举出 PluginStorage 中所有的key。

``` JavaScript
let ps = Application.PluginStorage
let itemCounts = ps.length; for(let i = 0;  i < itemCounts; ++i){     let itemKey = ps.Key(i) //如果要取到对应的value，使用ps.GetItem(itemKey)     console.log(itemKey) }
```

有关 PluginStorage 的详细方法和属性，可以参考 PluginStorage 方法和属性相关介绍

## 方法

| 序号 | 名称                                    | 说明                        |
|:----:|:----------------------------------------|:----------------------------|
|  1   | [clear](#PluginStorage.clear)           | 清空容器中的所有键值对。    |
|  2   | [getItem](#PluginStorage.getItem)       | 返回key对应的value。        |
|  3   | [key](#PluginStorage.key)               | 返回index对应的key。        |
|  4   | [removeItem](#PluginStorage.removeItem) | 删除容器中key对应的键值对。 |
|  5   | [setItem](#PluginStorage.setItem)       | 向容器中添加键值对。        |

## 属性

| 序号 | 名称                            | 说明                                |
|:----:|:--------------------------------|:------------------------------------|
|  1   | [length](#PluginStorage.length) | 返回PluginStorage容器中键值对的个数 |

## 成员方法

### PluginStorage.clear

清空容器中的所有键值对。

**语法**

**express.clear ()**

\*express是一个代表 PluginStorage 对象的变量。

**示例**

``` JavaScript
下列代码示例清空PluginStorage中的元素。
Application.PluginStorage.clear()
```

### PluginStorage.getItem

返回key对应的value。

**语法**

**express.getItem ( *key* )**

\*express是一个代表 PluginStorage 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型     | 说明                                                                                          |
|-------|-----------|--------------|-----------------------------------------------------------------------------------------------|
| *key* | 必选      | 整数、字符串 | key的数据类型只能是简单类型，具体来说可以是整数或字符串，推荐使用具有命名意义的字符串作为key. |

**示例**

``` JavaScript
下列代码获取key为count的value并在控制台输出。
let value = Application.PluginStorage.getItem("count")
if (value){
    console.log(value)
}
```

### PluginStorage.key

返回index对应的key。

**语法**

**express.key ( *index* )**

\*express是一个代表 PluginStorage 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                                                                                      |
|---------|-----------|----------|-------------------------------------------------------------------------------------------|
| *index* | 必选      | 整数     | index的起始值从0开始，常见的用法是先用length属性取到元素个数，然后用此方法来得到对应的key |

**示例**

``` JavaScript
下列代码获取容器中第1个元素的key并在控制台输出。
let count = Application.PluginStorage.length
if (count > 0){
    let firstkey = Application.PluginStorage.key(0)
    console.log(firstkey)                           
}
```

### PluginStorage.removeItem

删除容器中key对应的键值对。

**语法**

**express.removeItem ( *key* )**

\*express是一个代表 PluginStorage 对象的变量。

**参数**

| 名称  | 必选/可选 | 数据类型     | 说明                                          |
|-------|-----------|--------------|-----------------------------------------------|
| *key* | 必选      | 整数、字符串 | 如果容器中找不到key对应的键值对，则什么都不做 |

**示例**

``` JavaScript
下列代码用于删除key为"count"的键值对。
Application.PluginStorage.removeItem("count")
```

### PluginStorage.setItem

向容器中添加键值对。

**语法**

**express.setItem ( *key* , *value* )**

\*express是一个代表 PluginStorage 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型     | 说明                                                                                          |
|---------|-----------|--------------|-----------------------------------------------------------------------------------------------|
| *key*   | 必选      | 整数、字符串 | key的数据类型只能是简单类型，具体来说可以是整数或字符串，推荐使用具有命名意义的字符串作为key. |
| *value* | 必选      | Variant      | {description}                                                                                 |

**示例**

``` JavaScript
下列代码用于将key为"count"，值为5的键值对存入容器。
Application.PluginStorage.setItem("count", 5)
```

## 成员属性

### PluginStorage.length

返回PluginStorage容器中键值对的个数

**语法**

**express.length**

\*express是一个代表 PluginStorage 对象的变量。

> WPS 开发文档： [WPS 基础接口/加载项 API 参考/加载项数据/PluginStorage](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
