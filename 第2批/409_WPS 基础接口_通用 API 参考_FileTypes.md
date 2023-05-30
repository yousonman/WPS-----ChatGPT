# FileTypes 对象

## 简介

msoFileType 类型的值的集合，这些值决定在搜索期间返回的文件类型。

## 说明

所有搜索都只有一个 FileTypes 集合，因此在执行搜索之前清除 FileTypes 集合很重要，除非希望搜索以前搜索的文件类型。清除该集合的最简便方法是将 FileType 属性设置为要搜索的第一个文件类型。还可以使用 Remove 方法删除单个类型。要确定集合中每一项的文件类型，请使用 Item 方法返回 msoFileType 值。

## 方法

| 序号 | 名称                        | 说明                             |
|:----:|:----------------------------|:---------------------------------|
|  1   | [Add](#FileTypes.Add)       | 向文件搜索中新添一个文件类型。   |
|  2   | [Remove](#FileTypes.Remove) | 从集合中删除一个 FileType 对象。 |

## 属性

| 序号 | 名称                                  | 说明                                                                                                                             |
|:----:|:--------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------|
|  1   | [Application](#FileTypes.Application) | 获取一个 Application 对象，代表 FileTypes 对象的容器应用程序（可以使用 Automation 对象的此属性返回该对象的容器应用程序）。只读。 |
|  2   | [Creator](#FileTypes.Creator)         | 获取一个 32 位整数，指示创建 FileTypes 对象时所使用的应用程序。只读。                                                            |
|  3   | [Item](#FileTypes.Item)               | 获取一个指示要搜索的文件类型的值。只读。                                                                                         |

## 成员方法

### FileTypes.Add

向文件搜索中新添一个文件类型。

**语法**

**express.Add ( *FileType* )**

\*express是一个代表 FileTypes 对象的变量。

**参数**

| 名称       | 必选/可选 | 数据类型    | 说明                   |
|------------|-----------|-------------|------------------------|
| *FileType* | 必选      | MsoFileType | 指定要搜索的文件类型。 |

### FileTypes.Remove

从集合中删除一个 **FileType** 对象。

**语法**

**express.Remove ( *Index* )**

\*express是一个代表 FileTypes 对象的变量。

**参数**

| 名称    | 必选/可选 | 数据类型 | 说明                       |
|---------|-----------|----------|----------------------------|
| *Index* | 必选      | Long     | 要删除的文件类型的索引号。 |

## 成员属性

### FileTypes.Application

获取一个 **Application** 对象，代表 **FileTypes** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

\*express是一个代表 FileTypes 对象的变量。

### FileTypes.Creator

获取一个 32 位整数，指示创建 **FileTypes** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

\*express是一个代表 FileTypes 对象的变量。

**类型**

Long

### FileTypes.Item

获取一个指示要搜索的文件类型的值。只读。

**语法**

**express.Item**

\*express是一个代表 FileTypes 对象的变量。

> WPS 开发文档： [WPS 基础接口/通用 API 参考/FileTypes](https://qn.cache.wpscdn.cn/encs/doc/office_v19/index.htm)

------------------------------------------------------------------------
