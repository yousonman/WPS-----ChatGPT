# 文档格式转换

打开文档，将文档转换成PDF格式或OFD格式，再进行上传到服务器

``` JavaScript
_wps.openDoc({
    "docId" :"c2de1fcd1d3e4ac0b3cda1392c36c9",
    "uploadPath" :  "https://***",
    "fileName" :  "https://***",
    "suffix" : ".pdf ",
    "uploadAppendPath" :  "https://***",
    "uploadFieldName" : "file"
});
```

## 输入参数

| 参数名           | 数据类型 | 描述                                                       | 必选 |
|------------------|----------|---------------------------------------------------------------------------|------|
| *docId*            | String   | 文档ID，文档在数据库中的ID                                            | 是   |
| *fileName*         | String   | 获取服务器文档接口（不传即为新建空文档）                                  | 是   |
| *uploadPath*       | String   | 新建空文档文件名,不传默认新建文档以系统时间为名称                            | 是   |
| *suffix*           | String   | 要由Doc转到其他格式的后缀。包括：pdf、out、ofd三种。                              | 是   |
| *uploadAppendPath* | String   | 用户转其他文件格式时（如转PDF格式保存），OA助手会把转换格式后的文件，上载到这个路径上。不传默认为uploadPath的值 | 否   |
| *uploadFieldName*  | String   | 配合上传路径，包装上传form表单时，对应上传的字段域名称。不传默认为“file”   | 否   |

## 文档转换格式上传场景展现

![](服务器端图像/转换格式上传.gif)

## 业务流程（执行逻辑）

![](服务器端图像/文档转换上传.png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/文档格式转换](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E6%96%87%E6%A1%A3%E6%A0%BC%E5%BC%8F%E8%BD%AC%E6%8D%A2.html)

------------------------------------------------------------------------
