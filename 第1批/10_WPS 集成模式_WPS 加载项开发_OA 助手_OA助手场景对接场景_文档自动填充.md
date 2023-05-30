# 文档自动填充

打开一个包含书签的模板文件后，OA助手在后台调用templateDataUrl获取模板数据的json字符串（书签定义及值）填充到对应的模板位置中。支持文字、链接、图片三种填充类型数值。

**\_wps.fillTemplate** ( *params* )

``` JavaScript
_WpsStartUp([{
    "OpenDoc": {
        "docId": "123",
        "uploadPath": uploadPath,
        "fileName": filePath,
        "picPath" : GetDemoPngPath(),
        "copyUrl" : backupPath,
    }
}])
```

## 输入参数

| 参数名     | 数据类型 | 描述                                             | 必选 |
|------------|----------|--------------------------------------------------|------|
| *OpenDoc*    | String   | 方法名，对应于OA助手dispatcher支持的方法         | 是   |
| *docId*      | String   | 文档ID，OA助手用以标记文档的信息，以区分其他文档 | 是   |
| *uploadPath* | String   | 保存文档上传路径                                 | 是   |
| *fileName*   | String   | 打开的文档路径                                   | 是   |
| *picPath*    | String   | 插入图片的路径                                   | 否   |
| *copyUrl*    | String   | 备份的服务器路径                                 | 否   |

## 文档填充公文场景展现

![](服务器端图像/填充模板.gif)

## 业务流程（执行逻辑）

![](服务器端图像/填充模板.png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/文档自动填充](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E6%96%87%E6%A1%A3%E8%87%AA%E5%8A%A8%E5%A1%AB%E5%85%85.html)

------------------------------------------------------------------------
