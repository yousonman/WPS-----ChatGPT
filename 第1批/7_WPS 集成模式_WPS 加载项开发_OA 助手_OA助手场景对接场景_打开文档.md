# 打开文档

打开指定文档，可以通过参数指定保护模式，格式转换保存，设置用户名等

``` JavaScript
_WpsStartUp([{
    "OpenDoc": {
        "docId": "123",
        "uploadPath": uploadPath,
        "fileName": filePath,
        "docPassword": {
            "docPassword": docPassword
        }
    }
}])
```

| 参数名      | 数据类型 | 描述                                     | 必选 |
|-------------|----------|------------------------------------------|------|
| *OpenDoc*     | String   | 方法名，对应于OA助手dispatcher支持的方法 | 是   |
| *docId*       | String   | 文档ID                                   | 是   |
| *uploadPath*  | String   | 保存文档上传路径                         | 是   |
| *fileName*    | String   | 打开的文档路径                           | 是   |
| *docPassword* | String   | 文档密码                                 | 是   |

## 打开指定文档场景展现

![](服务器端图像/打开指定文件.gif)

## 业务流程（执行逻辑）

![](服务器端图像/文档流程图.png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/打开文档](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E6%89%93%E5%BC%80%E6%96%87%E6%A1%A3.html)

------------------------------------------------------------------------
