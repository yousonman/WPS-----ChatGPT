# 公文套红头

打开指定文档，将红头文件插入到当前文档的顶部，若模板中存在对应书签， 则会将当前文档的正文插入到书签所在位置

``` JavaScript
_WpsStartUp([{
    "OpenDoc": {
        "docId": "123",
        "fileName": filePath,
        "insertFileUrl": templateURL,
        "bkInsertFile": 'Content', //红头模板中填充正文的位置书签名
    }
}])
```

## 输入参数

| 参数名        | 数据类型 | 描述                                   | 必选 |
|---------------|----------|----------------------------------------|------|
| *OpenDoc*       | String   | 方法名，对应于OA助手dispatcher中的方法 | 是   |
| *docId*         | String   | 文档ID                                 | 是   |
| *fileName*      | String   | 打开的文档路径                         | 否   |
| *insertFileUrl* | String   | 指定的红头模板                         | 否   |
| *bkInsertFile*  | String   | 书签名，红头模板插入正文的位置         | 否   |

## 公文套红头场景展现

![](服务器端图像/套红头(OA-Web端).gif)

## 业务流程（执行逻辑）

![](服务器端图像/套红头(OA-Web端).png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/公文套红头](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E5%85%AC%E6%96%87%E5%A5%97%E7%BA%A2%E5%A4%B4.html)

------------------------------------------------------------------------
