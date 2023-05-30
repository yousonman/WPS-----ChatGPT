# 显示修订

打开文档，文档打开修订模式，通过参数指定是否显示痕迹

**\_wps.openDoc** ( *params* )

``` JavaScript
_WpsStartUp([{
    "OpenDoc": { // OpenDoc 对应于OA助手dispatcher中的方法名
        "uploadPath": uploadPath, // 保存文档上传路径
        "fileName": filePath,
        "userName": '李雷', //用户名
        "revisionCtrl": {
            "bOpenRevision": true, //true 打开修订，false关闭修订
            "bShowRevision": true //true 显示修订，false隐藏修订
        }
    }
}])
```

## 输入参数

| 参数名        | 数据类型 | 描述                                              | 必选 |
|---------------|----------|---------------------------------------------------|------|
| *docId*         | String   | 文档ID，文档在数据库中的ID                        | 是   |
| *fileName*      | String   | 获取服务器文档接口（不传即为新建空文档）          | 是   |
| *uploadPath*    | String   | 新建空文档文件名,不传默认新建文档以系统时间为名称 | 是   |
| *revisionCtrl*  | JSON     | 痕迹控制 ，不传正常打开                           | 是   |
| *bOpenRevision* | Boolean  | true(打开)false(关闭)修订                         | 是   |
| *bShowRevision* | Boolean  | true(显示)/false(关闭)痕迹                        | 是   |

## 文档修订模式场景展现

![](服务器端图像/打开修订模式.gif)

## 业务流程（执行逻辑）

![](服务器端图像/控制公文的修订显示.png)

> WPS 开发文档： [WPS 集成模式/WPS 加载项开发/OA 助手/OA助手场景对接场景/显示修订](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E9%9B%86%E6%88%90%E6%A8%A1%E5%BC%8F/WPS%20%E5%8A%A0%E8%BD%BD%E9%A1%B9%E5%BC%80%E5%8F%91/OA%20%E5%8A%A9%E6%89%8B/OA%E5%8A%A9%E6%89%8B%E5%9C%BA%E6%99%AF%E5%AF%B9%E6%8E%A5%E5%9C%BA%E6%99%AF/%E6%98%BE%E7%A4%BA%E4%BF%AE%E8%AE%A2.html)

------------------------------------------------------------------------
