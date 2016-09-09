
# Slice 对象
表示文档文件的切片。

|||
|:-----|:-----|
|**主机：**|PowerPoint 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|文件|
|**包含最后一次更改的版本**|1.1|

```
slice
```


## 成员


**属性**


|**名称**|**说明**|
|:-----|:-----|
|**[data](../../reference/shared/slice.data.md)**|获取文件切片的原始数据。|
|**[index](../../reference/shared/slice.index.md)**|获取文件切片的索引。|
|**[size](../../reference/shared/slice.size.md)**|获取以字节为单位的切片大小。|

## 备注

**Slice** 对象通过 [File.getSliceAsync](../../reference/shared/file.getsliceasync.md) 方法获得访问。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此对象。空的单元格表示相应的 Office 主机应用程序不支持此对象。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|


|||
|:-----|:-----|
|**在要求集中可用**|文件|
|**最低权限级别**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 PowerPoint 和 Word 的支持。|
|1.0|引入|
