
# File 对象
表示与 Office 外接程序关联的文档文件。

|||
|:-----|:-----|
|**主机：**|PowerPoint 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|文件|
|**包含最后一次更改的版本**|1.1|

```
file
```


## 成员


**属性**


|**名称**|**说明**|
|:-----|:-----|
|**[size](../../reference/shared/file.size.md)**|获取以字节为单位的文档文件大小。|
|**[sliceCount](../../reference/shared/file.slicecount.md)**|获取文件分为的切片数。|

**方法**


|**名称**|**说明**|
|:-----|:-----|
|**[closeAsync](../../reference/shared/file.closeasync.md)**|关闭文档文件。|
|**[getSliceAsync](../../reference/shared/file.getsliceasync.md)**|返回指定的切片。|

## 备注

使用传递给 **Document.getFileAsync** 方法的回调函数中的 [AsyncResult.value](../../reference/shared/asyncresult.value.md) 属性访问 [File](../../reference/shared/document.getfileasync.md) 对象。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此对象。空的单元格表示相应的 Office 主机应用程序不支持此对象。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|||||
|:-----|:-----|:-----|:-----|
||Office for Windows Desktop|Office Online（在浏览器中）|Office for iPad|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**在要求集中可用**|文件|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 PowerPoint 和 Word 的支持。|
|1.0|引入|
