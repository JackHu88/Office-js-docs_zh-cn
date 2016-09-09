
# Slice.data 属性
获取文件切片的原始数据。

|||
|:-----|:-----|
|**主机：**|PowerPoint 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|文件|
|**包含最后一次更改的版本**|1.1|

```
var sliceData = slice.data;
```


## 返回值

格式为 **Office.FileType.Text** ("text") 或 **Office.FileType.Compressed** ("compressed") 的文件切片的原始数据，格式由 _Document.getFileAsync_ 方法调用的 [fileType](../../reference/shared/document.getfileasync.md) 参数指定。


## 备注

"compressed"格式的文件将返回可以在需要时转换为 base64 编码的字符串的字节数组。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此属性。空的单元格表示相应的 Office 主机应用程序不支持此属性。

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



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 PowerPoint 和 Word 的支持。|
|1.0|引入|
