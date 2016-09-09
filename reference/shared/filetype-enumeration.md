
# FileType 枚举
指定返回文档的格式。

|||
|:-----|:-----|
|**主机：**|PowerPoint 和 Word|
|**包含最后一次更改的版本**|1.1|

```js
Office.FileType
```


## 成员


**值**


|**枚举**|**值**|**说明**|
|:-----|:-----|:-----|
|Office.FileType.Compressed|"compressed"|以字节数组形式返回 Office Open XML (OOXML) 格式的整个文档（.pptx 或 .docx）。|
|Office.FileType.Pdf|“pdf”|将 PDF 格式的整个文档作为字节数组返回。|
|Office.FileType.Text|"text"|只返回  **string** 形式的文档文本。（仅限 Word）|

## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 PowerPoint 和 Word 的支持。|
|1.1|添加了对另存为 PDF 的支持。|
|1.0|引入|
