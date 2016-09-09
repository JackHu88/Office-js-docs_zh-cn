
# DocumentMode 枚举
指定关联应用程序中的文档为只读，还是读写。 

|||
|:-----|:-----|
|**主机：**|Excel、PowerPoint、Project、Word|
|**在其中添加**|1.1|

```
Office.DocumentMode
```


## 成员


**值**


|**枚举**|**值**|**说明**|
|:-----|:-----|:-----|
|Office.DocumentMode.ReadOnly|"readOnly"|文档为只读。|
|Office.DocumentMode.ReadWrite|"readWrite"|可以读取和写入文档。|

## 备注

由 **Document** 对象的 [mode](../../reference/shared/document.md) 属性返回。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.0|引入|
