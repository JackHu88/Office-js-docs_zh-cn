
# InitializationReason 枚举
指定是刚刚插入外接程序，还是文档中已包含。 

|||
|:-----|:-----|
|**主机：**|Excel、Project、Word|
|**在其中添加**|1.0|

```
Office.InitializationReason
```


## 成员


**值**


|**枚举**|**值**|**说明**|
|:-----|:-----|:-----|
|Office.InitializationReason.Inserted|"inserted"|刚刚将 外接程序 插入文档中。|
|Office.InitializationReason.DocumentOpened|"documentOpened"|外接程序 已是打开的文档的一部分。|

## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.0|引入|
