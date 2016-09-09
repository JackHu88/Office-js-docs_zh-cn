
# AsyncResultStatus 枚举
指定异步调用的结果。 

|||
|:-----|:-----|
|**主机：**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**包含最后一次更改的版本**|1.1|

```
Office.AsyncResultStatus
```


## 成员


**值**


|**枚举**|**值**|**说明**|
|:-----|:-----|:-----|
|Office.AsyncResultStatus.Succeeded|"succeeded"|调用成功。|
|Office.AsyncResultStatus.Failed|"failed"|调用失败。|

## 备注

由 [AsyncResult](../../reference/shared/asyncresult.status.md) 对象的 [status](../../reference/shared/asyncresult.md) 属性返回。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。


有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|**适用于设备的 OWA**|**Office for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|Y|||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y||Y|||

|||
|:-----|:-----|
|**外接程序类型**|内容、任务窗格、Outlook|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|增加了对 Access 相关外接程序的支持。|
|1.0|引入|
