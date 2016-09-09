
# Bindings.document 属性
获取表示与此组绑定关联的文档的  **Document** 对象。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**包含最后一次更改的版本**|1.1|

```
var docObj = bindingsObj.document;
```


## 返回值

一个 [Document](../../reference/shared/bindings.document.md) 对象。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|添加了对代表 Access 相关内容外接程序中当前 Access 数据库的 **Document** 对象进行访问的权限。|
|1.0|引入|
