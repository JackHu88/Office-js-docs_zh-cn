
# Context.commerceAllowed 属性
获取外接程序是否将运行在允许链接到外部付款系统的平台上。

|||
|:-----|:-----|
|**主机：**|Excel 和 Word|
|**包含最后一次更改的版本**|1.1|

```
var allowCommerce = Office.context.commerceAllowed;
```


## 返回值

如果开发者可以在相应平台上的外接程序中显示销售或升级 UI，则返回 **True**；否则，返回 **False**。


## 备注

iOS 应用商店不支持提供其他付款系统的链接的应用程序和外接程序。但是，在 Windows 桌面上或在浏览器中（对于 Office Online）运行的 Office 外接程序不允许此类链接。如果您希望您的外接程序的 UI 提供除 iOS 以外的平台上的外部付款系统链接，您可以使用  **commerceAllowed** 属性控制链接何时显示。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for iPad**|
|:-----|:-----|
|**Excel**|Y|
|**PowerPoint**||
|**Word**|Y|

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
|1.1|引入。|
