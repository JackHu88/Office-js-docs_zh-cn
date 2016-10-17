
# <a name="activeview-enumeration"></a>ActiveView 枚举
指定文档活动视图的状态，例如，用户是否可以编辑文档。

|||
|:-----|:-----|
|**Office.js 版本中引入**|1.1|

|||
|:-----|:-----|
|**主机：**|PowerPoint|
|**在其中添加**|1.1|



```
Office.ActiveView
```


## <a name="members"></a>成员


**值**


|**枚举**|**值**|**Description**|
|:-----|:-----|:-----|
|Office.ActiveView.Read|"read"|主机应用程序的活动视图只允许用户阅读文档中的内容。|
|Office.ActiveView.Edit|"edit"|主机应用程序的活动视图允许用户编辑文档中的内容。|

## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|

|||
|:-----|:-----|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 PowerPoint 的支持。|
|1.1|引入|
