
# DocumentActiveViewChangedEventArgs.activeView 属性
获取  **ActiveView** 枚举值，该值指定文档活动视图的状态，例如，用户是否可以编辑文档。

|||
|:-----|:-----|
|**主机：**|PowerPoint|
|**在其中添加**|1.1|

```
var myView = eventArgsObj.activeView;
```


## 返回值

引发事件的视图的 [ActiveView](../../reference/shared/activeview-enumeration.md)。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|

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
|1.1|增加了对 Office for iPad 中 PowerPoint 的支持。|
|1.1|引入|
