

# SettingsChangedEventArgs.settings 属性
获取表示引发  **settingsChanged** 事件的设置的 **Settings** 对象。

|||
|:-----|:-----|
|**主机：**|Excel|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|Settings|
|**包含最后一次更改的版本**|1.0|

```js
var mySettings = eventArgsObj.settings;
```


## 返回值

一个 [Settings](../../reference/shared/document.settings.md) 对象，表示引发 [settingsChanged](../../reference/shared/settings.settingschangedevent.md) 事件的设置。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此属性。空的单元格表示相应的 Office 主机应用程序不支持此属性。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||

|||
|:-----|:-----|
|**在要求集中可用**|Settings|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|
