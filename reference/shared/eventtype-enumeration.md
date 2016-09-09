
# EventType 枚举
指定引发的事件的类型。由  **EventName**_EventArgs_ 对象的 **type** 属性返回。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint、Project、Word|
|**包含 Selection 最后一次更改的版本**|1.1|

```js
Office.EventType
```


## 成员


**值**


|枚举|值|说明|
|:-----|:-----|:-----|
|Office.EventType.ActiveViewChanged|"documentActiveViewChanged"|引发了 [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md) 事件。|
|Office.EventType.DocumentSelectionChanged|"documentSelectionChanged"|引发了 [Document.SelectionChanged](../../reference/shared/document.selectionchanged.event.md) 事件。|
|Office.EventType.BindingSelectionChanged|"bindingSelectionChanged"|引发了 [Binding.BindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) 事件。|
|Office.EventType.BindingDataChanged|"bindingDataChanged"|引发了 [Binding.BindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md) 事件。|
|Office.EventType.DataNodeDeleted|"nodeDeleted"|引发了 [CustomXmlPart.dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md) 事件。|
|Office.EventType.DataNodeInserted|"nodeInserted"|引发了 [CustomXmlPart.dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md) 事件。|
|Office.EventType.DataNodeReplaced|"nodeReplaced"|引发了 [CustomXmlPart.dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md) 事件。|
|Office.EventType.SettingsChanged|"settingsChanged"|引发了 [Settings.settingsChanged](../../reference/shared/settings.settingschangedevent.md) 事件。|

## 注解


 >**注意**：Project 的外接程序支持 **Office.EventType.ResourceSelectionChanged**、**Office.EventType.TaskSelectionChanged** 和 **Office.EventType.ViewSelectionChanged** 事件类型。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y||
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
|1.1| 为新 **Document.ActiveViewChanged** 事件添加了 Office.EventType.ActiveViewChanged 枚举。|
|1.0|引入|
