
# BindingSelectionChangedEventArgs.binding 属性
获取用于表示引发  **SelectionChanged** 事件的绑定的 **Binding** 对象。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**包含最后一次更改的版本**|1.1|

```
var myBinding = eventArgsObj.binding;
```


## 返回值

表示引发 [SelectionChanged](../../reference/shared/binding.md) 事件的 [Binding](../../reference/shared/binding.bindingselectionchangedevent.md) 对象。


## 支持详细信息


以下矩阵中的大写字母 Y 指示此属性在相应的 Office 主机应用程序中受到支持。空单元格指示 Office 主机应用程序不支持此属性。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 相关外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


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
|1.1|你可以立即针对 Access 内容外接程序中的 **SelectionChanged** 事件添加和删除事件处理程序。|
|1.0|引入|
