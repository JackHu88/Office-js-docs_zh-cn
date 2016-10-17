
# <a name="binding.removehandlerasync-method"></a>Binding.removeHandlerAsync 方法
从绑定移除指定事件类型的指定处理程序。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|在**要求集[中可用](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|BindingEvents|
|**包含最后一次更改的版本**|1.1|

```js
bindingObj.removeHandlerAsync(eventType [, options], callback);
```


## <a name="parameters"></a>参数



|**名称**|**类型**|**Description**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|指定要事件的类型。必需。||
| _options_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _handler_|**object**|指定要移除的处理程序的名称。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string** 或 **undefined**|在 **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。||

## <a name="callback-value"></a>回调值

在你传递给 _callback_ 参数的函数执行后，它会收到你可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **removeHandlerAsync** 方法的回调函数中，可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|始终返回 **undefined** 因为删除事件处理程序时，没有要检索的数据或对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果将用户定义的一个 **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## <a name="remarks"></a>备注

如果调用  _removeHandlerAsync_ 方法时省略了可选 **handler** 参数，则指定 _eventType_ 的所有事件处理程序都将移除。


## <a name="example"></a>示例

以下示例将删除  **BindingDataChanged** 事件名为 `onBindingDataChanged` 的处理程序。


```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(Office.EventType.BindingDataChanged, {handler:onBindingDataChanged});
}

```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|BindingEvents|
|**最低权限级别**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录





****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|增加了对 Access 相关外接程序的支持。|
|1.0|引入|
