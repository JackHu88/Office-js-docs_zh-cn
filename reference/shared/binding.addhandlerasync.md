
# <a name="binding.addhandlerasync-method"></a>Binding.addHandlerAsync 方法
将处理程序添加到指定事件类型的绑定。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|在**要求集[中可用](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|BindingEvents|
|**包含最后一次更改的版本**|1.1|

```
bindingObj.addHandlerAsync(eventType, handler [, options], callback);
```


## <a name="parameters"></a>参数



|**名称**|**类型**|**Description**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|指定要添加的事件的类型。必需。对于  **Binding** 对象事件， _eventType_ 参数可以指定为 **Office.EventType.BindingSelectionChanged**、 **Office.EventType.BindingDataChanged**，也可以为这些枚举对应的文本值。||
| _handler_|**object**|要添加的事件处理程序函数。||
| _options_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string** 或 **undefined**|在 **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。||

## <a name="callback-value"></a>回调值

在你传递给 _callback_ 参数的函数执行后，它会收到你可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **addHandlerAsync** 方法的回调函数中，可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|始终返回 **undefined**，这是因为添加事件处理程序时没有要检索的数据或对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果将用户定义的一个 **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## <a name="remarks"></a>备注

只要每个事件处理程序函数的名称是唯一的，您就可以为指定的  _eventType_ 添加多个事件处理程序。


## <a name="example"></a>示例

以下代码示例调用 [Office](../../reference/shared/office.select.md) 对象的 **select** 以访问 ID 为 MyBinding 的绑定，然后调用 **addHandlerAsync** 方法，以便为该绑定的 [bindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md) 事件添加事件处理程序函数。


```js
function addEventHandlerToBinding() {
    Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
}

function onBindingDataChanged(eventArgs) {
    write("Data has changed in binding: " + eventArgs.binding.id);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
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
|1.1|增加了对 Word Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|增加了对 Access 相关外接程序的支持。|
|1.0|引入|
