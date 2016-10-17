
# <a name="bindings.addfrompromptasync-method"></a>Bindings.addFromPromptAsync 方法
 显示可让用户指定要绑定的选择的 UI。

|||
|:-----|:-----|
|**主机：**|Access、Excel|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|不在集合中|
|**包含最后一次更改的版本**|1.1|

```
_bindingsObj.addFromPromptAsync(bindingType [, options], callback);
```


## <a name="parameters"></a>参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|指定要创建的绑定对象的类型。必需。如果所选的对象不能强制转换为指定类型，则返回 **null**。||
| _options_|**object**|指定以下任一 [可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _id_|**string**|指定用于标识新绑定对象的唯一名称。如果没有为 _id_ 参数传递任何实参，则会自动生成 [Binding.id](../../reference/shared/binding.id.md)。||
| _promptText_|**string**|指定显示在提示 UI 中且告诉用户选择内容的字符串。限制为 200 个字符。如果没有传递 _promptText_ 参数，则会显示“请进行选择”。||
| _sampleData_|[TableData](../../reference/shared/tabledata.md)|指定在提示 UI 中显示为可能由外接程序绑定的各种字段（列）的示例的示例数据的表格。**TableData** 对象中提供的标题指定字段选择 UI 中使用的标签。可选。**注意：**此参数仅在 Access 相关外接程序中使用。如果 Excel 相关外接程序中调用方法时提供了此参数，则会将其忽略。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string** 或 **undefined**|在 **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。||

## <a name="callback-value"></a>回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **addFromPromptAsync** 方法的回调函数中，可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|访问代表用户指定选区的 [Binding](../../reference/shared/binding.md) 对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果用户定义的一个 **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## <a name="remarks"></a>注解

将指定类型的绑定对象添加到 [Bindings](../../reference/shared/bindings.bindings.md) 集合，该集合将用所提供的 _id_ 进行标识。如果无法绑定指定选择，则该方法会失败。


## <a name="example"></a>示例




```js
function addBindingFromPrompt() {

    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'MyBinding', promptText: 'Select text to bind to.' }, function (asyncResult) {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|不在集合中|
|**最低权限级别**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 的支持。|
|1.1|在 Excel 相关外接程序中，你可以为包含表格数据的一系列单元格创建表绑定（将 _bindingType_ 作为 **Office.BindingType.Table** 传递），即使该数据未在 Excel UI 中作为表格添加到电子表格时也是如此操作（通过使用“**插入**” > “**表格**” > “**表格**”或“**开始**” > “**样式**” > “**套用表格格式**”命令实现）。|
|1.1|添加了对 Access 相关内容应用程序中表绑定的支持。 |
|1.1|添加了以 Excel 相关应用程序中表绑定的形式绑定到矩阵数据的支持。|
|1.0|引入|
