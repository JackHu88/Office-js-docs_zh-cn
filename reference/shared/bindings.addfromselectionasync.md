
# <a name="bindings.addfromselectionasync-method"></a>Bindings.addFromSelectionAsync 方法
将绑定添加到文档中的当前选择。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|MatrixBindings, TableBindings, TextBindings|
|**包含最后一次更改的版本**|1.1|

```
bindingsObj.addFromSelectionAsync(bindingType [, options], callback);
```


## <a name="parameters"></a>参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|指定要创建的绑定对象的类型。必需。如果所选的对象不能强制转换为指定类型，则返回 **null**。||
| _options_|**object**|指定以下任一 [可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _id_|**string**|指定用于标识新绑定对象的唯一名称。如果没有为 _id_ 参数传递任何实参，则会自动生成 [Binding.id](../../reference/shared/binding.id.md)。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string** 或 **undefined**|在 **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。||

## <a name="callback-value"></a>回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **addFromSelectionAsync** 方法的回调函数中，可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|访问代表用户指定选区的 [Binding](../../reference/shared/binding.md) 对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果用户定义的一个 **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## <a name="remarks"></a>注解

将指定类型的绑定对象添加到  **Bindings** 集合，该集合将用所提供的 _id_ 标识。


 >**注意**  在 Excel 中，如果调用传入现有绑定的 **Binding.id** 的 **addFromSelectionAsync** 方法，则会使用该绑定的 [Binding.type](../../reference/shared/binding.type.md)，并且无法通过为 _bindingType_ 参数指定不同值来更改其类型。如果需要使用现有 _id_ 并更改 _bindingType_，请先调用 [Bindings.releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md) 方法释放该绑定，然后调用 **addFromSelectionAsync** 方法重新建立新类型的绑定。


## <a name="example"></a>示例

将 [TextBinding](../../reference/shared/binding.textbinding.md) 添加到当前选择内容，其 **Binding.id** 为 “MyBinding”。


```js
function addBindingFromSelection() {
    Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'MyBinding' }, 
        function (asyncResult) {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    );
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
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|MatrixBindings, TableBindings, TextBindings|
|**最低权限级别**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|在 Excel 相关外接程序中，你可以为包含表格数据的一系列单元格创建表绑定（将 _bindingType_ 作为 **Office.BindingType.Table** 传递），即使该数据未作为表格添加到电子表格时也是如此操作（通过使用“**插入**” > “**表格**” > “**表格**”或“**开始**” > “**样式**” > “**套用表格格式**”命令实现）。|
|1.1|添加了对 Access 相关内容应用程序中表绑定的支持。 |
|1.0|引入|
