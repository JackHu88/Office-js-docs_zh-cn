

# <a name="office.select-method"></a>Office.select 方法
创建承诺以基于传入的选择器字符串返回绑定。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**在[要求集中可用](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**包含最后一次更改的版本**|1.1|

```js
Office.select(str, onError);
```


## <a name="parameters"></a>参数


_str_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**字符串**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;要为其解析和创建承诺的选择器字符串。

_onError_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**函数**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult**。可选。
    

## <a name="callback-value"></a>回调值

在你传递给 _onError_ 参数的函数执行后，它会收到你可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。如果操作失败，请使用 [AsyncResult.error](../../reference/shared/asyncresult.error.md) 属性，访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。


## <a name="remarks"></a>备注

**Office.select** 方法提供对 [Binding](../../reference/shared/binding.md) 对象承诺的访问权限，该承诺尝试在调用其任何异步方法时返回指定的绑定。

支持的格式：“bindings# _bindingId_”，它为 [id](../../reference/shared/binding.id.md) 为 `bindingId` 的绑定返回 **Binding** 对象。有关详细信息，请参阅 [Office 外接程序中的异步编程](../../docs/develop/asynchronous-programming-in-office-add-ins.md#asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings) 和 [绑定到文档或电子表格中的区域](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)。


 >**注意**：如果 **select** 方法目标成功返回一个 **Binding** 对象，该对象将只公开 [Binding](../../reference/shared/binding.md) 对象的以下四个方法：[getDataAsync](../../reference/shared/binding.getdataasync.md)、[setDataAsync](../../reference/shared/binding.setdataasync.md)、[addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) 和 [removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)。如果该目标无法返回 **Binding** 对象，则可以使用 _onError_ 回调访问 [asyncResult.error](../../reference/shared/asyncresult.error.md) 对象以获取详细信息。如果需要调用 **Binding** 对象的成员而不是 **select** 方法返回的 **Binding** 对象目标公开的四个方法，则应通过 [Document.bindings](../../reference/shared/document.bindings.md) 属性和 [Bindings.getByIdAsync](../../reference/shared/bindings.getbyidasync.md) 方法来使用 [getByIdAsync](../../reference/shared/bindings.getbyidasync.md) 方法，以检索 **Binding** 对象。


## <a name="example"></a>示例

以下代码示例使用 **select** 方法从 **Bindings** 集合检索 **id** 为“`cities`”的绑定，然后调用 [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) 方法以便为绑定的 [dataChanged](../../reference/shared/binding.bindingdatachangedevent.md) 事件添加事件处理程序。


```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}
```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**在要求集中可用**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**最低权限级别**|[ReadDocument（适用于 Open Office XML 的 ReadAllDocument）](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel和 Word 的支持|
|1.1|添加了 **select** 方法的使用，以返回在 Access 相关内容外接程序中创建的表绑定。|
|1.0|引入|
