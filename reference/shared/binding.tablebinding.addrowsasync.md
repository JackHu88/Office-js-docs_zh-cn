
# TableBinding.addRowsAsync 方法
将行和值添加到表中。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|TableBindings|
|**包含最后一次更改的版本**|1.1|

```js
bindingObj.addRowsAsync(rows, [,options], callback);
```


## 参数

_行_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**数组**

&nbsp;&nbsp;&nbsp;&nbsp;包含要添加到表中的一行或多行数据的数组的数组。 必需。
    
_选项_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**对象**

&nbsp;&nbsp;&nbsp;&nbsp;指定以下 [可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。
    
&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;类型：**数组、布尔值、null、数字、对象、字符串或未定义**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[AsyncResult](../../reference/shared/asyncresult.md) 对象中未经改动的返回的任何类型的用户定义项。 可选。<br/><br/>

_callback_<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;类型：**对象**
    
&nbsp;&nbsp;&nbsp;&nbsp;返回回调时调用的函数，其唯一的参数的类型为 [AsyncResult](../../reference/shared/asyncresult.md)。 可选。



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _rows_|**array**|包含要添加到表中的一行或多行数据的数组的数组。必需。||
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

在你传递给 _callback_ 参数的函数执行后，它会收到你可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **addRowsAsync** 方法的回调函数中，你可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|始终返回 **undefined**，这是因为没有要检索的对象或数据。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果您将用户定义的一个  **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## 备注

The success or failure of an  **addRowsAsync** operation is atomic. That is, the entire add rows operation must succeed, or it will be completely rolled back (and the **AsyncResult.status** property returned to the callback will report failure):


- 更新表时您作为  _data_ 参数传递的数组中的每行必须具有相同的列数。如果没有，整个操作都会失败。
    
- 数组中的每行和每个单元格都必须成功将该行和单元格添加到表中的新增行中。如果出于任何原因任意行或单元格未能添加成功，则整个操作将失败。
    
 **Excel Online 的其他标记**

对此方法的单个调用中，传递给  _rows_ 参数的值中的单元格总数不能超过 20,000。


## 示例




```js
function addRowsToTable() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        var binding = asyncResult.value;
        binding.addRowsAsync([["6", "k"], ["7", "j"]]);
    });
}

```




## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**在要求集中可用**|TableBindings|
|**最低权限级别**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel和 Word 的支持|
|1.1|增加了对在 Access 相关外接程序中写入表数据的支持。|
|1.0|引入|
