
# TableBinding.addColumnsAsync 方法
将列和值添加到表中。

|||
|:-----|:-----|
|**主机：**|Excel 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|TableBindings|
|**包含最后一次更改的版本**|1.0|

```
bindingObj.addColumnsAsync(data [, options], callback);
```


## 参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _data_|**array** 或 [TableData](../../reference/shared/tabledata.md)|包含要添加到表中的一行或多行数据的数组的数组（“矩阵”）或 **TableData** 对象。 必需。||
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **addColumnsAsync** 方法的回调函数中，您可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|始终返回 **undefined**，这是因为没有要检索的对象或数据。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果您将用户定义的一个  **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## 备注

若要添加指定数据和标题的值的一个或多个列，请将 **TableData** 对象以 _data_ 参数形式传递。若要添加仅指定数据的一个或多个列，请将数组的数组（“矩阵”）以 _data_ 参数形式传递。

**addColumnAsync** 操作的成功或失败是全有或全无的，即，整个添加列操作必须成功或者完全回滚（并且返回到回调的 **AsyncResult.status** 属性将报告失败）：


- 更新表时您作为 _data_ 参数传递的数组中的每行必须具有相同的行数。如果没有，整个操作都会失败。
    
- 数组中的每行和每个单元格都必须成功将该行或单元格添加到表中的新增列中。如果出于任何原因任意行或单元格未能添加成功，则整个操作将失败。
    
- 如果以数据参数形式传递 **TableData** 对象，标题行数必须匹配被更新的表的标题行数。
    
**Excel Online 的其他标记**

对此方法的单个调用中，传递给 **data** 参数的 _TableData_ 对象中的单元格总数不能超过 20,000。


## 示例

以下示例通过将 **TableData** 对象以 **addColumnsAsync** 方法的 _data_ 参数形式传递来将三行一列添加到具有 [id](../../reference/shared/binding.id.md) `"myTable"` 的绑定表中。 若要成功，被更新的表必须具有三行。


```js
// Add a column to a binding of type table by passing a TableData object.
function addColumns() {
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```

以下示例通过将数组的数组（“矩阵”）以 [addColumnsAsync](../../reference/shared/binding.id.md) 方法的 _data_ 参数形式传递来将三行一列添加到具有 **id**`myTable` 的绑定表中。若要成功，被更新的表必须具有三行。




```js
// Add a column to a binding of type table by passing an array of arrays.
function addColumns() {
    var myTable = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
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
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.0|引入|
