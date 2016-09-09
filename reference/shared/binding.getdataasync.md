
# Binding.getDataAsync 方法
返回绑定中包含的数据。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|MatrixBindings, TableBindings, TextBindings|
|**TableBindings 中的最后更改**|1.1|

```
bindingObj.getDataAsync([, options] , callback );
```


## 参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|指定如何强制设置数据。 ||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|指定是否返回具有应用了格式设置的值（如数字和日期）。||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|指定检索数据时是否必须应该筛选器。||
| _rows_|**Office.TableRange.ThisRow**| 指定预定义的字符串"thisRow"以获取当前所选行中的数据。|仅针对 Access 相关内容外接程序中的表绑定。|
| _startRow_|**number**|对于表或矩阵绑定，为绑定中的数据子集指定基于零的起始行。可选。 ||
| _startColumn_|**number**|对于表或矩阵绑定，为绑定中的数据的子集指定基于零的起始列。 ||
| _rowCount_|**number**|对于表或矩阵绑定，指定从  _startRow_ 偏移的列数。 ||
| _columnCount_|**number**|对于表或矩阵绑定，指定从  _startColumn_ 偏移的列数。||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给  **Binding.getDataAsync** 方法的回调函数中，您可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|访问指定绑定中的值。如果指定了 _coercionType_ 参数（且调用成功），则数据以[CoercionType](../../reference/shared/coerciontype-enumeration.md) 枚举主题中所述的格式返回。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果您将用户定义的一个  **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## 备注

如果省略了可选参数，则使用以下默认值（在适合数据的类型和格式的情况下）。



|**参数**|**默认**|
|:-----|:-----|
| _coercionType_|原始、非强制类型的绑定|
| _valueFormat_|无格式数据。|
| _filterType_|所有值（未经筛选）。|
| _startRow_|第一行。|
| _startColumn_|第一列。|
| _rowCount_|所有行。|
| _columnCount_|所有列。|
从 [MatrixBinding](../../reference/shared/binding.matrixbinding.md) 或 [TableBinding](../../reference/shared/binding.tablebinding.md) 调用时，如果指定了可选 **startRow**、_startColumn_、_rowCount_ 和 _columnCount_ 参数（且它们指定连续且有效的范围），_getDataAsync_ 方法将返回绑定值的子集。


## 示例




```
function showBindingData() {
    Office.select("bindings#MyBinding").getDataAsync(function (asyncResult) {
        write(asyncResult.value)
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



相对于具有标题行格式的数据，通过 _Binding.getDataAsync_ 方法使用 `"table"` 和 `"matrix"`**coercionType** 时有一个重要的行为差异，如以下两个示例中所示。这些代码示例演示了 [Binding.SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) 事件的事件处理程序功能。

如果指定  `"table"` _coercionType_，[TableData.rows](../../reference/shared/tabledata.rows.md) 属性（下列代码示例中的 `result.value.rows`）将返回一个仅包含表格正文行的数组。因此，它的第 0 行在表中将作为第一个非标题行。




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'table', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value.rows[0][0]); 
            } 
            else 
                write(result.error.message); 
    }); 
}     
// Function that writes to a div with id='message' on the page. 
function write(message){ 
    document.getElementById('message').innerText += message; 
}
```

但如果指定 `"matrix"` _coercionType_，那么以下代码示例中的  `result.value` 将返回一个第 0 行中包含表格标题的数组。如果表格标题包含多行，则所有行首先将作为单独的行包含在 `result.value` 矩阵中，然后再包含表格正文行。




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'matrix', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value[1][0]); 
            } 
            else 
                write(result.error.message); 
    }); 
}     
// Function that writes to a div with id='message' on the page. 
function write(message){ 
    document.getElementById('message').innerText += message; 
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
|**在要求集中可用**|MatrixBindings, TableBindings, TextBindings|
|**最低权限级别**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|增加了对 Access 相关外接程序中表绑定的支持。|
|1.0|引入|

## 另请参阅



#### 其他资源


[绑定到文档或电子表格中的区域](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
