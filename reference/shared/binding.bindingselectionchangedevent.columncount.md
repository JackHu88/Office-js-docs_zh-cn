
# <a name="bindingselectionchangedeventargs.columncount-property"></a>BindingSelectionChangedEventArgs.columnCount 属性
获取选择的列数。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**包含最后一次更改的版本**|1.1|

```
var colCount = eventArgsObj.columnCount;
```


## <a name="return-value"></a>返回值

所选的列数。如果只选择了一个单元格，则返回 1。


## <a name="remarks"></a>备注

如果用户选择了不连续的单元格，则返回此绑定内最后一个连续选区的计数。 

对于 Word，此属性只适用于 [BindingType](../../reference/shared/bindingtype-enumeration.md) 为"table"的绑定。如果绑定类型为"matrix"，将返回 **null**。此外，如果表格包含合并单元格，调用将失败，因为表的结构必须统一，此属性才能正确工作。


## <a name="example"></a>示例

以下示例向 [id](../../reference/shared/binding.bindingselectionchangedevent.md) 为 `myTable` 的绑定中添加 [SelectionChanged](../../reference/shared/binding.id.md) 事件的事件处理程序。当用户更改所选内容时，处理程序将显示所选内容中第一个单元格的坐标，以及所选的行数和列数。


```js
function addSelectionHandler() {
    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addHandlerAsync("bindingSelectionChanged", myHandler);
    });
}

// Display selection start coordinates and row/column count.
function myHandler(bArgs) {
    write("Selection start row/col: " + bArgs.startRow + "," + bArgs.startColumn);
    write("Selection row count: " + bArgs.rowCount);
    write("Selection col count: " + bArgs.columnCount);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此属性。空的单元格表示相应的 Office 主机应用程序不支持此属性。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


| |**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录





****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|你可以立即针对 Access 内容外接程序中的 **SelectionChanged** 事件添加和删除事件处理程序。|
|1.0|引入|
