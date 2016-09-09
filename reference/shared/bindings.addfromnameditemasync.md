
# Bindings.addFromNamedItemAsync 方法
将绑定添加到文档中的命名项。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|MatrixBindings, TableBindings, TextBindings|
|**最后更改**|1.1|

```
Office.context.document.bindings.addFromNamedItemAsync(itemName, bindingType [, options], callback);
```


## 参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _itemName_|**string**|已命名项目的名称。必需。||
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|指定要创建的绑定对象的类型。必需。 如果选择的对象不能强制为指定的类型，则返回  **null** 。||
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _id_|**string**|指定用于标识新绑定对象的唯一名称。如果没有为 _id_ 参数传递任何实参，则会自动生成 [Binding.id](../../reference/shared/binding.id.md)。||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **addFromNamedItemAsync** 方法的回调函数中，您可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|访问表示特定命名项目的 [Binding](../../reference/shared/binding.md) 对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果您将用户定义的一个  **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## 注解

 **对于 Excel**，_itemName_ 参数可引用已命名范围或表。

默认情况下，在 Excel 中添加表会为你添加的第一个表分配名称“Table1”，为你添加的第二个表分配名称“Table2”，以此类推。 若要在 Excel UI 中为表分配有意义的名称，请使用功能区的“**表格工具 | 设计**”选项卡上的“**表格名称**”属性。


 >**注意**  在 Excel 中，在指定表格作为命名项目时，必须完全限定该名称以便在表格名称中包括工作表名称，格式如下：`"Sheet1!Table1"`

 **对于 Word**，_itemName_ 参数会引用“**格式文本**”内容控件的“**标题**”属性。 （你不能绑定除“**格式文本**”内容控件之外的内容控件。）

默认情况下，不会向内容控件分配“**标题**”值。 若要在 Word UI 中分配有意义的名称，请从功能区的“**开发人员**”选项卡上的“**控件**”组中插入一个“**格式文本**”内容控件，并使用“**控件**”组中的“**属性**”命令显示“**内容控件属性**”对话框。 然后将内容控件的“**标题**”属性设置为需要从代码中引用的名称。


 >**注意**  在 Word 中，如果有多个“**格式文本**”内容控件具有同一“**标题**”属性值（名称），并且尝试使用此方法绑定到这些内容控件之一（通过将其名称指定为 _itemName_ 参数），操作将失败。


## 示例

以下示例将对 Excel 中 `myRange` 已命名项的绑定添加为“矩阵”绑定，并将绑定的 [id](../../reference/shared/binding.id.md) 分配为 `myMatrix`。


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

以下示例将对 Excel 中 `Table1` 已命名项的绑定添加为“表”绑定，并将绑定的 **id** 分配为 `myTable`.




```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("Table1", "table", {id:'myTable'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

以下示例在 Word 中创建了一个用于绑定到名为  `"FirstName"` 的格式文本内容控件的文本绑定，分配了 **id**`"firstName"`，并显示了相关信息。




```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
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
|1.1|在 Excel 相关外接程序中，你可以为包含表格数据的一系列单元格创建表绑定（将 _bindingType_ 作为 **Office.BindingType.Table** 传递），即使该数据未作为表格添加到电子表格时也是如此操作（通过使用“**插入**” > “**表格**” > “**表格**”或“**开始**” > “**样式**” > “**套用表格格式**”命令实现）。|
|1.1|添加了对 Access 相关内容应用程序中表绑定的支持。 |
|1.0|引入|

## 另请参阅



#### 其他资源


[绑定到文档或电子表格中的区域](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md#add-a-binding-to-a-named-item)
